import hashlib
import json
import os
import threading
from vertexai.generative_models import GenerationConfig

# Importation de ton instance Gemini et des paramètres de sécurité depuis settings.py
from config.settings import gemini_model, DEFAULT_SAFETY_SETTINGS, DEFAULT_MODEL

# Cache disque optionnel des reponses du modele (TRANSLATION_CACHE_DIR).
# Il evite de repayer un prompt deja envoye et rend les relances
# reproductibles : a prompt identique, reponse identique.
_CACHE_DIR = os.environ.get("TRANSLATION_CACHE_DIR")
_cache_lock = threading.Lock()


def _cache_path(kind, prompt):
    key = hashlib.sha256(f"{DEFAULT_MODEL}|{kind}|{prompt}".encode("utf-8")).hexdigest()
    return os.path.join(_CACHE_DIR, key[:2], key + ".json")


def _cache_get(kind, prompt):
    if not _CACHE_DIR:
        return None
    path = _cache_path(kind, prompt)
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return None


def _cache_put(kind, prompt, text):
    if not _CACHE_DIR or text is None:
        return
    path = _cache_path(kind, prompt)
    with _cache_lock:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump({"text": text}, fh, ensure_ascii=False)

def build_prompt(content, context, direction="Allemand ➔ Français", is_batch=False):
    """
    Construit le prompt dynamique pour Vertex AI en fonction du mode (batch ou unitaire).
    """
    if is_batch:
        # Transformation de la liste de textes en dictionnaire indexé pour garantir l'alignement
        payload = {str(i): txt for i, txt in enumerate(content)}
        
        prompt = f"""Tu es un traducteur expert professionnel.
Sens de traduction : {direction}

CONTEXTE ET RÈGLES :
{context}

TÂCHE :
Traduis le dictionnaire JSON suivant. Garde exactement les mêmes clés.
Ne renvoie QUE le JSON valide sans aucun formatage markdown autour.

TEXTES À TRADUIRE :
{json.dumps(payload, ensure_ascii=False)}
"""
    else:
        prompt = f"""Tu es un traducteur expert professionnel.
Sens de traduction : {direction}

CONTEXTE ET RÈGLES :
{context}

TÂCHE :
Traduis le texte suivant en conservant son formatage et ses balises (<b>, <i>, <u>).

TEXTE À TRADUIRE :
{content}
"""
    return prompt


def call_gemini_batch(prompt, expected_count):
    """
    Appelle Gemini avec un prompt demandant un retour en JSON (pour les lots).
    Retourne (liste_des_traductions, input_tokens, output_tokens).
    """
    generation_config = GenerationConfig(
        temperature=0.1,
        response_mime_type="application/json"
    )

    try:
        cached = _cache_get("batch", prompt)
        if cached is not None:
            text, in_tok, out_tok = cached["text"], 0, 0
        else:
            response = gemini_model.generate_content(
                prompt,
                generation_config=generation_config,
                safety_settings=DEFAULT_SAFETY_SETTINGS  # <-- Désactivation des filtres de sécurité
            )

            # Récupération sécurisée des compteurs de tokens
            in_tok = getattr(response.usage_metadata, 'prompt_token_count', 0) if hasattr(response, 'usage_metadata') else 0
            out_tok = getattr(response.usage_metadata, 'candidates_token_count', 0) if hasattr(response, 'usage_metadata') else 0

            # Vérification des réponses vides (bloquées ou plantées)
            if not response.candidates or not response.candidates[0].content.parts:
                print("⚠️ Gemini a renvoyé une réponse vide (Lot).")
                return None, in_tok, out_tok
            text = response.text

        data = json.loads(text)
        
        # Vérification de l'alignement (est-ce que l'IA a bien traduit tous les segments ?)
        translations = [data.get(str(i)) for i in range(expected_count)]
        
        if any(t is None for t in translations):
            print("⚠️ Lot incomplet détecté. Bascule automatique en mode unitaire.")
            return None, in_tok, out_tok

        _cache_put("batch", prompt, text)
        return translations, in_tok, out_tok
        
    except json.JSONDecodeError:
        print("⚠️ Erreur de formatage JSON renvoyé par Gemini. Bascule unitaire.")
        return None, 0, 0
    except Exception as e:
        print(f"⚠️ Erreur lors de l'appel BATCH à Vertex AI : {e}")
        return None, 0, 0  # <-- Évite le plantage 'cannot unpack non-iterable NoneType'


def call_gemini_fallback(prompt):
    """
    Appelle Gemini en mode texte simple (pour traiter les éléments un par un si le lot échoue).
    Retourne (texte_traduit, input_tokens, output_tokens).
    """
    generation_config = GenerationConfig(
        temperature=0.1
    )
    
    try:
        cached = _cache_get("single", prompt)
        if cached is not None:
            return cached["text"], 0, 0

        response = gemini_model.generate_content(
            prompt,
            generation_config=generation_config,
            safety_settings=DEFAULT_SAFETY_SETTINGS  # <-- Désactivation des filtres de sécurité
        )

        in_tok = getattr(response.usage_metadata, 'prompt_token_count', 0) if hasattr(response, 'usage_metadata') else 0
        out_tok = getattr(response.usage_metadata, 'candidates_token_count', 0) if hasattr(response, 'usage_metadata') else 0

        if not response.candidates or not response.candidates[0].content.parts:
            print("⚠️ Gemini a renvoyé une réponse vide (Fallback).")
            return None, in_tok, out_tok

        _cache_put("single", prompt, response.text.strip())
        return response.text.strip(), in_tok, out_tok
        
    except Exception as e:
        print(f"⚠️ Erreur lors de l'appel FALLBACK à Vertex AI : {e}")
        return None, 0, 0  # <-- Évite le plantage 'cannot unpack non-iterable NoneType'


# ==========================================================================
# Moteur de traduction par segments balises (utilise par tous les formats)
# ==========================================================================

import random
import re
import time
from concurrent.futures import ThreadPoolExecutor

from core.lang_detect import scores as _lang_scores
from core.text_utils import (glossary_entries, matching_entries, format_glossary_rules,
                              glossary_violations)

LANG = {
    "Allemand ➔ Français": ("de", "fr", "allemand", "français"),
    "Français ➔ Allemand": ("fr", "de", "français", "allemand"),
}

# Tarif Vertex AI de gemini-2.5-flash, en dollars par million de tokens.
PRICE_IN, PRICE_OUT = 0.30, 2.50

# Budget de caracteres par lot : assez pour que le modele voie le contexte
# (une diapositive, une section), assez peu pour que la reponse JSON reste
# courte et fiable.
BATCH_CHARS = 6000
BATCH_MAX_SEGMENTS = 60
MAX_WORKERS = 4

_RETRY_MARKERS = ("429", "resource exhausted", "resource_exhausted", "503", "unavailable",
                  "500", "internal", "deadline", "timeout", "temporarily")
MAX_RETRIES = 6

# Phrases d'assistant qu'un modele ajoute parfois a une traduction.
_CHATTY = ("voici la traduction", "here is the translation", "hier ist die übersetzung",
           "pourriez-vous", "veuillez fournir", "texte à traduire", "n'hésitez pas")


def _is_transient(exc):
    t = str(exc).lower()
    return any(m in t for m in _RETRY_MARKERS)


def generate_json(prompt, kind="segments"):
    """Appel JSON avec cache et reprise sur les erreurs temporaires.

    Renvoie (objet JSON ou None, tokens entree, tokens sortie).
    """
    cached = _cache_get(kind, prompt)
    if cached is not None:
        try:
            return json.loads(cached["text"]), 0, 0
        except ValueError:
            pass
    config = GenerationConfig(temperature=0.1, response_mime_type="application/json")
    in_tok = out_tok = 0
    for attempt in range(MAX_RETRIES):
        try:
            response = gemini_model.generate_content(
                prompt, generation_config=config, safety_settings=DEFAULT_SAFETY_SETTINGS)
            um = getattr(response, "usage_metadata", None)
            in_tok += getattr(um, "prompt_token_count", 0) or 0
            out_tok += getattr(um, "candidates_token_count", 0) or 0
            if not response.candidates or not response.candidates[0].content.parts:
                print("⚠️ Réponse vide du modèle.")
                return None, in_tok, out_tok
            text = response.text
            data = json.loads(text)
            _cache_put(kind, prompt, text)
            return data, in_tok, out_tok
        except ValueError:
            # JSON invalide (reponse tronquee) : une nouvelle tentative suffit
            # en general ; au-dela l'appelant redecoupe le lot.
            if attempt >= 1:
                return None, in_tok, out_tok
        except Exception as exc:
            if not _is_transient(exc) or attempt == MAX_RETRIES - 1:
                print(f"⚠️ Erreur Vertex AI : {exc}")
                return None, in_tok, out_tok
            delay = min(60.0, 4.0 * (2 ** attempt)) * (0.75 + 0.5 * random.random())
            print(f"⏳ Quota ou indisponibilité ({str(exc)[:60]}), nouvelle tentative dans {delay:.0f} s")
            time.sleep(delay)
    return None, in_tok, out_tok


def _segment_rules(direction, kind):
    src, tgt, src_name, tgt_name = LANG[direction]
    grammar = {
        "fr": "Français correct : contractions (du, des, aux, auprès des), accords, "
              "espace avant « : ; ? ! » selon l'usage français, guillemets « ».",
        "de": "Allemand correct : majuscule aux substantifs, déclinaisons, mots composés "
              "soudés selon l'usage, guillemets „ “.",
    }[tgt]
    house = ("« Filiale 4440 » se traduit « SPM 4440 »" if tgt == "fr"
             else "« SPM 4440 » et « supermarché 4440 » se traduisent « Filiale 4440 »")
    order_example = (" (« les articles NAT », pas « les NAT articles »)" if tgt == "fr"
                     else " (« die geprüften Filialen », pas « die Filialen geprüften »)")
    cited = ("« Frage 4.1.15 Stammdaten » → « Question 4.1.15 Données de base »" if tgt == "fr"
             else "« Question 4.1.15 Données de base » → « Frage 4.1.15 Stammdaten »")
    support = {
        "pptx": "une présentation PowerPoint : chaque texte doit tenir dans le même cadre "
                "que l'original",
        "docx": "un document Word",
        "xlsx": "un classeur Excel : les cellules restent des libellés courts",
    }.get(kind, "un document")
    return f"""Tu es traducteur professionnel. Tu traduis les segments de {support}
(documents d'audit interne d'une entreprise de distribution) du {src_name} vers le {tgt_name}.

RÈGLES
1. Tout texte rédigé en {src_name} est traduit en {tgt_name}. Un texte déjà en {tgt_name}
   est recopié À L'IDENTIQUE. Un segment peut mélanger les langues : traduis uniquement
   la partie en {src_name} et garde le reste tel quel. Ne traduis jamais vers le {src_name}.
2. Le texte dans une autre langue (souvent l'anglais : « Best Practices », « Medium Risk »,
   « Sales Forecast ») reste tel quel.
3. Ne traduis pas : noms de personnes, de lieux, d'entreprises et de magasins, codes,
   identifiants et références, même s'ils contiennent un mot traduisible (MID: 1598643,
   RG18, DR09, KW42, FOR 3, WBF4.c, L_FR_Logistik), noms de fichiers (avec « _ » ou une
   extension), noms d'outils et de logiciels, sigles sans équivalent établi, nombres,
   dates, montants et unités. Traduis en revanche les intitulés cités (titre de document,
   de chapitre, de question, de processus, de rubrique) : seul leur numéro ou code reste
   tel quel ({cited}).
4. Balises : <sN>…</sN> marque un passage dans un style particulier (gras, couleur…).
   Traduis d'abord la phrase avec l'ordre des mots naturel en {tgt_name}{order_example},
   puis remets chaque balise autour des mots TRADUITS
   correspondants, où qu'ils soient. N'invente aucune balise, n'en supprime aucune, ne
   les imbrique pas.
   <xN/> est un objet (champ, image, lien, renvoi) : garde-le exactement une fois, à
   l'endroit correspondant. <br/> est un retour à la ligne, <tab/> une tabulation,
   <p/> une fin de paragraphe : garde-les en même nombre, à l'endroit naturel.
5. Longueur : reste aussi concis que l'original. Un titre ou un libellé reste bref et
   nominal ; quand plusieurs traductions sont justes, choisis la plus courte.
6. Respecte la casse du début (majuscule ou minuscule), la ponctuation finale et les
   espaces en début et fin de segment.
7. Le glossaire est prioritaire sur tout autre choix ({house}).
8. {grammar}
9. Réponds UNIQUEMENT par un objet JSON {{"identifiant": "traduction"}} contenant tous les
   identifiants reçus, sans commentaire."""


def _needs_translation(plain, direction):
    """Faut-il envoyer ce texte au modele ?

    Non s'il ne contient aucune lettre, ou s'il est nettement dans la langue
    cible sans aucun indice de la langue source (titre allemand dans un
    document traduit vers l'allemand). Dans le doute, on envoie : la regle 1
    du prompt demande de recopier ce qui est deja dans la bonne langue.
    """
    if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]{2,}", plain):
        return False
    src, tgt = LANG[direction][:2]
    s = _lang_scores(plain)
    if s[tgt] >= 2.0 and s[src] < 0.5:
        return False
    return True


def _chunks(keys):
    out, cur, size = [], [], 0
    for k in keys:
        n = len(k)
        if cur and (size + n > BATCH_CHARS or len(cur) >= BATCH_MAX_SEGMENTS):
            out.append(cur)
            cur, size = [], 0
        cur.append(k)
        size += n
    if cur:
        out.append(cur)
    return out


def _batch_prompt(rules, glossary, payload, extra=""):
    return (rules + (("\n\n" + glossary.strip()) if glossary else "") + extra
            + "\n\nSEGMENTS (JSON, identifiant -> texte) :\n"
            + json.dumps(payload, ensure_ascii=False, indent=0))


_IDENT_RE = re.compile(r"(?<![\w.])(?=[\w.]*[_\d])(?=[\w.]*[A-Za-z])[A-Za-z0-9][\w.]*[A-Za-z0-9]")


def _identifiers(text):
    """Codes et identifiants : mots avec un tiret bas, ou melant lettres et
    chiffres (RG18, WBF4.c, L_FR_Logistik). Ils doivent rester intacts."""
    plain = re.sub(r"<[^>]+>", " ", text)
    return {m for m in _IDENT_RE.findall(plain) if len(m) >= 3}


def _missing_identifiers(src, dst):
    return sorted(i for i in _identifiers(src) if i not in re.sub(r"<[^>]+>", " ", dst))


def _went_backwards(src, dst, direction):
    """La traduction contient-elle plus de langue source que l'original ?
    Signe qu'un passage deja dans la langue cible a ete traduit a l'envers
    ("Inventaires" devenu "Bestandsaufnahmen" vers le francais)."""
    s_lang = LANG[direction][0]
    plain = lambda t: re.sub(r"<[^>]+>", " ", t)  # noqa: E731
    return _lang_scores(plain(dst))[s_lang] > _lang_scores(plain(src))[s_lang] + 0.9


def _not_translated(src, dst, direction):
    """Le modele a-t-il renvoye le texte presque tel quel ? Une vraie phrase
    dans la langue source dont la « traduction » reste nettement dans cette
    langue (vu avec un glossaire : seuls les termes imposes remplaces,
    « LIA est un Tool zur Bestellung von Artikeln... »)."""
    s_lang, t_lang = LANG[direction][:2]
    plain = lambda t: re.sub(r"<[^>]+>", " ", t)  # noqa: E731
    a, b = _lang_scores(plain(src)), _lang_scores(plain(dst))
    return a[s_lang] >= 4.0 and b[s_lang] >= 0.5 * a[s_lang] and b[s_lang] > b[t_lang]


def _looks_chatty(src, dst):
    low = dst.lower()
    return any(c in low and c not in src.lower() for c in _CHATTY)


def translate_segments(segments, direction, kind, glossary_df=None,
                       progress_callback=None, contexts=None):
    """Traduit une liste de Segment.

    Renvoie (resultats, stats) ou resultats[i] est la liste d'items a
    reecrire pour segments[i], ou None si le segment doit rester tel quel.
    contexts[i] (optionnel) : courte description de l'emplacement
    (« titre », « cellule de tableau »...), donnee au modele.
    """
    stats = {"input_tokens": 0, "output_tokens": 0, "calls": 0, "retried": 0,
             "fallback": 0, "skipped": 0, "segments": len(segments), "glossary_retry": 0}
    rules = _segment_rules(direction, kind)
    entries = glossary_entries(glossary_df, direction)
    plain = lambda t: re.sub(r"<[^>]+>", " ", t)  # noqa: E731

    # Dedoublonnage : un meme texte balise n'est traduit qu'une fois.
    tagged = [s.to_tagged() if s.has_text() else None for s in segments]
    todo = {}
    for i, (s, t) in enumerate(zip(segments, tagged)):
        if t is None or not _needs_translation(s.plain_text(), direction):
            continue
        todo.setdefault(t, []).append(i)
    stats["skipped"] = len(segments) - sum(len(v) for v in todo.values())

    keys = list(todo)
    ids = {t: "t%d" % n for n, t in enumerate(keys)}
    batches = _chunks(keys)
    translations = {}
    lock = threading.Lock()
    done = [0]
    extra_ctx = ("\n\nChaque segment est donné avec son emplacement (« où ») ; "
                 "traduis uniquement « texte ».") if contexts else ""

    def run(batch, depth=0):
        gl = format_glossary_rules(matching_entries(" ".join(plain(t) for t in batch), entries))
        payload = {}
        for t in batch:
            ctx = contexts[todo[t][0]] if contexts else None
            payload[ids[t]] = {"où": ctx, "texte": t} if ctx else t
        data, i_tok, o_tok = generate_json(_batch_prompt(rules, gl, payload, extra_ctx))
        with lock:
            stats["input_tokens"] += i_tok
            stats["output_tokens"] += o_tok
            stats["calls"] += 1
        got = {}
        if isinstance(data, dict):
            for t in batch:
                v = data.get(ids[t])
                if isinstance(v, dict):
                    v = v.get("texte")
                if isinstance(v, str):
                    got[t] = v
        missing = [t for t in batch if t not in got]
        if missing and len(batch) > 1 and depth < 4:
            # Lot incomplet ou JSON invalide : on redecoupe la partie manquante.
            mid = max(1, len(missing) // 2)
            for part in (missing[:mid], missing[mid:]):
                if part:
                    got.update(run(part, depth + 1))
        return got

    def run_top(batch):
        got = run(batch)
        with lock:
            translations.update(got)
            done[0] += 1
            if progress_callback:
                progress_callback(done[0], max(1, len(batches)))

    if batches:
        with ThreadPoolExecutor(max_workers=min(MAX_WORKERS, len(batches))) as ex:
            list(ex.map(run_top, batches))

    # Validation des balises ; une seconde demande pour les segments fautifs.
    results = [None] * len(segments)
    retry = []
    for t in keys:
        seg = segments[todo[t][0]]
        tr = translations.get(t)
        if tr is None or _looks_chatty(t, tr):
            retry.append((t, "pas de traduction exploitable"))
            continue
        items, problem = seg.parse(tr)
        if not problem and _missing_identifiers(t, tr):
            problem = "identifiants modifiés : %s" % ", ".join(_missing_identifiers(t, tr))
        if not problem and _went_backwards(t, tr, direction):
            problem = "une partie a été traduite vers le %s au lieu du %s" % (
                LANG[direction][2], LANG[direction][3])
        if not problem and _not_translated(t, tr, direction):
            problem = "la traduction est restée en %s" % LANG[direction][2]
        if not problem and entries:
            bad = glossary_violations(plain(t), plain(tr), entries)
            if bad:
                problem = "glossaire non respecté : " + ", ".join(
                    "« %s » doit devenir %s" % (e["shown"], " ou ".join("« %s »" % x for x in e["targets"]))
                    for e in bad)
                stats["glossary_retry"] += 1
        if problem:
            retry.append((t, problem))
            continue
        for i in todo[t]:
            results[i] = items

    def run_retry(job):
        t, problem = job
        seg = segments[todo[t][0]]
        gl = format_glossary_rules(matching_entries(plain(t), entries))
        extra = ("\n\nATTENTION : une première traduction de ce segment a été rejetée "
                 f"({problem}). Respecte strictement les règles 3, 4 et 7.")
        data, i_tok, o_tok = generate_json(_batch_prompt(rules, gl, {"t0": t}, extra), kind="retry")
        tr = data.get("t0") if isinstance(data, dict) else None
        if isinstance(tr, dict):
            tr = tr.get("texte")
        if not isinstance(tr, str) or _looks_chatty(t, tr):
            tr = translations.get(t)
        items = None
        fell_back = False
        if tr is not None:
            items, problem2 = seg.parse(tr)
            if problem2 and not problem2.startswith("style perdu"):
                fell_back = True
                items = seg.fallback(tr)
        with lock:
            stats["retried"] += 1
            stats["fallback"] += int(fell_back)
            stats["input_tokens"] += i_tok
            stats["output_tokens"] += o_tok
            stats["calls"] += 1
            if items is not None:
                for i in todo[t]:
                    results[i] = items

    if retry:
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
            list(ex.map(run_retry, retry))

    stats["total_cost_usd"] = (stats["input_tokens"] / 1e6 * PRICE_IN
                               + stats["output_tokens"] / 1e6 * PRICE_OUT)
    return results, stats


SHORTEN_RULES = """Tu es traducteur professionnel ({tgt_name}). Chaque segment ci-dessous est une
traduction déjà faite qui ne tient pas dans son cadre. Propose une version plus courte,
d'au plus « max » caractères (balises non comptées), avec le même sens et le même registre :
termes plus courts, tournures nominales, suppression des mots superflus. N'introduis
aucun sigle ni aucune abréviation absents du texte (pas de « DG » pour « direction
générale ») et n'abrège avec un point que si c'est l'usage courant. Garde exactement les balises (<sN>…</sN>, <xN/>, <br/>, <tab/>,
<p/>), les nombres, codes et noms propres, et la casse du début. Ne supprime aucune
information (date, jour, mention entre parenthèses, passage balisé) : raccourcis la
formulation, pas le contenu.
Réponds UNIQUEMENT par un objet JSON {{"identifiant": "version courte"}}."""


def shorten_segments(items, direction):
    """items : [(cle, segment, traduction balisee, max caracteres)] ->
    ({cle: items reconstruits}, tokens entree, tokens sortie)."""
    if not items:
        return {}, 0, 0
    tgt_name = LANG[direction][3]
    payload = {"t%d" % n: {"max": mx, "texte": tr} for n, (_, _, tr, mx) in enumerate(items)}
    prompt = (SHORTEN_RULES.format(tgt_name=tgt_name)
              + "\n\nSEGMENTS :\n" + json.dumps(payload, ensure_ascii=False, indent=0))
    data, i_tok, o_tok = generate_json(prompt, kind="shorten")
    out = {}
    if isinstance(data, dict):
        for n, (key, seg, _, _) in enumerate(items):
            v = data.get("t%d" % n)
            if isinstance(v, dict):
                v = v.get("texte")
            if isinstance(v, str) and _keeps_information(items[n][2], v):
                parsed, problem = seg.parse(v)
                if not problem:
                    out[key] = parsed
    return out, i_tok, o_tok


_NUM_RE = re.compile(r"\d+(?:[.,:/]\d+)*")


def _keeps_information(full, short):
    """Une version courte ne doit rien perdre : memes balises de style,
    memes nombres et dates, memes identifiants."""
    tags = lambda t: sorted(re.findall(r"<s(\d+)>|<x(\d+)/>", t))  # noqa: E731
    if tags(full) != tags(short):
        return False
    if sorted(_NUM_RE.findall(re.sub(r"<[^>]+>", " ", full))) != \
            sorted(_NUM_RE.findall(re.sub(r"<[^>]+>", " ", short))):
        return False
    words = lambda t: re.findall(r"[^\W\d_]+", re.sub(r"<[^>]+>", " ", t))  # noqa: E731
    n_full, n_short = len(words(full)), len(words(short))
    if n_short < (n_full if n_full <= 6 else 0.6 * n_full):
        return False    # un libelle court ne perd aucun mot
    abbr = re.compile(r"(?<![\w.])([^\W\d_]{1,7})\.(?=[\s/|,)]|$)")
    have = set(abbr.findall(re.sub(r"<[^>]+>", " ", full)))
    if any(w not in have for w in abbr.findall(re.sub(r"<[^>]+>", " ", short))):
        return False    # abreviation inventee ("Com.", "Planif.")
    return not _missing_identifiers(full, short)


def items_to_tagged(seg, items):
    """Items reconstruits -> texte balise (pour redemander une version courte)."""
    base = seg.base_style()
    out = []
    for it in items:
        if it[0] == "t":
            out.append(it[2] if it[1] == base else "<s%d>%s</s%d>" % (it[1], it[2], it[1]))
        elif it[0] == "x":
            out.append("<x%d/>" % it[1])
        else:
            out.append("<%s/>" % it[0])
    return "".join(out)