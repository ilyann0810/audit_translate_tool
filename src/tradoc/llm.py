# -*- coding: utf-8 -*-
"""
Phase 1 - Traduction des segments par Gemini, avec controle du prompt.

C'est ce que translateDocument ne permet pas. Quatre leviers, chacun
repondant a un defaut mesure en Phase 0 :

  1. glossaire metier      -> "Umsatz" ne devient plus "Volume des ventes"
  2. protection litterale  -> {kunde} reste {kunde}
  3. contrainte de longueur-> limite le debordement des zones de slide
  4. contexte par segment  -> un titre reste un titre, avec sa majuscule

Et deux garde-fous repris du PROMPT_CLAUDE_CODE.md :

  - groupement PAR DOCUMENT, pas par lot arbitraire : le modele doit voir
    le document entier pour etre coherent sur le ton et la terminologie
  - validation BLOQUANTE de l'alignement des identifiants : si l'ensemble
    des ids renvoyes differe de l'ensemble envoye, on ne persiste RIEN et
    on relance avec un lot plus petit
"""
from __future__ import annotations

import json
import os
import re


DEFAULT_MODEL = "gemini-2.5-flash"

# Ecart de longueur tolere avant de considerer qu'un segment risque de
# deborder. Le PROMPT_CLAUDE_CODE.md retient plus ou moins 15 %.
LENGTH_TOLERANCE = 0.15

# Meme expression que celle du comparateur de Phase 0 : ce qui est detecte
# comme placeholder doit etre exactement ce qui est protege.
PLACEHOLDER_RE = re.compile(
    r"\{[^{}\n]{1,40}\}"
    r"|\$\{[^{}\n]{1,40}\}"
    r"|%\([^)\n]{1,40}\)[sdfr]"
    r"|%[sdfr]"
    r"|<<[^<>\n]{1,40}>>"
    r"|\[\[[^\[\]\n]{1,40}\]\]"
)


# --------------------------------------------------------------------------
# Glossaire
# --------------------------------------------------------------------------

def load_glossary(path):
    """Glossaire metier, au format JSON {terme source: terme cible}.

    C'est le levier le moins couteux et le plus efficace : il corrige la
    terminologie sans rien changer au reste de la traduction.
    """
    if not path or not os.path.isfile(path):
        return {}
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def relevant_glossary(glossary, segments):
    """Ne transmet que les termes reellement presents dans le document.

    Un glossaire de 400 entrees envoye a chaque appel coute des tokens
    pour rien et dilue l'attention du modele.
    """
    blob = "\n".join(s.text for s in segments).lower()
    return {k: v for k, v in glossary.items() if k.lower() in blob}


# --------------------------------------------------------------------------
# Prompt
# --------------------------------------------------------------------------

SYSTEM = """Tu es traducteur professionnel de documents d'entreprise.

Tu traduis du {source} vers le {target}. Domaine : {domain}.

Tu recois un tableau JSON de segments. Chaque segment a un identifiant, un
texte, et une indication de l'endroit d'ou il vient dans le document.

REGLES ABSOLUES

1. Renvoie EXACTEMENT un objet par segment recu, avec le meme identifiant.
   N'en fusionne aucun, n'en decoupe aucun, n'en omets aucun, n'en ajoute
   aucun. L'ordre doit etre conserve.

2. Ne traduis JAMAIS le contenu de ces marqueurs, recopie-les caractere
   pour caractere : {{variable}}, ${{variable}}, %s, %d, %(nom)s,
   <<champ>>, [[champ]]. Le nom interne d'une variable n'est pas du texte :
   c'est une cle technique. Traduire {{kunde}} en {{client}} casse le
   document.

3. Longueur : reste a plus ou moins {tol} pour cent du nombre de caracteres
   de la source. Les zones de texte des presentations ne s'agrandissent
   pas. Si la traduction naturelle est plus longue, choisis une formulation
   plus courte plutot que de deborder.

4. Respecte la casse et la ponctuation de la source. Un segment qui
   commence par une majuscule garde sa majuscule. Un titre sans point
   final n'en prend pas.

5. Adapte le registre a l'endroit d'ou vient le segment :
   - un titre ou un libelle court reste nominal et bref
   - une cellule de tableau reste un libelle, pas une phrase
   - une note du presentateur peut etre plus parlee
   - un encadre garde son caractere d'avertissement

6. Un segment qui n'est qu'un nombre, un sigle, un nom propre ou une unite
   se recopie tel quel.

7. Le texte peut commencer ou finir par une espace : conserve-la, elle
   colle deux morceaux de phrase.

8. Certains textes contiennent des balises <s1>...</s1>, <s2>...</s2> qui
   marquent un style (gras, couleur, autre police). Garde chaque balise
   presente dans la source, autour des mots TRADUITS qui correspondent au
   passage balise. N'invente aucune balise, n'en supprime aucune, ne les
   imbrique pas. Si l'ordre des mots change, deplace la balise avec ses
   mots.

9. Un retour a la ligne \\n dans la source est voulu (titre sur deux
   lignes, element de liste) : conserve-le au meme endroit du sens.

10. Un texte deja en {target} est recopie a l'identique. Un segment qui
    melange les langues : ne traduis que la partie en {source}. Le texte
    dans une autre langue (souvent l'anglais) reste tel quel.

11. Codes, identifiants et references restent intacts, meme s'ils
    contiennent un mot traduisible (RG18, WBF4.c, L_FR_Logistik, KW42).
    Une reference de magasin suit l'usage de la langue cible : « Filiale
    4440 » en allemand, « SPM 4440 » en francais.
"""

USER = """Document : {name}
Segments a traduire : {count}
{glossary_block}
{segments}"""


def build_prompt(segments, source, target, domain, glossary, name):
    gl = relevant_glossary(glossary, segments)
    if gl:
        lines = "\n".join("  %s = %s" % (k, v) for k, v in sorted(gl.items()))
        glossary_block = (
            "\nGLOSSAIRE IMPOSE - utilise ces equivalents, pas d'autres :\n"
            "%s\n" % lines)
    else:
        glossary_block = ""

    payload = [
        {"id": s.uid, "ou": s.context_label(), "texte": s.text}
        for s in segments
    ]

    system = SYSTEM.format(
        source=source, target=target, domain=domain,
        tol=int(LENGTH_TOLERANCE * 100))
    user = USER.format(
        name=name, count=len(segments), glossary_block=glossary_block,
        segments=json.dumps(payload, ensure_ascii=False, indent=1))
    return system, user, gl


RESPONSE_SCHEMA = {
    "type": "ARRAY",
    "items": {
        "type": "OBJECT",
        "properties": {
            "id": {"type": "STRING"},
            "texte": {"type": "STRING"},
        },
        "required": ["id", "texte"],
    },
}


# --------------------------------------------------------------------------
# Appel
# --------------------------------------------------------------------------

class AlignmentError(RuntimeError):
    """Les identifiants renvoyes ne correspondent pas a ceux envoyes."""


def _client(project, location):
    from google import genai
    import os
    import json
    import tempfile

    # Adaptation Databricks : création du fichier ADC à partir du secret en mémoire
    gcp_key = os.environ.get("GCP_SERVICE_KEY")
    if gcp_key and not os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"):
        tmp_path = os.path.join(tempfile.gettempdir(), "databricks_gcp_key.json")
        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(gcp_key)
        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = tmp_path
        
        # Récupération automatique du project_id si non fourni
        if not project:
            project = json.loads(gcp_key).get("project_id")

    return genai.Client(vertexai=True, project=project, location=location)


# Backoff exponentiel sur les refus temporaires. C'est le piege n8 du
# PROMPT_CLAUDE_CODE.md, qui vise les quotas d'ecriture Google : il vaut
# aussi pour Vertex AI, qui renvoie 429 RESOURCE_EXHAUSTED des qu'un gros
# document sature le quota de tokens par minute. Sans cette reprise, un
# document long echoue au milieu et la traduction deja payee est perdue.
RETRY_STATUS = ("429", "RESOURCE_EXHAUSTED", "503", "UNAVAILABLE",
                "500", "INTERNAL", "DEADLINE_EXCEEDED")
MAX_RETRIES = 5
BASE_DELAY = 12.0


TOKENS = {"input": 0, "output": 0}


def _count_tokens(resp):
    um = getattr(resp, "usage_metadata", None)
    TOKENS["input"] += getattr(um, "prompt_token_count", 0) or 0
    TOKENS["output"] += getattr(um, "candidates_token_count", 0) or 0


def _is_transient(exc):
    text = str(exc)
    return any(s in text for s in RETRY_STATUS)


def _call(client, model, system, user, log=print):
    import random
    import time

    from google.genai import types

    config = types.GenerateContentConfig(
        system_instruction=system,
        response_mime_type="application/json",
        response_schema=RESPONSE_SCHEMA,
        temperature=0.2,
    )

    last = None
    for attempt in range(MAX_RETRIES):
        try:
            resp = client.models.generate_content(
                model=model, contents=user, config=config)
            _count_tokens(resp)
            return json.loads(resp.text)
        except ValueError as exc:
            # JSON invalide (reponse tronquee ou mal formee) : une nouvelle
            # tentative, puis on le traite comme un lot desaligne, que
            # l'appelant redecoupe au lieu d'abandonner le document.
            last = exc
            if attempt >= 1:
                raise AlignmentError("reponse JSON invalide : %s" % exc)
        except Exception as exc:
            last = exc
            if not _is_transient(exc) or attempt == MAX_RETRIES - 1:
                raise
            # Un peu d'aleatoire pour ne pas resynchroniser plusieurs
            # documents qui repartiraient tous en meme temps.
            delay = BASE_DELAY * (2 ** attempt) * (0.8 + 0.4 * random.random())
            log("    QUOTA : %s - nouvelle tentative dans %.0f s (%d/%d)"
                % (type(exc).__name__, delay, attempt + 1, MAX_RETRIES - 1))
            time.sleep(delay)
    raise last


def translate_batch(client, model, segments, source, target, domain,
                    glossary, name, log=print):
    """Traduit un lot et VERIFIE l'alignement avant de rendre la main.

    La verification n'est pas optionnelle. Un modele qui fusionne deux
    segments passe inapercu jusqu'a l'ouverture du document final, ou la
    moitie du texte est decalee.
    """
    system, user, _ = build_prompt(segments, source, target, domain,
                                   glossary, name)
    data = _call(client, model, system, user, log=log)

    got = {}
    for row in data:
        if isinstance(row, dict) and "id" in row and "texte" in row:
            got[row["id"]] = row["texte"]

    sent_ids = {s.uid for s in segments}
    if set(got) != sent_ids:
        missing = sorted(sent_ids - set(got))
        extra = sorted(set(got) - sent_ids)
        raise AlignmentError(
            "%d envoyes, %d renvoyes ; %d manquants, %d inattendus"
            % (len(sent_ids), len(got), len(missing), len(extra)))

    return got


# Budget de caracteres par appel. Le PROMPT_CLAUDE_CODE.md demande de
# grouper par document pour la coherence du ton et de la terminologie ;
# un document de 70 000 caracteres depasse largement ce qu'un seul appel
# peut rendre en sortie. On decoupe alors en tranches CONSECUTIVES, ce qui
# garde ensemble les slides voisines et donc l'essentiel du contexte.
CHARS_PER_CALL = 14000


def _chunk_by_chars(segments, budget):
    out, cur, size = [], [], 0
    for s in segments:
        n = len(s.text)
        if cur and size + n > budget:
            out.append(cur)
            cur, size = [], 0
        cur.append(s)
        size += n
    if cur:
        out.append(cur)
    return out


def _cache_key(model, source, target, domain, glossary, text):
    import hashlib
    raw = json.dumps([model, source, target, domain,
                      sorted((glossary or {}).items()), text],
                     ensure_ascii=False)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def translate_document(segments, source, target, domain, glossary,
                       project, location, model=DEFAULT_MODEL,
                       name="document", max_batch=None, log=print,
                       chars_per_call=CHARS_PER_CALL, cache_path=None):
    """Traduit les segments, en reutilisant ce qui est deja en cache.

    Le cache est la table silver.translations du PROMPT_CLAUDE_CODE.md en
    version fichier : un texte deja traduit avec le meme modele, les memes
    langues, le meme domaine et le meme glossaire n'est jamais repaye.
    """
    cache = {}
    if cache_path and os.path.isfile(cache_path):
        with open(cache_path, encoding="utf-8") as fh:
            cache = json.load(fh)

    keys = {s.uid: _cache_key(model, source, target, domain, glossary, s.text)
            for s in segments}
    hits = {s.uid: cache[keys[s.uid]] for s in segments if keys[s.uid] in cache}
    todo = [s for s in segments if s.uid not in hits]
    if hits:
        log("    cache : %d segments deja traduits, %d a traduire"
            % (len(hits), len(todo)))

    out, stats = ({}, {"calls": 0, "retries": 0, "max_depth": 0}) \
        if not todo else _translate_uncached(
            todo, source, target, domain, glossary, project, location,
            model, name, max_batch, log, chars_per_call)

    if cache_path and out:
        src = {s.uid: s.text for s in todo}
        for uid, text in out.items():
            # Un segment abandonne apres desalignement revient en langue
            # source : on ne le fige pas dans le cache.
            if text != src.get(uid):
                cache[keys[uid]] = text
        os.makedirs(os.path.dirname(os.path.abspath(cache_path)), exist_ok=True)
        # Deux traductions lancees en parallele partagent ce fichier : on
        # relit ce que l'autre a pu ecrire entre-temps avant d'enregistrer,
        # sinon la derniere a finir efface le travail de la premiere.
        if os.path.isfile(cache_path):
            try:
                with open(cache_path, encoding="utf-8") as fh:
                    cache = {**json.load(fh), **cache}
            except ValueError:
                pass
        tmp = cache_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(cache, fh, ensure_ascii=False)
        os.replace(tmp, cache_path)

    out.update(hits)
    return out, stats


def _translate_uncached(segments, source, target, domain, glossary,
                        project, location, model, name, max_batch, log,
                        chars_per_call):
    """Traduit tous les segments d'un document.

    Par defaut, un seul appel pour tout le document : c'est la consigne du
    PROMPT_CLAUDE_CODE.md, le modele doit voir l'ensemble pour rester
    coherent. En cas d'echec d'alignement, on redecoupe en deux et on
    recommence, jusqu'a isoler le lot fautif.
    """
    client = _client(project, location)
    out = {}
    stats = {"calls": 0, "retries": 0, "max_depth": 0}

    def run(batch, depth=0):
        stats["max_depth"] = max(stats["max_depth"], depth)
        if not batch:
            return
        try:
            stats["calls"] += 1
            out.update(translate_batch(client, model, batch, source, target,
                                       domain, glossary, name, log=log))
        except AlignmentError as exc:
            if len(batch) == 1:
                # Un seul segment et toujours desaligne : on le laisse en
                # langue source plutot que d'ecrire n'importe quoi.
                log("    ALIGNEMENT : segment %s abandonne (%s)"
                    % (batch[0].uid, exc))
                out[batch[0].uid] = batch[0].text
                return
            stats["retries"] += 1
            mid = len(batch) // 2
            log("    ALIGNEMENT : %s -> on redecoupe en %d + %d"
                % (exc, mid, len(batch) - mid))
            run(batch[:mid], depth + 1)
            run(batch[mid:], depth + 1)

    if max_batch:
        chunks = [segments[i:i + max_batch]
                  for i in range(0, len(segments), max_batch)]
    else:
        chunks = _chunk_by_chars(segments, chars_per_call)

    if len(chunks) > 1:
        log("    document decoupe en %d tranches consecutives" % len(chunks))
    for chunk in chunks:
        run(chunk)

    return out, stats


# --------------------------------------------------------------------------
# Raccourcissement (contrainte de longueur dure)
# --------------------------------------------------------------------------

SHORTEN_SYSTEM = """Tu es traducteur professionnel, langue cible : {target}.

Chaque segment recu est une traduction deja faite qui ne tient pas dans
son emplacement. Propose une version plus courte, d'au plus "max"
caracteres (balises non comptees), avec le meme sens et le meme registre.
Prefere les termes courts et les tournures nominales. N'introduis aucun
sigle ni aucune abreviation absents du texte (pas de « DG », pas de
« struct. ») ; n'abrege avec un point que si c'est l'usage courant. Garde exactement les balises <s1>...</s1>,
les variables {{...}} et %s, les nombres et les unites. Ne supprime aucune
information : ni date, ni jour, ni mention entre parentheses, ni passage
balise ; raccourcis la formulation, pas le contenu. Respecte la casse
du debut. Renvoie un objet par segment, avec le meme identifiant."""


def shorten(items, target, project, location, model=DEFAULT_MODEL,
            log=print):
    """items : [(id, traduction, max caracteres, contexte)] -> {id: texte}.

    C'est la relance "avec une contrainte plus dure" prevue par le
    PROMPT_CLAUDE_CODE.md pour les segments trop longs : on la fait apres
    avoir mesure le rendu reel, pas au jugé.
    """
    if not items:
        return {}
    client = _client(project, location)
    payload = [{"id": uid, "ou": ctx, "max": n, "texte": txt}
               for uid, txt, n, ctx in items]
    try:
        data = _call(client, model, SHORTEN_SYSTEM.format(target=target),
                     json.dumps(payload, ensure_ascii=False, indent=1), log=log)
    except Exception as exc:
        # Les versions courtes sont un plus : en cas d'echec, on garde les
        # traductions completes (reduites a l'ecriture si besoin).
        log("    raccourcissement impossible (%s) : traductions completes gardees" % exc)
        return {}
    wanted = {uid for uid, _, _, _ in items}
    return {row["id"]: row["texte"] for row in data
            if isinstance(row, dict) and row.get("id") in wanted
            and row.get("texte")}


# --------------------------------------------------------------------------
# Lecture d'images (texte illisible dans la couche texte)
# --------------------------------------------------------------------------

TRANSCRIBE_PROMPT = """Chaque image ci-dessous est un extrait d'une page de
document, precede de son identifiant. Transcris EXACTEMENT le texte visible
sur chaque image, sans le traduire, sans le corriger, sans rien ajouter.
Conserve les majuscules, la ponctuation, les chiffres et les unites.
Separe les lignes par \\n. Si une image ne contient pas de texte lisible,
renvoie une chaine vide pour elle. Renvoie un objet par image."""


def transcribe_images(items, project, location, model=DEFAULT_MODEL,
                      log=print, per_call=30):
    """Transcrit le texte d'images : [(id, octets PNG)] -> {id: texte}.

    Sert quand un PDF dessine du texte avec une police dont la table de
    caracteres est fausse : a l'ecran le texte est lisible, mais l'extraction
    donne du charabia. La seule source fiable est alors l'image elle-meme.
    """
    from google.genai import types

    client = _client(project, location)
    out = {}
    for i in range(0, len(items), per_call):
        batch = items[i:i + per_call]
        parts = [TRANSCRIBE_PROMPT]
        for uid, png in batch:
            parts.append("id: %s" % uid)
            parts.append(types.Part.from_bytes(data=png, mime_type="image/png"))
        config = types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=RESPONSE_SCHEMA, temperature=0.0)
        for attempt in range(MAX_RETRIES):
            try:
                resp = client.models.generate_content(
                    model=model, contents=parts, config=config)
                _count_tokens(resp)
                break
            except Exception as exc:
                if not _is_transient(exc) or attempt == MAX_RETRIES - 1:
                    raise
                import time
                time.sleep(BASE_DELAY * (2 ** attempt))
        wanted = {uid for uid, _ in batch}
        try:
            rows = json.loads(resp.text)
        except ValueError:
            log("    lecture visuelle : reponse illisible, zones laissees telles quelles")
            continue
        for row in rows:
            if isinstance(row, dict) and row.get("id") in wanted:
                out[row["id"]] = (row.get("texte") or "").strip()
    log("    lecture visuelle : %d zones transcrites" % len(out))
    return out


# --------------------------------------------------------------------------
# Controle qualite (task 5 du PROMPT_CLAUDE_CODE.md)
# --------------------------------------------------------------------------

def qa_flags(segment, translated):
    """Drapeaux automatiques, poses avant toute ecriture."""
    flags = []
    src, tgt = segment.text.strip(), translated.strip()

    if len(src) > 10 and len(tgt) > len(src) * (1 + LENGTH_TOLERANCE * 2):
        flags.append("OVERFLOW_RISK")

    if PLACEHOLDER_RE.findall(src) != PLACEHOLDER_RE.findall(tgt):
        flags.append("PLACEHOLDER_MISMATCH")

    if tgt == src and len(src) > 30:
        flags.append("UNTRANSLATED")

    if src and tgt and src[0].isupper() and tgt[0].islower():
        flags.append("CASE_LOST")

    if not tgt:
        flags.append("EMPTY")

    return flags