import hashlib
import json
import os
import threading
from vertexai.generative_models import GenerationConfig

from config.settings import gemini_model, DEFAULT_SAFETY_SETTINGS, DEFAULT_MODEL

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
    if is_batch:
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
Traduis le texte suivant en conservant son formatage et ses balises.

TEXTE À TRADUIRE :
{content}
"""
    return prompt

def call_gemini_batch(prompt, expected_count):
    generation_config = GenerationConfig(temperature=0.1, response_mime_type="application/json")
    try:
        cached = _cache_get("batch", prompt)
        if cached is not None:
            text, in_tok, out_tok = cached["text"], 0, 0
        else:
            response = gemini_model.generate_content(
                prompt, generation_config=generation_config, safety_settings=DEFAULT_SAFETY_SETTINGS
            )
            in_tok = getattr(response.usage_metadata, 'prompt_token_count', 0) if hasattr(response, 'usage_metadata') else 0
            out_tok = getattr(response.usage_metadata, 'candidates_token_count', 0) if hasattr(response, 'usage_metadata') else 0

            if not response.candidates or not response.candidates[0].content.parts:
                return None, in_tok, out_tok
            text = response.text

        data = json.loads(text)
        translations = [data.get(str(i)) for i in range(expected_count)]
        if any(t is None for t in translations):
            return None, in_tok, out_tok

        _cache_put("batch", prompt, text)
        return translations, in_tok, out_tok
    except Exception:
        return None, 0, 0

def call_gemini_fallback(prompt):
    generation_config = GenerationConfig(temperature=0.1)
    try:
        cached = _cache_get("single", prompt)
        if cached is not None:
            return cached["text"], 0, 0

        response = gemini_model.generate_content(
            prompt, generation_config=generation_config, safety_settings=DEFAULT_SAFETY_SETTINGS
        )
        in_tok = getattr(response.usage_metadata, 'prompt_token_count', 0) if hasattr(response, 'usage_metadata') else 0
        out_tok = getattr(response.usage_metadata, 'candidates_token_count', 0) if hasattr(response, 'usage_metadata') else 0

        if not response.candidates or not response.candidates[0].content.parts:
            return None, in_tok, out_tok

        _cache_put("single", prompt, response.text.strip())
        return response.text.strip(), in_tok, out_tok
    except Exception:
        return None, 0, 0

# ==========================================================================
# Moteur balisé
# ==========================================================================

import random
import re
import time
from concurrent.futures import ThreadPoolExecutor

from core.lang_detect import scores as _lang_scores
from core.text_utils import (glossary_entries, matching_entries, format_glossary_rules, glossary_violations)

LANG = {
    "Allemand ➔ Français": ("de", "fr", "allemand", "français"),
    "Français ➔ Allemand": ("fr", "de", "français", "allemand"),
}

PRICE_IN, PRICE_OUT = 0.30, 2.50
BATCH_CHARS = 6000
BATCH_MAX_SEGMENTS = 60
MAX_WORKERS = 8

_RETRY_MARKERS = ("429", "resource exhausted", "resource_exhausted", "503", "unavailable",
                  "500", "internal", "deadline", "timeout", "temporarily")
MAX_RETRIES = 6

_CHATTY = ("voici la traduction", "here is the translation", "hier ist die übersetzung",
           "pourriez-vous", "veuillez fournir", "texte à traduire", "n'hésitez pas")

def _is_transient(exc):
    t = str(exc).lower()
    return any(m in t for m in _RETRY_MARKERS)

def _json_config():
    """Reponse JSON, sans reflexion (thinking) : gemini-2.5-flash reflechit par
    defaut, ce qui double la duree et le cout de chaque appel sans ameliorer
    la traduction. Si le SDK installe ne connait pas thinking_config, on garde
    la configuration simple."""
    try:
        return GenerationConfig.from_dict({"temperature": 0.1,
                                           "response_mime_type": "application/json",
                                           "thinking_config": {"thinking_budget": 0}})
    except Exception:
        return GenerationConfig(temperature=0.1, response_mime_type="application/json")

def generate_json(prompt, kind="segments"):
    cached = _cache_get(kind, prompt)
    if cached is not None:
        try:
            return json.loads(cached["text"]), 0, 0
        except ValueError:
            pass
    config = _json_config()
    in_tok = out_tok = 0
    for attempt in range(MAX_RETRIES):
        try:
            response = gemini_model.generate_content(
                prompt, generation_config=config, safety_settings=DEFAULT_SAFETY_SETTINGS)
            um = getattr(response, "usage_metadata", None)
            in_tok += getattr(um, "prompt_token_count", 0) or 0
            # La reflexion est facturee comme de la sortie.
            out_tok += (getattr(um, "candidates_token_count", 0) or 0) \
                + (getattr(um, "thoughts_token_count", 0) or 0)
            if not response.candidates or not response.candidates[0].content.parts:
                return None, in_tok, out_tok
            text = response.text
            data = json.loads(text)
            _cache_put(kind, prompt, text)
            return data, in_tok, out_tok
        except ValueError:
            if attempt >= 1:
                return None, in_tok, out_tok
        except Exception as exc:
            if not _is_transient(exc) or attempt == MAX_RETRIES - 1:
                return None, in_tok, out_tok
            delay = min(60.0, 4.0 * (2 ** attempt)) * (0.75 + 0.5 * random.random())
            time.sleep(delay)
    return None, in_tok, out_tok

def _segment_rules(direction, kind):
    src, tgt, src_name, tgt_name = LANG[direction]
    grammar = {
        "fr": "Français correct : contractions (du, des, aux, auprès des), accords, espace avant « : ; ? ! ».",
        "de": "Allemand correct : majuscule aux substantifs, déclinaisons, mots composés soudés.",
    }[tgt]
    house = ("« Filiale 4440 » se traduit « SPM 4440 »" if tgt == "fr"
             else "« SPM 4440 » et « supermarché 4440 » se traduisent « Filiale 4440 »")
    support = {
        "pptx": "une présentation PowerPoint : chaque texte doit tenir dans le même cadre que l'original",
        "docx": "un document Word",
        "xlsx": "un classeur Excel : les cellules restent des libellés courts",
    }.get(kind, "un document")
    return f"""Tu es traducteur professionnel. Tu traduis les segments de {support}
(documents d'audit interne d'une entreprise de distribution) du {src_name} vers le {tgt_name}.

RÈGLES
1. Tout texte rédigé en {src_name} est traduit en {tgt_name}. Un texte déjà en {tgt_name}
   est recopié À L'IDENTIQUE. Un segment peut mélanger les langues : traduis uniquement
   la partie en {src_name} et garde le reste tel quel. Ne traduis jamais vers le {src_name}.
2. Le texte dans une autre langue (anglais) reste tel quel.
3. Ne traduis pas : noms propres, codes, identifiants (RG18, DR09, KW42, L_FR_Logistik), montants et unités.
4. Balises : <sN>…</sN> marque un style particulier. Conserve exactement toutes les balises.
   <xN/> est un objet : garde-le tel quel. <br/> est un saut de ligne.
5. Longueur : reste concis et direct.
6. Le glossaire est prioritaire ({house}).
7. {grammar}
8. Réponds UNIQUEMENT par un objet JSON {{"identifiant": "traduction"}} sans aucun commentaire."""

def _needs_translation(plain, direction):
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
    plain = re.sub(r"<[^>]+>", " ", text)
    return {m for m in _IDENT_RE.findall(plain) if len(m) >= 3}

def _missing_identifiers(src, dst):
    return sorted(i for i in _identifiers(src) if i not in re.sub(r"<[^>]+>", " ", dst))

def _went_backwards(src, dst, direction):
    s_lang = LANG[direction][0]
    plain = lambda t: re.sub(r"<[^>]+>", " ", t)
    return _lang_scores(plain(dst))[s_lang] > _lang_scores(plain(src))[s_lang] + 0.9

def _not_translated(src, dst, direction):
    s_lang, t_lang = LANG[direction][:2]
    plain = lambda t: re.sub(r"<[^>]+>", " ", t)
    a, b = _lang_scores(plain(src)), _lang_scores(plain(dst))
    if a[s_lang] >= 4.0 and b[s_lang] >= 0.5 * a[s_lang] and b[s_lang] > b[t_lang]:
        return True
    # Libelle court recopie tel quel alors qu'il est nettement dans la langue
    # source (« NAT Konzept zur O+G-Rettertüte ») : le seuil ci-dessus ne voit
    # pas les textes courts.
    return (" ".join(plain(dst).split()) == " ".join(plain(src).split())
            and a[s_lang] >= 1.5 and a[s_lang] > a[t_lang])

def _looks_chatty(src, dst):
    low = dst.lower()
    return any(c in low and c not in src.lower() for c in _CHATTY)

def _safe_progress(callback, cur, tot, msg):
    """Envoie la progression au callback sans risquer d'erreur de signature."""
    if not callback:
        return
    try:
        callback(cur, tot, msg)
    except TypeError:
        callback(cur, tot)

def translate_segments(segments, direction, kind, glossary_df=None,
                       progress_callback=None, contexts=None):
    stats = {"input_tokens": 0, "output_tokens": 0, "calls": 0, "retried": 0,
             "fallback": 0, "skipped": 0, "segments": len(segments), "glossary_retry": 0}
    rules = _segment_rules(direction, kind)
    entries = glossary_entries(glossary_df, direction)
    plain = lambda t: re.sub(r"<[^>]+>", " ", t)

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
            # Étape 1 : Progression des lots de 0 à 65 %
            pct = int((done[0] / max(1, len(batches))) * 65)
            _safe_progress(progress_callback, pct, 100, f"Traduction par l'IA : lot {done[0]}/{max(1, len(batches))}...")

    if batches:
        # Message immediat : la barre ne bouge qu'au retour du premier lot.
        _safe_progress(progress_callback, 1, 100,
                       f"🤖 Traduction par l'IA : {len(batches)} lot(s) envoyé(s)...")
        with ThreadPoolExecutor(max_workers=min(MAX_WORKERS, len(batches))) as ex:
            list(ex.map(run_top, batches))

    # Étape 2 : Validation des balises et filtrage des retries
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
            problem = "traduit à l'envers"
        if not problem and _not_translated(t, tr, direction):
            problem = "texte non traduit"
        if not problem and entries:
            bad = glossary_violations(plain(t), plain(tr), entries)
            if bad:
                problem = "glossaire non respecté"
                stats["glossary_retry"] += 1
        if problem:
            retry.append((t, problem))
            continue
        for i in todo[t]:
            results[i] = items

    # Étape 3 : Traitement transparent des retries (de 65 à 80 %)
    if retry:
        total_retries = len(retry)
        retry_done = [0]
        _safe_progress(progress_callback, 66, 100, f"🔍 Contrôle qualité : correction de {total_retries} segments...")

        def run_retry(job):
            t, problem = job
            seg = segments[todo[t][0]]
            gl = format_glossary_rules(matching_entries(plain(t), entries))
            extra = f"\n\nATTENTION : première traduction rejetée ({problem}). Respecte strictement les règles."
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
                retry_done[0] += 1
                pct = 65 + int((retry_done[0] / total_retries) * 15)
                _safe_progress(progress_callback, pct, 100, f"🔍 Contrôle qualité : segment corrigé {retry_done[0]}/{total_retries}...")
                if items is not None:
                    for i in todo[t]:
                        results[i] = items

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
            list(ex.map(run_retry, retry))

    stats["total_cost_usd"] = (stats["input_tokens"] / 1e6 * PRICE_IN + stats["output_tokens"] / 1e6 * PRICE_OUT)
    return results, stats

SHORTEN_RULES = """Tu es traducteur professionnel ({tgt_name}). Chaque segment ci-dessous est une
traduction déjà faite qui ne tient pas dans son cadre. Propose une version plus courte,
d'au plus « max » caractères (balises non comptées), avec le même sens.
Réponds UNIQUEMENT par un objet JSON {{"identifiant": "version courte"}}."""

def shorten_segments(items, direction):
    if not items:
        return {}, 0, 0
    tgt_name = LANG[direction][3]
    payload = {"t%d" % n: {"max": mx, "texte": tr} for n, (_, _, tr, mx) in enumerate(items)}
    prompt = (SHORTEN_RULES.format(tgt_name=tgt_name) + "\n\nSEGMENTS :\n" + json.dumps(payload, ensure_ascii=False, indent=0))
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
    tags = lambda t: sorted(re.findall(r"<s(\d+)>|<x(\d+)/>", t))
    if tags(full) != tags(short):
        return False
    if sorted(_NUM_RE.findall(re.sub(r"<[^>]+>", " ", full))) != sorted(_NUM_RE.findall(re.sub(r"<[^>]+>", " ", short))):
        return False
    words = lambda t: re.findall(r"[^\W\d_]+", re.sub(r"<[^>]+>", " ", t))
    n_full, n_short = len(words(full)), len(words(short))
    if n_short < (n_full if n_full <= 6 else 0.6 * n_full):
        return False
    return not _missing_identifiers(full, short)

def items_to_tagged(seg, items):
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