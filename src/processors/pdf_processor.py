import os
import re
import tempfile
import time

from tradoc import llm as tradoc_llm
from tradoc.pdf import translate_pdf
from core.text_utils import load_glossary_df
from core.translation_api import LANG, PRICE_IN, PRICE_OUT
from config.settings import DEFAULT_MODEL, DEFAULT_LOCATION, get_project_id


def translate_pdf_universal(input_path: str, output_path: str, glossary_path: str = None, 
                             progress_callback=None, direction="Allemand ➔ Français"):
    """
    Lanceur Databricks vers le moteur 'tradoc' avec suivi dynamique dans Streamlit.
    """
    start_time = time.time()
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    source_code, target_code, source_name, target_name = LANG[direction]

    # Glossaire dans le sens de la traduction
    glossary_dict = {}
    df = load_glossary_df(glossary_path) if glossary_path else None
    if df is not None and not df.empty:
        col_src, col_tgt = ("Allemand", "Francais") if source_code == "de" else ("Francais", "Allemand")
        glossary_dict = dict(zip(df[col_src].astype(str), df[col_tgt].astype(str)))

    # Activation systématique d'un cache disque pour éviter de repayer en cas de relance
    cache_dir = os.environ.get("TRANSLATION_CACHE_DIR") or os.path.join(tempfile.gettempdir(), "tradoc_cache")
    os.makedirs(cache_dir, exist_ok=True)
    cache_file = os.path.join(cache_dir, "pdf_cache.json")

    print(f"🚀 Lancement du moteur tradoc ({DEFAULT_MODEL}) : {source_code} ➔ {target_code}")
    tradoc_llm.TOKENS["input"] = tradoc_llm.TOKENS["output"] = 0

    # Le moteur tradoc annonce lui-meme ses etapes (analyse, traduction par
    # parties, ajustement, mise en page page par page) : on les relaie tel
    # quel. Les messages techniques du journal restent dans la console, sauf
    # l'attente sur quota, qui explique une pause a l'utilisateur.
    last = {"pct": 0}

    def pdf_progress(pct, label):
        last["pct"] = int(pct)
        if progress_callback:
            progress_callback(last["pct"], 100, label)

    def pdf_logger(msg):
        print(msg)
        m = re.search(r"QUOTA .* dans (\d+) s", str(msg))
        if m and progress_callback:
            progress_callback(last["pct"], 100,
                              f"⏳ Limite de l'IA atteinte : nouvelle tentative dans {m.group(1)} s...")

    try:
        stats = translate_pdf(
            in_path=input_path,
            out_path=output_path,
            source=source_name,
            target=target_name,
            domain="documents d'audit interne d'une entreprise de distribution",
            glossary=glossary_dict,
            project=get_project_id(),
            location=DEFAULT_LOCATION,
            model=DEFAULT_MODEL,
            log=pdf_logger,
            cache_path=cache_file,
            codes=(source_code, target_code),
            progress=pdf_progress,
        )

        if progress_callback:
            progress_callback(100, 100, "✅ PDF traduit avec succès !")

        exec_time = round(time.time() - start_time, 2)
        print(f"✅ PDF traduit avec succès ({exec_time}s)")
        in_tok, out_tok = tradoc_llm.TOKENS["input"], tradoc_llm.TOKENS["output"]
        return {
            "status": "success",
            "output_path": output_path,
            "time_seconds": exec_time,
            "input_tokens": in_tok,
            "output_tokens": out_tok,
            "total_cost_usd": in_tok / 1e6 * PRICE_IN + out_tok / 1e6 * PRICE_OUT,
            "stats": stats
        }

    except Exception as e:
        print(f"❌ Erreur critique : {e}")
        return {
            "status": "error",
            "output_path": None,
            "error_message": str(e)
        }