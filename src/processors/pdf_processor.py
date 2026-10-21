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

    # État partagé pour échelonner la progression de 0 à 100 %
    progress_state = {"val": 5}
    if progress_callback:
        progress_callback(5, 100, "🚀 Initialisation et analyse du PDF...")

    def smart_pdf_logger(msg):
        """Intercepte les logs de tradoc pour actualiser la barre de progression Streamlit."""
        print(msg)
        if not progress_callback:
            return

        txt = str(msg).strip()
        if not txt or txt == "-":
            return

        txt_lower = txt.lower()

        # 1. Progression par page
        m_page = re.search(r'page\s*(\d+)(?:\s*/\s*(\d+))?', txt, re.IGNORECASE)
        if m_page:
            cur = int(m_page.group(1))
            tot = int(m_page.group(2)) if m_page.group(2) else 50
            progress_state["val"] = min(70, 10 + int((cur / max(1, tot)) * 60))
            progress_callback(progress_state["val"], 100, f"📄 Traduction du PDF (page {cur}/{tot})...")
            return

        # 2. Réparation et injection des polices
        if "police" in txt_lower and "reparee" in txt_lower:
            progress_state["val"] = min(78, max(progress_state["val"], 72))
            progress_callback(progress_state["val"], 100, "🔧 Injection des caractères et polices vectorielles...")
            return

        # 3. Raccourcissement des blocs trop longs
        if "trop serres" in txt_lower or "versions plus courtes" in txt_lower:
            progress_state["val"] = 82
            progress_callback(progress_state["val"], 100, f"✂️ {txt}")
            return

        if "blocs raccourcis" in txt_lower:
            progress_state["val"] = 92
            progress_callback(progress_state["val"], 100, f"✅ {txt}")
            return

        # 4. Finalisation
        if "traduit avec succès" in txt_lower:
            progress_state["val"] = 98
            progress_callback(98, 100, "💾 Assemblage final du document PDF...")
            return

        # 5. Autres statuts textuels courts
        if len(txt) < 70 and not txt.startswith("http"):
            progress_state["val"] = min(95, progress_state["val"] + 1)
            progress_callback(progress_state["val"], 100, txt)

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
            log=smart_pdf_logger,  # Utilisation du logger connecté à Streamlit
            cache_path=cache_file,
            codes=(source_code, target_code)
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