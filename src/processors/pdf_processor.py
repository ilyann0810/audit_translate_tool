import os
import time
import json
from tradoc.pdf import translate_pdf
from core.text_utils import load_glossary_df 
from config.settings import gemini_model 

def translate_pdf_universal(input_path: str, output_path: str, glossary_path: str = None, progress_callback=None, direction="Allemand ➔ Français"):
    """
    Lanceur Databricks vers le moteur de traduction expert 'tradoc'.
    """
    start_time = time.time()
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    source_code = "de" if "Allemand" in direction else "fr"
    target_code = "fr" if "Français" in direction else "de"

    gcp_key = os.environ.get("GCP_SERVICE_KEY", "{}")
    project_id = json.loads(gcp_key).get("project_id", os.environ.get("GCP_PROJECT_ID"))

    glossary_dict = {}
    if glossary_path and os.path.isfile(glossary_path):
        try:
            df = load_glossary_df(glossary_path)
            if df is not None and not df.empty:
                keys = df.iloc[:, 0].astype(str).tolist()
                vals = df.iloc[:, 1].astype(str).tolist()
                glossary_dict = dict(zip(keys, vals))
        except Exception as e:
            print(f"⚠️ Impossible de charger le glossaire : {e}")

    print(f"🚀 Lancement du moteur tradoc (Gemini 2.5 Flash) : {source_code} ➔ {target_code}")

    try:
        stats = translate_pdf(
            in_path=input_path,
            out_path=output_path,
            source=source_code,
            target=target_code,
            domain="documents d'entreprise et supports de formation",
            glossary=glossary_dict,
            project=project_id,
            location="europe-west1",
            model="gemini-2.5-flash",
            log=print,
            codes=(source_code, target_code)
        )
        
        exec_time = round(time.time() - start_time, 2)
        print(f"✅ PDF traduit avec succès ({exec_time}s)")
        
        # --- CORRECTION DU BUG D'INTERFACE ---
        # Ajout des clés attendues par l'interface pour éviter le crash.
        return {
            "status": "success",
            "output_path": output_path,
            "time_seconds": exec_time,
            "input_tokens": 0,        # <-- Clé ajoutée
            "output_tokens": 0,       # <-- Clé ajoutée
            "total_cost_usd": 0.0,    # <-- Clé ajoutée
            "stats": stats
        }
        
    except Exception as e:
        print(f"❌ Erreur critique : {e}")
        return {
            "status": "error",
            "output_path": None,
            "error_message": str(e)
        }