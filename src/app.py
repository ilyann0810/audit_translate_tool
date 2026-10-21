import streamlit as st
import os
import tempfile
import threading
import time
import traceback
from core.router import process_translation

# Récupération sécurisée du gestionnaire de contexte Streamlit pour les threads
try:
    from streamlit.runtime.scriptrunner import add_script_run_ctx, get_script_run_ctx
except ImportError:
    try:
        from streamlit.scriptrunner import add_script_run_ctx, get_script_run_ctx
    except ImportError:
        add_script_run_ctx, get_script_run_ctx = None, None

st.set_page_config(page_title="Traducteur Documentaire", page_icon="🌍", layout="centered")

st.title("🌍 Traducteur Documentaire")
st.write("Déposez vos documents (Word, PowerPoint, Excel, PDF) pour une traduction respectant strictement la mise en forme et le glossaire métier.")

direction = st.selectbox(
    "Direction de la traduction :",
    options=["Allemand ➔ Français", "Français ➔ Allemand"],
    index=0
)

uploaded_files = st.file_uploader(
    "Document(s) à traduire (.docx, .pptx, .xlsx, .xlsm, .pdf)", 
    type=["docx", "pptx", "xlsx", "xlsm", "pdf"], 
    accept_multiple_files=True
)
glossary_file = st.file_uploader("Glossaire métier (Optionnel, .xlsx, .csv)", type=["xlsx", "csv"])

if uploaded_files and st.button("Lancer la traduction", type="primary"):
    total_files = len(uploaded_files)
    
    with tempfile.TemporaryDirectory() as temp_dir:
        glossary_path = None
        if glossary_file:
            glossary_path = os.path.join(temp_dir, glossary_file.name)
            with open(glossary_path, "wb") as f:
                f.write(glossary_file.getbuffer())

        main_ctx = get_script_run_ctx() if get_script_run_ctx else None

        for idx, uploaded_file in enumerate(uploaded_files, start=1):
            st.markdown(f"### 📄 Traitement ({idx}/{total_files}) : **{uploaded_file.name}**")
            
            input_path = os.path.join(temp_dir, uploaded_file.name)
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            prefix = "FR_" if direction == "Allemand ➔ Français" else "DE_"
            output_name = f"{prefix}{uploaded_file.name}"
            output_path = os.path.join(temp_dir, output_name)

            my_bar = st.progress(0, text="Préparation du document...")
            bar = {"pct": 0, "label": "Préparation du document...", "t0": time.time()}
            bar_lock = threading.Lock()
            stop_ticker = threading.Event()

            def show_bar():
                """Affiche l'étape en cours et le temps écoulé (appelé sous bar_lock)."""
                s = int(time.time() - bar["t0"])
                elapsed = f"{s // 60} min {s % 60:02d} s" if s >= 60 else f"{s} s"
                my_bar.progress(bar["pct"], text=f"{bar['label']}  ·  ⏳ {elapsed}")

            def update_progress(current, total, status_text=None):
                """Met à jour la barre avec prise en charge explicite du texte d'étape."""
                if total <= 0:
                    return
                if main_ctx is not None and add_script_run_ctx is not None:
                    add_script_run_ctx(threading.current_thread(), main_ctx)
                percent = min(100, max(0, int((current / total) * 100)))
                with bar_lock:
                    # Jamais de retour en arrière : les lots parallèles finissent dans le désordre.
                    bar["pct"] = max(bar["pct"], percent)
                    bar["label"] = status_text or "Traduction en cours..."
                    show_bar()

            def tick():
                """Fait avancer le compteur de temps même pendant un long appel à l'IA,
                pour que l'écran ne paraisse jamais figé."""
                while not stop_ticker.wait(1.0):
                    try:
                        with bar_lock:
                            show_bar()
                    except Exception:
                        return

            ticker = threading.Thread(target=tick, daemon=True)
            if main_ctx is not None and add_script_run_ctx is not None:
                add_script_run_ctx(ticker, main_ctx)
            ticker.start()

            try:
                try:
                    bilan = process_translation(
                        input_path,
                        output_path,
                        glossary_path,
                        progress_callback=update_progress,
                        direction=direction
                    )
                finally:
                    stop_ticker.set()
                    ticker.join()

                my_bar.empty()

                if not bilan or bilan.get("status") != "success" or not bilan.get("output_path"):
                    err_msg = bilan.get("error_message") if isinstance(bilan, dict) else "Retour processeur invalide"
                    raise RuntimeError(err_msg or "Traduction impossible sans détail d'erreur")

                with open(bilan["output_path"], "rb") as f:
                    translated_data = f.read()

                st.success(f"✅ **{uploaded_file.name}** traduit en {bilan['time_seconds']} secondes !")
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Tokens Entrée", f"{bilan['input_tokens']:,}")
                col2.metric("Tokens Sortie", f"{bilan['output_tokens']:,}")
                col3.metric("Coût Estimé", f"${bilan['total_cost_usd']:.4f}")

                st.download_button(
                    label=f"📥 Télécharger {output_name}",
                    data=translated_data,
                    file_name=output_name,
                    mime="application/octet-stream",
                    key=f"dl_{idx}_{uploaded_file.name}"
                )

            except Exception as e:
                my_bar.empty()
                err_text = str(e).strip() or repr(e)
                st.error(f"❌ Erreur sur {uploaded_file.name} : {err_text}")
                with st.expander("Détails techniques de l'erreur"):
                    st.code(traceback.format_exc())
            
            st.divider()