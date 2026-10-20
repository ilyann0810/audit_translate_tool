import streamlit as st
import os
import tempfile
from core.router import process_translation

# Configuration de la page
st.set_page_config(page_title="Traducteur Documentaire", page_icon="🌍", layout="centered")

st.title("🌍 Traducteur Documentaire")
st.write("Déposez vos documents (Word, PowerPoint, Excel, PDF) pour une traduction respectant strictement la mise en forme et le glossaire métier.")

# --- NOUVEAU : Sélection de la direction ---
direction = st.selectbox(
    "Direction de la traduction :",
    options=["Allemand ➔ Français", "Français ➔ Allemand"],
    index=0
)

# --- MODIFIÉ : Multi-fichiers (accept_multiple_files=True) ---

uploaded_files = st.file_uploader(
    "Document(s) à traduire (.docx, .pptx, .xlsx, .xlsm, .pdf)", 
    type=["docx", "pptx", "xlsx", "xlsm", "pdf"], 
    accept_multiple_files=True
)
glossary_file = st.file_uploader("Glossaire métier (Optionnel, .xlsx, .csv)", type=["xlsx", "csv"])

# Lancement du processus
if uploaded_files and st.button("Lancer la traduction", type="primary"):
    total_files = len(uploaded_files)
    
    with tempfile.TemporaryDirectory() as temp_dir:
        
        # 1. Sauvegarde du glossaire (une seule fois pour tous les fichiers)
        glossary_path = None
        if glossary_file:
            glossary_path = os.path.join(temp_dir, glossary_file.name)
            with open(glossary_path, "wb") as f:
                f.write(glossary_file.getbuffer())

        # 2. Boucle sur chaque document déposé
        for idx, uploaded_file in enumerate(uploaded_files, start=1):
            st.markdown(f"### 📄 Traitement ({idx}/{total_files}) : **{uploaded_file.name}**")
            
            input_path = os.path.join(temp_dir, uploaded_file.name)
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Préfixe du fichier de sortie dynamique selon la direction
            prefix = "FR_" if direction == "Allemand ➔ Français" else "DE_"
            output_name = f"{prefix}{uploaded_file.name}"
            output_path = os.path.join(temp_dir, output_name)

            # Initialisation de la barre de progression
            my_bar = st.progress(0, text="Connexion à l'IA...")
            
            def update_progress(current, total):
                """Fonction injectée dans les processeurs pour mettre à jour l'UI."""
                percent = int((current / total) * 100)
                my_bar.progress(percent, text=f"Traduction par l'IA : lot {current}/{total}...")

            try:
                # --- MODIFIÉ : Appel au routeur avec le paramètre direction ---
                bilan = process_translation(
                    input_path, 
                    output_path, 
                    glossary_path, 
                    progress_callback=update_progress,
                    direction=direction
                )

                # Nettoyage de la barre de progression à la fin
                my_bar.empty()

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
                    key=f"dl_{idx}_{uploaded_file.name}" # Clé unique obligatoire pour Streamlit
                )

            except Exception as e:
                st.error(f"❌ Erreur sur {uploaded_file.name} : {str(e)}")
            
            # Ligne de séparation entre chaque fichier
            st.divider()