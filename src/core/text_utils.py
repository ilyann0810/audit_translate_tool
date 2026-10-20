import os
import re
import pandas as pd

def load_glossary_df(glossary_path: str):
    """Charge le glossaire depuis un fichier Excel ou CSV et le nettoie."""
    if not glossary_path or not os.path.exists(glossary_path):
        print(f"⚠️ Glossaire non trouvé ou non spécifié : {glossary_path}")
        return None
    
    try:
        if glossary_path.endswith('.xlsx'):
            df = pd.read_excel(glossary_path)
        else:
            df = pd.read_csv(glossary_path)
        
        # Nettoyage des colonnes
        df.columns = [c.strip() for c in df.columns]
        
        if 'Allemand' not in df.columns or 'Francais' not in df.columns:
            print("⚠️ Le glossaire doit contenir les colonnes 'Allemand' et 'Francais'.")
            return None
            
        df = df.dropna(subset=['Allemand', 'Francais'])
        df['Allemand'] = df['Allemand'].astype(str).str.strip()
        df['Francais'] = df['Francais'].astype(str).str.strip()
        
        print(f"📖 Glossaire de {len(df)} lignes chargé.")
        return df
        
    except Exception as e:
        print(f"❌ Erreur lors du chargement du glossaire : {e}")
        return None

def clean_term_for_search(term: str, lang: str) -> str:
    """Supprime les articles (allemands ou français) au début des mots pour la recherche regex."""
    if lang == "de":
        return re.sub(r'^(die|der|das|den|dem|des)\s+', '', term, flags=re.IGNORECASE).strip()
    elif lang == "fr":
        return re.sub(r'^(le|la|les|l\'|un|une|des)\s+', '', term, flags=re.IGNORECASE).strip()
    return term

def get_matching_glossary_rules(batch_text: str, glossary_df, direction="Allemand ➔ Français") -> str:
    """Cherche les termes du glossaire présents dans le texte et génère la règle de prompt."""
    if glossary_df is None or glossary_df.empty:
        return ""
    
    matched_rules = []
    
    # Définition de la colonne source et cible selon la direction
    if direction == "Français ➔ Allemand":
        col_source, col_target, lang_code = 'Francais', 'Allemand', 'fr'
    else:
        col_source, col_target, lang_code = 'Allemand', 'Francais', 'de'
        
    # Tri temporaire par longueur du terme source (évite les remplacements partiels)
    temp_df = glossary_df.copy()
    temp_df['len'] = temp_df[col_source].str.len()
    temp_df = temp_df.sort_values(by='len', ascending=False)
    
    for _, row in temp_df.iterrows():
        source_term = row[col_source]
        target_term = row[col_target]
        
        search_term = clean_term_for_search(source_term, lang_code)
        pattern = r'\b' + re.escape(search_term) + r'\b'
        
        if re.search(pattern, batch_text, flags=re.IGNORECASE):
            matched_rules.append(f'- "{source_term}" -> "{target_term}"')
            
    if matched_rules:
        return "\nGLOSSAIRE MÉTIER OBLIGATOIRE :\n" + "\n".join(matched_rules) + "\n"
    
    return ""