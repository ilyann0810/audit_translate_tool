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
        df = df[(df['Allemand'] != "") & (df['Francais'] != "")]

        print(f"📖 Glossaire de {len(df)} lignes chargé.")
        return df

    except Exception as e:
        print(f"❌ Erreur lors du chargement du glossaire : {e}")
        return None

def clean_term_for_search(term: str, lang: str) -> str:
    """Supprime les articles (allemands ou français) au début des mots pour la recherche regex."""
    if lang == "de":
        return re.sub(r'^(die|der|das|den|dem|des|ein|eine)\s+', '', term, flags=re.IGNORECASE).strip()
    elif lang == "fr":
        return re.sub(r"^(le|la|les|un|une|des|du)\s+|^l['’]\s*", '', term, flags=re.IGNORECASE).strip()
    return term


# Terminaisons tolerees en fin de terme (pluriel, declinaison) :
# « supermarché » trouve « supermarchés », « Filiale » trouve « Filialen ».
_ENDINGS = {"fr": r"(?:s|x|es)?", "de": r"(?:e|n|en|er|ern|s|es)?"}


def _is_acronym(term: str) -> bool:
    """Sigle : « SPM », « GF », « DR », « FK25 ». Recherche sensible à la casse,
    sans terminaison (« DR » ne doit pas trouver « Dr. » ni « drei »)."""
    return bool(re.fullmatch(r"[A-ZÄÖÜ0-9][A-ZÄÖÜ0-9&./-]{0,7}", term)) and any(c.isalpha() for c in term)


def _term_pattern(term: str, lang: str, compounds: bool = False) -> str:
    """Expression régulière d'un terme du glossaire (article retiré)."""
    if _is_acronym(term):
        return r"(?<![\w-])" + re.escape(term) + r"s?(?![\w-])"
    words = term.split()
    body = r"\s+".join(re.escape(w) + _ENDINGS.get(lang, "") for w in words)
    if compounds and lang == "de" and len(term) >= 5 and len(words) == 1:
        # Mot composé allemand : « Abschriften » dans « Abschriftengründe »,
        # « Revision » dans « Revisionsauftrag ».
        return r"(?i)\w*" + re.escape(term) + r"\w*"
    return r"(?i)(?<!\w)" + body + r"(?!\w)"


def glossary_entries(glossary_df, direction="Allemand ➔ Français"):
    """Entrées du glossaire dans le sens de la traduction, regroupées par
    terme source.

    Un même terme source peut avoir plusieurs équivalents (« Filiale » ->
    « SPM » et « die Filiale » -> « le supermarché ») : ils sont tous admis.
    """
    if glossary_df is None or glossary_df.empty:
        return []
    if direction == "Français ➔ Allemand":
        col_s, col_t, ls, lt = 'Francais', 'Allemand', 'fr', 'de'
    else:
        col_s, col_t, ls, lt = 'Allemand', 'Francais', 'de', 'fr'
    groups = {}
    for _, row in glossary_df.iterrows():
        src, tgt = str(row[col_s]).strip(), str(row[col_t]).strip()
        key = clean_term_for_search(src, ls)
        if not key:
            continue
        g = groups.setdefault(key if _is_acronym(key) else key.lower(),
                              {"shown": src, "key": key, "targets": []})
        if tgt not in g["targets"]:
            g["targets"].append(tgt)
    out = []
    for g in groups.values():
        key = g["key"]
        out.append({
            "shown": g["shown"],
            "acronym": _is_acronym(key),
            "pattern": _term_pattern(key, ls),
            "loose": _term_pattern(key, ls, compounds=True),
            "targets": g["targets"],
            "target_patterns": [_term_pattern(clean_term_for_search(t, lt) or t, lt)
                                for t in g["targets"]],
        })
    # Termes longs d'abord : « die Filiale » avant « Filiale »
    out.sort(key=lambda e: -len(e["shown"]))
    return out


def matching_entries(text: str, entries, compounds: bool = True):
    """Entrées dont le terme source apparaît dans le texte."""
    return [e for e in entries if re.search(e["loose"] if compounds else e["pattern"], text)]


def format_glossary_rules(entries) -> str:
    if not entries:
        return ""
    lines = ['- "%s" -> %s' % (e["shown"], " ou ".join('"%s"' % t for t in e["targets"]))
             for e in entries]
    return ("\nGLOSSAIRE MÉTIER OBLIGATOIRE (aussi au pluriel, dans les mots composés, "
            "devant un numéro et dans les libellés courts ; s'il y a plusieurs équivalents, "
            "choisis celui qui convient au contexte) :\n" + "\n".join(lines) + "\n")


def get_matching_glossary_rules(batch_text: str, glossary_df, direction="Allemand ➔ Français") -> str:
    """Cherche les termes du glossaire présents dans le texte et génère la règle de prompt."""
    entries = glossary_entries(glossary_df, direction)
    return format_glossary_rules(matching_entries(batch_text, entries))


def glossary_violations(src: str, dst: str, entries):
    """Termes du glossaire présents dans le texte source (mot entier) mal
    rendus dans la traduction : aucun équivalent imposé n'y figure, ou le
    sigle source y est resté (« SPM 4443 » laissé tel quel vers l'allemand
    au lieu de « Filiale 4443 »)."""
    out = []
    for e in matching_entries(src, entries, compounds=False):
        found = any(re.search(p, dst) for p in e["target_patterns"])
        left = (e["acronym"] and re.search(e["pattern"], dst)
                and not any(re.search(e["pattern"], t) for t in e["targets"]))
        if not found or left:
            out.append(e)
    return out
