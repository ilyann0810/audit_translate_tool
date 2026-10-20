"""Mini glossaire au format du glossaire métier (colonnes Allemand / Francais).

Reproduit la forme du vrai fichier : sigles (SPM, GF), termes avec article
(« die Filiale », « le supermarché »), plusieurs entrées vers le même mot.

    python tests/make_glossary.py      -> tests/samples/glossaire_mini.xlsx
"""
import os

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))

ROWS = [
    ("Filiale", "SPM"),
    ("die Filiale", "le supermarché"),
    ("GF", "DR"),
    ("die Zentrale", "le siège"),
    ("Abschriften", "démarques"),
    ("die Kühlung", "le froid"),
    ("das Warenlager", "l'entrepôt"),
    ("Revision", "Audit interne"),
]

if __name__ == "__main__":
    out = os.path.join(HERE, "samples", "glossaire_mini.xlsx")
    pd.DataFrame(ROWS, columns=["Allemand", "Francais"]).to_excel(out, index=False)
    print(out)
