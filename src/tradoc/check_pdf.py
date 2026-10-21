# -*- coding: utf-8 -*-
"""
Phase 0 - Verification d'un PDF traduit.

Le comparateur OOXML ne s'applique pas ici : un PDF n'a ni paragraphes ni
runs, seulement des blocs de texte positionnes. On mesure donc ce qui est
verifiable sans structure :

  - le nombre de pages est-il conserve
  - les dimensions de page sont-elles inchangees
  - le volume de texte par page a-t-il explose ou disparu
  - reste-t-il du texte en langue source
  - les liens et les images ont-ils survecu

Usage :
    python -m tradoc.check_pdf source.pdf traduit.pdf
"""
from __future__ import annotations

import argparse
import os
import re


# Mots-outils tres frequents, propres a chaque langue. Leur presence dans
# le document traduit signale des zones restees en langue source. C'est
# grossier mais suffisant pour reperer une page entiere non traduite.
# Mots-outils propres a chaque langue. Les mots communs a deux langues
# sont exclus : "des" est un article allemand autant que francais, et le
# compter faisait passer "Verringerung des Aussenlaerms" pour du francais.
STOPWORDS = {
    "de": {"und", "der", "die", "das", "ist", "nicht", "eine", "einen",
           "einem", "einer", "werden", "wird", "wurde", "auch", "mit", "fuer",
           "für", "sich", "dass", "oder", "sind", "bei", "von", "zu", "im",
           "auf", "aus", "den", "dem", "zum", "zur", "über", "unter", "nach",
           "wie", "wir", "kann", "können", "hat", "haben", "nur", "noch",
           "mehr", "sehr", "schon", "durch", "gegen", "ohne", "um", "am",
           "als", "ein", "ihre", "ihr", "diese", "dieser", "sowie"},
    "fr": {"et", "le", "la", "les", "est", "pas", "une", "un", "du",
           "pour", "avec", "sur", "que", "qui", "dans", "sont", "par", "plus",
           "au", "aux", "ce", "cette", "ces", "nous", "vous", "il", "elle",
           "leur", "leurs", "être", "été", "mais", "ou", "où", "ne", "se",
           "sa", "son", "ses", "de", "d", "l", "à", "comme", "aussi"},
    "en": {"and", "the", "is", "not", "for", "with", "that", "are", "this",
           "from", "have", "been", "which", "of", "to", "in", "on", "by",
           "be", "or", "as", "it", "its", "their", "will", "can"},
}

# Lettres qui n'existent que dans une langue : un indice fort, meme sur un
# texte court ou sans mot-outil.
LETTERS = {"de": set("äöüßÄÖÜ"), "fr": set("éèêëàâçùûîïôœÉÈÊÀÇ"),
           "en": set()}

WORD_RE = re.compile(r"[A-Za-zÀ-ÿ]+")


def _lang_score(text, lang):
    words = [w.lower() for w in WORD_RE.findall(text)]
    if not words:
        return 0.0
    sw = STOPWORDS.get(lang, set())
    ratio = sum(1 for w in words if w in sw) / len(words)
    special = sum(1 for c in text if c in LETTERS.get(lang, ()))
    return ratio + 0.15 * min(1.0, special / 3.0)


def check(src_path, dst_path, source="de", target="fr"):
    import fitz  # PyMuPDF

    a = fitz.open(src_path)
    b = fitz.open(dst_path)

    res = {
        "src_pages": a.page_count,
        "dst_pages": b.page_count,
        "pages_match": a.page_count == b.page_count,
        "size_changes": [],
        "per_page": [],
        "src_chars": 0,
        "dst_chars": 0,
        "src_links": 0,
        "dst_links": 0,
        "src_images": 0,
        "dst_images": 0,
        "_src_xrefs": set(),
        "_dst_xrefs": set(),
        "src_vectors": 0,
        "dst_vectors": 0,
        "rasterised_pages": [],
        "suspect_pages": [],
    }

    n = min(a.page_count, b.page_count)
    for i in range(n):
        pa, pb = a[i], b[i]
        ta, tb = pa.get_text(), pb.get_text()
        res["src_chars"] += len(ta)
        res["dst_chars"] += len(tb)
        res["src_links"] += len(pa.get_links())
        res["dst_links"] += len(pb.get_links())

        # On compte les images UNIQUES, pas les references : un logo
        # repete sur seize pages compte pour une image, sinon la mesure
        # dit n'importe quoi.
        res["_src_xrefs"].update(im[0] for im in pa.get_images(full=True))
        res["_dst_xrefs"].update(im[0] for im in pb.get_images(full=True))

        # Rasterisation : la page perd ses dessins vectoriels et recoit une
        # seule image aux dimensions de la page. translateDocument aplatit
        # le fond avant de reposer le texte par-dessus.
        va, vb = len(pa.get_drawings()), len(pb.get_drawings())
        res["src_vectors"] += va
        res["dst_vectors"] += vb
        big = [im for im in pb.get_images(full=True)
               if im[2] >= pb.rect.width * 1.5 and im[3] >= pb.rect.height * 1.5]
        if va > 0 and vb == 0 and big:
            res["rasterised_pages"].append(i + 1)

        if (round(pa.rect.width), round(pa.rect.height)) != \
           (round(pb.rect.width), round(pb.rect.height)):
            res["size_changes"].append({
                "page": i + 1,
                "src": (round(pa.rect.width), round(pa.rect.height)),
                "dst": (round(pb.rect.width), round(pb.rect.height)),
            })

        ratio = (len(tb) / len(ta)) if ta else (1.0 if not tb else 99.0)
        residual = _lang_score(tb, source)
        arrived = _lang_score(tb, target)

        res["per_page"].append({
            "page": i + 1, "src_chars": len(ta), "dst_chars": len(tb),
            "ratio": ratio, "residual_source": residual,
            "target_score": arrived,
        })

        # Une page qui garde beaucoup de mots-outils de la langue source
        # et peu de la langue cible n'a pas ete traduite.
        if len(ta) > 80 and residual > arrived and residual > 0.03:
            res["suspect_pages"].append(i + 1)
        if len(ta) > 80 and len(tb) < len(ta) * 0.5:
            res["suspect_pages"].append(i + 1)

    res["src_images"] = len(res.pop("_src_xrefs"))
    res["dst_images"] = len(res.pop("_dst_xrefs"))
    res["suspect_pages"] = sorted(set(res["suspect_pages"]))
    a.close()
    b.close()
    return res


def render(res, name):
    L = []
    add = L.append
    add("## %s" % name)
    add("")
    ratio = (res["dst_chars"] / res["src_chars"]) if res["src_chars"] else 1.0

    add("| | source | traduit |")
    add("|---|---|---|")
    add("| pages | %d | %d |" % (res["src_pages"], res["dst_pages"]))
    add("| caracteres | %d | %d (x%.2f) |"
        % (res["src_chars"], res["dst_chars"], ratio))
    add("| liens | %d | %d |" % (res["src_links"], res["dst_links"]))
    add("| images uniques | %d | %d |" % (res["src_images"],
                                         res["dst_images"]))
    add("| dessins vectoriels | %d | %d |" % (res["src_vectors"],
                                              res["dst_vectors"]))
    add("")

    if not res["pages_match"]:
        add("**Le nombre de pages a change.** Un PDF traduit qui gagne ou "
            "perd des pages a debordé : le texte plus long a pousse le "
            "contenu.")
        add("")
    if res["size_changes"]:
        add("**%d pages ont change de dimensions.**" % len(res["size_changes"]))
        add("")
    if res["suspect_pages"]:
        add("**Pages suspectes** (texte reste en langue source, ou volume "
            "de texte effondre) : %s"
            % ", ".join(str(p) for p in res["suspect_pages"]))
        add("")
    if (res["pages_match"] and not res["size_changes"]
            and not res["suspect_pages"]):
        add("Pages, dimensions et couverture de traduction conformes.")
        add("")

    if res["dst_links"] < res["src_links"]:
        add("- %d liens perdus" % (res["src_links"] - res["dst_links"]))
    if res["rasterised_pages"]:
        add("**%d pages rasterisees.** Les graphiques vectoriels et les "
            "logos ont ete aplatis en une image par page, puis le texte "
            "traduit repose par-dessus. Le texte reste selectionnable, mais "
            "les visuels ne sont plus nets a l'agrandissement ni a "
            "l'impression haute definition."
            % len(res["rasterised_pages"]))
        add("")
    add("")

    worst = sorted(res["per_page"], key=lambda p: -p["ratio"])[:5]
    add("Pages ou le texte s'allonge le plus :")
    add("")
    add("| page | caracteres | expansion |")
    add("|---|---|---|")
    for p in worst:
        add("| %d | %d -> %d | x%.2f |"
            % (p["page"], p["src_chars"], p["dst_chars"], p["ratio"]))
    add("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("source")
    ap.add_argument("translated")
    ap.add_argument("--source-lang", default="de")
    ap.add_argument("--target-lang", default="fr")
    args = ap.parse_args()

    res = check(args.source, args.translated,
                args.source_lang, args.target_lang)
    print(render(res, os.path.basename(args.source)))


if __name__ == "__main__":
    main()