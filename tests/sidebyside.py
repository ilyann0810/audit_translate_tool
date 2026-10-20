# -*- coding: utf-8 -*-
"""
Images cote a cote original | traduction, page par page.

    python tests/sidebyside.py <etiquette> <motif de fichier> [pages...]

PPTX et DOCX sont exportes en PDF par PowerPoint / Word (le rendu que
verront les utilisateurs), puis chaque page est rendue en PNG.
Sortie : Outputs/<etiquette>/_png/<fichier>_p<N>.png
"""
import hashlib
import json
import os
import sys

import fitz

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)


def to_pdf(path, cache_dir):
    if path.lower().endswith(".pdf"):
        return path
    h = hashlib.sha1(open(path, "rb").read()).hexdigest()[:16]
    pdf = os.path.join(cache_dir, "render_%s.pdf" % h)
    if not os.path.isfile(pdf):
        from office_probe import probe_pptx, probe_docx
        fn = probe_pptx if path.lower().endswith(".pptx") else probe_docx
        fn(path, pdf)
    return pdf


def main():
    label, pattern = sys.argv[1], sys.argv[2].lower()
    pages = [int(x) for x in sys.argv[3:]]
    man = json.load(open(os.path.join(HERE, "manifest.json"), encoding="utf-8"))
    in_dir = os.path.normpath(os.path.join(HERE, man["input_dir"]))
    out_dir = os.path.join(ROOT, "Outputs", label)
    cache = os.path.join(ROOT, "Outputs", "_orig")
    os.makedirs(cache, exist_ok=True)
    png_dir = os.path.join(out_dir, "_png")
    os.makedirs(png_dir, exist_ok=True)
    for d in man["documents"]:
        if pattern not in d["file"].lower():
            continue
        a = fitz.open(to_pdf(os.path.join(in_dir, d["file"]), cache))
        b = fitz.open(to_pdf(os.path.join(out_dir, d["file"]), cache))
        for pno in (pages or range(1, a.page_count + 1)):
            if pno > a.page_count or pno > b.page_count:
                continue
            pa, pb = a[pno - 1], b[pno - 1]
            r = pa.rect
            zoom = 1100.0 / max(r.width, r.height)
            out = fitz.open()
            page = out.new_page(width=r.width * 2 + 10, height=r.height)
            page.show_pdf_page(fitz.Rect(0, 0, r.width, r.height), a, pno - 1)
            page.show_pdf_page(fitz.Rect(r.width + 10, 0, 2 * r.width + 10, r.height), b, pno - 1)
            page.draw_line((r.width + 5, 0), (r.width + 5, r.height), color=(1, 0, 0), width=2)
            name = "%s_p%02d.png" % (d["file"][:20].replace(" ", "_"), pno)
            page.get_pixmap(matrix=fitz.Matrix(zoom, zoom)).save(os.path.join(png_dir, name))
            print(os.path.join(png_dir, name))


if __name__ == "__main__":
    main()
