"""Controle de mise en page d'un PDF traduit, page par page.

    python tests/pdf_layout.py original.pdf traduit.pdf [--json sortie.json]

Signale, pour chaque page :
- chevauchements : deux lignes de texte qui se recouvrent, en plus de
  celles qui se recouvraient deja dans l'original (titre sur un grand
  chiffre decoratif, par exemple) ;
- hors page : texte qui depasse le bord de la page ;
- reductions : texte ecrit nettement plus petit que l'original au meme
  endroit (moins de 75 % de la taille d'origine).
"""
import argparse
import json
import sys

import fitz


def lines(page):
    out = []
    for b in page.get_text("dict")["blocks"]:
        if b.get("type") != 0:
            continue
        for ln in b["lines"]:
            spans = [s for s in ln["spans"] if s["text"].strip()]
            if not spans:
                continue
            r = fitz.Rect(ln["bbox"])
            if r.is_empty or r.width < 0.5:
                continue
            # boite resserree : hauteur des lettres, pas l'interligne
            size = max(s["size"] for s in spans)
            out.append({"rect": r, "size": size,
                        "text": "".join(s["text"] for s in spans).strip(),
                        "dir": tuple(round(x, 2) for x in ln["dir"])})
    return out


def _core(r, size):
    """Partie de la boite occupee par l'encre (sans l'interligne)."""
    pad = max(0.0, (r.height - size) / 2)
    return fitz.Rect(r.x0 + 0.5, r.y0 + pad + 0.15 * size, r.x1 - 0.5, r.y1 - pad - 0.1 * size)


def overlaps(ls):
    found = []
    for i in range(len(ls)):
        a = ls[i]
        ca = _core(a["rect"], a["size"])
        for j in range(i + 1, len(ls)):
            b = ls[j]
            if a["text"] == b["text"]:
                continue
            cb = _core(b["rect"], b["size"])
            inter = ca & cb
            if inter.is_empty:
                continue
            small = min(ca.get_area(), cb.get_area()) or 1
            if inter.get_area() / small > 0.12 and inter.height > 1 and inter.width > 2:
                found.append((a, b))
    return found


def _key(a, b):
    return tuple(sorted((a["text"][:25], b["text"][:25])))


def check(orig_path, tr_path):
    o, t = fitz.open(orig_path), fitz.open(tr_path)
    report = []
    for n in range(min(o.page_count, t.page_count)):
        po, pt = o[n], t[n]
        lo, lt = lines(po), lines(pt)
        base = len(overlaps(lo))
        ov = overlaps(lt)
        rect = pt.rect
        off = [l for l in lt if not rect.contains(l["rect"] + (1, 1, -1, -1))
               and l["dir"] == (1.0, 0.0)]
        small = []
        for l in lt:
            c = fitz.Point((l["rect"].x0 + l["rect"].x1) / 2, (l["rect"].y0 + l["rect"].y1) / 2)
            near = [x for x in lo if x["rect"].contains(c)]
            if near:
                ref = max(x["size"] for x in near)
                if l["size"] < 0.75 * ref:
                    small.append((l, ref))
        extra = len(ov) - base
        if extra > 0 or off or small:
            report.append({
                "page": n + 1,
                "chevauchements": max(0, extra),
                "exemples": ["%s  <->  %s" % (a["text"][:50], b["text"][:50]) for a, b in ov[:4]],
                "hors_page": [l["text"][:60] for l in off[:3]],
                "reductions": ["%.1f -> %.1f pt : %s" % (ref, l["size"], l["text"][:50])
                               for l, ref in small[:3]],
                "nb_reductions": len(small),
            })
    return report


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("original")
    ap.add_argument("traduit")
    ap.add_argument("--json")
    a = ap.parse_args()
    rep = check(a.original, a.traduit)
    for r in rep:
        print("p%-3d chevauchements %d, hors page %d, reductions %d"
              % (r["page"], r["chevauchements"], len(r["hors_page"]), r["nb_reductions"]))
        for e in r["exemples"][:2]:
            print("      ~ " + e)
        for e in r["reductions"][:1]:
            print("      - " + e)
    tot = sum(r["chevauchements"] for r in rep)
    print("TOTAL : %d pages signalees, %d chevauchements, %d reductions"
          % (len(rep), tot, sum(r["nb_reductions"] for r in rep)))
    if a.json:
        json.dump(rep, open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
