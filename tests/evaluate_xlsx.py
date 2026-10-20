# -*- coding: utf-8 -*-
"""
Comparaison d'un classeur traduit avec l'original.

compare() : structure seule (parties, images, graphiques, formules, styles
de cellules, styles du texte riche). Utilise par le test aller-retour.
eval_xlsx() : compare() plus le texte reste en langue source.
"""
import collections
import re
import zipfile

S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _strings(z):
    """Textes des cellules, commentaires et dessins, avec les styles riches."""
    from lxml import etree
    texts, rich = [], collections.Counter()
    for n in z.namelist():
        if not n.endswith(".xml"):
            continue
        x = z.read(n)
        if b"<si>" in x or b"inlineStr" in x or b"<comment " in x:
            root = etree.fromstring(x)
            for el in root.iter(S + "si", S + "is", S + "text"):
                texts.append("".join(t.text or "" for t in el.iter(S + "t")))
                for r in el.findall(S + "r"):
                    rpr = r.find(S + "rPr")
                    if rpr is not None:
                        rich[etree.tostring(rpr)] += 1
    return texts, rich


def compare(a, b):
    za, zb = zipfile.ZipFile(a), zipfile.ZipFile(b)
    out = {}
    na, nb = set(za.namelist()), set(zb.namelist())
    if na != nb:
        out["parties"] = {"perdues": sorted(na - nb), "ajoutees": sorted(nb - na)}
    for n in sorted(na & nb):
        if "/media/" in n and za.read(n) != zb.read(n):
            out.setdefault("images modifiees", []).append(n)
    for n in sorted(na & nb):
        if re.match(r"xl/worksheets/sheet\d+\.xml$", n):
            xa, xb = za.read(n).decode("utf8"), zb.read(n).decode("utf8")
            fa, fb = re.findall(r"<f[^>]*>[^<]*</f>", xa), re.findall(r"<f[^>]*>[^<]*</f>", xb)
            if fa != fb:
                out.setdefault("formules", []).append(n)
            sa, sb = re.findall(r'<c r="([A-Z]+\d+)"[^>]*?(?: s="(\d+)")?', xa), \
                re.findall(r'<c r="([A-Z]+\d+)"[^>]*?(?: s="(\d+)")?', xb)
            if sa != sb:
                out.setdefault("styles de cellule", []).append(n)
        if n.startswith("xl/charts/") or n == "xl/styles.xml" or n == "xl/workbook.xml":
            if n == "xl/charts/" or not n.endswith(".xml"):
                continue
            if n != "xl/workbook.xml" and n.startswith("xl/charts/"):
                ta = re.sub(r"<c:v>[^<]*</c:v>|<a:t>[^<]*</a:t>", "", za.read(n).decode("utf8"))
                tb = re.sub(r"<c:v>[^<]*</c:v>|<a:t>[^<]*</a:t>", "", zb.read(n).decode("utf8"))
                from lxml import etree

                def canon(x):
                    return etree.tostring(etree.fromstring(x.encode("utf8")), method="c14n")
                if canon(re.sub(r"^<\?xml[^>]*\?>\s*", "", ta)) != canon(re.sub(r"^<\?xml[^>]*\?>\s*", "", tb)):
                    out.setdefault("graphiques (hors texte)", []).append(n)
            elif za.read(n) != zb.read(n):
                out.setdefault("modifie", []).append(n)
    _, ra = _strings(za)
    _, rb = _strings(zb)
    if ra != rb:
        lost = ra - rb
        if lost:
            out["styles riches perdus"] = len(lost)
    return out


def eval_xlsx(orig, trans, direction, **_):
    import sys
    import os
    sys.path.insert(0, os.path.dirname(__file__))
    from evaluate import residue, SRC_LANG, TGT_LANG
    res = {"structure": compare(orig, trans), "residue": []}
    texts, _ = _strings(zipfile.ZipFile(trans))
    for t in texts:
        if residue(t, SRC_LANG[direction], TGT_LANG[direction]):
            res["residue"].append(t[:100])
    return res
