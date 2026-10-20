# -*- coding: utf-8 -*-
"""
Compare chaque document traduit a son original et compte les defauts.

    python tests/evaluate.py <etiquette> [--only motif] [--no-office] [--png]

Pour chaque document du manifeste present dans Outputs/<etiquette>/ :

PPTX
  - debordement mesure par PowerPoint : texte qui sort de sa forme alors
    qu'il tenait dans l'original, ou qui en sort davantage ;
  - tableaux qui grandissent (PowerPoint agrandit les lignes trop pleines)
    et tableaux qui sortent de la diapositive ;
  - styles : couleurs, gras, polices, tailles presentes dans l'original
    et disparues (ou apparues) dans la traduction, par paragraphe ;
  - mise en forme du cadre : ancrage vertical, marges, ajustement auto ;
  - texte encore en langue source ;
  - elements perdus : images, champs, liens.

DOCX
  - nombre de pages, images, champs, liens, zones de texte qui debordent
    (mesures par Word) ;
  - styles par paragraphe, texte en langue source, elements XML perdus.

PDF
  - rapport tradoc (debordements, reductions) et texte en langue source.

Le resultat va dans Outputs/<etiquette>/_eval.json et un resume s'affiche.
Les mesures de l'original sont mises en cache dans Outputs/_orig/.
"""
import argparse
import collections
import hashlib
import json
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))
sys.path.insert(0, HERE)

from core.lang_detect import scores  # noqa: E402

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

SRC_LANG = {"Allemand ➔ Français": "de", "Français ➔ Allemand": "fr"}
TGT_LANG = {"Allemand ➔ Français": "fr", "Français ➔ Allemand": "de"}

TOL = 2.0   # points de tolerance sur les mesures PowerPoint


# --------------------------------------------------------------------------
# Langue
# --------------------------------------------------------------------------

def residue(text, src, tgt):
    """Le texte est-il (encore) nettement en langue source ?"""
    s = scores(text)
    return s[src] >= 2.0 and s[src] > 1.5 * s[tgt]


# --------------------------------------------------------------------------
# PPTX : styles lus dans le XML
# --------------------------------------------------------------------------

def _rpr_sig(rpr):
    """Signature visuelle d'un run : couleur, gras, italique, souligne,
    taille, police, surlignage. None = herite."""
    if rpr is None:
        return {}
    sig = {}
    for k in ("b", "i", "u", "sz", "strike", "baseline", "cap"):
        if rpr.get(k) is not None:
            sig[k] = rpr.get(k)
    fill = rpr.find(A + "solidFill")
    if fill is not None and len(fill):
        c = fill[0]
        val = c.get("val") or ""
        mods = ",".join("%s=%s" % (m.tag.split("}")[1], m.get("val")) for m in c)
        sig["color"] = "%s:%s%s" % (c.tag.split("}")[1], val, ("(" + mods + ")") if mods else "")
    elif rpr.find(A + "gradFill") is not None:
        sig["color"] = "grad"
    hl = rpr.find(A + "highlight")
    if hl is not None and len(hl):
        sig["hl"] = hl[0].get("val")
    lat = rpr.find(A + "latin")
    if lat is not None:
        sig["font"] = lat.get("typeface")
    return sig


def _para_styles(p):
    """Repartition des caracteres par valeur de chaque propriete."""
    dist = collections.defaultdict(collections.Counter)
    total = 0
    for r in p.iter(A + "r"):
        t = r.find(A + "t")
        if t is None or not (t.text or "").strip():
            continue
        n = len(t.text.strip())
        total += n
        sig = _rpr_sig(r.find(A + "rPr"))
        for k in ("color", "b", "i", "u", "sz", "font", "hl", "strike", "baseline"):
            dist[k][sig.get(k, "-")] += n
    return dist, total


def _body_sig(txBody):
    bp = txBody.find(A + "bodyPr")
    if bp is None:
        return {}
    sig = {k: bp.get(k) for k in ("anchor", "lIns", "tIns", "rIns", "bIns", "wrap", "vert", "rot")
           if bp.get(k) is not None}
    for tag in ("normAutofit", "spAutoFit", "noAutofit"):
        el = bp.find(A + tag)
        if el is not None:
            sig["autofit"] = tag + ("(%s)" % el.get("fontScale") if el.get("fontScale") else "")
    return sig


def _ppr_sig(p):
    ppr = p.find(A + "pPr")
    if ppr is None:
        return {}
    sig = {k: ppr.get(k) for k in ("algn", "lvl", "marL", "indent") if ppr.get(k) is not None}
    for tag in ("lnSpc", "spcBef", "spcAft"):
        el = ppr.find(A + tag)
        if el is not None and len(el):
            sig[tag] = el[0].get("val")
    return sig


def _iter_txbodies(slide_xml_root):
    """(cle, txBody) pour chaque forme / cellule, cle stable entre versions."""
    for sp in slide_xml_root.iter():
        tag = sp.tag.split("}")[-1]
        if tag in ("sp",):
            nv = sp.find(".//" + P + "cNvPr")
            tb = sp.find(P + "txBody")
            if nv is not None and tb is not None:
                yield "sp%s" % nv.get("id"), tb
        elif tag == "graphicFrame":
            nv = sp.find(".//" + P + "cNvPr")
            tbl = sp.find(".//" + A + "tbl")
            if nv is None or tbl is None:
                continue
            for ri, tr in enumerate(tbl.findall(A + "tr")):
                for ci, tc in enumerate(tr.findall(A + "tc")):
                    tb = tc.find(A + "txBody")
                    if tb is not None:
                        yield "tb%s_r%dc%d" % (nv.get("id"), ri, ci), tb


def pptx_xml_index(path):
    from lxml import etree
    z = zipfile.ZipFile(path)
    out = {}
    names = sorted(n for n in z.namelist()
                   if re.match(r"ppt/(slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster|notesSlides/notesSlide)\d+\.xml$", n))
    for n in names:
        root = etree.fromstring(z.read(n))
        for key, tb in _iter_txbodies(root):
            paras = tb.findall(A + "p")
            out["%s|%s" % (n, key)] = {
                "body": _body_sig(tb),
                "paras": [{
                    "text": "".join(t.text or "" for t in p.iter(A + "t")),
                    "styles": _para_styles(p)[0],
                    "chars": _para_styles(p)[1],
                    "ppr": _ppr_sig(p),
                } for p in paras],
            }
    counts = collections.Counter()
    for n in z.namelist():
        if n.startswith("ppt/slides/slide") and n.endswith(".xml"):
            x = z.read(n).decode("utf8", "ignore")
            counts["pictures"] += x.count("<p:pic>") + x.count("<p:pic ")
            counts["fields"] += x.count("<a:fld ")
            counts["hlinks"] += x.count("hlinkClick")
            counts["shapes"] += x.count("<p:sp>") + x.count("<p:sp ")
            counts["paragraphs"] += x.count("<a:p>") + x.count("<a:p ")
    return out, counts


def _lost_values(o, t, min_share=0.15):
    """Valeurs de style qui couvraient une part notable du paragraphe
    d'origine et ont disparu (ou sont apparues) dans la traduction."""
    issues = []
    for k in set(o) | set(t):
        oc, tc = o.get(k, {}), t.get(k, {})
        ot, tt = sum(oc.values()) or 1, sum(tc.values()) or 1
        for v in set(oc) | set(tc):
            a, b = oc.get(v, 0) / ot, tc.get(v, 0) / tt
            if a >= min_share and b == 0:
                issues.append("%s=%s perdu" % (k, v))
            elif b >= min_share and a == 0:
                issues.append("%s=%s ajoute" % (k, v))
    return sorted(issues)


def eval_pptx(orig, trans, direction, office=True, cache_dir=None):
    src, tgt = SRC_LANG[direction], TGT_LANG[direction]
    res = {"overflow": [], "table_growth": [], "off_slide": [], "style": [],
           "frame": [], "para_format": [], "residue": [], "lost": []}

    oi, oc = pptx_xml_index(orig)
    ti, tc = pptx_xml_index(trans)
    for k in ("pictures", "fields", "hlinks", "shapes"):
        if tc[k] < oc[k]:
            res["lost"].append("%s : %d -> %d" % (k, oc[k], tc[k]))
    res["paragraphs"] = [oc["paragraphs"], tc["paragraphs"]]

    for key, o in oi.items():
        t = ti.get(key)
        if t is None:
            res["lost"].append("forme absente : " + key)
            continue
        if o["body"] != t["body"]:
            res["frame"].append({"key": key, "orig": o["body"], "trans": t["body"]})
        # Paragraphes : comparaison par position quand le nombre est inchange ;
        # sinon (phrases coupees par Entree, refondues en un paragraphe), on
        # compare les styles du cadre entier.
        if len(o["paras"]) == len(t["paras"]):
            pairs = zip(o["paras"], t["paras"])
        else:
            res.setdefault("merged", []).append({"key": key, "paras": [len(o["paras"]), len(t["paras"])]})

            def whole(paras):
                d = collections.defaultdict(collections.Counter)
                for pp in paras:
                    for k2, c in pp["styles"].items():
                        d[k2].update(c)
                return d
            lost = [x for x in _lost_values(whole(o["paras"]), whole(t["paras"]), 0.05)
                    if not x.startswith("sz=")]
            if lost:
                res["style"].append({"key": key, "text": o["paras"][0]["text"][:60], "issues": lost})
            pairs = []
        for op, tp in pairs:
            if not op["text"].strip():
                continue
            lost = _lost_values(op["styles"], tp["styles"])
            # La taille reduite pour faire tenir n'est pas une perte de style.
            lost = [x for x in lost if not x.startswith("sz=")]
            if lost:
                res["style"].append({"key": key, "text": op["text"][:60], "issues": lost})
            if op["ppr"] != tp["ppr"]:
                res["para_format"].append({"key": key, "text": op["text"][:50],
                                           "orig": op["ppr"], "trans": tp["ppr"]})
        for tp in t["paras"]:
            if "slides/slide" in key and residue(tp["text"], src, tgt):
                res["residue"].append(tp["text"][:100])

    if office:
        from office_probe import probe_pptx
        po = _probe_cached(orig, probe_pptx, cache_dir)
        pt = probe_pptx(trans)
        H = pt["slide_h"]
        for s_idx, (so, st) in enumerate(zip(po["slides"], pt["slides"])):
            omap = {it["id"]: it for it in so}
            for it in st:
                o = omap.get(it["id"])
                if o is None:
                    continue
                if "table" in it:
                    gh = sum(it["table"]["rows"]) - sum(o["table"]["rows"])
                    if gh > TOL:
                        # Grandir n'est un defaut que si le tableau recouvre
                        # alors une autre forme qu'il ne recouvrait pas.
                        ob = o["y"] + sum(o["table"]["rows"])
                        nb = it["y"] + sum(it["table"]["rows"])
                        hit = [x.get("text", x["name"])[:30] for x in st
                               if x["id"] != it["id"] and x["y"] >= ob - 1 and x["y"] < nb - 1
                               and min(x["x"] + x["w"], it["x"] + it["w"]) - max(x["x"], it["x"]) > 2]
                        if hit:
                            res["table_growth"].append({"slide": s_idx + 1, "id": it["id"],
                                                        "grow": round(gh, 1), "covers": hit[:3]})
                    if it["y"] + sum(it["table"]["rows"]) > H + TOL and \
                            o["y"] + sum(o["table"]["rows"]) <= H + TOL:
                        res["off_slide"].append({"slide": s_idx + 1, "id": it["id"]})
                    continue
                if it.get("other") or "bh" not in it or "bh" not in o:
                    continue
                ov_t = _overflow(it)
                ov_o = _overflow(o)
                if ov_t > ov_o + TOL:
                    why = _overflow_harm(it, o, st, so, pt)
                    if why:
                        res["overflow"].append({"slide": s_idx + 1, "id": it["id"],
                                                "excess": round(ov_t, 1), "orig_excess": round(ov_o, 1),
                                                "why": why, "text": it.get("text", "")[:60]})
    return res


def _overflow(it):
    """Ce qui depasse de la forme, en points (vertical, et horizontal
    quand le texte ne revient pas a la ligne)."""
    v = max(0.0, (it["by"] + it["bh"]) - (it["y"] + it["h"]),
            it["y"] - it["by"])
    h = 0.0
    if not it.get("wrap", -1) or it.get("wrap") == 0:
        h = max(0.0, (it["bx"] + it["bw"]) - (it["x"] + it["w"]), it["x"] - it["bx"])
    if it.get("rot"):
        return 0.0      # boites tournees : mesure non fiable
    return max(v, h)


def _rect(it):
    return (it["x"], it["y"], it["x"] + it["w"], it["y"] + it["h"])


def _text_rect(it):
    return (it["bx"], it["by"], it["bx"] + it["bw"], it["by"] + it["bh"])


def _inter(a, b):
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return w > TOL and h > TOL


def _overflow_harm(it, o, slide_t, slide_o, probe):
    """Le texte qui sort de sa forme gene-t-il ? Oui s'il sort d'une forme
    visible, du cadre qui la contient ou de la diapositive, ou s'il touche
    un element qu'il ne touchait pas dans l'original."""
    tr, sr = _text_rect(it), _rect(it)
    if it.get("framed"):
        return "sort d'une forme visible"
    if tr[3] > probe["slide_h"] + TOL or tr[1] < -TOL:
        return "sort de la diapositive"
    orig_tr = _text_rect(o)
    for x in slide_t:
        if x["id"] == it["id"]:
            continue
        xr = _rect(x) if "table" not in x else (x["x"], x["y"], x["x"] + x["w"],
                                                x["y"] + sum(x["table"]["rows"]))
        contains = xr[0] <= sr[0] + 1 and xr[2] >= sr[2] - 1 and xr[1] <= sr[1] + 1 and xr[3] >= sr[3] - 1
        if contains:
            if x.get("framed") and (tr[3] > xr[3] + TOL or tr[1] < xr[1] - TOL):
                return "sort du cadre « %s »" % x.get("text", x["name"])[:30]
            continue
        if _inter(tr, xr) and not _inter(orig_tr, xr):
            return "touche « %s »" % x.get("text", x["name"])[:30]
    return None


def _probe_cached(path, fn, cache_dir):
    h = hashlib.sha1(open(path, "rb").read()).hexdigest()[:16]
    p = os.path.join(cache_dir, "probe_%s.json" % h)
    if os.path.isfile(p):
        return json.load(open(p, encoding="utf-8"))
    res = fn(path)
    json.dump(res, open(p, "w", encoding="utf-8"))
    return res


# --------------------------------------------------------------------------
# DOCX
# --------------------------------------------------------------------------

def _docx_rpr_sig(rpr):
    if rpr is None:
        return {}
    sig = {}
    for tag in ("b", "i", "strike", "caps"):
        el = rpr.find(W + tag)
        if el is not None:
            sig[tag] = el.get(W + "val", "1")
    for tag, attr in (("color", "val"), ("sz", "val"), ("highlight", "val"),
                      ("u", "val"), ("vertAlign", "val")):
        el = rpr.find(W + tag)
        if el is not None:
            sig[tag] = el.get(W + attr)
    shd = rpr.find(W + "shd")
    if shd is not None:
        sig["shd"] = shd.get(W + "fill")
    rf = rpr.find(W + "rFonts")
    if rf is not None:
        sig["font"] = rf.get(W + "ascii") or rf.get(W + "hAnsi")
    st = rpr.find(W + "rStyle")
    if st is not None:
        sig["rstyle"] = st.get(W + "val")
    return sig


def docx_xml_index(path):
    from lxml import etree
    z = zipfile.ZipFile(path)
    out = []
    counts = collections.Counter()
    for n in sorted(z.namelist()):
        if not re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$", n):
            continue
        root = etree.fromstring(z.read(n))
        x = z.read(n).decode("utf8", "ignore")
        for tag in ("w:drawing", "w:fldChar", "w:instrText", "w:hyperlink", "w:footnoteReference",
                    "w:tab/", "w:br", "w:txbxContent", "w:bookmarkStart", "w:sdt>", "w:object",
                    "w:pict", "mc:AlternateContent"):
            counts[tag] += x.count("<" + tag)
        for i, p in enumerate(root.iter(W + "p")):
            dist = collections.defaultdict(collections.Counter)
            text = ""
            for r in p.findall(".//" + W + "r"):
                # runs des paragraphes imbriques (zones de texte) : comptes a part
                if r.getparent() is not p and r.getparent().tag != W + "hyperlink":
                    anc = r.getparent()
                    nested = False
                    while anc is not None and anc is not p:
                        if anc.tag == W + "p":
                            nested = True
                            break
                        anc = anc.getparent()
                    if nested:
                        continue
                t = "".join(x.text or "" for x in r.findall(W + "t"))
                text += t
                if not t.strip():
                    continue
                sig = _docx_rpr_sig(r.find(W + "rPr"))
                for k in ("color", "b", "i", "sz", "highlight", "shd", "font", "u", "rstyle", "vertAlign"):
                    dist[k][sig.get(k, "-")] += len(t.strip())
            out.append({"part": n, "idx": i, "text": text, "styles": dist})
    return out, counts


def eval_docx(orig, trans, direction, office=True, cache_dir=None):
    src, tgt = SRC_LANG[direction], TGT_LANG[direction]
    res = {"style": [], "residue": [], "lost": [], "office": {}}
    oi, oc = docx_xml_index(orig)
    ti, tc = docx_xml_index(trans)
    for k in oc:
        if tc[k] != oc[k]:
            res["lost"].append("%s : %d -> %d" % (k, oc[k], tc[k]))
    if len(oi) != len(ti):
        res["lost"].append("paragraphes : %d -> %d" % (len(oi), len(ti)))
    else:
        for o, t in zip(oi, ti):
            if not o["text"].strip():
                continue
            lost = _lost_values(o["styles"], t["styles"])
            if lost:
                res["style"].append({"text": o["text"][:60], "issues": lost})
    for t in ti:
        if residue(t["text"], src, tgt):
            res["residue"].append(t["text"][:100])
    if office:
        from office_probe import probe_docx
        po = _probe_cached(orig, probe_docx, cache_dir)
        pt = probe_docx(trans)
        res["office"] = {k: [po[k], pt[k]] for k in po if po[k] != pt[k]}
    return res


# --------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------

def eval_pdf(orig, trans, direction, **_):
    import fitz
    src, tgt = SRC_LANG[direction], TGT_LANG[direction]
    res = {"residue": [], "pages": []}
    a, b = fitz.open(orig), fitz.open(trans)
    res["pages"] = [a.page_count, b.page_count]
    for pno, page in enumerate(b):
        for blk in page.get_text("blocks"):
            if blk[6] == 0 and residue(blk[4], src, tgt):
                res["residue"].append("p%d: %s" % (pno + 1, re.sub(r"\s+", " ", blk[4])[:90]))
    return res


# --------------------------------------------------------------------------

def summarize(name, kind, r):
    if kind == "pptx":
        keys = ("overflow", "table_growth", "off_slide", "style", "frame", "para_format", "residue", "lost")
    elif kind == "docx":
        keys = ("style", "residue", "lost")
    else:
        keys = ("residue",)
    parts = ["%s=%d" % (k, len(r[k])) for k in keys]
    if r.get("structure"):
        parts.append("structure=%s" % r["structure"])
    if kind == "docx" and r.get("office"):
        parts.append("office=%s" % r["office"])
    return "%-48s %s" % (name[:48], "  ".join(parts))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("label")
    ap.add_argument("--only")
    ap.add_argument("--no-office", action="store_true")
    ap.add_argument("--manifest", default=os.path.join(HERE, "manifest.json"))
    args = ap.parse_args()

    man = json.load(open(args.manifest, encoding="utf-8"))
    in_dir = os.path.normpath(os.path.join(os.path.dirname(args.manifest), man["input_dir"]))
    out_dir = os.path.join(ROOT, "Outputs", args.label)
    cache_dir = os.path.join(ROOT, "Outputs", "_orig")
    os.makedirs(cache_dir, exist_ok=True)

    path_eval = os.path.join(out_dir, "_eval.json")
    allres = json.load(open(path_eval, encoding="utf-8")) if os.path.isfile(path_eval) else {}
    for d in man["documents"]:
        f = d["file"]
        if args.only and args.only.lower() not in f.lower():
            continue
        orig, trans = os.path.join(in_dir, f), os.path.join(out_dir, f)
        if not os.path.isfile(trans):
            continue
        kind = f.rsplit(".", 1)[-1].lower()
        if kind in ("xlsx", "xlsm"):
            from evaluate_xlsx import eval_xlsx
            fn = eval_xlsx
        else:
            fn = {"pptx": eval_pptx, "docx": eval_docx, "pdf": eval_pdf}[kind]
        r = fn(orig, trans, d["direction"], office=not args.no_office, cache_dir=cache_dir)
        allres[f] = r
        print(summarize(f, kind, r), flush=True)
    json.dump(allres, open(path_eval, "w", encoding="utf-8"), ensure_ascii=False, indent=1,
              default=lambda o: dict(o))


if __name__ == "__main__":
    main()
