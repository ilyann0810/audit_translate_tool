"""
Traduction XLSX / XLSM en gardant tout le classeur intact.

openpyxl reecrit le classeur entier et perd au passage ce qu'il ne connait
pas : images, graphiques, formes, certains formats conditionnels. Ici on
ne modifie que les parties XML qui contiennent du texte, et on recopie
tout le reste octet pour octet :

  - textes des cellules (xl/sharedStrings.xml et chaines en ligne), y
    compris le texte riche : chaque morceau garde sa police, sa couleur,
    son gras (segments balises, voir core/segments.py) ;
  - commentaires ;
  - zones de texte et formes (xl/drawings) ;
  - titres et libelles des graphiques (xl/charts).

Ne sont pas traduits : formules, noms de feuilles (les formules y font
reference), et tout texte utilise comme valeur dans une formule
(=NB.SI(A:A;"Ja") cesserait de compter si "Ja" devenait "Oui").
"""
import copy
import os
import re
import shutil
import tempfile
import time
import zipfile

from lxml import etree

from core.segments import Segment
from core.text_utils import load_glossary_df
from core.translation_api import LANG, PRICE_IN, PRICE_OUT, translate_segments

S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
S = "{%s}" % S_NS
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _style_key(rpr):
    return "-" if rpr is None else etree.tostring(rpr, encoding="unicode")


def string_item_segment(si):
    """Segment pour un <si> (chaine partagee), un <is> (chaine en ligne) ou
    le <text> d'un commentaire."""
    seg = Segment()
    seg.meta["si"] = si
    runs = si.findall(S + "r")
    if runs:
        seg.meta["rich"] = True
        for r in runs:
            t = r.find(S + "t")
            rpr = r.find(S + "rPr")
            text = t.text if t is not None and t.text else ""
            for k, part in enumerate(text.split("\n")):
                if k:
                    seg.add_break("br")
                seg.add_text(part, _style_key(rpr), rpr)
    else:
        t = si.find(S + "t")
        text = t.text if t is not None and t.text else ""
        for k, part in enumerate(text.split("\n")):
            if k:
                seg.add_break("br")
            seg.add_text(part, "-", None)
    return seg if seg.has_text() else None


def rebuild_string_item(seg, items):
    si = seg.meta["si"]
    keep = [c for c in si if c.tag not in (S + "t", S + "r")]   # rPh, phoneticPr
    for c in list(si):
        si.remove(c)

    # Retours a la ligne : caractere \n dans le texte, comme Excel.
    pieces = []     # (style_id, texte)
    for it in items:
        if it[0] == "t":
            if pieces and pieces[-1][0] == it[1]:
                pieces[-1] = (it[1], pieces[-1][1] + it[2])
            else:
                pieces.append((it[1], it[2]))
        elif it[0] in ("br", "p"):
            sid = pieces[-1][0] if pieces else seg.base_style()
            pieces.append((sid, "\n"))
        elif it[0] == "tab":
            sid = pieces[-1][0] if pieces else seg.base_style()
            pieces.append((sid, "\t"))

    def t_el(text):
        t = etree.Element(S + "t")
        t.text = text
        if text != text.strip() or "\n" in text:
            t.set(XML_SPACE, "preserve")
        return t

    if not seg.meta.get("rich"):
        si.insert(0, t_el("".join(p[1] for p in pieces)))
    else:
        for k, (sid, text) in enumerate(pieces):
            r = etree.Element(S + "r")
            rpr = seg.styles.get(sid)
            if rpr is not None:
                r.append(copy.deepcopy(rpr))
            r.append(t_el(text))
            si.insert(k, r)
    for c in keep:
        si.append(c)


def part_types(z):
    """Type de contenu de chaque partie, d'apres [Content_Types].xml."""
    root = etree.fromstring(z.read("[Content_Types].xml"))
    ns = "{http://schemas.openxmlformats.org/package/2006/content-types}"
    defaults = {d.get("Extension").lower(): d.get("ContentType") for d in root.findall(ns + "Default")}
    out = {}
    for n in z.namelist():
        out[n] = defaults.get(n.rsplit(".", 1)[-1].lower(), "")
    for o in root.findall(ns + "Override"):
        out[o.get("PartName").lstrip("/")] = o.get("ContentType")
    return out


_FORMULA_STR = re.compile(r'"((?:[^"]|"")*)"')


_REF_RE = re.compile(r"(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!)?"
                     r"\$?([A-Z]{1,3})\$?(\d+)?(?::\$?([A-Z]{1,3})\$?(\d+)?)?")


def _col_num(col):
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def sheet_names_by_part(z):
    """Partie XML de chaque feuille -> nom de la feuille."""
    wb = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid_target = {r.get("Id"): r.get("Target") for r in rels}
    out = {}
    for sh in wb.iter(S + "sheet"):
        rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rid_target.get(rid, "")
        path = target.lstrip("/") if target.startswith("/") else "xl/" + target
        out[path] = sh.get("name")
    return out


def formula_protection(z, sheet_parts):
    """Textes et cellules a ne pas traduire a cause des formules.

    Une formule qui contient du texte ("Ja") compare des cellules a ce
    texte : traduire l'un sans l'autre fausse le calcul. On protege donc
    les chaines litterales des formules ET toutes les cellules des plages
    lues par ces formules (=NB.SI(C2:C5;"Ja") protege C2:C5).

    Renvoie (chaines protegees, {nom de feuille: [(col1, lig1, col2, lig2)]}).
    """
    literals, ranges = set(), {}
    names = sheet_names_by_part(z)
    for n in sheet_parts:
        here = names.get(n)
        x = z.read(n).decode("utf8", "ignore")
        for f in re.findall(r"<f[^>]*>([^<]*)</f>", x):
            f = f.replace("&quot;", '"').replace("&amp;", "&").replace("&apos;", "'")
            lits = _FORMULA_STR.findall(f)
            if not lits:
                continue
            for m in lits:
                if m.strip():
                    literals.add(m.replace('""', '"').strip())
            body = _FORMULA_STR.sub('""', f)
            for m in _REF_RE.finditer(body):
                sheet = (m.group(1) or "").replace("''", "'") or m.group(2) or here
                c1, r1, c2, r2 = m.group(3), m.group(4), m.group(5), m.group(6)
                if not r1 and not c2:
                    continue        # mot isole (nom de fonction), pas une reference
                col1 = _col_num(c1)
                col2 = _col_num(c2) if c2 else col1
                row1 = int(r1) if r1 else 1
                row2 = int(r2) if r2 else (row1 if r1 and not c2 else 1048576)
                ranges.setdefault(sheet, []).append((col1, row1, col2, row2))
    return literals, ranges


def _in_ranges(ref, ranges):
    m = re.match(r"([A-Z]+)(\d+)$", ref or "")
    if not m:
        return False
    col, row = _col_num(m.group(1)), int(m.group(2))
    return any(c1 <= col <= c2 and r1 <= row <= r2 for c1, r1, c2, r2 in ranges)


def translate_xlsx_universal(input_path: str, output_path: str, glossary_path: str = None,
                             progress_callback=None, direction="Allemand ➔ Français"):
    start_time = time.time()
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    # Dossier de travail propre a cet appel : deux traductions simultanees
    # d'un fichier du meme nom ne doivent pas partager leurs copies.
    work = tempfile.mkdtemp(prefix="traduction_")
    ext = os.path.splitext(input_path)[1]
    tmp_in = os.path.join(work, "in" + ext)
    tmp_out = os.path.join(work, "out" + ext)
    shutil.copy(input_path, tmp_in)
    tgt_lang = LANG[direction][1]
    glossary_df = load_glossary_df(glossary_path)

    from processors.pptx_processor import paragraph_segment as dml_segment
    from processors.pptx_processor import rebuild_paragraph as dml_rebuild

    zin = zipfile.ZipFile(tmp_in)
    names = zin.namelist()
    types = part_types(zin)
    sheets = [n for n in names if types.get(n, "").endswith("worksheet+xml")]
    protected, prot_ranges = formula_protection(zin, sheets)
    sheet_by_part = sheet_names_by_part(zin)
    # Chaines partagees lues par une formule : index proteges.
    protected_sst = set()
    for n in sheets:
        rng = prot_ranges.get(sheet_by_part.get(n))
        if not rng:
            continue
        x = zin.read(n).decode("utf8", "ignore")
        for ref, idx in re.findall(r'<c r="([A-Z]+\d+)"[^>]*t="s"[^>]*>\s*<v>(\d+)</v>', x):
            if _in_ranges(ref, rng):
                protected_sst.add(int(idx))

    roots = {}          # partie -> arbre XML modifie
    segments, contexts, kinds = [], [], []

    def add(seg, ctx, kind):
        if seg is None:
            return
        plain = seg.plain_text().strip()
        if plain in protected:
            return      # texte utilise par une formule : on n'y touche pas
        segments.append(seg)
        contexts.append(ctx)
        kinds.append(kind)

    for n in names:
        ct = types.get(n, "")
        if ct.endswith("sharedStrings+xml"):
            root = etree.fromstring(zin.read(n))
            roots[n] = root
            for k, si in enumerate(root.findall(S + "si")):
                if k not in protected_sst:
                    add(string_item_segment(si), "cellule", "si")
        elif n in sheets:
            data = zin.read(n)
            if b"inlineStr" not in data:
                continue
            root = etree.fromstring(data)
            roots[n] = root
            rng = prot_ranges.get(sheet_by_part.get(n), [])
            for c in root.iter(S + "c"):
                if c.get("t") == "inlineStr" and not _in_ranges(c.get("r"), rng):
                    isel = c.find(S + "is")
                    if isel is not None:
                        add(string_item_segment(isel), "cellule", "si")
        elif ct.endswith("spreadsheetml.comments+xml"):
            root = etree.fromstring(zin.read(n))
            roots[n] = root
            for txt in root.iter(S + "text"):
                add(string_item_segment(txt), "commentaire", "si")
        elif ct.endswith("officedocument.drawing+xml") or ct.endswith("drawingml.chart+xml"):
            root = etree.fromstring(zin.read(n))
            roots[n] = root
            ctx = "zone de texte" if ct.endswith("drawing+xml") and "chart" not in ct else "titre de graphique"
            for p in root.iter(A + "p"):
                add(dml_segment(p), ctx, "dml")
            if ct.endswith("chart+xml"):
                for v in root.iter(C + "v"):
                    par = v.getparent()
                    if par is not None and par.tag == C + "pt" and v.text \
                            and re.search(r"[^\W\d_]", v.text):
                        seg = Segment()
                        seg.add_text(v.text, "-", None)
                        seg.meta["v"] = v
                        add(seg, "libellé de graphique", "v")

    print(f"⚡ Traitement XLSX de {len(segments)} textes...")
    results, stats = translate_segments(segments, direction, "xlsx", glossary_df,
                                        progress_callback, contexts)

    for seg, items, kind in zip(segments, results, kinds):
        if items is None:
            continue
        if kind == "si":
            rebuild_string_item(seg, items)
        elif kind == "dml":
            dml_rebuild(seg, items, tgt_lang)
        else:
            seg.meta["v"].text = "".join(it[2] for it in items if it[0] == "t")

    # Ecriture : parties modifiees regenerees, le reste recopie tel quel.
    with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename in roots:
                data = etree.tostring(roots[info.filename], xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            else:
                data = zin.read(info.filename)
            zout.writestr(info, data)
    zin.close()

    shutil.copy(tmp_out, output_path)
    shutil.rmtree(work, ignore_errors=True)

    print(f"🔤 Segments : {stats['segments']}, envoyés : {stats['segments'] - stats['skipped']}, "
          f"redemandés : {stats['retried']}, sans styles locaux : {stats['fallback']}, "
          f"textes protégés par des formules : {len(protected) + len(protected_sst)}")
    return {
        "status": "success", "output_path": output_path,
        "time_seconds": round(time.time() - start_time, 2),
        "input_tokens": stats["input_tokens"], "output_tokens": stats["output_tokens"],
        "total_cost_usd": stats["input_tokens"] / 1e6 * PRICE_IN + stats["output_tokens"] / 1e6 * PRICE_OUT,
    }