# -*- coding: utf-8 -*-
"""
Mesures prises par Office lui-meme (Windows + PowerPoint/Word installes).

Seul PowerPoint sait ou tombe vraiment chaque ligne de texte : on lui
demande, forme par forme, la boite reellement occupee par le texte
(BoundTop/BoundHeight) et la hauteur reelle des lignes de tableau, qui
grandissent quand le texte ne tient plus. Word donne le nombre de pages.
Les deux exportent aussi un PDF, rendu ensuite en PNG pour l'oeil.

Outil de test uniquement : il n'est pas utilise par l'application.
"""
import json
import os
import sys

import pythoncom
import win32com.client

MSO_GROUP = 6
MSO_TRUE = -1


def _text_shapes(shapes, path=()):
    for i in range(1, shapes.Count + 1):
        sh = shapes.Item(i)
        key = path + (sh.Id,)
        if sh.Type == MSO_GROUP:
            yield from _text_shapes(sh.GroupItems, key)
            continue
        yield key, sh


def probe_pptx(path, pdf_out=None):
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    prs = app.Presentations.Open(os.path.abspath(path), True, False, False)
    out = {"slides": []}
    try:
        for s_idx in range(1, prs.Slides.Count + 1):
            slide = prs.Slides.Item(s_idx)
            items = []
            for key, sh in _text_shapes(slide.Shapes):
                rec = {"id": "/".join(map(str, key)), "name": sh.Name,
                       "x": sh.Left, "y": sh.Top, "w": sh.Width, "h": sh.Height}
                try:
                    rec["framed"] = bool(sh.Fill.Visible == MSO_TRUE or sh.Line.Visible == MSO_TRUE)
                except Exception:
                    rec["framed"] = True
                try:
                    has_table = sh.HasTable == MSO_TRUE
                except Exception:
                    has_table = False
                if has_table:
                    tb = sh.Table
                    rec["table"] = {
                        "rows": [tb.Rows.Item(r).Height for r in range(1, tb.Rows.Count + 1)],
                    }
                    cells = []
                    for r in range(1, tb.Rows.Count + 1):
                        for c in range(1, tb.Columns.Count + 1):
                            try:
                                tr = tb.Cell(r, c).Shape.TextFrame.TextRange
                                if tr.Text.strip():
                                    cells.append({"r": r, "c": c, "text": tr.Text[:200],
                                                  "bw": tr.BoundWidth, "bh": tr.BoundHeight})
                            except Exception:
                                pass
                    rec["cells"] = cells
                    items.append(rec)
                    continue
                try:
                    if sh.HasTextFrame != MSO_TRUE or not sh.TextFrame.HasText:
                        rec["other"] = True
                        items.append(rec)
                        continue
                    tf = sh.TextFrame
                    tr = tf.TextRange
                    rec.update({
                        "text": tr.Text[:300],
                        "bx": tr.BoundLeft, "by": tr.BoundTop,
                        "bw": tr.BoundWidth, "bh": tr.BoundHeight,
                        "autosize": tf.AutoSize, "wrap": tf.WordWrap,
                        "rot": sh.Rotation,
                    })
                    try:
                        rec["autofit2"] = sh.TextFrame2.AutoSize
                    except Exception:
                        pass
                except Exception:
                    continue
                items.append(rec)
            out["slides"].append(items)
        out["slide_w"] = prs.PageSetup.SlideWidth
        out["slide_h"] = prs.PageSetup.SlideHeight
        if pdf_out:
            prs.SaveAs(os.path.abspath(pdf_out), 32)  # ppSaveAsPDF
    finally:
        prs.Close()
        app.Quit()
    return out


def probe_docx(path, pdf_out=None):
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    doc = app.Documents.Open(os.path.abspath(path), False, True, False)
    try:
        doc.Repaginate()
        out = {
            "pages": doc.ComputeStatistics(2),
            "words": doc.ComputeStatistics(0),
            "inline_shapes": doc.InlineShapes.Count,
            "shapes": doc.Shapes.Count,
            "fields": doc.Fields.Count,
            "hyperlinks": doc.Hyperlinks.Count,
            "tables": doc.Tables.Count,
        }
        # Zones de texte qui debordent : Word le signale par Overflowing.
        overflow = []
        for i in range(1, doc.Shapes.Count + 1):
            sh = doc.Shapes.Item(i)
            try:
                if sh.TextFrame.HasText and sh.TextFrame.Overflowing:
                    overflow.append(sh.TextFrame.TextRange.Text[:120])
            except Exception:
                pass
        out["overflowing_textboxes"] = overflow
        if pdf_out:
            doc.SaveAs2(os.path.abspath(pdf_out), 17)  # wdFormatPDF
    finally:
        doc.Close(False)
        app.Quit()
    return out


if __name__ == "__main__":
    src, dst_json = sys.argv[1], sys.argv[2]
    pdf = sys.argv[3] if len(sys.argv) > 3 else None
    fn = probe_pptx if src.lower().endswith(".pptx") else probe_docx
    res = fn(src, pdf)
    with open(dst_json, "w", encoding="utf-8") as fh:
        json.dump(res, fh, ensure_ascii=False, indent=1)
