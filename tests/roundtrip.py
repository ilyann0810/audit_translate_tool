# -*- coding: utf-8 -*-
"""
Test aller-retour, sans appel au modele.

Chaque segment est "traduit" par lui-meme : serialise en texte balise,
relu par Segment.parse puis reecrit dans le document. Le document produit
doit etre identique a l'original (memes textes, memes styles, memes
objets). Toute difference est un defaut d'extraction ou de reconstruction,
independant de la qualite de traduction.

    python tests/roundtrip.py [fichiers...]     (par defaut : manifeste + samples)
"""
import glob
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "src"))
sys.path.insert(0, HERE)
os.environ.setdefault("GCP_PROJECT_ID", "roundtrip-sans-appel")


def fake_translate(segments, direction, kind, glossary_df=None, progress_callback=None,
                   contexts=None):
    results = []
    for s in segments:
        if not s.has_text():
            results.append(None)
            continue
        items, problem = s.parse(s.to_tagged())
        if problem:
            print("   !! aller-retour du balisage :", problem, "|", s.to_tagged()[:80])
        results.append(items)
    return results, {"input_tokens": 0, "output_tokens": 0, "calls": 0, "retried": 0,
                     "fallback": 0, "skipped": 0, "segments": len(segments)}


def main():
    import core.translation_api as api
    api.translate_segments = fake_translate
    from processors import pptx_processor, docx_processor
    pptx_processor.translate_segments = fake_translate
    pptx_processor.shorten_segments = lambda items, direction: ({}, 0, 0)
    docx_processor.translate_segments = fake_translate
    try:
        from processors import xlsx_processor
        xlsx_processor.translate_segments = fake_translate
    except Exception:
        xlsx_processor = None
    import evaluate

    files = sys.argv[1:]
    if not files:
        man = json.load(open(os.path.join(HERE, "manifest.json"), encoding="utf-8"))
        in_dir = os.path.normpath(os.path.join(HERE, man["input_dir"]))
        files = [os.path.join(in_dir, d["file"]) for d in man["documents"]]
        files += sorted(glob.glob(os.path.join(HERE, "samples", "*.*")))
    out_dir = os.path.join(ROOT, "Outputs", "_roundtrip")
    os.makedirs(out_dir, exist_ok=True)
    for f in files:
        ext = f.rsplit(".", 1)[-1].lower()
        if ext not in ("pptx", "docx", "xlsx"):
            continue
        dst = os.path.join(out_dir, os.path.basename(f))
        print("==", os.path.basename(f))
        if ext == "pptx":
            pptx_processor.translate_pptx_universal(f, dst, None, direction="Allemand ➔ Français")
            r = evaluate.eval_pptx(f, dst, "Allemand ➔ Français", office=False)
            bad = {k: v for k, v in r.items() if k not in ("residue", "paragraphs") and v}
        elif ext == "docx":
            docx_processor.translate_docx_universal(f, dst, None, direction="Allemand ➔ Français")
            r = evaluate.eval_docx(f, dst, "Allemand ➔ Français", office=False)
            bad = {k: v for k, v in r.items() if k not in ("residue", "office") and v}
        else:
            xlsx_processor.translate_xlsx_universal(f, dst, None, direction="Allemand ➔ Français")
            import evaluate_xlsx
            bad = evaluate_xlsx.compare(f, dst)
        # Textes : chaque paragraphe doit etre identique.
        diffs = text_diffs(f, dst, ext)
        if diffs:
            bad["texte"] = diffs[:5] + (["... %d au total" % len(diffs)] if len(diffs) > 5 else [])
        if bad:
            for k, v in bad.items():
                print("   %s : %s" % (k, json.dumps(v, ensure_ascii=False, default=str)[:600]))
        else:
            print("   identique")


def text_diffs(a, b, ext):
    import zipfile
    import re
    za, zb = zipfile.ZipFile(a), zipfile.ZipFile(b)
    out = []
    for n in za.namelist():
        if not n.endswith(".xml") or n not in zb.namelist():
            continue
        xa, xb = za.read(n).decode("utf8", "ignore"), zb.read(n).decode("utf8", "ignore")
        pat = r"<(?:a|w):t(?: [^>]*)?>([^<]*)</(?:a|w):t>" if ext != "xlsx" else r"<t(?: [^>]*)?>([^<]*)</t>"
        import html
        ta = html.unescape("".join(re.findall(pat, xa)))
        tb = html.unescape("".join(re.findall(pat, xb)))
        if ta != tb:
            i = next((k for k, (x, y) in enumerate(zip(ta, tb)) if x != y), min(len(ta), len(tb)))
            out.append("%s : ...%r... != ...%r..." % (n, ta[max(0, i - 30):i + 30], tb[max(0, i - 30):i + 30]))
    return out


if __name__ == "__main__":
    main()
