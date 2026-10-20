# -*- coding: utf-8 -*-
"""
Banc d'essai : traduit les documents du manifeste avec le code de src/.

    python tests/run_suite.py <etiquette> [--only motif] [--manifest fichier]

Les sorties vont dans Outputs/<etiquette>/. Les reponses du modele sont
mises en cache dans .cache/translations : une relance ne repaie que les
prompts qui ont change, et deux versions du code comparees sur le meme
cache recoivent les memes traductions.
"""
import argparse
import json
import os
import sys
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "src")


def load_env_local():
    """Charge GCP_PROJECT_ID depuis un .env.local situe plus haut."""
    d = ROOT
    while True:
        p = os.path.join(d, ".env.local")
        if os.path.isfile(p):
            for line in open(p, encoding="utf-8"):
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip().strip('"\''))
            return
        parent = os.path.dirname(d)
        if parent == d:
            return
        d = parent


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("label")
    ap.add_argument("--only", default=None)
    ap.add_argument("--manifest", default=os.path.join(HERE, "manifest.json"))
    ap.add_argument("--glossary", default=None, help="glossaire Excel / CSV (Allemand, Francais)")
    args = ap.parse_args()
    args.manifest = os.path.abspath(args.manifest)
    if args.glossary:
        args.glossary = os.path.abspath(args.glossary)

    load_env_local()
    os.environ.setdefault("TRANSLATION_CACHE_DIR",
                          os.path.join(ROOT, ".cache", "translations"))
    os.makedirs(os.environ["TRANSLATION_CACHE_DIR"], exist_ok=True)
    sys.path.insert(0, SRC)
    os.chdir(SRC)
    from core.router import process_translation

    man = json.load(open(args.manifest, encoding="utf-8"))
    in_dir = os.path.normpath(os.path.join(os.path.dirname(args.manifest), man["input_dir"]))
    out_dir = os.path.join(ROOT, "Outputs", args.label)
    os.makedirs(out_dir, exist_ok=True)

    summary = []
    for d in man["documents"]:
        if args.only and args.only.lower() not in d["file"].lower():
            continue
        src = os.path.join(in_dir, d["file"])
        dst = os.path.join(out_dir, d["file"])
        print("\n===", d["file"], "|", d["direction"], flush=True)
        t0 = time.time()
        try:
            res = process_translation(src, dst, args.glossary, direction=d["direction"])
            ok = res.get("status") == "success"
            msg = "" if ok else res.get("error_message", "")
        except Exception as exc:
            traceback.print_exc()
            ok, msg = False, str(exc)
        summary.append((d["file"], ok, round(time.time() - t0), msg))

    print("\n" + "-" * 60)
    for f, ok, t, msg in summary:
        print("%-4s %4ds  %s %s" % ("ok" if ok else "FAIL", t, f[:60], msg[:80]))


if __name__ == "__main__":
    main()
