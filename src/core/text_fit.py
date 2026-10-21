# -*- coding: utf-8 -*-
"""
Mesure de texte : combien de place prend un paragraphe dans un cadre.

PowerPoint ne recalcule pas la mise en page a l'ouverture d'un fichier :
si la traduction est plus longue, le texte sort du cadre. Il faut donc
savoir AVANT d'ecrire si la traduction tiendra, et sinon de combien
reduire.

On simule le retour a la ligne mot par mot avec les largeurs reelles des
caracteres de la police (PyMuPDF lit les polices installees). La mesure
n'est jamais parfaite (police absente du serveur, crenage...), c'est
pourquoi les processeurs comparent toujours la traduction a l'ORIGINAL
mesure de la meme facon : l'erreur de mesure est la meme des deux cotes.

Polices : recherche par nom de famille dans les dossiers de polices du
systeme. Sans la police (police d'entreprise non installee, serveur Linux
nu), on prend une police de silhouette proche : condensee -> Arial Narrow
ou Helvetica resserree, empattements -> Times, sinon Helvetica (memes
largeurs qu'Arial, incluse dans PyMuPDF donc toujours disponible).
"""
import glob
import os
import re
import threading

import fitz

_FONT_DIRS = [
    os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
    "/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.fonts"),
    os.path.expanduser("~/.local/share/fonts"), "/Library/Fonts", "/System/Library/Fonts",
]

_STYLE_WORDS = re.compile(
    r"\b(regular|bold|italic|oblique|book|normal|roman|demibold|semibold|medium|"
    r"light|black|heavy|thin|extrabold|ultrabold)\b", re.I)

_index = None
_index_lock = threading.Lock()
_font_cache = {}


def _build_index():
    idx = {}
    for d in _FONT_DIRS:
        if not d or not os.path.isdir(d):
            continue
        for path in glob.glob(os.path.join(d, "**", "*"), recursive=True):
            if not path.lower().endswith((".ttf", ".otf", ".ttc")):
                continue
            try:
                f = fitz.Font(fontfile=path)
            except Exception:
                continue
            fam = _STYLE_WORDS.sub("", f.name or "").strip().lower()
            fam = re.sub(r"\s+", " ", fam)
            if not fam:
                continue
            key = (bool(f.is_bold), bool(f.is_italic))
            # A style egal, le fichier "normal" (sans Light/Semibold) gagne.
            weight_word = re.search(r"light|semibold|demibold|medium|black|heavy|thin",
                                    f.name or "", re.I)
            prev = idx.setdefault(fam, {}).get(key)
            if prev is None or (prev[1] and not weight_word):
                idx[fam][key] = (path, bool(weight_word))
    return idx


def _get_index():
    global _index
    with _index_lock:
        if _index is None:
            _index = _build_index()
    return _index


_BASE14 = {  # (gras, italique) -> police incluse dans PyMuPDF
    "sans": {(False, False): "helv", (True, False): "hebo", (False, True): "heit", (True, True): "hebi"},
    "serif": {(False, False): "tiro", (True, False): "tibo", (False, True): "tiit", (True, True): "tibi"},
    "mono": {(False, False): "cour", (True, False): "cobo", (False, True): "coit", (True, True): "cobi"},
}


class MeasureFont:
    """Police de mesure, avec un facteur de largeur quand c'est une
    police de remplacement plus large ou plus etroite que l'originale."""

    def __init__(self, font, factor=1.0):
        self.font = font
        self.factor = factor

    def width(self, text, size):
        try:
            return self.font.text_length(text, fontsize=size) * self.factor
        except Exception:
            return len(text) * size * 0.5 * self.factor


def get_font(family, bold=False, italic=False):
    """Police de mesure pour une famille et un style."""
    key = ((family or "").lower(), bool(bold), bool(italic))
    if key in _font_cache:
        return _font_cache[key]
    fam = re.sub(r"\s+", " ", _STYLE_WORDS.sub("", family or "")).strip().lower()
    idx = _get_index()
    mf = None
    styles = idx.get(fam)
    if styles:
        path = (styles.get((bold, italic)) or styles.get((bold, False))
                or styles.get((False, False)) or next(iter(styles.values())))[0]
        try:
            f = fitz.Font(fontfile=path)
            # Gras demande mais seule la version normale existe : le gras
            # synthetique de PowerPoint elargit d'environ 5 %.
            factor = 1.05 if bold and not f.is_bold else 1.0
            mf = MeasureFont(f, factor)
        except Exception:
            mf = None
    if mf is None:
        mf = _substitute(fam, bold, italic, idx)
    _font_cache[key] = mf
    return mf


def _substitute(fam, bold, italic, idx):
    narrow = any(k in fam for k in ("cond", "narrow", "compressed"))
    if narrow:
        for cand in ("arial narrow", "liberation sans narrow"):
            st = idx.get(cand)
            if st:
                path = (st.get((bold, italic)) or st.get((False, False)))[0]
                return MeasureFont(fitz.Font(fontfile=path))
        return MeasureFont(fitz.Font(_BASE14["sans"][(bold, italic)]), 0.82)
    if any(k in fam for k in ("courier", "mono", "consolas")):
        kind = "mono"
    elif any(k in fam for k in ("times", "georgia", "garamond", "cambria", "serif", "roman",
                                "minion", "palatino", "book antiqua")) and "sans" not in fam:
        kind = "serif"
    else:
        kind = "sans"
    factor = 1.0
    # Calibri/Carlito et polices "humanistes" recentes : nettement plus
    # etroites qu'Arial.
    if any(k in fam for k in ("calibri", "carlito", "aptos")):
        for cand in ("carlito", "calibri"):
            st = idx.get(cand)
            if st:
                path = (st.get((bold, italic)) or st.get((False, False)))[0]
                return MeasureFont(fitz.Font(fontfile=path))
        factor = 0.9
    return MeasureFont(fitz.Font(_BASE14[kind][(bold, italic)]), factor)


# --------------------------------------------------------------------------
# Mise en page
# --------------------------------------------------------------------------

LINE_FACTOR = 1.2   # hauteur d'une ligne en interligne simple, en tailles de police

_WORD_RE = re.compile(r"\S+\s*|\s+")


class Run:
    __slots__ = ("text", "size", "font")

    def __init__(self, text, size, font):
        self.text, self.size, self.font = text, size, font


class Para:
    """Paragraphe a mesurer.

    runs : liste de Run ; line_spacing : ("pct", 1.0) ou ("pts", 14.0) ;
    space_before / space_after en points ; indent : retrait gauche (pts).
    """

    def __init__(self, runs, line_spacing=("pct", 1.0), space_before=0.0,
                 space_after=0.0, indent=0.0, first_indent=0.0):
        self.runs = runs
        self.line_spacing = line_spacing
        self.space_before = space_before
        self.space_after = space_after
        self.indent = indent
        self.first_indent = first_indent


def _tokens(runs, scale):
    """Mots (avec leur espace de fin) et retours forces, avec leur largeur."""
    out = []
    for r in runs:
        size = r.size * scale
        parts = r.text.split("\n")
        for k, part in enumerate(parts):
            if k:
                out.append(("br", 0.0, 0.0, size))
            for w in _WORD_RE.findall(part):
                stripped = w.rstrip()
                wid = r.font.width(stripped, size)
                full = r.font.width(w, size) if w != stripped else wid
                out.append(("w", wid, full, size))
    return out


def layout(paras, width, scale=1.0, spacing_reduction=0.0, wrap=True):
    """Hauteur totale, nombre de lignes par paragraphe, largeur maximale.

    Un mot plus long que la ligne est coupe (comme PowerPoint), ce qui
    compte une ligne de plus par largeur depassee.
    """
    total_h = 0.0
    lines_per_para = []
    max_w = 0.0
    for i, p in enumerate(paras):
        toks = _tokens(p.runs, scale)
        sizes = [t[3] for t in toks] or [max((r.size for r in p.runs), default=10.0) * scale]
        avail = max(1.0, width - p.indent)
        lines, cur_w, line_sizes, cur_sizes = 0, 0.0, [], []
        first = True

        def flush():
            nonlocal lines, cur_w, cur_sizes, first
            lines += 1
            line_sizes.append(max(cur_sizes) if cur_sizes else (sizes[0] if sizes else 10.0))
            cur_w, cur_sizes, first = 0.0, [], False

        for kind, wid, full, size in toks:
            if kind == "br":
                flush()
                continue
            lim = avail - (p.first_indent if first else 0.0)
            if wrap and cur_w > 0 and cur_w + wid > lim:
                max_w = max(max_w, cur_w)
                flush()
                lim = avail
            if wrap and wid > lim and cur_w == 0:
                # mot trop long : PowerPoint le coupe en plusieurs lignes
                n = int(wid // max(lim, 1.0))
                for _ in range(n):
                    cur_sizes.append(size)
                    flush()
                wid = wid - n * lim
                full = wid
            cur_w += full
            cur_sizes.append(size)
            max_w = max(max_w, cur_w)
        if cur_sizes or not line_sizes:
            flush()
        lines_per_para.append(lines)

        ls_kind, ls_val = p.line_spacing
        if ls_kind == "pts":
            h = ls_val * len(line_sizes)
        else:
            pct = max(0.5, ls_val - spacing_reduction)
            h = sum(s * LINE_FACTOR * pct for s in line_sizes)
        total_h += h
        if i > 0:
            total_h += p.space_before * (1.0 - spacing_reduction)
        if i < len(paras) - 1:
            total_h += p.space_after * (1.0 - spacing_reduction)
    return total_h, lines_per_para, max_w