# -*- coding: utf-8 -*-
"""
Traduction de PDF en gardant le style.

Google translateDocument aplatit chaque page en une image puis repose le
texte par-dessus, sans gras, dans une police generique, avec un filigrane.
Ici on remplace le texte sur place :

  1. extraction des blocs, avec pour chaque passage sa police, sa taille,
     sa couleur, sa graisse ; l'alignement, l'interligne et l'orientation
     du bloc sont deduits de la position des lignes
  2. les passages dont le style differe du reste du bloc sont balises
     (<s1>...</s1>) ; le modele traduit en gardant les balises autour des
     mots correspondants, ce qui conserve un mot en gras ou en couleur au
     milieu d'une phrase
  3. suppression du texte d'origine SANS remplissage : photos, aplats de
     couleur et dessin vectoriel restent intacts sous le texte
  4. mise en page de la traduction par le moteur HTML de PyMuPDF
     (insert_htmlbox), qui gere les styles melanges, la justification, la
     rotation et la reduction de taille quand le texte est plus long

Polices : celles embarquees dans un PDF sont presque toujours tronquees aux
seuls caracteres du document. Un PDF allemand n'a ni e accent aigu ni c
cedille. On reprend la police embarquee quand elle couvre la traduction,
sinon une police Windows de meme famille et de meme graisse.

Limites : pas d'OCR (un PDF scanne n'a pas de texte a remplacer), et le
texte incruste dans une image (graphique exporte en bitmap) reste tel quel.
"""
from __future__ import annotations

import html
import math
import os
import re
import unicodedata

import fitz

from .llm import (DEFAULT_MODEL, qa_flags, shorten, transcribe_images,
                  translate_document)

# En dessous de cette echelle, un bloc est juge trop serre : on redemande
# au modele une traduction plus courte avant de l'ecrire.
SHORTEN_BELOW = 0.85


# Un PDF ne stocke pas de mots mais des glyphes positionnes, et MuPDF
# devine les espaces d'apres les ecarts. Les deux erreurs existent :
#   - titre crene serre      : "Ei nf uehr ung"          (espaces en trop)
#   - etiquettes sans espace : "wennichdafuerfinanziell" (espaces manquants)
# On lit donc chaque page avec et sans la devinette, et on garde par bloc
# la version dont les mots ont une longueur plausible.
FLAGS_GUESS = fitz.TEXTFLAGS_DICT
FLAGS_RAW = fitz.TEXTFLAGS_DICT | fitz.TEXT_INHIBIT_SPACES
NATURAL_WORD_LEN = 6.0

# Polices dont la table de caracteres est fausse (frequent dans les
# captures d'ecran exportees) : le texte extrait est du charabia du type
# '<")1)I#.2I(0="81'. Le traduire puis le redessiner detruirait l'image.
# Le %, le + ou le ? sont frequents dans du vrai texte ("0% 20% 40%") : les
# compter faisait prendre l'axe d'un graphique pour du charabia, puis
# marquait sa police comme fausse pour toute la page.
CODE_CHARS = set('"#$=<>@{}|\\^_`~')

# Reduction maximale acceptee avant de considerer qu'un bloc deborde.
MIN_SCALE = 0.62

FONT_DIR = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")

# Familles de remplacement, par (gras, italique). Le choix vise la
# silhouette : condense pour condense, empattements pour empattements.
SYSTEM_FONTS = {
    "calibri": ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
    "arial": ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    "narrow": ("ARIALN.TTF", "ARIALNB.TTF", "ARIALNI.TTF", "ARIALNBI.TTF"),
    "serif": ("georgia.ttf", "georgiab.ttf", "georgiai.ttf", "georgiaz.ttf"),
    "times": ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
    "sans": ("segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf"),
    "mono": ("cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"),
}

# Demi-gras disponibles. Ailleurs, un demi-gras devient une graisse normale,
# plus proche visuellement qu'un gras plein.
SEMIBOLD_FONTS = {
    "sans": ("seguisb.ttf", "seguisbi.ttf"),
    "mono": ("cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"),
}

BULLETS = set("•·▪◦‣–—*○●□■»›")
TAG_RE = re.compile(r"<s(\d+)>(.*?)</s\1>", re.S)
ANY_TAG_RE = re.compile(r"</?s\d+>")


# --------------------------------------------------------------------------
# Modele
# --------------------------------------------------------------------------

class Style:
    __slots__ = ("font", "size", "color", "bold", "italic", "script",
                 "semibold", "deco")

    def __init__(self, span, script=None):
        self.script = script    # "sub", "sup" ou None
        # Souligne / barre / surligne dessines a part : (("u", couleur), ...)
        self.deco = span.get("_deco", ())
        self.font = span["font"]
        self.size = round(span["size"], 1)
        self.color = span["color"]
        name = self.font.lower()
        # Un demi-gras remplace par un gras plein alourdit tout un
        # paragraphe : "DemosEF-SemiBold" est un corps de texte, pas un
        # titre. On le garde a part.
        self.semibold = any(w in name for w in
                            ("semibold", "demibold", "semi-bold", "demi"))
        self.bold = not self.semibold and (
            bool(span["flags"] & 16)
            or any(w in name for w in ("bold", "black", "heavy")))
        self.italic = bool(span["flags"] & 2) or any(
            w in name for w in ("italic", "oblique", "ita"))

    def key(self):
        return (self.font, self.size, self.color, self.bold, self.italic,
                self.script, self.semibold, self.deco)

    def has(self, kind):
        return next((c for k, c in self.deco if k == kind), None)


class Block:
    """Un bloc de texte et tout ce qu'il faut pour le redessiner."""

    __slots__ = ("uid", "page", "rect", "line_rects", "styles", "base",
                 "text", "align", "rotate", "line_height", "ocr", "room",
                 "angle", "tilt_lines", "origin", "hard_lines", "room_bottom",
                 "decos", "hang", "category", "style", "box", "runs", "uniform_style")

    def __init__(self, uid, page, rect, line_rects, styles, base, text,
                 align, rotate, line_height, ocr=False):
        self.ocr = ocr              # texte a lire sur l'image
        self.origin = None          # point de base du premier caractere
        self.hard_lines = False     # chaque ligne finit par un retour voulu
        self.room_bottom = None     # bas de l'espace libre sous le bloc
        self.decos = []             # traits / aplats de decoration a effacer
        self.hang = None            # retrait suspendu d'une liste a puces (pt)
        self.room = rect           # cadre d'ecriture, eventuellement elargi
        self.angle = 0.0            # texte en biais : angle en degres
        self.tilt_lines = []        # en biais : [(centre, texte, taille)]
        self.uid = uid
        self.page = page
        self.rect = rect
        self.line_rects = line_rects
        self.styles = styles        # id balise -> Style ; 0 = style de base
        self.base = base
        self.text = text            # texte envoye au modele, balises incluses
        self.align = align
        self.rotate = rotate
        self.line_height = line_height
        # Champs attendus par llm.py.
        self.category = "PDF"
        self.style = {"size_pt": base.size}
        self.box = {"w_pt": rect.width, "h_pt": rect.height}
        self.runs = len(styles)
        self.uniform_style = len(styles) == 1

    def context_label(self):
        w, h = (self.rect.height, self.rect.width) if self.rotate in (90, 270) \
            else (self.rect.width, self.rect.height)
        kind = "titre" if len(self.text) <= 60 and self.base.size >= 14 \
            else "texte"
        return "%s dans un cadre de %d x %d points" % (kind, w, h)


# --------------------------------------------------------------------------
# Extraction
# --------------------------------------------------------------------------

def _rotation(line):
    dx, dy = line["dir"]
    if abs(dx - 1) < 0.01:
        return 0
    if abs(dy + 1) < 0.01:
        return 90
    if abs(dy - 1) < 0.01:
        return 270
    if abs(dx + 1) < 0.01:
        return 180
    return None


def _mean_word_len(text):
    words = re.findall(r"[^\W\d_]+", text)
    return sum(map(len, words)) / len(words) if words else NATURAL_WORD_LEN


def _block_text(b):
    return " ".join("".join(s["text"] for s in ln["spans"])
                    for ln in b["lines"])


def _garbled(text):
    chars = [c for c in text if not c.isspace()]
    if not chars:
        return False
    code = sum(c in CODE_CHARS for c in chars)
    if code >= 3 and code > 0.12 * len(chars):
        return True
    # Autre signature : longues suites d'un meme caractère (",,,,,,,,,,")
    # dans un bloc ou les lettres sont minoritaires.
    letters = sum(c.isalpha() for c in chars)
    return bool(re.search(r"(.)\1{9,}", text)) and letters < 0.5 * len(chars)


def _script(span, line):
    """Indice ou exposant : plus petit que la ligne et decale verticalement."""
    big = max(line["spans"], key=lambda s: s["size"])
    if span["size"] > 0.85 * big["size"]:
        return None
    dy = span["origin"][1] - big["origin"][1]
    if dy > 0.08 * big["size"]:
        return "sub"
    if dy < -0.08 * big["size"] or span["flags"] & 1:
        return "sup"
    return None


def _has_letters(s):
    return any(unicodedata.category(c).startswith("L") for c in s)


# Au-dela de cet ecart (en tailles de police) entre deux caracteres d'une
# meme ligne, on considere que ce sont deux cellules ou deux colonnes. Un
# espace justifie depasse rarement une taille de police.
GAP_SPLIT = 1.6


def _split_line(line):
    """Coupe une ligne MuPDF (format rawdict) aux grands vides horizontaux.

    MuPDF met parfois sur une meme ligne des morceaux que la mise en page
    separe nettement : le libelle et les montants d'une ligne de tableau,
    deux pieds de page. Redessines d'un seul tenant, les montants quittent
    leurs colonnes. On mesure l'ecart entre caracteres visibles (les
    espaces inseres par MuPDF ne comptent pas) et on coupe la ou il
    depasse GAP_SPLIT tailles de police.

    Renvoie des lignes au format dict (spans avec "text").
    """
    horizontal = abs(line["dir"][0] - 1) < 0.01
    segs, cur, last_x1 = [], None, None

    for sp in line["spans"]:
        for ch in sp["chars"]:
            c = ch["c"]
            visible = not c.isspace()
            x0, x1 = ch["bbox"][0], ch["bbox"][2]
            if (horizontal and visible and cur is not None and last_x1 is not None
                    and x0 - last_x1 > max(GAP_SPLIT * sp["size"], 6)):
                segs.append(cur)
                cur, last_x1 = None, None
            if cur is None:
                if not visible:
                    continue
                cur = {"dir": line["dir"], "wmode": line.get("wmode", 0),
                       "spans": []}
            if not cur["spans"] or cur["spans"][-1]["_src"] is not sp:
                cur["spans"].append({
                    "_src": sp, "text": "", "font": sp["font"],
                    "size": sp["size"], "color": sp["color"],
                    "flags": sp["flags"], "origin": ch["origin"],
                    "bbox": list(ch["bbox"])})
            span = cur["spans"][-1]
            span["text"] += c
            span.setdefault("chars", []).append(ch)
            bb = span["bbox"]
            span["bbox"] = [min(bb[0], ch["bbox"][0]), min(bb[1], ch["bbox"][1]),
                            max(bb[2], ch["bbox"][2]), max(bb[3], ch["bbox"][3])]
            if visible:
                last_x1 = x1
    if cur is not None:
        segs.append(cur)

    out = []
    for seg in segs:
        if not "".join(s["text"] for s in seg["spans"]).strip():
            continue
        r = fitz.Rect(seg["spans"][0]["bbox"])
        for sp in seg["spans"][1:]:
            r |= fitz.Rect(sp["bbox"])
        seg["bbox"] = tuple(r)
        out.append(seg)
    return out


def _as_dict_block(b):
    """Bloc rawdict -> bloc dict, avec les lignes coupees aux grands vides."""
    lines = [seg for ln in b["lines"] for seg in _split_line(ln)]
    return {"type": 0, "bbox": b["bbox"], "lines": lines}


def _hidden_spans(page):
    """Textes presents dans le PDF mais invisibles a l'ecran.

    Vestiges d'une edition precedente (texte blanc sur blanc), couche texte
    invisible d'un scan, texte recouvert par une forme : le traduire puis
    le redessiner par-dessus tout le ferait apparaitre.

    On s'appuie sur le journal de peinture de la page (get_bboxlog), qui
    liste dans l'ordre tout ce qui est peint : formes remplies, traits (un
    bandeau de couleur est parfois un trait tres epais), images, texte. Un
    texte est cache s'il est en mode invisible, s'il est recouvert en bonne
    partie par une peinture opaque posee apres lui, ou s'il a la couleur de
    ce qui est peint juste dessous. Renvoie [(rect, texte)].
    """
    log = page.get_bboxlog()
    draw = {d["seqno"]: d for d in page.get_drawings()}
    # Une image a fond transparent (schema detoure) ne cache rien la ou
    # elle est transparente : seules les images sans masque comptent.
    masked = {tuple(round(v) for v in im["bbox"]) for im in page.get_image_info()
              if im.get("has-mask")}

    def alpha(d, key):
        # Une opacite de 0.0 est une vraie valeur (aplat invisible pose par
        # PowerPoint derriere chaque zone de texte) : « or 1 » la prenait
        # pour une absence de valeur, donc pour un aplat opaque, et le
        # texte pose dessus passait pour cache.
        v = d.get(key)
        return 1.0 if v is None else v

    def colour(j):
        kind, d = log[j][0], draw.get(j)
        if kind == "fill-image" or d is None:
            return None                     # image : couleur inconnue
        if kind == "fill-path" and alpha(d, "fill_opacity") > 0.99:
            return d.get("fill")
        if kind == "stroke-path" and alpha(d, "stroke_opacity") > 0.99:
            return d.get("color")
        return None

    def visible(j):
        kind, d = log[j][0], draw.get(j)
        if kind == "fill-image" or d is None:
            return True
        if kind == "fill-path":
            return alpha(d, "fill_opacity") > 0.05
        return alpha(d, "stroke_opacity") > 0.05

    def opaque(j):
        kind, d = log[j][0], draw.get(j)
        if kind == "fill-image":
            return tuple(round(v) for v in log[j][1]) not in masked
        return (kind == "fill-path" and d is not None and d.get("fill")
                and alpha(d, "fill_opacity") > 0.99)

    out = []
    for sp in page.get_texttrace():
        i, r = sp["seqno"], fitz.Rect(sp["bbox"])
        if r.is_empty or i >= len(log):
            continue
        text = "".join(chr(c[0]) for c in sp["chars"])
        hidden = sp.get("type") == 3 or log[i][0] == "ignore-text"
        # Hors de la page (reperes de mise en page de PowerPoint poses
        # au-dessus du bord) : jamais affiche.
        if not hidden:
            inside = (r & page.rect).get_area()
            hidden = inside < 0.5 * r.get_area()
        if not hidden:
            covered = sum((fitz.Rect(log[j][1]) & r).get_area()
                          for j in range(i + 1, len(log))
                          if opaque(j) and fitz.Rect(log[j][1]).intersects(r))
            hidden = covered > 0.25 * r.get_area()
        if not hidden:
            mid = (r.tl + r.br) / 2
            under = [j for j in range(i)
                     if log[j][0] in ("fill-path", "stroke-path", "fill-image")
                     and fitz.Rect(log[j][1]).contains(mid) and visible(j)]
            if under:
                bg = colour(under[-1])
            else:
                bg = (1, 1, 1)              # papier
            col = sp.get("color") or (0, 0, 0)
            hidden = bg is not None and len(bg) == len(col) and max(
                abs(a - b) for a, b in zip(col, bg)) < 0.08
        if hidden:
            out.append((r, text))
    return out


def _is_hidden(line, hidden):
    t = "".join(sp["text"] for sp in line["spans"]).replace(" ", "")
    lr = fitz.Rect(line["bbox"])
    mid = (lr.tl + lr.br) / 2
    return any(r.contains(mid) and t and t in ht.replace(" ", "")
               for r, ht in hidden)


def _split_columns(lines):
    """Separe les lignes posees cote a cote dans des colonnes distinctes.

    MuPDF range parfois dans un meme bloc des lignes de colonnes
    differentes situees a la meme hauteur : les deux pieds de page d'une
    double page, les cellules d'une ligne de tableau. Recollees, elles
    donnent "PAGE 14 La ventilation La ventilation PAGE 15" redessine d'un
    seul tenant. Deux lignes ne restent ensemble que si elles se
    chevauchent horizontalement.
    """
    groups = []
    for ln in lines:
        x0, x1 = ln["bbox"][0], ln["bbox"][2]
        for g in groups:
            overlap = min(x1, g["x1"]) - max(x0, g["x0"])
            if overlap > 0.3 * max(1.0, min(x1 - x0, g["x1"] - g["x0"])):
                g["lines"].append(ln)
                g["x0"], g["x1"] = min(g["x0"], x0), max(g["x1"], x1)
                break
        else:
            groups.append({"lines": [ln], "x0": x0, "x1": x1})
    out = []
    for g in groups:
        r = fitz.Rect(g["lines"][0]["bbox"])
        for ln in g["lines"][1:]:
            r |= fitz.Rect(ln["bbox"])
        out.append({"type": 0, "bbox": tuple(r), "lines": g["lines"]})
    return out


def _line_text(ln):
    return "".join(sp["text"] for sp in ln["spans"])


def _line_size(ln):
    return max((sp["size"] for sp in ln["spans"] if sp["text"].strip()), default=0.0)


def _is_bullet(t):
    t = t.strip()
    return len(t) == 1 and (t in BULLETS or t == "-" or unicodedata.category(t) == "So")


def _span_style(sp):
    return (sp["font"], round(sp["size"], 1), sp["color"], sp.get("_deco", ()))


def _parallel_lines(cur, nxt):
    """Deux libelles empiles, pas un paragraphe qui continue : les deux
    lignes commencent au meme endroit par le meme style (« Taux de
    conversion | Payant : » en gras), et le style change exactement au
    passage a la ligne (fin en normal, debut en gras). Un paragraphe qui
    revient a la ligne continue dans le style de la fin de ligne."""
    a = [s for s in cur["spans"] if s["text"].strip()]
    b = [s for s in nxt["spans"] if s["text"].strip()]
    if len(a) < 2 or not b:
        return False
    first_a, last_a, first_b = _span_style(a[0]), _span_style(a[-1]), _span_style(b[0])
    return (first_a == first_b and last_a != first_b
            and abs(cur["bbox"][0] - nxt["bbox"][0]) < 2)


def _hanging_indent(lines):
    """Retrait suspendu d'une liste dont la puce est dans la ligne
    (« ● Texte… » puis lignes suivantes alignees sur « Texte ») : distance
    entre la puce et le debut du texte, ou None. Les lignes sans puce
    doivent toutes tomber sur ce retrait, sinon ce n'est pas une liste."""
    marks, cont = [], []
    for ln in lines:
        spans = [s for s in ln["spans"] if s["text"].strip()]
        if not spans:
            continue
        t0 = spans[0]["text"].lstrip()
        if _is_bullet(t0[0]):
            text_x = None
            chars = [c for c in spans[0].get("chars", ()) if not c["c"].isspace()]
            if len(chars) > 1:
                text_x = chars[1]["bbox"][0]
            elif len(spans) > 1:
                text_x = spans[1]["bbox"][0]
            if text_x is None:
                return None
            marks.append((ln["bbox"][0], text_x))
        else:
            cont.append(ln["bbox"][0])
    if not marks:
        return None
    bx = min(m[0] for m in marks)
    hangs = sorted(tx - b for b, tx in marks)
    h = hangs[len(hangs) // 2]
    size = max((_line_size(ln) for ln in lines), default=10)
    if h <= 0.2 * size or h > 6 * size:
        return None
    if any(abs(c - (bx + h)) > 2.5 for c in cont):
        return None
    return h


SIZE_JUMP = 1.35    # ecart de taille entre deux lignes qui separe deux blocs
INLINE_JUMP = 3.0   # ecart de taille qui separe deux passages d'une meme ligne


def _split_line_sizes(ln):
    """Coupe une ligne entre deux passages de tailles sans commune mesure :
    « Septembre 2026 | Nom » (24 pt) suivi du chiffre decoratif « 02 »
    (300 pt) sur la meme ligne de base. Une fin de titre plus petite
    (« (seulement Magasin) ») reste dans la ligne."""
    spans = ln["spans"]
    parts, cur, last = [], [], None
    for sp in spans:
        if sp["text"].strip():
            if last and max(sp["size"], last) / max(1e-3, min(sp["size"], last)) > INLINE_JUMP:
                parts.append(cur)
                cur = []
            last = sp["size"]
        cur.append(sp)
    parts.append(cur)
    if len(parts) == 1:
        return [ln]
    out = []
    for p in parts:
        boxes = [fitz.Rect(sp["bbox"]) for sp in p if sp["text"].strip()]
        if not boxes:
            continue
        r = boxes[0]
        for bx in boxes[1:]:
            r |= bx
        out.append(dict(ln, spans=p, bbox=tuple(r)))
    return out


def _merge_continuations(cands, shapes):
    """Recolle une phrase que MuPDF a coupee en plusieurs blocs, une ligne
    par bloc (texte centre sous une image : « Preisfindung in Abhaengigkeit
    von » / « Marktbegebenheiten und » / « Mitbewerbern »). Traduits
    separement, ces morceaux donnent des fragments sans syntaxe.

    Deux blocs se suivent si : meme police, taille et couleur ; le second
    commence a un interligne sous le premier ; ils sont alignes (meme
    marge gauche ou meme centre) ; aucun trait ni bord de cadre ne passe
    entre eux ; et la phrase continue (le premier finit sur un mot en
    minuscule sans ponctuation finale, ou le second commence en minuscule).
    """
    def style(ln):
        sp = next((s for s in ln["spans"] if s["text"].strip()), None)
        return (sp["font"], round(sp["size"], 1), sp["color"]) if sp else None

    def continues(a_txt, b_txt):
        a, b = a_txt.rstrip(), b_txt.lstrip()
        if not a or not b or a[-1] in ".:!?;":
            return False
        last = re.findall(r"[^\W\d_]+", a)
        return b[0].islower() or bool(last and last[-1][0].islower())

    order = sorted(range(len(cands)), key=lambda i: (fitz.Rect(cands[i]["bbox"]).y0,
                                                     fitz.Rect(cands[i]["bbox"]).x0))
    alive = set(range(len(cands)))
    for i in order:
        if i not in alive:
            continue
        while True:
            a = cands[i]
            if _rotation(a["lines"][0]) != 0 or a.get("_ocr"):
                break
            la = a["lines"][-1]
            ra, sa = fitz.Rect(la["bbox"]), _line_size(la)
            best = None
            for j in alive:
                if j == i:
                    continue
                b = cands[j]
                lb = b["lines"][0]
                rb = fitz.Rect(lb["bbox"])
                if (_rotation(lb) != 0 or b.get("_ocr") or style(lb) != style(la)
                        or not (0 <= rb.y0 - ra.y1 <= 0.6 * sa)):
                    continue
                same_left = abs(rb.x0 - ra.x0) < 2
                same_mid = abs((rb.x0 + rb.x1) - (ra.x0 + ra.x1)) / 2 < 3
                if not (same_left or same_mid):
                    continue
                if not continues(_line_text(la), _line_text(lb)):
                    continue
                gap = fitz.Rect(min(ra.x0, rb.x0), ra.y1, max(ra.x1, rb.x1), rb.y0)
                if any(s.x1 > gap.x0 and s.x0 < gap.x1 and
                       (gap.y0 - 0.5 <= s.y0 <= gap.y1 + 0.5 or gap.y0 - 0.5 <= s.y1 <= gap.y1 + 0.5)
                       for s in shapes):
                    continue
                best = j
                break
            if best is None:
                break
            b = cands[best]
            r = fitz.Rect(a["bbox"]) | fitz.Rect(b["bbox"])
            cands[i] = dict(a, lines=a["lines"] + b["lines"], bbox=tuple(r))
            alive.discard(best)
    return [cands[i] for i in sorted(alive)]


def _rgb_int(c):
    if isinstance(c, int):
        return c
    c = list(c)[:3] + [0] * (3 - len(c)) if c else [0, 0, 0]
    if len(c) == 1:
        c = c * 3
    return (int(round(c[0] * 255)) << 16) | (int(round(c[1] * 255)) << 8) | int(round(c[2] * 255))


def _near_white(c):
    return all(((c >> s) & 255) > 240 for s in (16, 8, 0))


def _page_decorations(page):
    """Traits fins et aplats qui pourraient souligner, barrer ou surligner
    un passage. PowerPoint et Word les exportent comme des dessins a part,
    poses sous ou sur le texte : ils ne suivent pas une traduction."""
    out = []
    for d in page.get_drawings():
        r = fitz.Rect(d["rect"])
        if r.width < 2:
            continue
        # Aplat invisible (opacite nulle) : PowerPoint en pose derriere
        # chaque zone de texte.
        fill = d.get("fill") if (d.get("fill_opacity") is None
                                 or d.get("fill_opacity") > 0.05) else None
        stroke = d.get("color") if (d.get("stroke_opacity") is None
                                    or d.get("stroke_opacity") > 0.05) else None
        if "f" not in (d.get("type") or ""):
            fill = None
        if "s" not in (d.get("type") or ""):
            stroke = None
        col = fill if fill is not None else stroke
        if col is None:
            continue
        d = dict(d, fill=fill)
        thick = r.height
        if fill is None:
            thick = max(r.height, d.get("width") or 0)
            r = fitz.Rect(r.x0, r.y0 - thick / 2, r.x1, r.y1 + thick / 2)
        out.append({"rect": r, "thick": thick, "color": _rgb_int(col),
                    "fill": d.get("fill") is not None})
    return out


def _tag_decorations(cands, decos):
    """Attache a chaque passage les decorations qui le concernent.

    Trait fin juste sous la ligne de base : souligne ; a mi-hauteur des
    lettres : barre ; aplat ajuste a la hauteur du texte : surligne. Le
    passage (ou les passages contigus) doit couvrir l'essentiel du trait :
    une bordure de tableau, plus large que le texte de la case, n'est pas
    un soulignement, et une case coloree n'est pas un surlignage.
    """
    # On travaille a la lettre : un souligne couvre souvent une partie
    # seulement d'un passage de meme style (« Food Aktion » dans « Wegfall
    # der Planung fuer Food Aktion! »).
    lines = [ln for b in cands for ln in b["lines"]]
    chars = [(ch, sp) for ln in lines for sp in ln["spans"] for ch in sp.get("chars", ())
             if not ch["c"].isspace()]
    if not chars or not decos:
        return
    for d in decos:
        r = d["rect"]
        thin = d["thick"] <= 3.0 and r.height <= 3.5
        # Un trait, pas un point d'une trame de fond : long et allonge.
        if thin and (r.width < 6 or r.width < 4 * max(r.height, d["thick"], 0.5)):
            continue
        if not thin and not (d["fill"] and not _near_white(d["color"])):
            continue
        cy = (r.y0 + r.y1) / 2
        hits = {"u": [], "s": [], "h": []}
        for ch, sp in chars:
            cb = fitz.Rect(ch["bbox"])
            size = sp["size"] or 10
            cx = (cb.x0 + cb.x1) / 2
            if not (r.x0 - 0.5 <= cx <= r.x1 + 0.5):
                continue
            if thin:
                dy = cy - ch["origin"][1]
                if -0.08 * size <= dy <= 0.35 * size:
                    hits["u"].append(ch)
                elif -0.6 * size <= dy <= -0.15 * size:
                    hits["s"].append(ch)
            else:
                ov = min(cb.y1, r.y1) - max(cb.y0, r.y0)
                if ov >= 0.8 * cb.height and r.height <= 1.6 * cb.height:
                    hits["h"].append(ch)
        for kind, lst in hits.items():
            if not lst:
                continue
            x0 = min(ch["bbox"][0] for ch in lst)
            x1 = max(ch["bbox"][2] for ch in lst)
            need = 0.7 if kind in "us" else 0.85
            if x1 - x0 < need * r.width:
                continue
            for ch in lst:
                ch["_deco"] = tuple(sorted(set(ch.get("_deco", ())) | {(kind, d["color"])}))
                ch["_deco_rect"] = fitz.Rect(r)
            break
    for ln in lines:
        ln["spans"] = [part for sp in ln["spans"] for part in _split_span_deco(sp, ln)]


def _split_span_deco(sp, ln):
    """Coupe un passage la ou sa decoration change. Les espaces prennent la
    decoration de la lettre qui les precede, sauf en fin de passage."""
    chars = sp.get("chars")
    if not chars or not any(ch.get("_deco") for ch in chars):
        return [sp]
    keys, last = [], ()
    for i, ch in enumerate(chars):
        if ch["c"].isspace():
            nxt = next((c for c in chars[i + 1:] if not c["c"].isspace()), None)
            k = last if nxt is not None and nxt.get("_deco", ()) == last else ()
        else:
            k = ch.get("_deco", ())
            last = k
        keys.append(k)
    out, start = [], 0
    for i in range(1, len(chars) + 1):
        if i == len(chars) or keys[i] != keys[start]:
            part = chars[start:i]
            boxes = [fitz.Rect(c["bbox"]) for c in part]
            bb = boxes[0]
            for b in boxes[1:]:
                bb |= b
            new = dict(sp, text="".join(c["c"] for c in part), chars=part,
                       origin=part[0]["origin"], bbox=list(bb))
            new.pop("_deco", None)
            new.pop("_deco_rects", None)
            if keys[start]:
                new["_deco"] = keys[start]
                new["_deco_rects"] = list({tuple(c["_deco_rect"]) for c in part
                                           if c.get("_deco_rect") is not None})
                new["_deco_rects"] = [fitz.Rect(t) for t in new["_deco_rects"]]
            out.append(new)
            start = i
    return out


def _invisible_rects(page):
    """Aplats sans encre : opacite nulle et pas de contour visible.
    PowerPoint en pose un derriere chaque zone de texte, a la taille exacte
    du texte ; le journal de peinture les liste comme des formes, et un
    cadre d'ecriture « contenu » dans l'un d'eux ne pouvait plus s'elargir."""
    out = []
    for d in page.get_drawings():
        fo = d.get("fill_opacity")
        so = d.get("stroke_opacity")
        typ = d.get("type") or ""
        fill_seen = "f" in typ and (fo is None or fo > 0.05)
        stroke_seen = "s" in typ and (so is None or so > 0.05)
        if not fill_seen and not stroke_seen:
            out.append(fitz.Rect(d["rect"]))
    return out


def _matches(r, rects, tol=1.0):
    return any(abs(r.x0 - o.x0) < tol and abs(r.y0 - o.y0) < tol and
               abs(r.x1 - o.x1) < tol and abs(r.y1 - o.y1) < tol for o in rects)


def _page_shapes(page):
    """Surfaces dessinees (aplats, traits, images), comme dans widen()."""
    W = page.rect.width
    ghost = _invisible_rects(page)
    out = []
    for kind, r in page.get_bboxlog():
        r = fitz.Rect(r)
        if (kind in ("fill-path", "stroke-path", "fill-image") and r.width < 0.95 * W
                and not (kind == "fill-path" and _matches(r, ghost))):
            out.append(r)
    return out


def _split_reading(block):
    """Separe un bloc en morceaux qui se lisent chacun d'un seul tenant.

    MuPDF regroupe parfois des textes sans rapport : sur une page de titre,
    le grand chiffre « 02 », le titre et la ligne « Septembre 2026 | Nom »,
    dans un ordre qui remonte la page. Recomposes en un paragraphe, ils
    donnent « Septembre 2026 | Nom 02 Titre » sur une ligne. On coupe :
    - avant une ligne placee plus haut que la precedente ;
    - entre deux lignes de tailles tres differentes (titre / sous-titre) ;
    - autour d'une ligne faite d'une seule puce. La puce n'appartient alors
      a aucun bloc : elle reste dessinee a sa place, et le texte de l'element
      commence a sa propre ligne au lieu de remonter a celle de la puce.
    """
    groups, cur, prev = [], [], None
    lines = [part for ln in block["lines"] for part in _split_line_sizes(ln)]
    for ln in lines:
        t = _line_text(ln)
        if not t.strip():
            continue
        if _is_bullet(t):
            if cur:
                groups.append(cur)
            cur, prev = [], None
            continue
        # Debut d'une liste apres un texte sans puce (titre « Points a
        # noter : » puis « • ... ») : la liste forme son propre bloc, avec
        # son retrait suspendu.
        if (cur and _is_bullet(t.lstrip()[:1] or "x")
                and not any(_is_bullet(_line_text(c).lstrip()[:1] or "x") for c in cur)):
            groups.append(cur)
            cur, prev = [], None
        if prev is not None:
            pr, r = fitz.Rect(prev["bbox"]), fitz.Rect(ln["bbox"])
            s0, s1 = _line_size(prev), _line_size(ln)
            above = r.y1 < pr.y0 + 0.5 * pr.height
            jump = s0 > 0 and s1 > 0 and max(s0, s1) / min(s0, s1) > SIZE_JUMP
            if above or jump:
                groups.append(cur)
                cur = []
        cur.append(ln)
        prev = ln
    if cur:
        groups.append(cur)
    if len(groups) == 1 and groups[0] == block["lines"]:
        return [block]
    out = []
    for g in groups:
        r = fitz.Rect(g[0]["bbox"])
        for ln in g[1:]:
            r |= fitz.Rect(ln["bbox"])
        out.append(dict(block, bbox=tuple(r), lines=g))
    return out


def _split_bullets(block, bullets):
    """Un element de liste par bloc quand les puces sont dessinees a part.

    La puce reste a sa place d'origine. Si toute la liste formait un seul
    bloc, un element traduit plus long qu'une ligne decalerait tous les
    suivants, qui ne seraient plus en face de leur puce."""
    lines = block["lines"]
    if len(lines) < 2:
        return [block]
    cuts = [i for i in range(1, len(lines))
            if _has_bullet(lines[i]["bbox"], bullets, _line_size(lines[i]) or 10)]
    if not cuts:
        return [block]
    out, start = [], 0
    for c in cuts + [len(lines)]:
        part = lines[start:c]
        r = fitz.Rect(part[0]["bbox"])
        for ln in part[1:]:
            r |= fitz.Rect(ln["bbox"])
        out.append(dict(block, bbox=tuple(r), lines=part))
        start = c
    return out


def _alignment(lines, rect):
    """Deduit l'alignement de la position des lignes dans le bloc."""
    if len(lines) < 2:
        return "left"
    tol = max(1.5, 0.01 * rect.width)
    x0 = [ln["bbox"][0] for ln in lines]
    x1 = [ln["bbox"][2] for ln in lines]
    mid = [(a + b) / 2 for a, b in zip(x0, x1)]
    left = max(x0) - min(x0) < tol
    right = max(x1) - min(x1) < tol
    # Justifie : la grande majorite des lignes (hors fins de paragraphe)
    # touchent la marge droite.
    inner = x1[:-1]
    full = len(lines) >= 3 and sum(
        1 for v in inner if max(inner) - v < tol) >= 0.75 * len(inner)
    if left and full:
        return "justify"
    if left:
        return "left"
    if right:
        return "right"
    if max(mid) - min(mid) < tol:
        return "center"
    return "left"


def _separator(cur, nxt, cur_w, block_w, room=None, first_w=None, size=10.0):
    """Ce qui relie deux lignes : rien, une espace, ou un vrai retour.

    room : place restee libre a droite de la ligne courante ; first_w :
    largeur du premier mot de la ligne suivante. Si ce mot tenait dans la
    place libre, le retour n'etait pas automatique : il est voulu (element
    de liste, libelle, ligne de tableau), meme sans ponctuation finale.
    """
    a, b = cur.rstrip(), nxt.lstrip()
    if not a or not b:
        return " "
    if b[0] in BULLETS:
        return "\n"
    if a in BULLETS:
        return " "
    if a.endswith("-") and len(a) > 1 and a[-2].isalpha():
        # Cesure allemande : "Ent-" + "wicklung" se recolle sans tiret,
        # "Open-" + "Access" garde son tiret.
        return "\x00" if b[0].islower() else ""
    # Une ligne nettement plus courte que le bloc est une fin voulue
    # (titre sur deux lignes, fin de paragraphe), pas un retour a la ligne
    # automatique.
    if cur_w < 0.6 * block_w and a[-1] in ".:!?;)":
        return "\n"
    if room is not None and first_w is not None:
        # Ligne courte : un retour automatique ne la laisse ainsi que devant
        # un mot vraiment trop long. Ligne longue : marge plus prudente.
        need = 1.1 * first_w if cur_w < 0.6 * block_w else 1.4 * first_w + 0.6 * size
        if room > need:
            return "\n"
    return " "


def _bullet_boxes(page):
    """Puces dessinees a part (un caractere seul : ●, ○, ■, -...) : une
    ligne precedee d'une puce commence un nouvel element de liste."""
    out = []
    for b in page.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            if _is_bullet(_line_text(ln)):
                out.append(fitz.Rect(ln["bbox"]))
    return out


def _has_bullet(line_bbox, bullets, size):
    r = fitz.Rect(line_bbox)
    cy = (r.y0 + r.y1) / 2
    for bb in bullets:
        if bb.y0 - 1 <= cy <= bb.y1 + 1 and bb.x1 <= r.x0 + 1 and r.x0 - bb.x1 < 4 * size:
            return True
    return False


def extract_blocks(doc):
    blocks = []
    n = 0

    for pno in range(doc.page_count):
        page = doc[pno]
        guessed = {tuple(round(v) for v in b["bbox"]): _as_dict_block(b)
                   for b in page.get_text("rawdict", flags=FLAGS_GUESS)["blocks"]
                   if b.get("type") == 0}
        candidates = []
        tainted = set()     # polices qui ont produit du charabia sur la page
        hidden = _hidden_spans(page)
        bullets = _bullet_boxes(page)
        for raw in page.get_text("rawdict", flags=FLAGS_RAW)["blocks"]:
            if raw.get("type") != 0 or not raw["lines"]:
                continue
            b = _as_dict_block(raw)
            if hidden:
                b["lines"] = [ln for ln in b["lines"] if not _is_hidden(ln, hidden)]
            if not b["lines"]:
                continue
            alt = guessed.get(tuple(round(v) for v in raw["bbox"]))
            if alt is not None and (
                    abs(_mean_word_len(_block_text(alt)) - NATURAL_WORD_LEN)
                    <= abs(_mean_word_len(_block_text(b)) - NATURAL_WORD_LEN)):
                b = alt
            # Charabia : la police a une table de caracteres fausse. Le
            # texte est lisible a l'ecran, on le lira donc sur l'image.
            garbled = _garbled(_block_text(b))
            if garbled:
                tainted.update(sp["font"] for ln in b["lines"]
                               for sp in ln["spans"] if sp["text"].strip())
            pieces = (_split_columns(b["lines"])
                      if _rotation(b["lines"][0]) == 0 else [b])
            pieces = [q for pc in pieces
                      for q in (_split_reading(pc)
                                if _rotation(pc["lines"][0]) == 0 else [pc])]
            pieces = [q for pc in pieces
                      for q in (_split_bullets(pc, bullets)
                                if bullets and _rotation(pc["lines"][0]) == 0 else [pc])]
            for pc in pieces:
                pc["_ocr"] = garbled
            candidates.extend(pieces)
        candidates = _merge_continuations(candidates, _page_shapes(page))
        _tag_decorations(candidates, _page_decorations(page))

        # Un bout de charabia trop court pour etre reconnu ("DT`") vient de
        # la meme police fautive : il passe aussi par la lecture visuelle.
        for b in candidates:
            fonts_here = {sp["font"] for ln in b["lines"]
                          for sp in ln["spans"] if sp["text"].strip()}
            if fonts_here and fonts_here <= tainted:
                b["_ocr"] = True

        for b in candidates:
            rot = _rotation(b["lines"][0])
            angle = 0.0
            if rot is None:
                # Texte en biais (pastille, tampon). Le moteur HTML ne sait
                # tourner que par quarts de tour : ce bloc sera redessine
                # ligne par ligne avec insert_text, a l'angle d'origine.
                dx, dy = b["lines"][0]["dir"]
                angle = math.degrees(math.atan2(-dy, dx))
                rot = 0
                # Un effet de contour dessine parfois la meme ligne deux
                # fois au meme endroit : on ne la garde qu'une fois.
                uniq = []
                for ln in b["lines"]:
                    t = "".join(sp["text"] for sp in ln["spans"]).strip()
                    c = fitz.Rect(ln["bbox"]).tl
                    if t and not any(t == u[0] and abs(c - u[1]) < 1.5
                                     for u in uniq):
                        uniq.append((t, c, ln))
                b = dict(b, lines=[u[2] for u in uniq])

            rect = fitz.Rect(b["bbox"])
            lines = [ln for ln in b["lines"]
                     if "".join(s["text"] for s in ln["spans"]).strip()]
            if not lines:
                continue

            # Style de base : celui qui porte le plus de caracteres.
            weight = {}
            for ln in lines:
                for sp in ln["spans"]:
                    if sp["text"].strip():
                        st = Style(sp, _script(sp, ln))
                        weight[st.key()] = weight.get(st.key(), 0) + len(sp["text"])
            if not weight:
                continue
            base_key = max(weight, key=weight.get)

            ids = {base_key: 0}
            styles = {}
            parts = []          # (id de style, texte)
            seps = []           # separateurs entre lignes
            block_w = rect.height if rot in (90, 270) else rect.width

            for i, ln in enumerate(lines):
                for sp in ln["spans"]:
                    if not sp["text"]:
                        continue
                    st = Style(sp, _script(sp, ln))
                    k = st.key() if sp["text"].strip() else base_key
                    if k not in ids:
                        ids[k] = len(ids)
                    styles.setdefault(ids[k], st)
                    parts.append((ids[k], sp["text"]))
                if i + 1 < len(lines):
                    cur = "".join(s["text"] for s in ln["spans"])
                    nxt = "".join(s["text"] for s in lines[i + 1]["spans"])
                    lb = ln["bbox"]
                    cur_w = (lb[3] - lb[1]) if rot in (90, 270) else (lb[2] - lb[0])
                    room = first_w = None
                    if rot == 0:
                        nb = lines[i + 1]["bbox"]
                        words = nxt.split()
                        if words and nxt.strip():
                            room = rect.x1 - lb[2]
                            first_w = (nb[2] - nb[0]) * len(words[0]) / len(nxt.strip())
                    size_here = max((sp["size"] for sp in ln["spans"]), default=10.0)
                    if angle or (rot == 0 and (_has_bullet(lines[i + 1]["bbox"], bullets, size_here)
                                               or _is_bullet(nxt.strip()[:1] or "x")
                                               or _parallel_lines(ln, lines[i + 1]))):
                        sep = "\n"
                    else:
                        sep = _separator(cur, nxt, cur_w, block_w, room, first_w, size_here)
                    seps.append(sep)
                    if sep == "\x00":
                        # retire le tiret de cesure du dernier morceau
                        sid, txt = parts[-1]
                        parts[-1] = (sid, txt.rstrip()[:-1])
                    elif sep:
                        parts.append((0, sep))

            ocr = bool(b.get("_ocr"))
            plain = "".join(t for _, t in parts)
            if not ocr and (not _has_letters(plain) or len(plain.strip()) < 2):
                continue        # nombres, symboles : on laisse l'original

            styles.setdefault(0, Style(lines[0]["spans"][0]))
            if styles[0].key() != base_key:
                for ln in lines:
                    for sp in ln["spans"]:
                        if Style(sp, _script(sp, ln)).key() == base_key:
                            styles[0] = Style(sp, _script(sp, ln))
                            break

            # Fusion des passages consecutifs de meme style, puis balisage.
            merged = []
            for sid, txt in parts:
                if merged and merged[-1][0] == sid:
                    merged[-1] = (sid, merged[-1][1] + txt)
                else:
                    merged.append((sid, txt))
            tagged = "".join(
                txt if sid == 0 or not txt.strip()
                else "<s%d>%s</s%d>" % (sid, txt, sid)
                for sid, txt in merged)
            tagged = re.sub(r"[ \t]{2,}", " ", tagged)

            # Interligne mesure : ecart entre lignes rapporte a la taille.
            size = styles[0].size or 10
            if len(lines) >= 2:
                if rot in (90, 270):
                    gaps = [abs(lines[i + 1]["bbox"][0] - lines[i]["bbox"][0])
                            for i in range(len(lines) - 1)]
                else:
                    gaps = [lines[i + 1]["bbox"][1] - lines[i]["bbox"][1]
                            for i in range(len(lines) - 1)]
                gaps = [g for g in gaps if g > 0]
                lh = (sorted(gaps)[len(gaps) // 2] / size) if gaps else 1.2
                lh = min(max(lh, 0.95), 2.0)
            else:
                lh = 1.15

            if ocr:
                # Le texte viendra de l'image : un seul style, celui qui
                # domine, puisqu'on ne sait pas quels mots il couvrait.
                styles = {0: styles[0]}
                tagged = ""

            n += 1
            blocks.append(Block(
                uid="p%02db%03d" % (pno, n),
                page=pno,
                rect=rect,
                line_rects=[fitz.Rect(ln["bbox"]) for ln in lines],
                styles=styles,
                base=styles[0],
                text=tagged.strip(),
                align=_alignment(lines, rect) if rot == 0 else "left",
                rotate=rot,
                line_height=lh,
                ocr=ocr,
            ))
            first = next((sp for sp in lines[0]["spans"] if sp["text"].strip()), None)
            if first is not None and not ocr:
                blocks[-1].origin = (first["origin"][0], first["origin"][1])
            blocks[-1].hard_lines = bool(seps) and all(x == "\n" for x in seps)
            blocks[-1].decos = [r for ln in lines for sp in ln["spans"]
                                for r in sp.get("_deco_rects", [])]
            if rot == 0 and not angle:
                blocks[-1].hang = _hanging_indent(lines)
                if blocks[-1].hang:
                    # Alignement du texte de la liste, puce mise a part :
                    # les lignes a puce commencent plus a gauche que les
                    # suivantes, ce qui faisait croire a un texte centre.
                    h = blocks[-1].hang
                    shifted = [dict(ln, bbox=(ln["bbox"][0] + h,) + tuple(ln["bbox"][1:]))
                               if _is_bullet(_line_text(ln).lstrip()[:1] or "x") else ln
                               for ln in lines]
                    al = _alignment(shifted, fitz.Rect(rect.x0 + h, rect.y0, rect.x1, rect.y1))
                    blocks[-1].align = al if al in ("left", "justify") else "left"
            if angle:
                blk = blocks[-1]
                blk.angle = angle
                blk.align = "center"
                blk.tilt_lines = [
                    ((fitz.Rect(ln["bbox"]).tl + fitz.Rect(ln["bbox"]).br) / 2,
                     "".join(sp["text"] for sp in ln["spans"]).strip(),
                     max(sp["size"] for sp in ln["spans"]))
                    for ln in lines]

    return blocks


def widen(doc, blocks):
    """Elargit le cadre d'ecriture des libelles d'une ligne.

    Le cadre d'origine colle aux glyphes de la langue source : "Neuf" dans
    une cellule large. Une traduction plus longue ("Neubau") y serait
    reduite ou debordee alors que la cellule a de la place. On etend le
    cadre, dans le sens de l'alignement, jusqu'au premier obstacle : bord
    de la forme qui contient le texte, forme ou image voisine, autre bloc
    de texte, marge de la page. Les paragraphes de plusieurs lignes gardent
    leur largeur, pour ne pas changer leurs coupures de ligne.
    """
    by_page = {}
    for b in blocks:
        by_page.setdefault(b.page, []).append(b)

    for pno, items in by_page.items():
        page = doc[pno]
        W = page.rect.width
        # Journal de peinture plutot que get_drawings : un bandeau de
        # couleur est parfois un trait tres epais, dont get_drawings donne
        # une boite de hauteur nulle. Le journal donne la surface reelle,
        # epaisseur comprise, et inclut les images.
        shapes = []
        # Les soulignes / barres / surlignes suivront le texte traduit : ce
        # ne sont ni des cadres ni des obstacles.
        deco_rects = [fitz.Rect(r) for b in items for r in b.decos]

        def is_deco(r):
            return any(abs(r.x0 - d.x0) < 1.5 and abs(r.x1 - d.x1) < 1.5
                       and r.y0 >= d.y0 - 2 and r.y1 <= d.y1 + 2 for d in deco_rects)

        ghost = _invisible_rects(page)
        for kind, r in page.get_bboxlog():
            r = fitz.Rect(r)
            if (kind not in ("fill-path", "stroke-path", "fill-image")
                    or r.width >= 0.95 * W or r.height <= 0.5 or is_deco(r)
                    or (kind == "fill-path" and _matches(r, ghost))):
                continue
            if kind == "stroke-path":
                # La boite d'un trait inclut une demi-epaisseur au-dela de
                # chaque extremite. Avec des bouts coupes net, le cas
                # courant pour un bandeau, l'encre s'arrete avant : sans
                # cette correction, la fin d'un titre blanc deborde du
                # bandeau et disparait sur le blanc de la page.
                if r.width >= r.height:
                    r = fitz.Rect(r.x0 + r.height / 2, r.y0,
                                  r.x1 - r.height / 2, r.y1)
                else:
                    r = fitz.Rect(r.x0, r.y0 + r.width / 2,
                                  r.x1, r.y1 - r.width / 2)
            shapes.append(r)
        others = [fitz.Rect(b.rect) for b in items]
        # Le texte laisse tel quel (horaires, nombres, sigles) n'a pas de
        # bloc, mais reste un obstacle : « 15:10-15:25 » a droite d'un
        # sous-titre de l'ordre du jour.
        # Au niveau des passages, pas des lignes : MuPDF met parfois dans
        # une meme ligne la date d'une page de titre et le chiffre geant
        # « 04 » ; le cadre de la ligne commence alors en haut de la page
        # et la date n'etait plus vue comme un obstacle sous le titre.
        text_lines = []
        for tb in page.get_text("dict")["blocks"]:
            for ln in tb.get("lines", []):
                for sp in ln["spans"]:
                    if sp["text"].strip():
                        text_lines.append(fitz.Rect(sp["bbox"]))
        others.extend(text_lines)
        H = page.rect.height
        for b in items:
            b.room_bottom = _free_below(b, text_lines, shapes, H)

        for b in items:
            # Libelle d'une ligne, ou pile de lignes toutes terminees par un
            # retour voulu (liste, libelles empiles) : chaque ligne peut
            # s'allonger vers la droite. Un paragraphe qui revient a la
            # ligne automatiquement garde sa largeur.
            if b.rotate != 0 or b.angle or (len(b.line_rects) != 1 and not b.hard_lines):
                continue
            r = fitz.Rect(b.rect)
            h = r.height

            def vmatch(o):
                return min(o.y1, r.y1) - max(o.y0, r.y0) > 0.3 * h

            right, left = W - 18, 18.0
            for o in others:
                if o == r or not vmatch(o):
                    continue
                if o.x0 >= r.x1 - 0.5:
                    right = min(right, o.x0 - 3)
                if o.x1 <= r.x0 + 0.5:
                    left = max(left, o.x1 + 3)
            for o in shapes:
                if not vmatch(o):
                    continue
                if o.x0 <= r.x0 + 1 and o.x1 >= r.x1 - 1:
                    # forme qui contient le texte : on reste dedans
                    right = min(right, o.x1 - 2)
                    left = max(left, o.x0 + 2)
                elif o.x0 >= r.x1 - 0.5:
                    right = min(right, o.x0 - 2)
                elif o.x1 <= r.x0 + 0.5:
                    left = max(left, o.x1 + 2)

            room = fitz.Rect(r)
            if b.align == "right":
                room.x0 = min(r.x0, left)
            elif b.align == "center":
                d = max(0.0, min(r.x0 - left, right - r.x1))
                room.x0, room.x1 = r.x0 - d, r.x1 + d
            else:
                room.x1 = max(r.x1, right)
            b.room = room


def _free_below(b, text_lines, shapes, page_h):
    """Bas de l'espace libre sous un bloc : jusqu'au prochain texte ou
    dessin qui chevauche sa largeur, sans sortir de la forme qui le contient
    (case de tableau, encadre) ni de la marge basse de la page. Un texte
    traduit plus long peut y prendre une ligne de plus au lieu d'etre
    reduit. Renvoie None pour un texte tourne ou en biais."""
    if b.rotate != 0 or b.angle:
        return None
    r = fitz.Rect(b.rect)
    size = b.base.size or 10
    bottom = page_h - 18
    inside = [o for o in shapes if o.x0 <= r.x0 + 1 and o.x1 >= r.x1 - 1
              and o.y0 <= r.y0 + 1 and o.y1 >= r.y1 - 1]
    for o in inside:
        bottom = min(bottom, o.y1 - 0.3 * size)
    for o in list(text_lines) + [s for s in shapes if s not in inside]:
        if o.x1 <= r.x0 + 1 or o.x0 >= r.x1 - 1:
            continue                    # a cote, pas en dessous
        if o.y0 >= r.y1 - 1:
            # On garde au moins un tiers de l'ecart d'origine avec ce qui
            # suit : un titre qui descend ne doit pas coller a la date.
            gap = o.y0 - r.y1
            bottom = min(bottom, o.y0 - max(0.15 * size, gap / 3))
    return max(r.y1, bottom)


def _ink_columns(page, rect, zoom):
    """Colonnes de pixels qui contiennent de l'encre sombre (du texte)."""
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=rect,
                          colorspace=fitz.csGRAY)
    w, h, buf = pix.width, pix.height, pix.samples
    cols = [False] * w
    for r in range(h):
        row = buf[r * w:(r + 1) * w]
        for c in range(w):
            if not cols[c] and row[c] < 110:
                cols[c] = True
    return cols


def _rows(rects):
    """Regroupe des rectangles de ligne en rangees (meme hauteur)."""
    rows = []
    for r in sorted(rects, key=lambda r: r.y0):
        for row in rows:
            u = row[0]
            if min(u.y1, r.y1) - max(u.y0, r.y0) > 0.5 * min(u.height, r.height):
                row[0] = u | r
                break
        else:
            rows.append([fitz.Rect(r)])
    return [row[0] for row in rows]


# Bande blanche minimale, en tailles de police, qui separe deux cellules
# sur l'image. Un espace entre deux mots en fait environ un quart.
CELL_GAP = 1.0


def _split_on_whitespace(page, row, size):
    """Coupe une rangee la ou l'image montre une bande sans encre."""
    zoom = max(2.0, min(8.0, 36.0 / max(size, 1)))
    # Bande centrale seulement : la ligne peut effleurer le filet horizontal
    # d'une cellule, qui mettrait de l'encre dans toutes les colonnes et
    # masquerait les blancs. Le corps des lettres, lui, est au milieu.
    band = fitz.Rect(row.x0, row.y0 + 0.3 * row.height,
                     row.x1, row.y1 - 0.25 * row.height)
    ink = _ink_columns(page, band, zoom)
    min_gap = CELL_GAP * size * zoom
    segs, start, blank = [], None, 0
    for i, dark in enumerate(ink + [False] * int(min_gap + 1)):
        if dark:
            if start is None:
                start = i
            blank = 0
        elif start is not None:
            blank += 1
            if blank >= min_gap:
                segs.append((start, i - blank + 1))
                start, blank = None, 0
    out = []
    for a, b in segs:
        x0 = row.x0 + a / zoom - 0.8
        x1 = row.x0 + b / zoom + 0.8
        out.append(fitz.Rect(max(row.x0, x0), row.y0, min(row.x1, x1), row.y1))
    return out


def split_unreadable_by_cells(doc, blocks):
    """Decoupe les zones a lire sur image en cellules, d'apres l'image.

    Dans une capture d'ecran exportee avec une police fautive, la couche
    texte ment sur tout : les caracteres sont faux et leurs largeurs aussi,
    si bien que le vide entre un libelle et ses montants disparait. MuPDF
    regroupe en outre plusieurs rangees d'un tableau dans un meme bloc.

    On se fie donc a ce que l'oeil voit : chaque rangee est coupee aux
    bandes sans encre d'au moins CELL_GAP tailles de police, puis les
    morceaux de rangees consecutives alignes dans la meme colonne sont
    reunis, pour qu'un libelle sur deux lignes reste un seul bloc et garde
    son contexte a la traduction.
    """
    out = []
    for b in blocks:
        if not b.ocr or b.rotate != 0:
            out.append(b)
            continue
        page = doc[b.page]
        size = b.base.size or 6

        cells = []          # [rect union, [rects de ligne]]
        for row in _rows(b.line_rects):
            for seg in _split_on_whitespace(page, row, size):
                for c in cells:
                    u = c[0]
                    h_overlap = min(u.x1, seg.x1) - max(u.x0, seg.x0)
                    v_gap = seg.y0 - u.y1
                    if (h_overlap > 0.6 * min(u.width, seg.width)
                            # lignes serrees : leurs boites se chevauchent
                            and -0.5 * seg.height < v_gap < 0.8 * seg.height
                            and c[1][-1].y0 < seg.y0):
                        c[0] = u | seg
                        c[1].append(seg)
                        break
                else:
                    cells.append([fitz.Rect(seg), [seg]])

        if len(cells) <= 1:
            out.append(b)
            continue
        for i, (r, lines) in enumerate(cells):
            out.append(Block("%s_%d" % (b.uid, i), b.page, r, lines,
                             b.styles, b.base, "", "left", 0,
                             b.line_height, ocr=True))
    return out


def read_unreadable(doc, blocks, project, location, model, log=print):
    """Transcrit sur l'image les blocs dont la couche texte est du charabia.

    Les blocs sans texte lisible a l'issue de la lecture sont retires : on
    ne touche pas a ce qu'on ne sait pas lire.
    """
    todo = [b for b in blocks if b.ocr]
    if not todo:
        return blocks
    items = []
    for b in todo:
        # Agrandissement pour que la ligne fasse ~36 pixels de haut.
        zoom = max(2.0, min(8.0, 36.0 / max(b.base.size, 1)))
        clip = fitz.Rect(b.rect) + (-1, -1, 1, 1)
        pix = doc[b.page].get_pixmap(matrix=fitz.Matrix(zoom, zoom),
                                     clip=clip)
        items.append((b.uid, pix.tobytes("png")))
    read = transcribe_images(items, project, location, model, log=log)

    kept = []
    for b in blocks:
        if b.ocr:
            txt = read.get(b.uid, "")
            if not _has_letters(txt):
                continue
            b.text = txt
        kept.append(b)
    return kept


# --------------------------------------------------------------------------
# Polices
# --------------------------------------------------------------------------

_PORTABLE = {
    "calibri": ("carlito", "calibri"),
    "arial": ("liberation sans", "arimo", "arial"),
    "narrow": ("liberation sans narrow", "arial narrow"),
    "serif": ("liberation serif", "tinos", "georgia", "dejavu serif"),
    "times": ("liberation serif", "tinos", "times new roman"),
    "sans": ("dejavu sans", "liberation sans", "noto sans", "segoe ui"),
    "mono": ("liberation mono", "cousine", "dejavu sans mono", "courier new"),
}
_BASE14 = {"serif": "ti", "times": "ti", "mono": "co"}
_portable_cache = {}


def _portable_font(family, bold, italic):
    """Octets d'une police de remplacement disponible partout."""
    key = (family, bool(bold), bool(italic))
    if key in _portable_cache:
        return _portable_cache[key]
    buf = None
    try:
        from core.text_fit import _get_index
        idx = _get_index()
        for fam in _PORTABLE.get(family, ()):
            st = idx.get(fam)
            if st:
                path = (st.get((bool(bold), bool(italic))) or st.get((bool(bold), False))
                        or st.get((False, False)))[0]
                with open(path, "rb") as fh:
                    buf = fh.read()
                break
    except Exception:
        buf = None
    if buf is None:
        pre = _BASE14.get(family, "he")
        suffix = {(False, False): "ro" if pre == "ti" else ("ur" if pre == "co" else "lv"),
                  (True, False): "bo", (False, True): "it", (True, True): "bi"}[(bool(bold), bool(italic))]
        name = pre + suffix if not (pre == "he" and suffix == "lv") else "helv"
        try:
            buf = fitz.Font(name).buffer
        except Exception:
            buf = None
    _portable_cache[key] = buf
    return buf


class FontBook:
    """Trouve, pour chaque style, une police qui contient tous les caracteres
    du texte a ecrire, en preferant la police d'origine."""

    def __init__(self, doc):
        self.embedded = {}      # nom sans prefixe de sous-ensemble -> [(buf, Font)]
        self.cache = {}
        self.ink = {}           # (police, caractere) -> dessine vraiment ?
        self.story = {}         # police -> utilisable par insert_htmlbox ?
        seen = set()
        for page in doc:
            for x in page.get_fonts(full=True):
                xref, base = x[0], x[3]
                if xref in seen:
                    continue
                seen.add(xref)
                name = base.split("+")[-1]
                try:
                    buf = doc.extract_font(xref)[3]
                    if not buf:
                        continue
                    font = fitz.Font(fontbuffer=buf)
                except Exception:
                    continue
                self.embedded.setdefault(name, []).append((buf, font))
        self.used = {"embarquee": 0, "systeme": 0, "base": 0}
        self._repair(doc)

    def _repair(self, doc):
        """Polices embarquees sans table de caracteres (export PowerPoint /
        Word) : on les reconstruit, accents compris (voir fontfix). Calcule
        une fois par document, avant que le texte d'origine soit efface."""
        fixed = getattr(doc, "_tradoc_repaired", None)
        if fixed is None:
            fixed = {}
            try:
                from .fontfix import repair_fonts
                fixed = repair_fonts(self.embedded, doc, self._donor_buf, log=print)
            except Exception as exc:
                print("  reparation des polices impossible : %s" % exc)
            try:
                doc._tradoc_repaired = fixed
            except Exception:
                pass
        for name, buf in fixed.items():
            try:
                self.embedded[name] = [(buf, fitz.Font(fontbuffer=buf))] + self.embedded[name]
            except Exception:
                continue

    def _neighbours(self, name):
        """Polices embarquees de la meme famille et de la meme pente, a une
        graisse d'ecart au plus ; a egalite, la plus grasse."""
        from .fontfix import style_of
        fam, weight, italic = style_of(name)
        out = []
        for other in self.embedded:
            f, w, it = style_of(other)
            if other != name and f == fam and it == italic and abs(w - weight) <= 1:
                out.append((abs(w - weight), -w, other))
        return [n for _, _, n in sorted(out)]

    def _donor_buf(self, name):
        """Police systeme de meme graisse et de meme pente, qui fournit les
        accents et signes absents d'une police embarquee."""
        from .fontfix import style_of
        _, weight, italic = style_of(name)
        family = self._family(name)
        bold, semibold = weight >= 7, weight in (5, 6)
        if semibold and family in SEMIBOLD_FONTS and family != "mono":
            fname = SEMIBOLD_FONTS[family][int(italic)]
        else:
            fname = SYSTEM_FONTS[family][int(bold) + 2 * int(italic)]
        path = os.path.join(FONT_DIR, fname)
        try:
            with open(path, "rb") as fh:
                return fh.read()
        except OSError:
            return _portable_font(family, bold or semibold, italic)

    @staticmethod
    def _family(name):
        low = name.lower()
        if "calibri" in low:
            return "calibri"
        if any(k in low for k in ("cond", "narrow", "interstate", "compressed")):
            return "narrow"
        if "arial" in low or "helvetica" in low:
            return "arial"
        if "times" in low:
            return "times"
        if "courier" in low or "mono" in low:
            return "mono"
        if any(k in low for k in ("serif", "georgia", "garamond", "minion",
                                  "cambria", "demos", "roman", "antiqua")):
            if "sans" not in low:
                return "serif"
        return "sans"

    def _draws(self, buf, ch):
        """Le caractere produit-il reellement de l'encre ?

        On ne peut pas se fier a la table de la police : un sous-ensemble
        embarque garde souvent l'entree d'un caractere (has_glyph repond
        oui, la boite englobante est remplie) tout en ayant vide son dessin.
        Le caractere s'affiche alors comme un trou : "D velopper". Seul un
        rendu reel tranche.
        """
        key = (id(buf), ch)
        if key not in self.ink:
            tmp = fitz.open()
            page = tmp.new_page(width=40, height=40)
            try:
                page.insert_font(fontname="T", fontbuffer=buf)
                page.insert_text((6, 30), ch, fontname="T", fontsize=28)
                pix = page.get_pixmap(dpi=72, colorspace=fitz.csGRAY)
                self.ink[key] = any(v < 140 for v in pix.samples)
            except Exception:
                self.ink[key] = False
            tmp.close()
        return self.ink[key]

    def _story_ok(self, buf, font):
        """Le moteur HTML accepte-t-il cette police ?

        insert_htmlbox ne signale rien quand il ne sait pas utiliser une
        police : il ecrit le texte en Noto Serif, sa police interne. Un PDF
        peut embarquer plusieurs variantes d'une meme police, dont certaines
        passent et d'autres non. On ecrit donc un caractere que la police
        possede, et on relit le nom de la police reellement utilisee.
        """
        key = id(buf)
        if key not in self.story:
            probe = next((c for c in "aeinorstAEINORST"
                          if font.has_glyph(ord(c)) and self._draws(buf, c)),
                         None)
            if probe is None:
                # Police de signes (icones) : un de ses propres caracteres.
                try:
                    cps = font.valid_codepoints()
                except Exception:
                    cps = []
                probe = next((chr(cp) for cp in cps[:50]
                              if cp > 32 and self._draws(buf, chr(cp))), None)
            ok = False
            if probe:
                tmp = fitz.open()
                page = tmp.new_page(width=200, height=60)
                arch = fitz.Archive()
                arch.add(buf, "T.font")
                try:
                    page.insert_htmlbox(
                        page.rect,
                        '<p style="font-family:T;font-size:20px">%s</p>' % probe,
                        css="@font-face{font-family:T;src:url(T.font);}",
                        archive=arch)
                    used = {sp["font"] for b in page.get_text("dict")["blocks"]
                            if b.get("type") == 0
                            for ln in b["lines"] for sp in ln["spans"]}
                    ok = bool(used) and not any(
                        u.startswith(("Noto", "Droid", "Charis")) for u in used)
                except Exception:
                    ok = False
                tmp.close()
            self.story[key] = ok
        return self.story[key]

    def _covers(self, buf, font, text):
        if not self._story_ok(buf, font):
            return False
        return all(font.has_glyph(ord(c)) and self._draws(buf, c)
                   for c in set(text)
                   if not c.isspace() and unicodedata.category(c)[0] != "C")

    def pick(self, style, text):
        """Renvoie (octets de police, origine)."""
        for buf, font in self.embedded.get(style.font, []):
            if self._covers(buf, font, text):
                self.used["embarquee"] += 1
                return buf, "embarquee"
        # Police d'origine inutilisable : graisse voisine de la meme famille
        # (LidlFontPro-SemiBold -> LidlFontPro-Bold / -Medium).
        for name in self._neighbours(style.font):
            for buf, font in self.embedded[name]:
                if self._covers(buf, font, text):
                    self.used["embarquee"] += 1
                    return buf, "embarquee"

        family = self._family(style.font)
        if style.semibold and family in SEMIBOLD_FONTS:
            fname = SEMIBOLD_FONTS[family][int(style.italic)]
        else:
            files = SYSTEM_FONTS[family]
            fname = files[int(style.bold) + 2 * int(style.italic)]
        path = os.path.join(FONT_DIR, fname)
        if path not in self.cache:
            try:
                with open(path, "rb") as fh:
                    self.cache[path] = fh.read()
            except OSError:
                self.cache[path] = None
        if self.cache[path]:
            self.used["systeme"] += 1
            return self.cache[path], "systeme"

        # Hors Windows (serveur Linux) : police equivalente installee
        # (Liberation, Carlito, DejaVu...), sinon police incluse dans
        # PyMuPDF, dans la bonne graisse : le gras ne doit pas se perdre.
        buf = _portable_font(family, style.bold or style.semibold, style.italic)
        if buf:
            self.used["systeme"] += 1
            return buf, "systeme"
        self.used["base"] += 1
        return None, "base"


# --------------------------------------------------------------------------
# Reecriture
# --------------------------------------------------------------------------

_WORD_RE = re.compile(r"[^\W\d_]{7,}")


def _hyphenate(text, hyph):
    """Insere des cesures douces dans les mots longs.

    Sans elles, un mot compose allemand qui ne tient pas sur la ligne part
    entier a la ligne suivante, et la justification ecarte enormement les
    mots restants : "Wirtschaftsakteure       und". Le moteur HTML coupe
    aux cesures douces et n'affiche le tiret que si la coupure a lieu.
    """
    return _WORD_RE.sub(lambda m: hyph.inserted(m.group(0), hyphen="\u00ad"),
                        text)


def _hex(c):
    return "#%06x" % (c or 0)


def _split_tags(translated, block):
    """Decoupe la traduction en passages (id de style, texte).

    Si les balises renvoyees ne correspondent pas a celles envoyees, on les
    retire et tout le bloc prend le style de base : on perd le gras local
    plutot que d'appliquer un style au mauvais mot.
    """
    sent = set(int(m) for m in re.findall(r"<s(\d+)>", block.text))
    got = set(int(m) for m in re.findall(r"<s(\d+)>", translated))
    # Une balise orpheline (ouverte sans fermeture, ou mal imbriquee) reste
    # visible une fois les paires correctes retirees.
    leftover = ANY_TAG_RE.search(TAG_RE.sub("", translated)) is not None
    if sent != got or leftover or not got <= set(block.styles):
        return [(0, ANY_TAG_RE.sub("", translated))], bool(sent)

    out, pos = [], 0
    for m in TAG_RE.finditer(translated):
        if m.start() > pos:
            out.append((0, translated[pos:m.start()]))
        out.append((int(m.group(1)), m.group(2)))
        pos = m.end()
    if pos < len(translated):
        out.append((0, translated[pos:]))
    if ANY_TAG_RE.search("".join(t for _, t in out)):
        return [(0, ANY_TAG_RE.sub("", translated))], True
    return out, False


def _render(page, block, pieces, fonts, hyph=None):
    """Compose le HTML du bloc et l'ecrit dans son cadre."""
    # Cesures pour le texte courant compose sur plusieurs lignes, pas pour
    # les titres ni les libelles d'une ligne.
    letters = [c for c in "".join(t for _, t in pieces) if c.isalpha()]
    caps = letters and sum(c.isupper() for c in letters) > 0.7 * len(letters)
    # Les cesures ne servent qu'a eviter les trous du texte justifie. Dans
    # une etiquette alignee a gauche ou a droite, elles ne font qu'ajouter
    # des tirets que l'original n'avait pas ("sou-tenir").
    if (hyph is not None and block.align == "justify" and not caps
            and block.rotate == 0 and block.base.size < 16):
        pieces = [(sid, _hyphenate(txt, hyph)) for sid, txt in pieces]
    arch = fitz.Archive()
    css = ["p{margin:0;padding:0;}"]
    registered = {}

    def family_for(sid, text):
        st = block.styles.get(sid, block.base)
        buf, origin = fonts.pick(st, text)
        if buf is None:
            return ("Helvetica", st)
        key = id(buf)
        if key not in registered:
            fam = "F%d" % len(registered)
            arch.add(buf, fam + ".font")
            css.append("@font-face{font-family:%s;src:url(%s.font);}"
                       % (fam, fam))
            registered[key] = fam
        return (registered[key], st)

    # Texte de chaque style, pour verifier la couverture des glyphes.
    by_style = {}
    for sid, txt in pieces:
        by_style[sid] = by_style.get(sid, "") + txt

    base_fam, base_st = family_for(0, by_style.get(0, "") or "a")
    fams = {0: (base_fam, base_st)}
    for sid in by_style:
        if sid != 0:
            fams[sid] = family_for(sid, by_style[sid])

    def deco_css(st):
        lines = [w for k_, w in (("u", "underline"), ("s", "line-through")) if st.has(k_) is not None]
        return "text-decoration:%s;" % " ".join(lines) if lines else ""

    def span_html(sid, txt, k, marks, extra=""):
        esc = html.escape(txt).replace("\n", "<br>")
        if sid == 0 and not extra:
            return esc
        fam, st = fams[sid]
        shift = {"sub": "vertical-align:sub;",
                 "sup": "vertical-align:super;"}.get(st.script, "")
        return ('<span style="font-family:%s;font-size:%.2fpx;color:%s;%s%s%s'
                'font-weight:normal;font-style:normal">%s</span>'
                % (fam, st.size * k, _hex(marks.get(sid, st.color)), shift,
                   deco_css(st), extra, esc))

    # Liste a puces avec retrait suspendu : un paragraphe par element, la
    # puce dans la marge, le texte qui commence au retrait d'origine.
    paragraphs = None
    if getattr(block, "hang", None):
        paragraphs, cur = [], []
        for sid, txt in pieces:
            parts = txt.split("\n")
            for i, part in enumerate(parts):
                if i:
                    paragraphs.append(cur)
                    cur = []
                if part:
                    cur.append((sid, part))
        paragraphs.append(cur)
        paragraphs = [p for p in paragraphs if "".join(t for _, t in p).strip()]

    bullet_w = {}

    def _advance(st, text, size):
        buf, _ = fonts.pick(st, text)
        try:
            f = fitz.Font(fontbuffer=buf) if buf else fitz.Font("helv")
            return f.text_length(text, fontsize=size)
        except Exception:
            return 0.5 * size * len(text)

    def bullet_gap(sid, ch, k, marks):
        """Espaces insecables apres la puce pour que le texte tombe au
        retrait d'origine. Le moteur HTML ignore letter-spacing et les
        marges en ligne : on met n espaces a la taille du texte, puis un
        dernier plus petit pour l'appoint (il n'agrandit pas la ligne)."""
        key = (sid, ch, round(k, 3))
        if key not in bullet_w:
            st = fams[sid][1]
            gap = max(0.0, block.hang * k - _advance(st, ch, st.size * k))
            size = base_st.size * k
            sp_w = _advance(base_st, " ", size) or 0.25 * size
            n = int(gap // sp_w)
            rest = gap - n * sp_w
            bullet_w[key] = (n, size * rest / sp_w if rest > 0.3 else 0.0)
        n, small = bullet_w[key]
        out = span_html(0, " " * n, k, marks, " ") if n else ""
        if small:
            fam = fams[0][0]
            out += ('<span style="font-family:%s;font-size:%.2fpx"> </span>'
                    % (fam, small))
        return out

    def p_open(k, marks, extra=""):
        return ('<p style="font-family:%s;font-size:%.2fpx;color:%s;%s'
                'line-height:%.2f;text-align:%s;font-weight:normal;'
                'font-style:normal;%s">'
                % (base_fam, base_st.size * k, _hex(marks.get(0, base_st.color)),
                   deco_css(base_st), block.line_height, block.align, extra))

    def make_html(k, marks=None):
        """HTML du bloc, tailles multipliees par k. marks : {id de style:
        couleur} pour reperer ou tombent les passages surlignes."""
        marks = marks or {}
        if paragraphs is None:
            body = "".join(span_html(sid, txt, k, marks) for sid, txt in pieces)
            return p_open(k, marks) + body + "</p>"
        out = []
        hang = block.hang * k
        for para in paragraphs:
            items = list(para)
            # premier caractere visible : une puce ?
            j = next((i for i, (_, t) in enumerate(items) if t.strip()), None)
            first = items[j][1].lstrip() if j is not None else ""
            if first and _is_bullet(first[0]):
                sid = items[j][0]
                rest = first[1:].lstrip()
                head = span_html(sid, first[0], k, marks) + bullet_gap(sid, first[0], k, marks)
                items = items[:j] + ([(sid, rest)] if rest else []) + items[j + 1:]
                body = head + "".join(span_html(s, t, k, marks) for s, t in items[j:])
                style = "padding-left:%.2fpx;text-indent:-%.2fpx" % (hang, hang)
            else:
                body = "".join(span_html(s, t, k, marks) for s, t in items)
                style = "padding-left:%.2fpx" % hang
            out.append(p_open(k, marks, style) + body + "</p>")
        return "".join(out)

    css_text = "\n".join(css)

    # Un peu d'air : le cadre d'origine colle aux glyphes allemands.
    r = fitz.Rect(block.room)
    if block.rotate in (90, 270):
        r.y0 -= 2
        r.y1 += 2
    else:
        if block.room == block.rect:
            r.x1 += 2       # un peu d'air seulement si le cadre n'a pas ete
                            # deja ajuste a un contenant par widen()
        r.y1 += 0.25 * base_st.size

    # Echelle : on ne se fie pas a la reponse d'insert_htmlbox, qui annonce
    # parfois que tout tient (marge 0, echelle 1) alors qu'il a laisse
    # tomber la derniere ligne : « ... vers « » sans la suite. On mesure
    # la hauteur reelle du texte dans un cadre tres haut, a la largeur du
    # cadre, et on cherche la plus grande taille qui tient.
    rot = block.rotate in (90, 270)
    box_w, box_h = (r.height, r.width) if rot else (r.width, r.height)
    # Le cadre d'origine colle aux lettres ; une ligne composee est un peu
    # plus haute (interligne). Sans cette tolerance, un libelle d'une
    # ligne qui tient etait reduit de quelques pour cent. Elle reste bien
    # en dessous d'une ligne : une ligne en trop est toujours detectee.
    tol = (max(block.line_height, 1.0) - 1.0 + 0.45) * base_st.size
    box_h += tol
    k, fits = _fit_scale(make_html, css_text, arch, box_w, box_h)
    multi = len(block.line_rects) > 1
    if (k < 0.99 and (multi or k < 0.8) and not rot and block.room_bottom
            and block.room_bottom > r.y1 + 1):
        # Espace libre sous le bloc : une ligne de plus plutot qu'un texte
        # plus petit que l'original. Un titre d'une ligne reste sur une
        # ligne tant que la reduction est legere.
        k2, fits2 = _fit_scale(make_html, css_text, arch, box_w,
                               block.room_bottom - r.y0)
        if k2 > k + 0.02:
            k, fits = k2, fits2
            r = fitz.Rect(r.x0, r.y0, r.x1, block.room_bottom)
    doc_html = make_html(k)

    r = _align_first_line(page, block, r, doc_html, css_text, arch)
    # Un peu de marge en bas : la mesure garantit deja que le texte tient,
    # cette marge evite seulement qu'un arrondi fasse tomber une ligne.
    slack = 0.6 * base_st.size * k
    if block.rotate == 0:
        r = fitz.Rect(r.x0, r.y0, r.x1, r.y1 + slack)
    elif block.rotate == 90:
        r = fitz.Rect(r.x0, r.y0, r.x1 + slack, r.y1)
    elif block.rotate == 270:
        r = fitz.Rect(r.x0 - slack, r.y0, r.x1, r.y1)
    else:
        r = fitz.Rect(r.x0, r.y0 - slack, r.x1, r.y1)
    _draw_highlights(page, block, r, make_html, k, css_text, arch)
    page.insert_htmlbox(r, doc_html, css=css_text, archive=arch,
                        scale_low=0, rotate=block.rotate)
    return k, fits


def _draw_highlights(page, block, r, make_html, k, css_text, arch):
    """Surlignage : le moteur HTML ne sait pas peindre un fond derriere un
    passage. On compose le bloc sur une page brouillon en donnant a chaque
    passage surligne une couleur temoin, on releve ou tombent ces passages,
    et on peint les aplats sur la vraie page avant d'y ecrire le texte."""
    hl = {sid: st.has("h") for sid, st in block.styles.items() if st.has("h") is not None}
    if not hl or block.rotate != 0:
        return
    marks = {sid: 0x010203 + 7 * i for i, sid in enumerate(hl)}
    tmp = fitz.open()
    try:
        pg = tmp.new_page(width=page.rect.width, height=page.rect.height)
        pg.insert_htmlbox(r, make_html(k, marks), css=css_text, archive=arch, scale_low=0)
        back = {v: hl[sid] for sid, v in marks.items()}
        for b in pg.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                for sp in ln["spans"]:
                    if sp["color"] in back and sp["text"].strip():
                        c = back[sp["color"]]
                        rgb = (((c >> 16) & 255) / 255, ((c >> 8) & 255) / 255, (c & 255) / 255)
                        page.draw_rect(fitz.Rect(sp["bbox"]), color=None, fill=rgb, width=0)
    finally:
        tmp.close()


def _natural_height(html_text, css_text, arch, width):
    """Hauteur occupee par le HTML compose a la largeur donnee."""
    tall = 20000.0
    tmp = fitz.open()
    try:
        pg = tmp.new_page(width=width + 10, height=tall + 10)
        spare, _ = pg.insert_htmlbox(fitz.Rect(0, 0, width, tall), html_text,
                                     css=css_text, archive=arch, scale_low=1)
        return tall - spare if spare >= 0 else float("inf")
    finally:
        tmp.close()


def _fit_scale(make_html, css_text, arch, width, height):
    """(echelle, tient) : la plus grande echelle <= 1 a laquelle le texte
    tient dans le cadre. Sous MIN_SCALE, on continue a reduire pour ne
    rien perdre, mais le bloc est signale comme debordant."""
    def ok(k):
        return _natural_height(make_html(k), css_text, arch, width) <= height + 0.5
    if ok(1.0):
        return 1.0, True
    lo, hi, fits = MIN_SCALE, 1.0, True
    if not ok(MIN_SCALE):
        lo, hi, fits = 0.25, MIN_SCALE, False
        if not ok(lo):
            return lo, False
    for _ in range(7):
        mid = (lo + hi) / 2
        if ok(mid):
            lo = mid
        else:
            hi = mid
    return lo, fits


def _align_first_line(page, block, r, doc_html, css_text, arch):
    """Decale le cadre pour que la premiere ligne reecrite tombe sur la
    ligne d'origine.

    Le moteur HTML place la premiere ligne un peu plus bas (et un peu plus
    a droite) que le texte d'origine : 2 points sur du 9 pt, 4 sur un titre
    de 28 pt. Un soulignement ou un filet dessine sous le texte se retrouve
    alors en travers des lettres. On rend le bloc sur une page brouillon,
    on mesure ou tombe le premier caractere, et on corrige.
    """
    if block.origin is None or block.rotate != 0:
        return r
    try:
        tmp = fitz.open()
        scratch = tmp.new_page(width=page.rect.width, height=page.rect.height)
        spare, _ = scratch.insert_htmlbox(r, doc_html, css=css_text, archive=arch,
                                          scale_low=MIN_SCALE)
        if spare < 0:
            tmp.close()
            return r
        got = None
        for b in scratch.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                for sp in ln["spans"]:
                    if sp["text"].strip():
                        got = sp["origin"]
                        break
                if got:
                    break
            if got:
                break
        tmp.close()
        if got is None:
            return r
        size = block.base.size or 10
        dy = block.origin[1] - got[1]
        dx = block.origin[0] - got[0] if block.align in ("left", "justify") else 0.0
        if abs(dy) > 0.6 * size:
            dy = 0.0        # le texte a change de ligne : pas de recalage
        if abs(dx) > 0.6 * size:
            dx = 0.0
        return fitz.Rect(r.x0 + dx, r.y0 + dy, r.x1 + dx, r.y1 + dy)
    except Exception:
        return r


def _drawing_key(d):
    r = d["rect"]
    return (round(r.x0), round(r.y0), round(r.x1), round(r.y1),
            len(d["items"]))


def _replay(page, d):
    """Redessine un trace tel que get_drawings() l'a decrit."""
    shape = page.new_shape()
    for it in d["items"]:
        op = it[0]
        if op == "l":
            shape.draw_line(it[1], it[2])
        elif op == "re":
            shape.draw_rect(it[1])
        elif op == "qu":
            shape.draw_quad(it[1])
        elif op == "c":
            shape.draw_bezier(it[1], it[2], it[3], it[4])
    kind = d.get("type") or "f"
    cap = d.get("lineCap") or 0
    if isinstance(cap, (tuple, list)):     # (debut, fin, tirets) selon versions
        cap = cap[0]
    shape.finish(
        fill=d.get("fill") if "f" in kind else None,
        color=d.get("color") if "s" in kind else None,
        width=d.get("width") or 1,
        lineCap=int(cap),
        lineJoin=int(d.get("lineJoin") or 0),
        dashes=d.get("dashes") or None,
        even_odd=bool(d.get("even_odd")),
        closePath=bool(d.get("closePath")))
    shape.commit()


def _still_visible(page, sp):
    """Le texte se voit-il encore ? On cherche sa couleur dans le rendu."""
    c = sp["color"] or 0
    rgb = ((c >> 16) & 255, (c >> 8) & 255, c & 255)
    pix = page.get_pixmap(dpi=144, clip=fitz.Rect(sp["bbox"]),
                          colorspace=fitz.csRGB, alpha=False)
    buf, n = pix.samples, 0
    for i in range(0, len(buf), 3):
        if (abs(buf[i] - rgb[0]) + abs(buf[i + 1] - rgb[1])
                + abs(buf[i + 2] - rgb[2])) < 90:
            n += 1
            if n >= 4:
                return True
    return False


def _reemit_text(page, rect, fonts, done, hidden=()):
    """Reecrit a l'identique le texte d'origine reste dans une zone.

    Quand une forme est repeinte, le texte non traduit qui etait pose
    dessus (un montant dans une ligne de tableau grisee) se retrouve
    dessous et disparait. On le redessine par-dessus, avec sa police, sa
    taille, sa couleur et sa position d'origine.
    """
    for b in page.get_text("dict", clip=rect)["blocks"]:
        for ln in b.get("lines", []):
            dx, dy = ln["dir"]
            angle = math.degrees(math.atan2(-dy, dx))
            for sp in ln["spans"]:
                if not sp["text"].strip():
                    continue
                if _is_hidden({"spans": [sp], "bbox": sp["bbox"]}, hidden):
                    continue        # un vestige cache doit le rester
                if _still_visible(page, sp):
                    continue        # pas recouvert : le reecrire l'epaissirait
                key = (round(sp["origin"][0], 1), round(sp["origin"][1], 1),
                       sp["text"])
                if key in done:
                    continue
                done.add(key)
                buf, _ = fonts.pick(Style(sp), sp["text"])
                name = "R%d" % len(done)
                if buf:
                    page.insert_font(fontname=name, fontbuffer=buf)
                else:
                    name = "helv"
                c = sp["color"] or 0
                color = (((c >> 16) & 255) / 255, ((c >> 8) & 255) / 255,
                         (c & 255) / 255)
                o = fitz.Point(sp["origin"])
                page.insert_text(o, sp["text"], fontname=name,
                                 fontsize=sp["size"], color=color,
                                 morph=(o, fitz.Matrix(angle)) if angle else None)


def _restore_colours(page, before, images):
    """Repeint les formes dont la couleur a change pendant la suppression.

    apply_redactions retire le texte en laissant les dessins, mais quand un
    texte et l'aplat qui est dessous partagent un reglage de couleur, ce
    reglage part avec le texte : une cellule d'en-tete verte ressort noire.

    Repeindre la forme seule la poserait PAR-DESSUS tout ce qui avait ete
    dessine apres elle : le carre bleu d'un logo recouvrirait ses lettres.
    On repeint donc la forme puis, dans l'ordre d'origine, chaque trace qui
    la recouvrait. Si une image ou un trace semi-transparent est au-dessus,
    on renonce : une image ne se rejoue pas ainsi, et une transparence
    rejouee s'assombrit. Mieux vaut une couleur fausse qu'un logo masque.
    """
    after = {}
    for d in page.get_drawings():
        after.setdefault(_drawing_key(d), []).append(d)

    def changed(d):
        cands = after.get(_drawing_key(d), [])
        return bool(cands) and d.get("fill") and not any(
            a.get("fill") and max(abs(x - y) for x, y in
                                  zip(d["fill"], a["fill"])) <= 0.05
            for a in cands)

    fixed = skipped = 0
    areas = []              # surfaces repeintes, ou le texte doit revenir
    for d in before:
        if not changed(d):
            continue
        r = fitz.Rect(d["rect"])
        stack = [d] + [e for e in before
                       if e["seqno"] > d["seqno"]
                       and fitz.Rect(e["rect"]).intersects(r)]
        transparent = any((e.get("fill_opacity") or 1) < 0.999
                          or (e.get("stroke_opacity") or 1) < 0.999
                          for e in stack)
        if transparent or any(r.intersects(im) for im in images):
            skipped += 1
            continue
        for e in stack:
            _replay(page, e)
            areas.append(fitz.Rect(e["rect"]))
        fixed += 1
    return fixed, skipped, areas


def _norm(text):
    return re.sub(r"\s+", " ", ANY_TAG_RE.sub("", text or "")).strip()


# Un espace normal mesure environ 0,28 fois la taille de police. Etalonne
# sur une colonne justifiee etroite : sans cesure les ecarts montent a 1,0
# fois la taille (trous visibles), avec cesures ils restent sous 0,72.
HOLE_RATIO = 0.85


def _holes(page, block):
    """Trous dans le texte redessine : ecart entre mots anormalement large.

    L'ecart est rapporte a la taille de police REELLEMENT rendue, qui peut
    etre plus petite que l'originale si le bloc a ete reduit pour tenir.
    """
    clip = fitz.Rect(block.room) + (0, 0, 3, block.base.size or 10)
    # Taille de chaque ligne : la zone peut meler un titre et un sous-titre,
    # et rapporter l'ecart du titre a la taille du sous-titre fabrique de
    # faux trous. Les numeros (bloc, ligne) de "words" et de "dict"
    # designent les memes lignes.
    blocks = page.get_text("dict", clip=clip)["blocks"]
    line_size = {}
    for bi, b in enumerate(blocks):
        for li, ln in enumerate(b.get("lines", [])):
            sz = [sp["size"] for sp in ln["spans"] if sp["text"].strip()]
            if sz:
                line_size[(bi, li)] = max(sz)
    lines = {}
    for w in page.get_text("words", clip=clip):
        lines.setdefault((w[5], w[6]), []).append(w)
    worst = (0.0, "")
    for key, ws in lines.items():
        size = line_size.get(key)
        if not size:
            continue
        ws.sort(key=lambda w: w[0])
        for a, b in zip(ws, ws[1:]):
            ratio = (b[0] - a[2]) / size
            if ratio > worst[0]:
                worst = (ratio, " ".join(w[4] for w in ws))
    if worst[0] > HOLE_RATIO:
        return "%s (ecart %.1f x la taille)" % (
            worst[1].replace("­", ""), worst[0])
    return None


def _balance(words, n):
    """Repartit des mots en n lignes de longueurs proches."""
    total = sum(len(w) for w in words) + len(words)
    lines, cur = [], []
    for w in words:
        cur.append(w)
        if len(lines) < n - 1 and len(" ".join(cur)) >= total / n:
            lines.append(" ".join(cur))
            cur = []
    lines.append(" ".join(cur))
    return [l for l in lines if l] or [""]


def _render_tilted(page, block, translated, fonts):
    """Redessine un bloc en biais, ligne par ligne, a l'angle d'origine.

    Chaque ligne traduite est centree sur le centre de la ligne source,
    dans la meme police, couleur et inclinaison. Si la traduction est
    nettement plus longue, la taille baisse juste assez pour garder la
    largeur d'origine plus 10 %.
    """
    src = block.tilt_lines
    text = ANY_TAG_RE.sub("", translated).replace("\u00ad", "")
    tgt = [t.strip() for t in text.split("\n") if t.strip()]
    if len(tgt) != len(src):
        tgt = _balance(" ".join(tgt).split(), len(src))
        tgt += [""] * (len(src) - len(tgt))

    a = math.radians(block.angle)
    u = fitz.Point(math.cos(a), -math.sin(a))      # sens de lecture
    n = fitz.Point(math.sin(a), math.cos(a))       # vers le bas des lettres
    c = block.base.color or 0
    color = (((c >> 16) & 255) / 255, ((c >> 8) & 255) / 255, (c & 255) / 255)
    scale = 1.0

    for i, ((center, src_txt, size), txt) in enumerate(zip(src, tgt)):
        if not txt:
            continue
        buf, _ = fonts.pick(block.base, txt)
        if buf:
            name = "T%s%d" % (block.uid.replace("_", ""), i)
            page.insert_font(fontname=name, fontbuffer=buf)
            font = fitz.Font(fontbuffer=buf)
        else:
            name, font = "helv", fitz.Font("helv")
        w0 = font.text_length(src_txt, fontsize=size)
        w = font.text_length(txt, fontsize=size)
        if w0 and w > 1.1 * w0:
            scale = min(scale, 1.1 * w0 / w)
            size *= 1.1 * w0 / w
            w = font.text_length(txt, fontsize=size)
        origin = center - u * (w / 2) + n * (0.35 * size)
        page.insert_text(origin, txt, fontname=name, fontsize=size,
                         color=color, morph=(origin, fitz.Matrix(block.angle)))
    return scale


def rewrite(doc, blocks, translations, log=print, hyph=None):
    stats = {"written": 0, "shrunk": 0, "overflow": [], "tag_fallback": 0,
             "recoloured": 0, "unchanged": 0, "issues": []}
    fonts = FontBook(doc)

    by_page = {}
    for b in blocks:
        new = translations.get(b.uid, "")
        if not new.strip():
            continue
        # Un texte que la traduction laisse identique (nom propre, sigle,
        # nombre, libelle deja dans la langue cible) n'est PAS redessine :
        # les glyphes d'origine restent, avec leur police et leur position
        # exactes. Le redessiner ne pourrait que l'abimer.
        if _norm(new) == _norm(b.text):
            stats["unchanged"] += 1
            continue
        by_page.setdefault(b.page, []).append(b)

    for pno, items in sorted(by_page.items()):
        page = doc[pno]
        before = page.get_drawings()
        hidden = _hidden_spans(page)
        images = [fitz.Rect(i["bbox"]) for i in page.get_image_info()]
        for b in items:
            for lr in b.line_rects:
                page.add_redact_annot(lr, fill=False)
        page.apply_redactions(
            images=fitz.PDF_REDACT_IMAGE_NONE,
            graphics=fitz.PDF_REDACT_LINE_ART_NONE,
            text=fitz.PDF_REDACT_TEXT_REMOVE)
        fixed, skipped, areas = _restore_colours(page, before, images)
        # Le texte d'origine reste pose sur les surfaces repeintes (un
        # montant dans une ligne grisee) est passe dessous : on le reecrit
        # par-dessus, une fois que tout est repeint.
        done = set()
        for area in areas:
            _reemit_text(page, area, fonts, done, hidden)
        stats["recoloured"] += fixed
        stats["recolour_skipped"] = stats.get("recolour_skipped", 0) + skipped

        # Soulignes, barres et surlignes d'origine : ils sont redessines
        # avec le texte traduit, sous les bons mots. On n'efface que le
        # dessin entierement couvert par la zone de la decoration.
        decos = [r for b in items if not b.angle for r in b.decos]
        if decos:
            for r in decos:
                page.add_redact_annot(fitz.Rect(r) + (-0.6, -0.6, 0.6, 0.6), fill=False)
            page.apply_redactions(
                images=fitz.PDF_REDACT_IMAGE_NONE,
                graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                text=fitz.PDF_REDACT_TEXT_NONE)
            stats["decorations"] = stats.get("decorations", 0) + len(decos)

        for b in items:
            if b.angle:
                _render_tilted(page, b, translations[b.uid], fonts)
                stats["written"] += 1
                continue
            pieces, fell_back = _split_tags(translations[b.uid], b)
            stats["tag_fallback"] += int(fell_back)
            scale, fits = _render(page, b, pieces, fonts, hyph)
            stats["written"] += 1
            dst = _norm(translations[b.uid])
            if scale < 0.999:
                stats["shrunk"] += 1
            if scale < 0.75:
                stats["issues"].append({"page": pno + 1, "kind": "fortement reduit",
                                        "text": "%s (x%.2f)" % (dst[:70], scale)})
            if not fits:
                stats["overflow"].append({
                    "uid": b.uid, "page": pno + 1,
                    "src": _norm(b.text)[:60], "dst": dst[:60]})
                stats["issues"].append({"page": pno + 1, "kind": "deborde",
                                        "text": dst[:80]})
            hole = _holes(page, b)
            if hole:
                stats["issues"].append({"page": pno + 1, "kind": "trous",
                                        "text": hole[:80]})

    stats["fonts"] = fonts.used
    return stats


_NUM_RE = re.compile(r"\d+(?:[.,:/]\d+)*")


def keeps_information(full, short):
    """Une version courte ne doit rien perdre : memes passages stylés
    (un mot en rouge ou en gras porte souvent l'information), memes
    nombres et dates, memes retours a la ligne, et presque tous les mots."""
    tags = lambda t: sorted(re.findall(r"<s(\d+)>", t))  # noqa: E731
    if tags(full) != tags(short):
        return False
    if sorted(_NUM_RE.findall(full)) != sorted(_NUM_RE.findall(short)):
        return False
    if not enough_words(full, short):
        return False
    if new_abbreviations(full, short):
        return False
    return full.count("\n") == short.count("\n")


_ABBR_RE = re.compile(r"(?<![\w.])([^\W\d_]{1,7})\.(?=[\s/|,)]|$)")


def new_abbreviations(full, short):
    """Mots abreges par un point ("Com.", "Planif.") absents de la version
    complete : la version courte a invente des abreviations."""
    plain = lambda t: re.sub(r"<[^>]+>", " ", t)  # noqa: E731
    have = set(_ABBR_RE.findall(plain(full)))
    return [w for w in _ABBR_RE.findall(plain(short)) if w not in have]


def enough_words(full, short):
    """Un libelle court (6 mots ou moins) ne perd aucun mot ("Mer 18.11.26"
    ne devient pas "18.11.26") ; un texte plus long en garde au moins 60 %."""
    words = lambda t: re.findall(r"[^\W\d_]+", re.sub(r"<[^>]+>", " ", t))  # noqa: E731
    n_full, n_short = len(words(full)), len(words(short))
    if n_full <= 6:
        return n_short >= n_full
    return n_short >= 0.6 * n_full


def tighten(doc, blocks, translations, target, project, location, model,
            hyph=None, log=print):
    """Mesure chaque bloc traduit et fait raccourcir ceux qui ne tiennent pas.

    La consigne de longueur du prompt ne suffit pas : le modele ne voit pas
    la largeur reelle du cadre ni la chasse de la police. On met donc en
    page chaque traduction sur une page brouillon ; si elle doit etre
    reduite sous SHORTEN_BELOW, on redemande une version d'au plus N
    caracteres, N etant deduit de la reduction mesuree. La version courte
    n'est gardee que si elle tient mieux.
    """
    fonts = FontBook(doc)
    scratch = fitz.open()
    pages = {}

    def measure(b, text):
        # PyMuPDF detache les objets page deja ouverts des qu'on ajoute une
        # page au document : on garde le numero et on rouvre la page.
        if b.page not in pages:
            r = doc[b.page].rect
            scratch.new_page(width=r.width, height=r.height)
            pages[b.page] = scratch.page_count - 1
        pieces, _ = _split_tags(text, b)
        scale, fits = _render(scratch[pages[b.page]], b, pieces, fonts, hyph)
        return scale if fits else scale * 0.5

    todo, scales = [], {}
    for b in blocks:
        new = translations.get(b.uid, "")
        if not new.strip() or _norm(new) == _norm(b.text) or b.angle:
            continue
        sc = measure(b, new)
        if sc < SHORTEN_BELOW and len(_norm(new)) > 6:
            limit = max(4, int(len(_norm(new)) * sc * 0.95))
            todo.append((b.uid, new, limit, b.context_label()))
            scales[b.uid] = sc
    if not todo:
        scratch.close()
        return 0

    log("    %d blocs trop serres : demande de versions plus courtes"
        % len(todo))
    short = shorten(todo, target, project, location, model, log=log)
    by_uid = {b.uid: b for b in blocks}
    better = 0
    for uid, txt in short.items():
        if not keeps_information(translations[uid], txt):
            continue
        if measure(by_uid[uid], txt) > scales[uid] + 0.03:
            translations[uid] = txt
            better += 1
    scratch.close()
    log("    %d blocs raccourcis" % better)
    return better


def _language_issues(blocks, translations, src_code, tgt_code):
    """Texte encore dans la langue source apres traduction."""
    from .check_pdf import _lang_score
    out = []
    for b in blocks:
        new = _norm(translations.get(b.uid, ""))
        if len(new.split()) < 4:
            continue
        s, t = _lang_score(new, src_code), _lang_score(new, tgt_code)
        if s >= 0.08 and s > t:
            kind = ("non traduit" if new == _norm(b.text)
                    else "reste en langue source")
            out.append({"page": b.page + 1, "kind": kind, "text": new[:80]})
    return out


def translate_pdf(in_path, out_path, source, target, domain, glossary,
                  project, location="global", model=DEFAULT_MODEL,
                  log=print, cache_path=None, codes=("de", "fr")):
    doc = fitz.open(in_path)
    blocks = extract_blocks(doc)
    tagged = sum(1 for b in blocks if not b.uniform_style)
    unreadable = sum(1 for b in blocks if b.ocr)
    log("  %d blocs sur %d pages, dont %d a styles melanges, %d illisibles"
        % (len(blocks), doc.page_count, tagged, unreadable))
    if unreadable:
        blocks = split_unreadable_by_cells(doc, blocks)
        blocks = read_unreadable(doc, blocks, project, location, model, log)
    widen(doc, blocks)

    try:
        import pyphen
        hyph = pyphen.Pyphen(lang={"de": "de_DE", "fr": "fr_FR",
                                   "en": "en_US"}.get(codes[1], codes[1]),
                             left=3, right=3)
    except Exception:
        hyph = None

    translations, call_stats = translate_document(
        blocks, source, target, domain, glossary, project, location,
        model, name=os.path.basename(in_path), log=log,
        cache_path=cache_path)

    flagged = []
    for b in blocks:
        new = translations.get(b.uid, b.text)
        fl = qa_flags(b, new)
        if fl:
            flagged.append({"uid": b.uid, "flags": fl,
                            "src": b.text, "dst": new})

    tighten(doc, blocks, translations, target, project, location, model,
            hyph, log)
    stats = rewrite(doc, blocks, translations, log=log, hyph=hyph)
    stats["issues"] = _language_issues(blocks, translations, *codes) \
        + stats["issues"]
    try:
        doc.subset_fonts()
    except Exception:
        pass
    doc.save(out_path, garbage=4, deflate=True)
    doc.close()

    log("  polices : %(embarquee)d d'origine, %(systeme)d de remplacement, "
        "%(base)d de secours" % stats["fonts"])
    if stats["tag_fallback"]:
        log("  %d blocs ont perdu leur style local (balises non rendues)"
            % stats["tag_fallback"])

    stats["blocks"] = len(blocks)
    stats["flagged"] = flagged
    stats["calls"] = call_stats
    return stats