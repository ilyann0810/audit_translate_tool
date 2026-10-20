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
                 "semibold")

    def __init__(self, span, script=None):
        self.script = script    # "sub", "sup" ou None
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
                self.script, self.semibold)


class Block:
    """Un bloc de texte et tout ce qu'il faut pour le redessiner."""

    __slots__ = ("uid", "page", "rect", "line_rects", "styles", "base",
                 "text", "align", "rotate", "line_height", "ocr", "room",
                 "angle", "tilt_lines",
                 "category", "style", "box", "runs", "uniform_style")

    def __init__(self, uid, page, rect, line_rects, styles, base, text,
                 align, rotate, line_height, ocr=False):
        self.ocr = ocr              # texte a lire sur l'image
        self.room = rect            # cadre d'ecriture, eventuellement elargi
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

    def colour(j):
        kind, d = log[j][0], draw.get(j)
        if kind == "fill-image" or d is None:
            return None                     # image : couleur inconnue
        if kind == "fill-path" and (d.get("fill_opacity") or 1) > 0.99:
            return d.get("fill")
        if kind == "stroke-path" and (d.get("stroke_opacity") or 1) > 0.99:
            return d.get("color")
        return None

    def opaque(j):
        kind, d = log[j][0], draw.get(j)
        if kind == "fill-image":
            return tuple(round(v) for v in log[j][1]) not in masked
        return (kind == "fill-path" and d is not None and d.get("fill")
                and (d.get("fill_opacity") or 1) > 0.99)

    out = []
    for sp in page.get_texttrace():
        i, r = sp["seqno"], fitz.Rect(sp["bbox"])
        if r.is_empty or i >= len(log):
            continue
        text = "".join(chr(c[0]) for c in sp["chars"])
        hidden = sp.get("type") == 3 or log[i][0] == "ignore-text"
        if not hidden:
            covered = sum((fitz.Rect(log[j][1]) & r).get_area()
                          for j in range(i + 1, len(log))
                          if opaque(j) and fitz.Rect(log[j][1]).intersects(r))
            hidden = covered > 0.25 * r.get_area()
        if not hidden:
            mid = (r.tl + r.br) / 2
            under = [j for j in range(i)
                     if log[j][0] in ("fill-path", "stroke-path", "fill-image")
                     and fitz.Rect(log[j][1]).contains(mid)]
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


def _separator(cur, nxt, cur_w, block_w):
    """Ce qui relie deux lignes : rien, une espace, ou un vrai retour."""
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
    return " "


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
            for pc in pieces:
                pc["_ocr"] = garbled
            candidates.extend(pieces)

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
                    sep = ("\n" if angle
                           else _separator(cur, nxt, cur_w, block_w))
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
        for kind, r in page.get_bboxlog():
            r = fitz.Rect(r)
            if (kind not in ("fill-path", "stroke-path", "fill-image")
                    or r.width >= 0.95 * W or r.height <= 0.5):
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

        for b in items:
            if b.rotate != 0 or b.angle or len(b.line_rects) != 1:
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
    body = []
    for sid, txt in pieces:
        esc = html.escape(txt).replace("\n", "<br>")
        if sid == 0:
            body.append(esc)
            continue
        fam, st = family_for(sid, by_style[sid])
        shift = {"sub": "vertical-align:sub;",
                 "sup": "vertical-align:super;"}.get(st.script, "")
        body.append(
            '<span style="font-family:%s;font-size:%.1fpx;color:%s;%s'
            'font-weight:normal;font-style:normal">%s</span>'
            % (fam, st.size, _hex(st.color), shift, esc))

    doc_html = (
        '<p style="font-family:%s;font-size:%.1fpx;color:%s;'
        'line-height:%.2f;text-align:%s;font-weight:normal;'
        'font-style:normal">%s</p>'
        % (base_fam, base_st.size, _hex(base_st.color), block.line_height,
           block.align, "".join(body)))
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

    spare, scale = page.insert_htmlbox(
        r, doc_html, css=css_text, archive=arch,
        scale_low=MIN_SCALE, rotate=block.rotate)
    if spare >= 0:
        return scale, True

    # insert_htmlbox n'ecrit RIEN quand le texte ne tient pas a l'echelle
    # minimale. On force l'ecriture sans limite plutot que de perdre le
    # texte, et on signale le debordement.
    spare, scale = page.insert_htmlbox(
        r, doc_html, css=css_text, archive=arch,
        scale_low=0, rotate=block.rotate)
    return scale, False


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