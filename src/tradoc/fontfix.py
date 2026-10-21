"""Reparation des polices embarquees dans un PDF.

Un PDF exporte par PowerPoint ou Word embarque ses polices TrueType sous
forme de sous-ensembles « CID » : les dessins des lettres sont la, mais la
table qui relie un caractere a son dessin (cmap) a ete retiree. Pour un
logiciel qui veut ecrire du texte, la police semble vide, et le traducteur
se rabattait sur une autre police (Segoe UI, DejaVu...) aux largeurs
differentes : textes plus larges, retours a la ligne decales, chevauchements.

Ce module reconstruit une police utilisable :

1. Table des caracteres : chaque caractere affiche dans le PDF donne son
   numero de dessin (page.get_texttrace). Toutes les lettres de l'original
   redeviennent disponibles.
2. Lettres absentes du sous-ensemble (les accents francais dans un document
   allemand, et inversement) :
   - lettre accentuee : lettre de base de la police + accent. L'accent est
     pris dans la police elle-meme quand une autre lettre le porte deja
     (le trema de « ä » sert pour « ë »), sinon dans une police systeme,
     a la bonne echelle et a la bonne hauteur ;
   - lettre de base absente : prise dans une autre graisse de la meme
     famille (penchee pour une italique), sinon dans la police systeme ;
   - ponctuation absente (« », ’, …) : police systeme, a l'echelle.

Les lettres de la police d'origine ne sont jamais modifiees.
"""
import io
import unicodedata

import fitz
from fontTools.ttLib import TTFont, newTable
from fontTools.ttLib.tables._c_m_a_p import CmapSubtable
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.pens.cu2quPen import Cu2QuPen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.boundsPen import BoundsPen
from fontTools.pens.recordingPen import RecordingPen, DecomposingRecordingPen, replayRecording

# Caracteres a garantir : ASCII, Latin-1, et la typographie courante en
# francais et en allemand.
REPERTOIRE = ("".join(chr(c) for c in range(0x20, 0x7F))
              + "".join(chr(c) for c in range(0xA0, 0x100))
              + "ŒœŸ‘’‚“”„–—…€•‹›  ")

# Diacritique combinant -> forme isolee (presente dans presque toutes les polices)
MARKS = {"́": "´", "̀": "`", "̂": "ˆ", "̈": "¨",
         "̧": "¸", "̃": "˜", "̊": "˚", "̌": "ˇ"}
BELOW = {"̧"}

ITALIC_SHEAR = 0.2      # ~11 degres : pente d'une italique de labeur
MIN_OWN_LETTERS = 16    # minuscules propres (sur 26) pour reparer une police

_WEIGHTS = [("thin", 1), ("extralight", 2), ("ultralight", 2), ("light", 3), ("book", 4),
            ("regular", 4), ("normal", 4), ("medium", 5), ("semibold", 6), ("demibold", 6),
            ("demi", 6), ("extrabold", 8), ("ultrabold", 8), ("bold", 7), ("black", 9),
            ("heavy", 9)]


def style_of(name):
    """« LidlFontPro-SemiBoldItalic » -> (« lidlfontpro », 6, True)."""
    fam, _, sty = name.partition("-")
    low = sty.lower()
    italic = "italic" in low or "oblique" in low
    weight = 4
    for key, w in _WEIGHTS:
        if key in low:
            weight = w
            break
    return fam.lower(), weight, italic


# --------------------------------------------------------------------------
# Contours
# --------------------------------------------------------------------------

def _bounds(rec):
    bp = BoundsPen(None)
    replayRecording(rec, bp)
    return bp.bounds


def _transform(rec, t):
    out = RecordingPen()
    replayRecording(rec, TransformPen(out, t))
    return out.value


def _contours(rec):
    out, cur = [], []
    for op in rec:
        cur.append(op)
        if op[0] in ("closePath", "endPath"):
            out.append(cur)
            cur = []
    if cur:
        out.append(cur)
    return out


class Source:
    """Une police dont on peut lire les dessins : sfnt (TrueType / OpenType)
    ou CFF nu (polices internes de PyMuPDF)."""

    def __init__(self, buf, cmap=None):
        self.fitz = fitz.Font(fontbuffer=buf)
        self.cff = None
        try:
            self.tt = TTFont(io.BytesIO(buf), lazy=False)
            self.upm = self.tt["head"].unitsPerEm
            self.gs = self.tt.getGlyphSet()
            self.order = self.tt.getGlyphOrder()
            if cmap is None:
                cmap = {}
                if "cmap" in self.tt:
                    for cp, gname in self.tt.getBestCmap().items():
                        cmap[chr(cp)] = gname
            else:
                cmap = {c: self.order[g] for c, g in cmap.items() if 0 <= g < len(self.order)}
            self.adv = {g: self.tt["hmtx"][g][0] for g in self.order}
        except Exception:
            # CFF nu : noms de glyphes standard Adobe
            from fontTools.cffLib import CFFFontSet
            from fontTools import agl
            self.tt = None
            cff = CFFFontSet()
            cff.decompile(io.BytesIO(buf), None)
            top = cff[cff.fontNames[0]]
            self.gs = top.CharStrings
            self.upm = int(round(1 / top.FontMatrix[0])) if top.FontMatrix[0] else 1000
            cmap = {}
            for gname in self.gs.keys():
                u = agl.toUnicode(gname)
                if len(u) == 1:
                    cmap.setdefault(u, gname)
            self.adv = None
        self.cmap = cmap
        self._rec = {}

    def has(self, ch):
        g = self.cmap.get(ch)
        if g is None:
            return False
        if ch.isspace():
            return True
        rec = self.record(g)
        return bool(rec)

    def record(self, gname):
        if gname not in self._rec:
            rec = DecomposingRecordingPen(self.gs)
            try:
                self.gs[gname].draw(rec)
                self._rec[gname] = rec.value
            except Exception:
                self._rec[gname] = []
        return self._rec[gname]

    def glyph(self, ch):
        return self.record(self.cmap[ch])

    def advance(self, ch):
        g = self.cmap[ch]
        if self.adv is not None:
            return self.adv[g]
        return self.fitz.glyph_advance(ord(ch)) * self.upm

    def height(self, chars):
        """Hauteur du premier caractere disponible (hauteur d'x, de capitale)."""
        for c in chars:
            if self.has(c):
                b = _bounds(self.glyph(c))
                if b:
                    return b[3]
        return None


# --------------------------------------------------------------------------
# Reconstruction
# --------------------------------------------------------------------------

class Repair:
    def __init__(self, name, target, siblings, donor):
        self.name = name
        self.t = target                     # Source de la police a reparer
        self.fam, self.weight, self.italic = style_of(name)
        # Autres graisses de la famille, de la plus proche a la plus lointaine
        # Au-dela d'une graisse d'ecart, la difference se voit dans le mot
        # (un « q » gras au milieu d'un texte normal) : la police systeme de
        # la bonne graisse est alors plus discrete.
        sib = []
        for sname, src in siblings:
            fam, w, it = style_of(sname)
            if fam != self.fam or sname == name or abs(w - self.weight) > 1:
                continue
            if it == self.italic:
                sib.append((abs(w - self.weight), 0.0, src))
            elif self.italic and not it:
                sib.append((abs(w - self.weight) + 0.5, ITALIC_SHEAR, src))
        sib.sort(key=lambda x: x[0])
        self.siblings = [(shear, src) for _, shear, src in sib]
        self.donor = donor
        self.added = {}                     # caractere -> (contours, avance)
        self.origin = {}                    # caractere -> provenance (journal)
        xh_t = target.height("xvwzunmo")
        xh_d = donor.height("xvwzunmo") if donor else None
        k = (xh_t / target.upm) / (xh_d / donor.upm) if xh_t and xh_d else 1.0
        self.k = max(0.8, min(1.25, k))     # taille relative de la police systeme

    # -- lettre de base -----------------------------------------------------
    def base(self, ch):
        """(contours, avance) dans les unites de la police cible, ou None."""
        if ch in self.added:
            return self.added[ch]
        if self.t.has(ch):
            return self.t.glyph(ch), self.t.advance(ch)
        for shear, src in self.siblings:
            if src.has(ch):
                s = self.t.upm / src.upm
                return (_transform(src.glyph(ch), (s, 0, shear * s, s, 0, 0)),
                        src.advance(ch) * s)
        return None

    def _dotless(self, ch):
        """« i » / « j » sans leur point, pour y poser un accent."""
        dl = {"i": "ı", "j": "ȷ"}.get(ch)
        if dl and self.base(dl):
            return self.base(dl)
        got = self.base(ch)
        if not got:
            return None
        rec, adv = got
        parts = _contours(rec)
        b = [_bounds(p) for p in parts]
        if len(parts) < 2 or not all(b):
            return got
        # Le point est le contour le plus haut
        top = max(range(len(parts)), key=lambda i: b[i][1])
        keep = [op for i, p in enumerate(parts) if i != top for op in p]
        return keep, adv

    # -- accent -------------------------------------------------------------
    def _own_mark(self, mark, upper):
        """Accent deja present dans la police (ou une graisse soeur) : pris
        dans une lettre composee qui le porte (« ä » = « a » + trema).
        Renvoie (contours de l'accent, ecart vertical a la lettre, source)."""
        for shear, src in [(0.0, self.t)] + self.siblings:
            if src.tt is None or shear:
                continue
            glyf = src.tt["glyf"]
            for c, gname in src.cmap.items():
                d = unicodedata.normalize("NFD", c)
                if len(d) != 2 or d[1] != mark or d[0].isupper() != upper:
                    continue
                g = glyf[gname]
                if not g.isComposite() or d[0] not in src.cmap:
                    continue
                base_g = src.cmap[d[0]]
                comps = [x for x in g.components if x.glyphName != base_g]
                if len(comps) != 1 or len(g.components) != 2:
                    continue
                cmp_ = comps[0]
                rec = _transform(src.record(cmp_.glyphName), (1, 0, 0, 1, cmp_.x, cmp_.y))
                bb = _bounds(src.glyph(d[0]))
                mb = _bounds(rec)
                if not (bb and mb):
                    continue
                s = self.t.upm / src.upm
                rec = _transform(rec, (s, 0, 0, s, 0, 0))
                gap = (mb[1] - bb[3]) * s if mark not in BELOW else (bb[1] - mb[3]) * s
                return rec, gap
        return None

    def _donor_mark(self, ch, mark):
        """Accent pris dans la police systeme, a l'echelle de la police cible.
        L'ecart a la lettre est celui de la lettre accentuee de la police
        systeme (« é » comparee a « e »)."""
        d = self.donor
        sp = MARKS.get(mark)
        base = unicodedata.normalize("NFD", ch)[0]
        if not d or not sp or not d.has(sp):
            return None
        s = self.t.upm / d.upm * self.k
        rec = _transform(d.glyph(sp), (s, 0, 0, s, 0, 0))
        mb = _bounds(rec)
        if not mb:
            return None
        gap = None
        if d.has(ch) and d.has(base):
            pb, bb = _bounds(d.glyph(ch)), _bounds(d.glyph(base))
            if pb and bb:
                if mark in BELOW:
                    # bas de l'accent sous la lettre
                    gap = (bb[1] - pb[1]) * s - (mb[3] - mb[1])
                else:
                    gap = (pb[3] - bb[3]) * s - (mb[3] - mb[1])
        if gap is None:
            gap = 0.08 * self.t.upm
        return rec, gap

    def compose(self, ch):
        d = unicodedata.normalize("NFD", ch)
        if len(d) != 2 or d[1] not in MARKS:
            return None
        b, mark = d[0], d[1]
        got = self._dotless(b) if b in "ij" and mark not in BELOW else self.base(b)
        if not got:
            return None
        brec, adv = got
        bb = _bounds(brec)
        if not bb:
            return None
        m = self._own_mark(mark, b.isupper()) or self._donor_mark(ch, mark)
        if not m:
            return None
        mrec, gap = m
        mb = _bounds(mrec)
        if mark in BELOW:
            dy = bb[1] - gap - mb[3]
        else:
            # pour i/j sans point : poser l'accent au-dessus de la hauteur d'x
            top = bb[3]
            if b in "ij":
                top = max(top, self.t.height("xvwzunmo") or top)
            dy = top + gap - mb[1]
        cx_b = (bb[0] + bb[2]) / 2
        cy_m = (mb[1] + mb[3]) / 2 + dy
        cy_b = (bb[1] + bb[3]) / 2
        shear = ITALIC_SHEAR if self.italic else 0.0
        dx = cx_b - (mb[0] + mb[2]) / 2 + shear * (cy_m - cy_b)
        return brec + _transform(mrec, (1, 0, 0, 1, dx, dy)), adv

    def whole(self, ch):
        """Caractere entier de la police systeme, a l'echelle."""
        d = self.donor
        if not d or not d.has(ch):
            return None
        # La police systeme est choisie dans la meme pente : pas de penche ici.
        s = self.t.upm / d.upm * self.k
        return _transform(d.glyph(ch), (s, 0, 0, s, 0, 0)), d.advance(ch) * s

    def space(self, ch):
        sp = self.base(" ")
        w = sp[1] if sp else 0.25 * self.t.upm
        return [], (w * 0.5 if ch in "  " else w)

    def run(self, chars=REPERTOIRE):
        for ch in chars:
            if self.t.has(ch):
                continue
            how = None
            if ch in (" ", " ", " ", " "):
                got, how = self.space(ch), "espace"
            else:
                # Lettre accentuee : d'abord la lettre de base et l'accent de
                # la police elle-meme (« Ö » = « O » + trema de « Ä »).
                got, how = self.compose(ch), "lettre + accent"
                if not got:
                    got, how = self.base(ch), "graisse voisine"
                if not got:
                    got, how = self.whole(ch), "police systeme"
            if got:
                self.added[ch] = got
                self.origin[ch] = how
        return self


def build_font(src_buf, cmap_gid, repair):
    """Police TrueType complete : table des caracteres + lettres ajoutees."""
    tt = TTFont(io.BytesIO(src_buf))
    order = list(tt.getGlyphOrder())
    glyf, hmtx = tt["glyf"], tt["hmtx"]
    mapping = {}
    for ch, gid in cmap_gid.items():
        if 0 <= gid < len(order):
            mapping[ord(ch)] = order[gid]
    for ch, (rec, adv) in repair.added.items():
        name = "uni%04X" % ord(ch)
        pen = TTGlyphPen(None)
        replayRecording(rec, Cu2QuPen(pen, max_err=1.0, reverse_direction=False))
        g = pen.glyph()
        g.recalcBounds(glyf)
        glyf[name] = g
        if name not in order:
            order.append(name)
        hmtx[name] = (int(round(adv)), int(getattr(g, "xMin", 0) or 0))
        mapping[ord(ch)] = name
    tt.setGlyphOrder(order)
    glyf.glyphOrder = order

    cm = newTable("cmap")
    cm.tableVersion = 0
    cm.tables = []
    bmp = {cp: g for cp, g in mapping.items() if cp <= 0xFFFF}
    for pid, eid in ((0, 3), (3, 1)):
        st = CmapSubtable.newSubtable(4)
        st.platformID, st.platEncID, st.language = pid, eid, 0
        st.cmap = dict(bmp)
        cm.tables.append(st)
    tt["cmap"] = cm

    if "post" not in tt:
        post = newTable("post")
        post.formatType = 3.0
        post.italicAngle = -11.0 if repair.italic else 0.0
        post.underlinePosition, post.underlineThickness = -100, 50
        post.isFixedPitch = 0
        post.minMemType42 = post.maxMemType42 = post.minMemType1 = post.maxMemType1 = 0
        tt["post"] = post
    if "name" not in tt:
        nm = newTable("name")
        nm.names = []
        fam, _, sty = repair.name.partition("-")
        for nid, val in ((1, fam), (2, sty or "Regular"), (4, repair.name), (6, repair.name)):
            nm.setName(val, nid, 3, 1, 0x409)
        tt["name"] = nm
    out = io.BytesIO()
    tt.save(out)
    return out.getvalue()


def char_maps(doc):
    """{nom de police: {caractere: numero de dessin}} lus dans le texte."""
    maps = {}
    for page in doc:
        try:
            trace = page.get_texttrace()
        except Exception:
            continue
        for sp in trace:
            m = maps.setdefault(sp["font"].split("+")[-1], {})
            for c in sp["chars"]:
                u, g = c[0], c[1]
                if u > 0 and g > 0 and chr(u) not in m:
                    m[chr(u)] = g
    return maps


def repair_fonts(embedded, doc, donor_for, log=None):
    """embedded : {nom: [(octets, fitz.Font)]}. Renvoie {nom: octets reparés}
    pour chaque police sans table de caracteres utilisable.

    donor_for(nom) -> octets d'une police systeme de meme graisse / pente.
    """
    maps = char_maps(doc)
    sources = {}
    for name, lst in embedded.items():
        buf, font = lst[0]
        cmap = maps.get(name)
        if not cmap or font.has_glyph(ord("a")) or font.has_glyph(ord("e")):
            continue            # police deja utilisable, ou rien a reconstruire
        try:
            src = Source(buf, cmap)
            if src.tt is None or "glyf" not in src.tt:
                continue
            sources[name] = (buf, cmap, src)
        except Exception:
            continue
    out = {}
    siblings = [(n, s[2]) for n, s in sources.items()]
    donors = {}
    for name, (buf, cmap, src) in sources.items():
        try:
            dbuf = donor_for(name)
            if dbuf not in donors:
                donors[dbuf] = Source(dbuf) if dbuf else None
            rep = Repair(name, src, siblings, donors[dbuf]).run()
            # Une police dont presque toutes les lettres viendraient de la
            # police systeme (icones, police citee pour deux caracteres) ne
            # recoit pas de lettres : la police systeme seule est plus
            # homogene. Elle recoit quand meme sa table de caracteres, pour
            # que ses propres signes (icones) restent utilisables.
            own = sum(1 for c in "abcdefghijklmnopqrstuvwxyz"
                      if src.has(c) or rep.origin.get(c) == "graisse voisine")
            if own < MIN_OWN_LETTERS:
                rep.added, rep.origin = {}, {}
            out[name] = build_font(buf, cmap, rep)
            if log:
                kinds = {}
                for how in rep.origin.values():
                    kinds[how] = kinds.get(how, 0) + 1
                log("  police %s reparee : %d caracteres d'origine, %s"
                    % (name, len(cmap), ", ".join("%d %s" % (v, k) for k, v in kinds.items())))
        except Exception as exc:
            if log:
                log("  police %s non reparee : %s" % (name, exc))
    return out