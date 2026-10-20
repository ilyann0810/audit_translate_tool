"""
Traduction DOCX en gardant la forme et le style.

Chaque paragraphe (corps, tableaux, zones de texte, en-tetes, pieds de
page, notes, commentaires) devient un segment balise (core/segments.py) :

  - les runs de meme apparence sont regroupes ; chaque groupe garde son
    rPr COMPLET (couleur, surlignage, taille, police, style de caractere),
    recopie tel quel sur le texte traduit ;
  - ce qui n'est pas du texte reste en place sous forme d'objet : images,
    zones de texte, champs (instruction de champ, TOC, numero de page),
    renvois de notes et de commentaires, symboles, texte supprime en mode
    revision, formules ;
  - les liens hypertexte et les insertions en mode revision restent des
    conteneurs : le texte traduit est remis dedans ;
  - les signets places en debut ou en fin de paragraphe y restent.

Les paragraphes imbriques (zone de texte dans un paragraphe) sont des
segments a part : l'objet qui les contient est deplace, jamais recopie,
ce qui garde leur traduction.
"""
import copy
import os
import re
import shutil
import tempfile
import time

import docx
from lxml import etree

from core.segments import Segment
from core.text_utils import load_glossary_df
from core.translation_api import LANG, PRICE_IN, PRICE_OUT, translate_segments

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

LOCALE = {"fr": "fr-FR", "de": "de-DE"}

# Parties a traduire (fin du type de contenu).
_TEXT_PARTS = ("document.main+xml", "header+xml", "footer+xml", "footnotes+xml",
               "endnotes+xml", "comments+xml")

# Conteneurs dont le contenu est du texte du paragraphe.
_CONTAINERS = {W + "hyperlink", W + "ins", W + "moveTo", W + "smartTag", W + "customXml"}

# Marqueurs sans largeur.
_MARKERS = {W + "bookmarkStart", W + "bookmarkEnd", W + "commentRangeStart",
            W + "commentRangeEnd", W + "permStart", W + "permEnd",
            W + "moveFromRangeStart", W + "moveFromRangeEnd", W + "moveToRangeStart",
            W + "moveToRangeEnd"}

# Contenu de run qui n'est pas du texte a traduire.
_RUN_TEXT = {W + "t"}
_RUN_DROP = {W + "lastRenderedPageBreak"}

_NEUTRAL_RPR = {W + "lang", W + "noProof", W + "rtl"}


def _style_key(rpr):
    if rpr is None:
        return "-"
    el = copy.deepcopy(rpr)
    for child in list(el):
        if child.tag in _NEUTRAL_RPR:
            el.remove(child)
    for a in list(el.attrib):
        del el.attrib[a]
    return etree.tostring(el, encoding="unicode")


class _Builder:
    """Construit le segment d'un paragraphe en parcourant son contenu."""

    def __init__(self, p):
        self.seg = Segment()
        self.seg.meta["p"] = p
        self.lead, self.trail = [], []
        self.pending_markers = []
        self.pending_objects = []
        self.seen_text = False
        self.in_field = False       # entre fldChar begin et separate

    def _flush_objects(self):
        # Les objets consecutifs forment un seul objet, sauf s'ils ne sont
        # pas dans le meme conteneur (lien, insertion...).
        group, chain = [], None
        for el, ch in self.pending_objects:
            if group and ch != chain:
                self.seg.add_object((group, chain))
                group = []
            group.append(el)
            chain = ch
        if group:
            self.seg.add_object((group, chain))
        self.pending_objects = []

    def _flush_markers(self, at_end=False):
        if not self.pending_markers:
            return
        if not self.seen_text:
            self.lead.extend(el for el, _ in self.pending_markers)
        elif at_end:
            self.trail.extend(el for el, _ in self.pending_markers)
        else:
            self.pending_objects.extend(self.pending_markers)
        self.pending_markers = []

    def object(self, el, containers=()):
        self._flush_markers()
        self.pending_objects.append((el, tuple(containers)))

    def text(self, txt, rpr, containers):
        if not txt:
            return
        self._flush_markers()
        self._flush_objects()
        self.seen_text = True
        key = (_style_key(rpr), tuple(id(c) for c in containers))
        self.seg.add_text(txt, key, (rpr, containers))
        if containers:
            sid = self.seg.items[-1][1]
            self.seg.meta.setdefault("no_base", set()).add(sid)

    def brk(self, kind, el=None, rpr=None, containers=()):
        self._flush_markers()
        self._flush_objects()
        if kind == "br":
            self.seg.meta.setdefault("br", el)
        self.seg.add_break(kind)
        # le style du retour suit le texte voisin
        self.seg.meta.setdefault("break_rpr", rpr)

    def run(self, r, containers):
        rpr = r.find(W + "rPr")
        children = [c for c in r if c.tag != W + "rPr"]
        # Run de champ ou d'objet pur : on garde le run d'origine entier.
        simple_obj = all(c.tag not in _RUN_TEXT and c.tag not in (W + "tab", W + "br", W + "cr")
                         for c in children)
        has_fld = any(c.tag in (W + "fldChar", W + "instrText") for c in children)
        for c in children:
            if c.tag == W + "fldChar":
                t = c.get(W + "fldCharType")
                if t == "begin":
                    self.in_field = True
                elif t in ("separate", "end"):
                    self.in_field = False
        if self.in_field and not has_fld:
            # Instruction de champ sur plusieurs runs : objet.
            self.object(r, containers)
            return
        if simple_obj or has_fld:
            if children and all(c.tag in _RUN_DROP for c in children):
                return
            self.object(r, containers)
            return
        for c in children:
            tag = c.tag
            if tag == W + "t":
                self.text(c.text or "", rpr, containers)
            elif tag == W + "tab":
                self.brk("tab", rpr=rpr)
            elif tag in (W + "br", W + "cr"):
                if c.get(W + "type") in ("page", "column"):
                    self.object(self._wrap(c, rpr), containers)
                else:
                    self.brk("br", el=c, rpr=rpr)
            elif tag == W + "noBreakHyphen":
                self.text("‑", rpr, containers)
            elif tag == W + "softHyphen" or tag in _RUN_DROP:
                continue
            else:
                self.object(self._wrap(c, rpr), containers)

    @staticmethod
    def _wrap(child, rpr):
        # Le run d'emballage n'est cree qu'a l'ecriture : l'extraction ne
        # doit rien modifier, un paragraphe peut rester tel quel.
        return ("wrap", child, rpr)

    def walk(self, parent, containers=()):
        for child in parent:
            tag = child.tag
            if tag == W + "pPr":
                continue
            if tag == W + "r":
                self.run(child, containers)
            elif tag in _CONTAINERS:
                self.walk(child, containers + (child,))
            elif tag == W + "sdt":
                content = child.find(W + "sdtContent")
                if content is not None and content.find(W + "r") is not None \
                        and content.find(W + "p") is None:
                    self.walk(content, containers + (child,))
                else:
                    self.object(child, containers)
            elif tag in _MARKERS:
                self.pending_markers.append((child, tuple(containers)))
            elif tag == W + "proofErr":
                continue
            else:
                self.object(child, containers)

    def finish(self):
        self._flush_markers(at_end=True)
        self._flush_objects()
        self.seg.meta["lead"] = self.lead
        self.seg.meta["trail"] = self.trail
        return self.seg


def paragraph_segment(p):
    b = _Builder(p)
    b.walk(p)
    seg = b.finish()
    return seg if seg.has_text() else None


def _shell(container):
    """Copie vide d'un conteneur (lien, insertion, controle de contenu)."""
    sh = etree.Element(container.tag, attrib=dict(container.attrib), nsmap=container.nsmap)
    if container.tag == W + "sdt":
        for c in container:
            if c.tag != W + "sdtContent":
                sh.append(copy.deepcopy(c))
        etree.SubElement(sh, W + "sdtContent")
    return sh


def _inner(shell):
    return shell.find(W + "sdtContent") if shell.tag == W + "sdt" else shell


def rebuild_paragraph(seg, items, target_lang):
    p = seg.meta["p"]
    ppr = p.find(W + "pPr")
    for child in list(p):
        if child is not ppr:
            p.remove(child)
    locale = LOCALE.get(target_lang)

    # Chaque element a inserer, avec la chaine de conteneurs ou le placer.
    out = []        # (element, conteneurs)

    def make_run(rpr, *contents):
        r = etree.Element(W + "r")
        if rpr is not None:
            rp = copy.deepcopy(rpr)
            lang = rp.find(W + "lang")
            if locale and lang is not None and lang.get(W + "val"):
                lang.set(W + "val", locale)
            r.append(rp)
        for c in contents:
            r.append(c)
        return r

    def text_el(txt):
        t = etree.Element(W + "t")
        t.text = txt
        if txt != txt.strip():
            t.set(XML_SPACE, "preserve")
        return t

    last = seg.styles.get(seg.base_style(), (None, ()))
    for it in items:
        kind = it[0]
        if kind == "t":
            rpr, containers = seg.styles.get(it[1], (None, ()))
            for k, piece in enumerate(it[2].split("\n")):
                if k:
                    out.append((make_run(rpr, etree.Element(W + "br")), containers))
                if piece:
                    out.append((make_run(rpr, text_el(piece)), containers))
            last = (rpr, containers)
        elif kind == "tab":
            out.append((make_run(last[0], etree.Element(W + "tab")), last[1]))
        elif kind == "br":
            src = seg.meta.get("br")
            br = copy.deepcopy(src) if src is not None else etree.Element(W + "br")
            out.append((make_run(last[0], br), last[1]))
        elif kind == "x":
            group, chain = seg.objects[it[1]]
            for el in group:
                if isinstance(el, tuple):
                    _, child, rpr = el
                    r = etree.Element(W + "r")
                    if rpr is not None:
                        r.append(copy.deepcopy(rpr))
                    r.append(child)     # deplace, pas de copie : garde l'identite
                    el = r
                out.append((el, chain))

    for el in seg.meta.get("lead", []):
        p.append(el)
    # Regroupement des runs consecutifs d'un meme conteneur. Deux liens
    # identiques (meme cible) se suivant forment un seul lien : l'ordre des
    # mots de la traduction peut avoir separe les morceaux d'un meme lien.
    def sig(chain):
        return tuple((c.tag, tuple(sorted(c.attrib.items()))) for c in chain)

    open_shells = {}
    prev_chain = ()
    for el, chain in out:
        if sig(chain) != sig(prev_chain):
            open_shells = {}
        parent = p
        for depth, cont in enumerate(chain):
            key = (depth, sig(chain[:depth + 1]))
            if key not in open_shells:
                sh = _shell(cont)
                _inner(parent).append(sh) if parent is not p else p.append(sh)
                open_shells[key] = sh
            parent = open_shells[key]
        if parent is p:
            p.append(el)
        else:
            _inner(parent).append(el)
        prev_chain = chain
    for el in seg.meta.get("trail", []):
        p.append(el)


def _context(p, part_kind):
    if part_kind != "corps":
        return part_kind
    anc = p.getparent()
    while anc is not None:
        if anc.tag == W + "txbxContent":
            return "zone de texte"
        if anc.tag == W + "tc":
            return "cellule de tableau"
        anc = anc.getparent()
    ppr = p.find(W + "pPr")
    st = ppr.find(W + "pStyle") if ppr is not None else None
    name = (st.get(W + "val") or "").lower() if st is not None else ""
    if any(k in name for k in ("heading", "berschrift", "titre", "title", "titel")):
        return "titre"
    return "paragraphe"


_PART_KIND = {"document.main+xml": "corps", "header+xml": "en-tête", "footer+xml": "pied de page",
              "footnotes+xml": "note de bas de page", "endnotes+xml": "note de fin",
              "comments+xml": "commentaire"}


def _set_default_language(doc, src_lang, tgt_lang):
    """Langue de verification : celle de la traduction, dans les styles."""
    try:
        root = doc.styles.element
    except Exception:
        return
    loc = LOCALE.get(tgt_lang)
    for lang in root.iter(W + "lang"):
        v = lang.get(W + "val")
        if v and v.lower().startswith(src_lang) and loc:
            lang.set(W + "val", loc)


def translate_docx_universal(input_path: str, output_path: str, glossary_path: str = None,
                             progress_callback=None, direction="Allemand ➔ Français"):
    """Fonction principale appelée par le routeur Streamlit."""
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

    src_lang, tgt_lang = LANG[direction][:2]
    doc = docx.Document(tmp_in)
    glossary_df = load_glossary_df(glossary_path)

    segments, contexts = [], []
    blob_parts = []
    chart_values = []
    for part in doc.part.package.iter_parts():
        ct = part.content_type
        kind = next((v for k, v in _PART_KIND.items() if ct.endswith(k)), None)
        is_chart = ct.endswith("drawingml.chart+xml")
        if kind is None and not is_chart:
            continue
        if hasattr(part, "_element"):
            root = part._element
        else:
            try:
                root = etree.fromstring(part.blob)
            except Exception:
                continue
            blob_parts.append((part, root))
        if is_chart:
            # Graphique : paragraphes DrawingML (titres) et caches de libelles.
            from processors.pptx_processor import paragraph_segment as dml_segment
            for p in root.iter(A + "p"):
                seg = dml_segment(p)
                if seg is not None:
                    seg.meta["dml"] = True
                    segments.append(seg)
                    contexts.append("titre de graphique")
            for v in root.iter(C + "v"):
                if v.getparent() is not None and v.getparent().tag == C + "pt" and v.text \
                        and re.search(r"[^\W\d_]", v.text):
                    seg = Segment()
                    seg.add_text(v.text, "-", None)
                    seg.meta["v"] = v
                    chart_values.append(seg)
            continue
        # Liste figee avant toute modification (les zones de texte sont des
        # paragraphes imbriques dans d'autres).
        for p in list(root.iter(W + "p")):
            seg = paragraph_segment(p)
            if seg is not None:
                segments.append(seg)
                contexts.append(_context(p, kind))
        # Formes DrawingML (zones de texte "wps") : texte en w:p deja couvert.
    n_main = len(segments)
    segments += chart_values
    contexts += ["libellé de graphique"] * len(chart_values)

    print(f"⚡ Traitement DOCX de {len(segments)} paragraphes...")
    results, stats = translate_segments(segments, direction, "docx", glossary_df,
                                        progress_callback, contexts)

    from processors.pptx_processor import rebuild_paragraph as dml_rebuild
    for seg, items in zip(segments[:n_main], results[:n_main]):
        if items is None:
            continue
        if seg.meta.get("dml"):
            dml_rebuild(seg, items, tgt_lang)
        else:
            rebuild_paragraph(seg, items, tgt_lang)
    for seg, items in zip(segments[n_main:], results[n_main:]):
        if items is not None:
            seg.meta["v"].text = "".join(it[2] for it in items if it[0] == "t")
    for part, root in blob_parts:
        part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _set_default_language(doc, src_lang, tgt_lang)

    doc.save(tmp_out)
    shutil.copy(tmp_out, output_path)
    shutil.rmtree(work, ignore_errors=True)

    print(f"🔤 Segments : {stats['segments']}, envoyés : {stats['segments'] - stats['skipped']}, "
          f"redemandés : {stats['retried']}, sans styles locaux : {stats['fallback']}")
    return {
        "status": "success",
        "output_path": output_path,
        "time_seconds": round(time.time() - start_time, 2),
        "input_tokens": stats["input_tokens"],
        "output_tokens": stats["output_tokens"],
        "total_cost_usd": stats["input_tokens"] / 1e6 * PRICE_IN + stats["output_tokens"] / 1e6 * PRICE_OUT,
    }
