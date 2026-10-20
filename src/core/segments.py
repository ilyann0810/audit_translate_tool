# -*- coding: utf-8 -*-
"""
Segment balise : le format d'echange entre un document et le modele.

Un paragraphe (PPTX, DOCX, XLSX) est une suite de morceaux de texte ayant
chacun un style (police, couleur, gras...), entrecoupee d'objets qui ne se
traduisent pas : champ (numero de page), image, renvoi de note, lien,
tabulation, retour a la ligne.

On l'envoie au modele sous cette forme :

    Die <s1>Umweltpraemie</s1> gilt ab 2025<x1/>.

  - le style le plus present du paragraphe est le style de base : son texte
    n'a pas de balise, ce qui limite le nombre de balises a manipuler ;
  - chaque autre style a une balise <sN>...</sN> que le modele deplace avec
    les mots traduits correspondants ;
  - chaque objet a une balise <xN/> qu'il doit garder telle quelle ;
  - <br/> est un retour a la ligne, <tab/> une tabulation, <p/> une
    separation de paragraphes (phrase coupee par Entree).

Le style lui-meme (le XML complet du run) reste cote Python : le modele ne
voit qu'un numero. A la reconstruction, chaque morceau traduit recoit une
COPIE INTEGRALE du style d'origine. Aucune propriete ne peut se perdre :
couleur de theme et ses variations, surlignage, police complexe, espacement,
langue, effet de texte...

Si la reponse ne respecte pas les balises (balise inventee, objet perdu,
balises imbriquees), parse() le signale et l'appelant redemande ou retombe
sur une version sans styles locaux mais avec tous les objets.
"""
import re

TOKEN_RE = re.compile(r"<s(\d+)>|</s(\d+)>|<x(\d+)/>|<br/>|<tab/>|<p/>")


class Segment:
    """Un paragraphe (ou un groupe de paragraphes) a traduire.

    items : liste de tuples
        ("t", style_id, texte)   morceau de texte
        ("x", obj_id)            objet a conserver
        ("br",) ("tab",) ("p",)  retour a la ligne, tabulation, paragraphe
    styles : style_id -> charge utile opaque (XML du run, etc.)
    objects : obj_id -> charge utile opaque (element XML a remettre)
    """

    def __init__(self):
        self.items = []
        self.styles = {}
        self.objects = {}
        self._style_keys = {}
        self.meta = {}

    # -- construction ------------------------------------------------------

    def add_text(self, text, style_key, payload):
        """style_key : signature hashable du style (deux runs de meme
        signature sont fusionnes) ; payload : ce qu'il faudra recopier."""
        if not text:
            return
        sid = self._style_keys.get(style_key)
        if sid is None:
            sid = len(self._style_keys)
            self._style_keys[style_key] = sid
            self.styles[sid] = payload
        if self.items and self.items[-1][0] == "t" and self.items[-1][1] == sid:
            self.items[-1] = ("t", sid, self.items[-1][2] + text)
        else:
            self.items.append(("t", sid, text))

    def add_object(self, payload):
        oid = len(self.objects) + 1
        self.objects[oid] = payload
        self.items.append(("x", oid))
        return oid

    def add_break(self, kind):
        assert kind in ("br", "tab", "p")
        self.items.append((kind,))

    # -- lecture -----------------------------------------------------------

    def plain_text(self):
        out = []
        for it in self.items:
            if it[0] == "t":
                out.append(it[2])
            elif it[0] in ("br", "p"):
                out.append("\n")
            elif it[0] == "tab":
                out.append("\t")
        return "".join(out)

    def base_style(self):
        """Style le plus present. Les styles de meta["no_base"] (texte dans
        un lien hypertexte, par exemple) ne le deviennent qu'en dernier
        recours : le texte non balise de la traduction prend ce style."""
        weight = {}
        avoid = self.meta.get("no_base", ())
        for it in self.items:
            if it[0] == "t":
                w = len(it[2].strip()) * (0.001 if it[1] in avoid else 1.0)
                weight[it[1]] = weight.get(it[1], 0) + w
        if not weight:
            return 0
        return max(weight, key=lambda k: (weight[k], -k))

    def has_text(self):
        return any(it[0] == "t" and it[2].strip() for it in self.items)

    # -- serialisation -----------------------------------------------------

    def to_tagged(self):
        base = self.base_style()
        out = []
        for it in self.items:
            kind = it[0]
            if kind == "t":
                sid, txt = it[1], it[2]
                if sid == base or not txt.strip():
                    out.append(txt)
                else:
                    # Les espaces restent hors balise : le modele les gere
                    # mieux, et un mot balise ne colle pas a son voisin.
                    m = re.match(r"^(\s*)(.*?)(\s*)$", txt, re.S)
                    out.append("%s<s%d>%s</s%d>%s" % (m.group(1), sid, m.group(2),
                                                     sid, m.group(3)))
            elif kind == "x":
                out.append("<x%d/>" % it[1])
            elif kind == "br":
                out.append("<br/>")
            elif kind == "tab":
                out.append("<tab/>")
            elif kind == "p":
                out.append("<p/>")
        tagged = "".join(out)
        # Deux balises identiques consecutives separees par un blanc : on
        # les fusionne ("<s1>A</s1> <s1>B</s1>" -> "<s1>A B</s1>").
        tagged = re.sub(r"</s(\d+)>(\s*)<s\1>", r"\2", tagged)
        return tagged

    def counts(self):
        c = {"x": sorted(it[1] for it in self.items if it[0] == "x"),
             "br": sum(1 for it in self.items if it[0] == "br"),
             "tab": sum(1 for it in self.items if it[0] == "tab"),
             "p": sum(1 for it in self.items if it[0] == "p")}
        return c

    # -- relecture de la traduction ------------------------------------------

    def parse(self, translated):
        """Traduction balisee -> (items, probleme ou None).

        Les items ont la meme forme que self.items. Le probleme decrit ce
        qui ne va pas ; dans ce cas les items renvoyes sont quand meme
        utilisables (fallback()).
        """
        if translated is None:
            return None, "vide"
        base = self.base_style()
        items, stack, problem = [], [], None
        pos = 0
        seen_x = []
        for m in TOKEN_RE.finditer(translated):
            if m.start() > pos:
                sid = stack[-1] if stack else None
                items.append(("t", sid, translated[pos:m.start()]))
            pos = m.end()
            tok = m.group(0)
            if m.group(1):
                sid = int(m.group(1))
                if sid not in self.styles or sid == base:
                    problem = problem or "style inconnu s%d" % sid
                if stack:
                    problem = problem or "balises imbriquees"
                stack.append(sid if sid in self.styles else base)
            elif m.group(2):
                if not stack:
                    problem = problem or "fermeture orpheline"
                else:
                    stack.pop()
            elif m.group(3):
                oid = int(m.group(3))
                if oid not in self.objects:
                    problem = problem or "objet inconnu x%d" % oid
                    continue
                seen_x.append(oid)
                items.append(("x", oid))
            elif tok == "<br/>":
                items.append(("br",))
            elif tok == "<tab/>":
                items.append(("tab",))
            elif tok == "<p/>":
                items.append(("p",))
        if pos < len(translated):
            sid = stack[-1] if stack else None
            items.append(("t", sid, translated[pos:]))
        # Texte hors balise : style de base, sauf un simple blanc entre deux
        # passages, qui prend le style du texte qui le precede (un espace
        # surligne ou souligne doit le rester, un espace hors lien aussi).
        prev = None
        for k, it in enumerate(items):
            if it[0] != "t":
                continue
            if it[1] is None:
                sid = prev if (not it[2].strip() and prev is not None) else base
                items[k] = ("t", sid, it[2])
            prev = items[k][1]
        if stack:
            problem = problem or "balise non fermee"
        want = self.counts()
        if sorted(seen_x) != want["x"]:
            problem = problem or "objets %s au lieu de %s" % (sorted(seen_x), want["x"])
        got_p = sum(1 for it in items if it[0] == "p")
        if got_p != want["p"]:
            problem = problem or "%d separations de paragraphe au lieu de %d" % (got_p, want["p"])
        got_tab = sum(1 for it in items if it[0] == "tab")
        if got_tab != want["tab"]:
            problem = problem or "%d tabulations au lieu de %d" % (got_tab, want["tab"])
        # Balise de style perdue : le mot en gras ou en couleur ne l'est
        # plus. Pas bloquant, mais a redemander une fois.
        used = {it[1] for it in items if it[0] == "t" and it[2].strip()}
        wanted = {it[1] for it in self.items if it[0] == "t" and it[2].strip()}
        if not problem and wanted - used:
            problem = "style perdu %s" % sorted(wanted - used)
        items = [it for it in items if not (it[0] == "t" and it[2] == "")]
        return items, problem

    def fallback(self, translated):
        """Version sure quand les balises sont inutilisables.

        Texte entierement dans le style de base, objets remis a leur
        position relative d'origine, separations de paragraphes et
        tabulations dans le nombre d'origine. Les styles locaux sont perdus
        mais rien n'est detruit ni ajoute.
        """
        base = self.base_style()
        clean = TOKEN_RE.sub(lambda m: {"<br/>": "\n", "<tab/>": "\t", "<p/>": "\x00"}.get(m.group(0), ""),
                             translated or "")
        # Position relative de chaque element non textuel dans l'original.
        total = max(1, sum(len(it[2]) for it in self.items if it[0] == "t"))
        acc, anchors = 0, []
        for it in self.items:
            if it[0] == "t":
                acc += len(it[2])
            else:
                anchors.append((acc / total, it))
        text = clean.replace("\x00", " ").replace("\n", " ").replace("\t", " ")
        text = re.sub(r" {2,}", " ", text)
        n = len(text)
        items, last = [], 0
        for frac, it in anchors:
            cut = int(round(frac * n))
            # On coupe a une frontiere de mot.
            while 0 < cut < n and text[cut - 1] != " " and text[cut] != " ":
                cut += 1
            cut = max(cut, last)
            if cut > last:
                items.append(("t", base, text[last:cut]))
            items.append(it)
            last = cut
        if last < n:
            items.append(("t", base, text[last:]))
        return items
