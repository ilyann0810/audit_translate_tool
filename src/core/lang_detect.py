# -*- coding: utf-8 -*-
"""
Detection de langue legere (allemand / francais / anglais), sans dependance.

Sert a deux choses :
  - ne pas envoyer au modele un texte deja dans la langue cible (un titre
    allemand dans un document qu'on traduit vers l'allemand), ce qui evite
    que le modele le "retraduise" dans l'autre sens ;
  - verifier apres traduction qu'il ne reste pas de texte en langue source.

Principe : mots-outils propres a chaque langue (articles, prepositions,
pronoms, auxiliaires), lettres propres a chaque langue (ß, ä, ö, ü / é, è,
ç, à...), et terminaisons typiques des mots longs. Les mots communs a
plusieurs langues ("die" anglais, "des" allemand et francais) sont exclus
des listes. Un texte trop court ou sans indice renvoie None : on ne
devine pas.
"""
import re

_DE_WORDS = set("""
und oder aber nicht kein keine keinen keiner der die das dem den des ein eine
einen einem einer eines ist sind war waren wird werden wurde wurden worden
hat haben hatte hatten sein ihr ihre ihren seine seinen mit ohne fuer für
von vom zum zur bei beim aus auf nach über ueber unter vor hinter zwischen
durch gegen um wie wenn dass daß ob auch noch nur schon sehr mehr weniger
sowie bzw ggf inkl sich wir sie ihnen uns dieser diese dieses diesem diesen
jeder jede jedes alle allen werden kann können koennen muss müssen sollen soll
darf dürfen wurde hier dort heute gemäß gemaess bzgl anhand innerhalb
außerhalb ausserhalb sowohl als also bereits zudem jedoch dabei damit dazu
davon dafür hierbei wobei welche welcher welches zwar ob bis seit während
waehrend wegen trotz statt laut je pro im ins am zu es ich du er
""".split())

_FR_WORDS = set("""
et ou mais ne pas plus le la les un une du des au aux ce cet cette ces
est sont était étaient sera seront été être avoir ont avait
avec sans pour par dans sur sous chez vers entre selon depuis pendant
qui que quoi dont où lors leur leurs son sa ses notre nos votre vos
il ils elle elles nous vous on se ne y en très aussi encore déjà
afin ainsi alors donc car comme si non tout tous toute toutes
cela ceci celui celle ceux celles même autre autres chaque plusieurs
doit doivent peut peuvent fait faire mise mis lorsque puis

""".split())

_EN_WORDS = set("""
the and or not of to in on for with from by at is are was were be been
this that these those which who what when where how it its their our
your we you they he she has have had will would can could should may
""".split())

# Lettres caracteristiques. Le "é" existe dans des noms propres allemands
# mais reste un tres bon indice de francais dans un texte d'entreprise.
_DE_CHARS = set("äöüßÄÖÜ")
_FR_CHARS = set("éèêëàâçùûîïôœÉÈÊÀÇÔ")

_DE_SUFFIX = re.compile(r"(ung|ungen|keit|heit|lich|isch|schaft|chen|tion[e]n|"
                        r"ierung|ieren|bar|weise)$")
_FR_SUFFIX = re.compile(r"(ement|ements|tion|tions|eur|eurs|euse|ique|iques|"
                        r"aire|aires|ité|ités|ée|ées|ais|ait|aient|ons|ez)$")

_TOKEN = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿŒœß]+(?:['’][A-Za-zÀ-ÖØ-öø-ÿ]+)?")


def scores(text):
    """Renvoie {'de': x, 'fr': y, 'en': z} : indices accumules par langue."""
    s = {"de": 0.0, "fr": 0.0, "en": 0.0}
    if not text:
        return s
    for tok in _TOKEN.findall(text):
        low = tok.lower()
        # Elision francaise : l'audit, d'un, qu'il
        if "'" in low or "’" in low:
            head = re.split(r"['’]", low)[0]
            if head in ("l", "d", "qu", "c", "s", "n", "j", "jusqu", "lorsqu"):
                s["fr"] += 1.0
                low = re.split(r"['’]", low)[1]
        if low in _DE_WORDS:
            s["de"] += 1.0
        if low in _FR_WORDS:
            s["fr"] += 1.0
        if low in _EN_WORDS:
            s["en"] += 1.0
        if any(c in _DE_CHARS for c in tok):
            s["de"] += 1.5
        if any(c in _FR_CHARS for c in tok):
            s["fr"] += 1.5
        if len(low) >= 6:
            if _DE_SUFFIX.search(low):
                s["de"] += 0.5
            if _FR_SUFFIX.search(low):
                s["fr"] += 0.5
        # Substantif allemand compose : long mot capitalise sans accent
        # francais ("Kommissionierungskontrollen").
        if len(tok) >= 13 and tok[0].isupper() and not any(c in _FR_CHARS for c in tok):
            s["de"] += 0.5
    return s


def detect(text, min_score=1.5):
    """'de', 'fr', 'en' ou None si le texte ne donne pas assez d'indices."""
    s = scores(text)
    best = max(s, key=s.get)
    ordered = sorted(s.values(), reverse=True)
    if ordered[0] < min_score or ordered[0] < 1.5 * ordered[1]:
        return None
    return best


def contains_language(text, lang, min_score=1.5):
    """Le texte contient-il des indices nets de cette langue ?

    Plus large que detect() : un titre mixte "Wochenkennzahlen | Mise a
    jour en cours" contient du francais meme si l'allemand y est aussi.
    """
    return scores(text)[lang] >= min_score
