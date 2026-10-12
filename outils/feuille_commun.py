# -*- coding: utf-8 -*-
# Outils communs aux feuilles de reponses des seances 5 et 6.
# Meme principe que feuille2/3/4.py (reponses en bleu encadrees, inserees
# apres une ancre), mais l'insertion se fait sur l'arbre XML et non par
# decoupage de texte : les enonces de ces seances contiennent des zones de
# texte imbriquees, qu'un decoupage par expressions regulieres couperait
# au mauvais endroit.
import os
import shutil
import tempfile
import unicodedata
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
BLEU = "1F4E79"


def q(tag):
    return "{%s}%s" % (W, tag)


def norm(t):
    """Texte compare sans accents, sans casse, apostrophes et espaces unifies."""
    t = unicodedata.normalize("NFKD", t)
    t = "".join(c for c in t if not unicodedata.combining(c))
    for a in "’‘ʼ`":
        t = t.replace(a, "'")
    t = t.replace(" ", " ").replace(" ", " ")
    return " ".join(t.lower().split())


class Feuille:
    def __init__(self, src):
        self.work = tempfile.mkdtemp(prefix="feuille_")
        with zipfile.ZipFile(src) as z:
            self.names = z.namelist()
            z.extractall(self.work)
        self.doc = os.path.join(self.work, "word", "document.xml")
        self.tree = etree.parse(self.doc)
        self.body = self.tree.getroot().find(q("body"))
        self.paras = []
        for el in self.body:
            if el.tag == q("p"):
                self.paras.append((el, norm(self._texte(el))))
        self.n = 0
        self.retires = 0

    @staticmethod
    def _texte(p):
        out = []
        for t in p.iter(q("t")):
            # on ignore le texte des zones de texte ancrees dans le paragraphe
            if any(a.tag == q("txbxContent") for a in t.iterancestors()):
                continue
            out.append(t.text or "")
        return "".join(out)

    def ancre(self, fragment, occurrence=1):
        f = norm(fragment)
        k = 0
        for el, t in self.paras:
            if f in t:
                k += 1
                if k == occurrence:
                    return el
        raise SystemExit("ANCRE INTROUVABLE (%d) : %s" % (occurrence, fragment))

    def premier(self):
        return self.paras[0][0]

    def _p(self, xml):
        return etree.fromstring('<w:p xmlns:w="%s">%s</w:p>' % (W, xml))

    @staticmethod
    def _esc(t):
        return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def reponse(self, lignes):
        ppr = ('<w:pPr><w:pBdr>'
               '<w:top w:val="single" w:sz="4" w:space="2" w:color="%s"/>'
               '<w:left w:val="single" w:sz="18" w:space="6" w:color="%s"/>'
               '<w:bottom w:val="single" w:sz="4" w:space="2" w:color="%s"/>'
               '<w:right w:val="single" w:sz="4" w:space="4" w:color="%s"/>'
               '</w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="F2F7FB"/>'
               '<w:spacing w:before="80" w:after="140" w:line="250" w:lineRule="auto"/>'
               '<w:ind w:left="0" w:firstLine="0"/><w:jc w:val="left"/>'
               '<w:rPr><w:color w:val="%s"/><w:sz w:val="21"/></w:rPr></w:pPr>') % ((BLEU,) * 5)
        rpr = '<w:rPr><w:b/><w:color w:val="%s"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>' % BLEU
        runs = []
        for k, l in enumerate(lignes):
            if k:
                runs.append("<w:r>" + rpr + "<w:br/></w:r>")
            runs.append("<w:r>" + rpr + '<w:t xml:space="preserve">' + self._esc(l) + "</w:t></w:r>")
        return self._p(ppr + "".join(runs))

    def identite(self, texte):
        rpr = '<w:rPr><w:b/><w:color w:val="%s"/><w:sz w:val="28"/></w:rPr>' % BLEU
        return self._p('<w:pPr><w:spacing w:before="0" w:after="160"/><w:jc w:val="center"/>' + rpr + '</w:pPr>'
                       '<w:r>' + rpr + '<w:t xml:space="preserve">' + self._esc(texte) + '</w:t></w:r>')

    def apres(self, ancre, *elements):
        """Insere les elements, dans l'ordre, juste apres le paragraphe ancre."""
        ref = ancre
        for e in elements:
            ref.addnext(e)
            ref = e
            self.n += 1

    @staticmethod
    def _cadre_vide(p):
        """Paragraphe encadre laisse vide par l'enonce pour y ecrire la reponse."""
        if p is None or p.tag != q("p"):
            return False
        if p.find(".//" + q("pBdr")) is None or p.find(".//" + q("sectPr")) is not None:
            return False
        for tag in ("t", "drawing", "pict", "object", "br"):
            if p.find(".//" + q(tag)) is not None:
                return False
        return True

    def repondre(self, fragment, lignes, occurrence=1):
        r = self.reponse(lignes)
        self.apres(self.ancre(fragment, occurrence), r)
        # le cadre vide prevu par l'enonce ferait croire a une question sans
        # reponse : on le retire, la reponse encadree le remplace
        suivant = r.getnext()
        while self._cadre_vide(suivant):
            a_retirer = suivant
            suivant = suivant.getnext()
            self.body.remove(a_retirer)
            self.retires += 1

    def nettoyer_fin(self):
        """Retire les paragraphes vides en fin de document : pousses par les
        reponses inserees, ils creeraient une derniere page blanche."""
        n = 0
        while True:
            enfants = [e for e in self.body if e.tag != q("sectPr")]
            if not enfants:
                break
            d = enfants[-1]
            if d.tag != q("p") or d.find(".//" + q("sectPr")) is not None:
                break
            if any(d.find(".//" + q(t)) is not None for t in ("t", "drawing", "pict", "object")):
                break
            self.body.remove(d)
            n += 1
        return n

    def enregistrer(self, out):
        self.tree.write(self.doc, xml_declaration=True, encoding="UTF-8", standalone=True)
        etree.parse(self.doc)  # le XML produit doit rester bien forme
        if os.path.exists(out):
            os.remove(out)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for n in self.names:
                z.write(os.path.join(self.work, n.replace("/", os.sep)), n)
        shutil.rmtree(self.work, ignore_errors=True)
        print("OK : %s | %d octets | %d insertions | %d cadres vides retires" % (out, os.path.getsize(out), self.n, self.retires))
