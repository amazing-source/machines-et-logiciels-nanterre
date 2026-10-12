# -*- coding: utf-8 -*-
import re, html, zipfile, os, shutil
import xml.dom.minidom as md

SRC = r"C:\Users\utilisateur\Desktop\Cours\Environnement Numerique\seance-3\supports\Feuille_Exercices_Excel_ML_Nanterre.docx"
OUT = r"C:\Users\utilisateur\Desktop\Cours\Environnement Numerique\seance-3\SEANCE 3 - Reponses - [Prénom].docx"
WORK = r"C:\Users\utilisateur\AppData\Local\Temp\travail"

if os.path.isdir(WORK):
    shutil.rmtree(WORK)
os.makedirs(WORK)
with zipfile.ZipFile(SRC) as z:
    names = z.namelist()
    z.extractall(WORK)

p = os.path.join(WORK, "word", "document.xml")
xml = open(p, encoding="utf-8").read()

PPR = ('<w:pPr><w:pBdr>'
       '<w:top w:val="single" w:sz="4" w:space="2" w:color="1F4E79"/>'
       '<w:left w:val="single" w:sz="18" w:space="6" w:color="1F4E79"/>'
       '<w:bottom w:val="single" w:sz="4" w:space="2" w:color="1F4E79"/>'
       '<w:right w:val="single" w:sz="4" w:space="4" w:color="1F4E79"/>'
       '</w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="F2F7FB"/>'
       '<w:spacing w:before="80" w:after="140" w:line="240" w:lineRule="auto"/>'
       '<w:rPr><w:color w:val="1F4E79"/><w:sz w:val="21"/></w:rPr></w:pPr>')
RPR = '<w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="21"/></w:rPr>'
RPR_MONO = '<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:b/><w:color w:val="C00000"/><w:sz w:val="20"/></w:rPr>'


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def para(lines):
    runs = []
    for k, l in enumerate(lines):
        if k:
            runs.append("<w:r>" + RPR + "<w:br/></w:r>")
        rpr = RPR_MONO if l.startswith("=") else RPR
        runs.append("<w:r>" + rpr + '<w:t xml:space="preserve">' + esc(l) + "</w:t></w:r>")
    return "<w:p>" + PPR + "".join(runs) + "</w:p>"


spans = []
for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", xml, re.S):
    body = re.sub(r"<[^>]+>", "", m.group(0))
    spans.append((m.start(), m.end(), html.unescape(body).strip()))


def anchor(needle):
    for (s, e, t) in spans:
        if needle in t:
            return e
    raise SystemExit("ANCRE INTROUVABLE : " + needle)


recap = [
    "RESULTATS OBTENUS — classeur NOM_Prenom_TD_Excel.xlsx",
    "Exercice 1 : total des depenses 746,79 € HT / 896,15 € TTC (TVA 20 %).",
    "Exercice 2 : total general 358 unites ; montants TTC 1 980 € / 5 100 € / 3 312 € ; grille de remises 10/20/30 % complete par une seule recopie croisee.",
    "Exercice 3 : 8 inscrits, 6 presents a l'oral ; 4 admis et 4 ajournes.",
    "Exercice 4 : total facture 145,50 €, identique par SOMME (E20) et par SOMMEPROD (E21) — la verification croisee est bonne.",
    "Exercice 5 : valeur totale du stock 530,20 € ; 1 reference en rupture (Casque audio), 3 en alerte, 2 correctes.",
    "Exercice 6 : moyenne ponderee 11,20 -> mention Passable, decision Admis ; 4 notes renseignees sur 5 matieres.",
    "Aucune cellule en erreur (#N/A, #VALEUR!, #DIV/0!) dans l'ensemble du classeur.",
]

rep_ex3 = [
    "Pour Martin et Richard, qui n'ont pas de note d'oral, la fonction MOYENNE ne renvoie pas d'erreur : elle ignore purement et simplement les cellules vides. Martin obtient donc 8,00 et Richard 17,00, c'est-a-dire leur seule note d'ecrit recopiee.",
    "C'est un piege : la valeur affichee n'est plus une moyenne ecrit/oral, elle n'est pas comparable a celle des autres etudiants. Consequence concrete ici, Richard est declare « Admis » avec 17,00 alors qu'il ne s'est pas presente a l'oral.",
    "A retenir : une cellule vide n'est pas un zero pour MOYENNE. Si l'absence devait compter 0, il faudrait ecrire =SOMME(B5:C5)/2 et non =MOYENNE(B5:C5). C'est un choix de regle de jury, pas un detail technique — et c'est exactement pour reperer ce cas que NBVAL est utilise en D15 : 6 presents pour 8 inscrits.",
]

rep_ex6 = [
    "NBVAL renvoie 4 alors que le tableau compte 5 matieres : la note de Communication manque. La moyenne calculee n'est donc pas fiable.",
    "Le probleme vient du denominateur : le coefficient 1 de Communication est malgre tout compte dans le total des coefficients (B10 = 10), alors que sa note ponderee vaut 0. Tout se passe comme si l'etudiant avait eu 0 dans cette matiere.",
    "L'effet est loin d'etre negligeable : la moyenne tombe a 11,20 au lieu de 12,44, et la mention passe de « Assez Bien » a « Passable ».",
    "L'ecart entre NBVAL (4) et le nombre de lignes du tableau (5) est precisement le signal d'alerte qui permet de reperer la donnee manquante AVANT de tirer une conclusion d'un chiffre faux.",
]

rep_reflexion = [
    "Il faut diviser par la somme des seuls coefficients des matieres effectivement notees, et non par le total de tous les coefficients :",
    "=SOMMEPROD(B5:B9;C5:C9)/SOMMEPROD((C5:C9<>\"\")*B5:B9)",
    "Le numerateur additionne les produits coefficient x note ; les cellules vides y valent 0 et ne faussent rien. Le denominateur, lui, ne retient que les coefficients dont la note est renseignee : le test (C5:C9<>\"\") produit une suite de VRAI/FAUX convertie en 1/0, qui annule le coefficient des matieres non notees.",
    "Resultat : 12,44 au lieu de 11,20, soit la mention « Assez Bien ». Cette formule est placee en D15 de la feuille « Exercice 6 », sous l'intitule « Moyenne corrigee », a cote de la moyenne brute pour que la comparaison soit visible.",
    "Remarque : si la regle du jury voulait au contraire qu'une matiere non notee compte pour 0 (absence injustifiee), c'est la formule actuelle de D11 qu'il faut conserver. Les deux formules sont justes, elles repondent simplement a deux regles differentes — d'ou l'interet de les afficher toutes les deux.",
]

ident = ('<w:p><w:pPr><w:spacing w:before="0" w:after="160"/><w:jc w:val="center"/>'
         '<w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="28"/></w:rPr></w:pPr>'
         '<w:r><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="28"/></w:rPr>'
         '<w:t xml:space="preserve">[NOM Prénom] \u2014 [N° étudiant]</w:t></w:r></w:p>')

inserts = [
    (anchor("Observez ce qu'il se passe pour les"), para(rep_ex3)),
    (anchor("qu'en d"), para(rep_ex6)),
    (anchor("Comment pourriez-vous corriger la formule"), para(rep_reflexion)),
    (spans[0][1], ident + para(recap)),
]

for pos, txt in sorted(inserts, key=lambda x: -x[0]):
    xml = xml[:pos] + txt + xml[pos:]

open(p, "w", encoding="utf-8").write(xml)
md.parse(p)

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n in names:
        z.write(os.path.join(WORK, n.replace("/", os.sep)), n)

print("OK :", OUT, os.path.getsize(OUT), "octets,", len(inserts), "insertions")
