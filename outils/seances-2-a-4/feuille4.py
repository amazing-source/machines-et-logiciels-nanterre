# -*- coding: utf-8 -*-
import re, html, zipfile, os, shutil
import xml.dom.minidom as md

SRC = r"C:\Users\utilisateur\Desktop\Cours\Environnement Numerique\seance-4\supports\Feuille_Exercices_Excel_Approfondissement_ML_Nanterre.docx"
OUT = r"C:\Users\utilisateur\Desktop\Cours\Environnement Numerique\seance-4\SEANCE 4 - Reponses - [Prénom].docx"
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


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def para(lines):
    runs = []
    for k, l in enumerate(lines):
        if k:
            runs.append("<w:r>" + RPR + "<w:br/></w:r>")
        runs.append("<w:r>" + RPR + '<w:t xml:space="preserve">' + esc(l) + "</w:t></w:r>")
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
    "TRAVAIL REALISE — classeur NOM_Prenom_TD_Excel_Approfondissement.xlsx",
    "Le classeur contient les 6 feuilles d'exercices d'origine plus 3 feuilles de tableaux croises dynamiques creees pour l'exercice 3 et l'exercice 5.",
]

ex1 = [
    "Fait sur les 300 lignes (A4:J304) :",
    "• tri a deux niveaux : Region de A a Z, puis Montant du plus grand au plus petit ;",
    "• filtres automatiques actives sur tout le tableau, puis filtre cumule Categorie = Informatique ET Montant > 500 € (filtre numerique « Superieur a ») ;",
    "• barres de donnees bleues sur la colonne Montant (I5:I304) ;",
    "• regle de seuil : les montants superieurs a 2 000 € ressortent en rouge ;",
    "• jeu de 3 icones (feux tricolores) sur la colonne Marge (J5:J304) ;",
    "• regle par formule =$J5<0,15 appliquee a A5:J304, qui surligne la LIGNE ENTIERE quand la marge est critique. Le $ bloque la colonne J tandis que le numero de ligne reste relatif : c'est ce qui permet a une seule regle de traiter les 300 lignes.",
]

ex2 = [
    "Fait sur les 200 commandes :",
    "• liste deroulante sur la colonne Statut (I5:I204), source $L$5:$L$8, avec alerte d'erreur bloquante ;",
    "• validation Nombre entier compris entre 1 et 5 sur la Note satisfaction (J5:J204), avec message de saisie « Note entre 1 et 5 » et message d'erreur explicite ;",
    "• graphique en SECTEURS sur la synthese par categorie (N5:O10), avec titre et etiquettes en pourcentage ;",
    "• graphique en HISTOGRAMME sur la synthese mensuelle, montrant l'evolution du montant total mois par mois.",
    "Remarque technique : la synthese mensuelle occupe Q:S, avec une colonne intermediaire « Debut du mois » qui ne doit pas etre tracee. Elle a donc ete retiree des series du graphique.",
]

ex3 = [
    "Trois tableaux croises dynamiques ont ete construits sur de nouvelles feuilles, a partir des 400 commandes (A4:I404) :",
    "• feuille « TCD Region x Categorie » : Lignes = Region, Colonnes = Categorie, Valeurs = Somme du Montant, avec totaux generaux en ligne et en colonne, et le champ Vendeur place en zone de FILTRE pour permettre le filtrage interactif ;",
    "• feuille « TCD Vendeur » : Lignes = Vendeur, Valeurs = NOMBRE de ID Commande (et non Somme : le calcul a ete change dans les parametres du champ de valeurs), trie par ordre alphabetique INVERSE des noms ;",
    "• feuille « TCD Stock » (exercice 5) : Lignes = Entrepot, Colonnes = Categorie, Valeurs = Somme de la Valeur stock.",
]

ex4 = [
    "Resultats obtenus sur les 250 employes :",
    "• G257 — NB.SI.ENS : 3 cadres du service Commercial travaillent a Paris.",
    "• G258 — BDSOMME (zone de criteres K262:L263, Service = Informatique ET Anciennete > 5) : masse salariale mensuelle de 93 110,00 €.",
    "• G259 — BDMOYENNE (zone de criteres K266:K267, Statut = Cadre) : salaire mensuel moyen de 5 297,17 €.",
    "• I5:I254 — SIERREUR : prime exprimee en pourcentage du salaire annuel, =SIERREUR(H5/(G5*12);\"N/A\"). Pour le premier employe : 11,1 %.",
    "• Filtre avance (Donnees > Avance, criteres C270:G271 : Informatique + Cadre + salaire > 4 000 €), extraction copiee a partir de A274 : 12 lignes extraites.",
    "• Feuille protegee par mot de passe : TD2026Excel.",
]

ex5 = [
    "Fait sur les 350 references reparties dans 4 entrepots :",
    "• tri Entrepot croissant puis Valeur stock decroissante ; filtres automatiques actives et filtre pose sur une seule categorie (Informatique) ;",
    "• barres de donnees sur la Valeur stock (H5:H354) ;",
    "• regle par formule =$E5<$F5 sur A5:I354 : la ligne entiere se surligne des que la quantite passe sous le seuil mini ;",
    "• I5:I354 — ratio quantite / seuil avec =SIERREUR(E5/F5;\"Seuil non defini\"), puis jeu de 3 icones sur cette colonne ;",
    "• G357 — NB.SI.ENS, references en rupture (quantite = 0) dans l'entrepot de Paris : 0.",
    "Verification faite : le jeu de donnees ne contient AUCUNE quantite egale a zero, sur les 350 references et tous entrepots confondus. Le resultat 0 n'est donc pas une erreur de formule mais bien la realite du fichier — c'est le genre de resultat qu'il faut confirmer avant de conclure.",
]

ex6 = [
    "Fait sur les 300 etudiants :",
    "• validation de type Liste sur la colonne Session (H5:H304), restreinte aux valeurs 1 et 2 ;",
    "• validation decimale comprise entre 0 et 20 sur les colonnes Note 1 et Note 2 (E5:F304) ;",
    "• G5:G304 — moyenne des deux notes avec =SIERREUR(MOYENNE(E5:F5);\"Absent\"). Pour le premier etudiant (6 et 11,9) : 8,95. Les etudiants absents aux deux epreuves affichent « Absent » au lieu d'une erreur #DIV/0! ;",
    "• G311 — BDMOYENNE avec la zone de criteres A307:A308 : la moyenne generale de la filiere Informatique est de 11,19 ;",
    "• graphique en secteurs (effectif par filiere) et histogramme (moyenne par filiere) a partir de la synthese A314:C319 ;",
    "• feuille protegee par mot de passe : TD2026Excel.",
]

ident = ('<w:p><w:pPr><w:spacing w:before="0" w:after="160"/><w:jc w:val="center"/>'
         '<w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="28"/></w:rPr></w:pPr>'
         '<w:r><w:rPr><w:b/><w:color w:val="1F4E79"/><w:sz w:val="28"/></w:rPr>'
         '<w:t xml:space="preserve">[NOM Prénom] \u2014 [N° étudiant]</w:t></w:r></w:p>')

# ancres : fragments sans accent ni apostrophe typographique, pour etre surs
inserts = [
    (anchor("pour surligner la ligne"), para(ex1)),
    (anchor("Mettez en forme les deux graphiques"), para(ex2)),
    (anchor("Triez ce second TCD par ordre inverse"), para(ex3)),
    (anchor("accidentelle des"), para(ex4)),
    (anchor("Construisez un TCD (nouvelle feuille)"), para(ex5)),
    (anchor("la feuille avec un mot de passe de votre choix"), para(ex6)),
    (spans[0][1], ident + para(recap)),
]

for pos, txt in sorted(inserts, key=lambda a: -a[0]):
    xml = xml[:pos] + txt + xml[pos:]

open(p, "w", encoding="utf-8").write(xml)
md.parse(p)

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n in names:
        z.write(os.path.join(WORK, n.replace("/", os.sep)), n)

print("OK :", OUT, os.path.getsize(OUT), "octets,", len(inserts), "insertions")
