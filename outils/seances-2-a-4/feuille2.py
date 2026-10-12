# -*- coding: utf-8 -*-
import re, html, zipfile, os, shutil

SRC = r"C:\Users\utilisateur\Desktop\Cours\Environnement Numerique\seance-2\supports\SEANCE 2 TTX.docx"
DST_DIR = r"C:\Users\utilisateur\Desktop\Cours\Environnement Numerique\seance-2"
OUT = os.path.join(DST_DIR, "SEANCE 2 - Reponses - [Prénom].docx")
WORK = r"C:\Users\utilisateur\AppData\Local\Temp\travail"

if os.path.isdir(WORK):
    shutil.rmtree(WORK)
os.makedirs(WORK)
with zipfile.ZipFile(SRC) as z:
    names = z.namelist()
    z.extractall(WORK)

p = os.path.join(WORK, "word", "document.xml")
xml = open(p, encoding="utf-8").read()

PPR = ('<w:pPr><w:widowControl w:val="0"/><w:pBdr>'
       '<w:top w:val="single" w:sz="4" w:space="2" w:color="1F4E79"/>'
       '<w:left w:val="single" w:sz="4" w:space="4" w:color="1F4E79"/>'
       '<w:bottom w:val="single" w:sz="4" w:space="2" w:color="1F4E79"/>'
       '<w:right w:val="single" w:sz="4" w:space="4" w:color="1F4E79"/>'
       '</w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="F2F7FB"/>'
       '<w:spacing w:before="60" w:after="120" w:line="240" w:lineRule="auto"/>'
       '<w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:color w:val="1F4E79"/>'
       '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr>')
RPR = '<w:rPr><w:rFonts w:cstheme="minorHAnsi"/><w:b/><w:color w:val="1F4E79"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def para(lines):
    runs = []
    for k, l in enumerate(lines):
        if k:
            runs.append("<w:r>" + RPR + "<w:br/></w:r>")
        runs.append("<w:r>" + RPR + '<w:t xml:space="preserve">' + esc(l) + "</w:t></w:r>")
    return "<w:p>" + PPR + "".join(runs) + "</w:p>"


# --- decoupe en paragraphes avec leurs positions ---
spans = []
for m in re.finditer(r"<w:p\b[^>]*>.*?</w:p>", xml, re.S):
    body = re.sub(r"<[^>]+>", "", m.group(0))
    spans.append((m.start(), m.end(), html.unescape(body).strip()))


def anchor(needle, occurrence=1, after_index=0):
    """retourne la position de fin du paragraphe contenant needle"""
    seen = 0
    for (s, e, t) in spans:
        if s < after_index:
            continue
        if needle in t:
            seen += 1
            if seen == occurrence:
                return e
    raise SystemExit("ANCRE INTROUVABLE : " + needle)


A = {}
A["instances"] = [
    "3 instances remplacées.",
    "Détail : « une préparation aux pistaches » et « 30 g de pistaches hachées » sont bien devenues « noisettes ». La troisième occurrence, « plus pistaché », n'est PAS trouvée par une recherche de « pistache » : l'accent final en fait un mot différent pour Word. Il a fallu une seconde recherche sur « pistaché » pour la remplacer par « noisetté ».",
    "Vérification finale : plus aucune occurrence de la chaîne « istach » dans le document. C'est le piège classique du Rechercher/Remplacer — on croit avoir tout traité alors qu'une forme dérivée subsiste.",
]
A["style_defaut"] = [
    "Normal.",
    "Le document source n'utilise aucun style de titre : absolument tout le texte est en Normal. C'est précisément le défaut que la suite de l'exercice corrige — sans styles, le document n'a aucune structure exploitable (pas de mode Plan, pas de table des matières automatique, pas de navigation).",
    "Conséquence pratique, et c'est un piège : puisque le titre est en Normal, modifier « Normal » pour le centrer en gras 18 aurait centré et grossi TOUT le document. J'ai donc d'abord appliqué le style « Titre » à cette ligne, puis modifié ce style-là (centré, gras, 18, espacement de 4 cm avant).",
]
A["observez_normal"] = [
    "Tout le corps du document change d'un seul coup : police Arial, texte justifié, interligne 1,15 et 8 points d'espace après chaque paragraphe.",
    "C'est la démonstration de l'intérêt des styles : comme tous les paragraphes héritent de Normal, une seule modification remplace des dizaines de mises en forme manuelles. C'est aussi pour cette raison que Normal est le style le plus sensible du document — le modifier se répercute partout.",
]
A["saut_page"] = [
    "Chaque paragraphe en Titre 1 bascule immédiatement en haut d'une nouvelle page. Le document est désormais organisé à raison d'une grande section par page, soit 9 pages au total : page de titre, sommaire, puis une page par section.",
    "L'effet est global et automatique parce qu'il est porté par le STYLE et non par des sauts de page insérés à la main : si j'ajoute une section plus tard, elle démarrera d'elle-même sur une page neuve, et si je change d'avis il suffit de décocher l'option une seule fois.",
]
A["heritage"] = [
    "Non — et c'est contre-intuitif. Les autres titres ne changent pas automatiquement.",
    "Vérification faite dans le volet des styles : Titre 1, Titre 2 et Titre 3 sont tous les trois basés sur « Normal », et non les uns sur les autres. Il n'existe donc aucune chaîne d'héritage entre Titre 1 et Titre 3 : modifier la couleur de Titre 1 laisse Titre 2 et Titre 3 sur la couleur d'accentuation du thème.",
    "Deux solutions pour obtenir le bleu foncé partout : soit modifier le style parent commun, « Normal », mais cela colorerait aussi tout le corps du texte ; soit fixer la couleur sur chacun des trois styles de titre. C'est la seconde que j'ai retenue.",
    "À retenir : l'héritage ne fonctionne qu'entre un style et le style sur lequel il est « basé ». Ici la chaîne passe par Normal, pas par Titre 1.",
]
A["titre4"] = [
    "Aucun. La recherche par style ne renvoie 0 paragraphe en Titre 4 : le document n'en contient pas, la hiérarchie s'arrête au Titre 3.",
    "L'intérêt de la manipulation est de savoir chercher par MISE EN FORME et non par texte : dans Rechercher avancé, on laisse le champ de recherche vide et on précise Format > Style > Titre 4.",
]
A["para_vides"] = [
    "Parce qu'ils créent une mise en page fragile et faussent la structure du document :",
    "• l'espacement devient irrégulier et incontrôlable — on ne peut plus le régler proprement par le style (espace avant / après) ;",
    "• à la moindre modification du texte, les blancs se déplacent et laissent des trous en milieu de page ou des pages à moitié vides ;",
    "• ils polluent le mode Plan, la table des matières et les exports PDF ou EPUB ;",
    "• ils cassent l'accessibilité : un lecteur d'écran annonce chaque paragraphe vide, ce qui rend la lecture pénible pour une personne malvoyante.",
    "La règle professionnelle est donc : jamais d'appui sur Entrée pour aérer un document, l'espacement se règle dans le style.",
]
A["niveau"] = [
    "Sa mise en forme change instantanément et intégralement. Le paragraphe « DÉFINIR LES LIENS ENTRE HUMAIN ET MACHINE » adopte toutes les caractéristiques de Titre 1 : corps plus grand, couleur bleu foncé, et surtout le saut de page avant défini à l'exercice 1 — il démarre donc maintenant sur une page neuve. Sa numérotation suit aussi : il devient « SECTION III : ».",
    "Le texte n'a pas bougé d'un seul caractère. Seul son NIVEAU hiérarchique a changé, et toute la mise en forme a suivi automatiquement. C'est exactement ce que le mode Plan permet de faire : raisonner en structure, pas en apparence.",
]
A["maj_tdm"] = [
    "Le sommaire se met à jour tout seul : l'intitulé « Conclusion » devient « Conclusion générale » dans la table des matières, et les numéros de page sont recalculés au passage.",
    "C'est tout l'intérêt d'une table des matières générée par CHAMP plutôt que tapée à la main : elle reste cohérente avec le corps du document quoi qu'on modifie. Seule contrainte : Word ne la rafraîchit pas en continu, il faut penser à faire clic droit > Mettre à jour les champs > Mettre à jour toute la table avant chaque impression ou export.",
]

# --- identite en tete de document ---
ident = ("<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"120\"/><w:jc w:val=\"center\"/>"
         "<w:rPr><w:b/><w:color w:val=\"1F4E79\"/><w:sz w:val=\"28\"/></w:rPr></w:pPr>"
         "<w:r><w:rPr><w:b/><w:color w:val=\"1F4E79\"/><w:sz w:val=\"28\"/></w:rPr>"
         "<w:t xml:space=\"preserve\">[NOM Prénom] \u2014 [N° étudiant]</w:t></w:r></w:p>")

# --- insertions, de la fin vers le debut pour ne pas decaler les positions ---
inserts = []
inserts.append((anchor("Notez le nombre d"), para(A["instances"])))
inserts.append((anchor("Quel est son style par d"), para(A["style_defaut"])))
i_obs1 = anchor("Qu'observez-vous sur l'ensemble du texte")
inserts.append((i_obs1, para(A["observez_normal"])))
inserts.append((anchor("Quel est l'effet imm"), para(A["saut_page"])))
inserts.append((anchor("changent-ils automatiquement de couleur"), para(A["heritage"])))
inserts.append((anchor("En trouvez-vous"), para(A["titre4"])))
inserts.append((anchor("Pourquoi la pr"), para(A["para_vides"])))
inserts.append((anchor("Qu'advient-il de sa mise en forme"), para(A["niveau"])))
# le second "Qu'observez-vous" (exercice 5), apres la mise a jour des champs
inserts.append((anchor("Qu'observez-vous", occurrence=1, after_index=i_obs1), para(A["maj_tdm"])))

# identite : juste apres le tout premier paragraphe
inserts.append((spans[0][1], ident))

for pos, txt in sorted(inserts, key=lambda x: -x[0]):
    xml = xml[:pos] + txt + xml[pos:]

open(p, "w", encoding="utf-8").write(xml)

import xml.dom.minidom as md
md.parse(p)

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n in names:
        z.write(os.path.join(WORK, n.replace("/", os.sep)), n)

print("OK :", OUT, os.path.getsize(OUT), "octets,", len(inserts), "insertions")
