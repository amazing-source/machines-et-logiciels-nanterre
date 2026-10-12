# -*- coding: utf-8 -*-
# Seance 5 : feuille d'exercices completee (activite 1 + projet personnel).
# Les tailles de fichiers citees sont relevees sur les fichiers reels.
# Usage : python feuille5.py <enonce.docx> <sortie.docx> <identite.json>
import json
import os
import re
import sys

from feuille_commun import Feuille

SRC, OUT = sys.argv[1], sys.argv[2]
ID = json.load(open(sys.argv[3], encoding="utf-8"))
NB = "\u00a0"
ML = ID["dossier_rendu"]
PROJET = "PROJET_MULTIMEDIA_" + ID["fichier_numero"]


def typo(t):
    t = re.sub(r" ([?:;!»%])", NB + r"\1", t)
    t = t.replace("« ", "«" + NB)
    t = re.sub(r"(\d) (\d{3})", r"\1" + NB + r"\2", t)
    return t


def taille(nom):
    """Taille d'un fichier reel, en Ko ou Mo (base 1000), a la francaise."""
    o = os.path.getsize(os.path.join(ML, nom))
    if o >= 1_000_000:
        return ("%.1f Mo" % (o / 1_000_000)).replace(".", ",")
    return "%d Ko" % round(o / 1000)


EXP = "SEANCE5_EXPORTS"
t_pptx = taille("SEANCE5_MULTIMEDIA.pptx")
t_ppsx = taille(os.path.join(EXP, "SEANCE5_MULTIMEDIA.ppsx"))
t_pdf = taille(os.path.join(EXP, "SEANCE5_MULTIMEDIA.pdf"))
t_hd = taille(os.path.join(EXP, "SEANCE5_MULTIMEDIA_HD_720p.mp4"))
t_fhd = taille(os.path.join(EXP, "SEANCE5_MULTIMEDIA_FullHD_1080p.mp4"))
o_hd = os.path.getsize(os.path.join(ML, EXP, "SEANCE5_MULTIMEDIA_HD_720p.mp4"))
o_fhd = os.path.getsize(os.path.join(ML, EXP, "SEANCE5_MULTIMEDIA_FullHD_1080p.mp4"))
plus = round(100 * (o_fhd / o_hd - 1))

f = Feuille(SRC)


def rep(fragment, *lignes, occurrence=1):
    f.repondre(fragment, [typo(l) for l in lignes], occurrence)


f.apres(f.premier(),
        f.identite("%s \u2014 %s" % (ID["nom_prenom"], ID["numero"])),
        f.reponse([typo(l) for l in (
            "FICHIERS RENDUS",
            "• Activité 1 : SEANCE5_MULTIMEDIA.pptx (16 diapositives), avec cette feuille d'exercices.",
            "• Projet : %s.xlsx (le classeur de données) et %s.pptx (la présentation, 6 diapositives)." % (PROJET, PROJET),
            "Tous sont enregistrés dans le dossier MACHINES ET LOGICIELS créé en séance 1.",
        )]))

# ------------------------------------------------------------------ EXERCICE 1
rep("Revenir en mode d'affichage normal",
    "Fait :",
    "• nouvelle présentation créée à partir du modèle Office « Facette », enregistrée au format .pptx sous le nom SEANCE5_MULTIMEDIA dans MACHINES ET LOGICIELS ;",
    "• les 15 diapositives de « Numérique & Environnement » copiées dans ce fichier : elles ont pris le thème de destination ;",
    "• onglet Création : variante verte du thème, en accord avec le sujet ;",
    "• dispositions : « Diapositive de titre » pour la première, « Titre et contenu » pour les autres. Les titres des diapositives 11 à 14 n'étaient que de simples zones de texte dans l'original : je les ai replacés dans l'espace réservé du titre pour qu'ils suivent le masque ;",
    "• mise en page : les objets de l'original débordaient sur le décor du thème, à droite ; je les ai recalés et j'ai harmonisé les tailles de texte ;",
    "• masque des diapositives : jeu de couleurs « Vert » choisi dans le groupe Arrière-plan, style d'arrière-plan n° 1 (fond blanc), pied de page et numéro de diapositive agrandis en corps 11.",
    "Ce que j'en retiens : une modification faite dans le masque se répercute sur toutes les diapositives, c'est ce qui garantit l'unité visuelle.",
    "Modes d'affichage explorés : Normal (travail courant), Plan (le texte seul), Trieuse de diapositives (vue d'ensemble, réorganisation), Page de notes (commentaires de l'orateur), Lecture. Côté masques : masque des diapositives, masque du document, masque des pages de notes.")

# ------------------------------------------------------------------ EXERCICE 2
rep("Transformez ensuite le texte des diapositives 10",
    "Fait :",
    "• pied de page « Année universitaire 2026-2027 » et numéro de diapositive, avec « Appliquer partout » (16 diapositives) ;",
    "• diapositive 2 : lien https://pix.fr/ sur le mot « Pix » et lien https://pix.fr/competences sur « Protection et sécurité » ;",
    "• diapositive 3 : SmartArt « Liste de blocs simple », 9 blocs, un par thématique, saisis dans le volet Texte ;",
    "• diapositive 8 : icône d'ampoule placée devant « Le saviez-vous ? » ;",
    "• diapositive 5 : zone de texte insérée à droite de la figure, avec le texte demandé, puis style rapide « Effet modéré » dans la couleur d'accentuation 1 ; figure et texte réalignés côte à côte ;",
    "• diapositive 9 : les trois puces remplacées par un graphique en secteurs (30 %, 30 %, 40 %) saisi dans la feuille Excel associée ; titre, légende à droite et étiquettes en pourcentage vérifiés ;",
    "• diapositives 10 à 14 de l'original (devenues 10 et 12 à 15 après l'insertion de l'exercice 3) : texte converti en SmartArt « Liste de blocs verticale » par Accueil > Paragraphe > Convertir en graphique SmartArt, le thème à gauche, les éco-gestes à droite.")

# ------------------------------------------------------------------ EXERCICE 3
rep("Insérer un titre et sa source",
    "Fait :",
    "• diapositive 8 : image d'empreinte carbone (pictogramme d'une empreinte de pied marquée CO2) insérée à droite du texte, avec sa source en légende ;",
    "• nouvelle diapositive insérée après la 10, disposition « Contenu avec légende » : la même image y est collée et redimensionnée par un coin, donc sans déformation (proportions verrouillées) ;",
    "• titre : « L'empreinte carbone : la trace de nos usages » ; source : « Carbon footprint icon », Tommaso.sansone91, Wikimedia Commons, licence CC0 (domaine public).",
    "J'ai retenu une image du domaine public pour pouvoir la réutiliser sans restriction ; la source est citée malgré tout.")

# ------------------------------------------------------------------ EXERCICE 4
rep("Visualiser l'effet de la transition",
    "Fait : transition « Pousser » (la nouvelle diapositive pousse l'ancienne hors de l'écran), durée 1 seconde, passage au clic, puis « Appliquer partout ».",
    "Les 16 diapositives ont la même transition : une seule transition, sobre, évite l'effet de catalogue.")

# ------------------------------------------------------------------ EXERCICE 5
rep("des animations ajoutées avec l'aperçu",
    "Fait, avec l'effet d'apparition « Fondu » :",
    "• diapositive 5 : la figure apparaît au clic ; la zone de texte s'affiche ensuite toute seule (Démarrer : Après la précédente ; Délai : 4 s) ;",
    "• diapositive 7 : la première zone de texte apparaît au clic ; la seconde, le commentaire du graphique, suit automatiquement (Après la précédente ; Délai : 4 s).",
    "La diapositive 5 ne comportant qu'une seule zone de texte en plus du titre, c'est la figure qui y joue le rôle du premier élément.",
    "Dans le volet Animation, on retrouve bien l'ordre : le chiffre 1 pour l'effet au clic, puis une horloge pour l'effet « après la précédente ».")

# ------------------------------------------------------------------ EXERCICE 6
rep("Exportez ensuite au format vidéo",
    "Fait, dans le sous-dossier SEANCE5_EXPORTS :",
    "• .ppsx (%s) : même contenu que le .pptx, mais un double-clic lance directement le diaporama en plein écran, sans l'interface de modification ; transitions et animations conservées ;" % t_ppsx,
    "• .pdf (%s) : format figé et universel, une page par diapositive. Les options permettent d'encadrer les diapositives, d'ajouter les pages de notes ou de placer plusieurs diapositives par page. En revanche les transitions et les animations sont perdues : chaque diapositive est exportée dans son état final ; seuls les liens hypertextes restent actifs ;" % t_pdf,
    "• .mp4 HD 720p (%s) et Full HD 1080p (%s) : la vidéo conserve transitions, animations et minutage, mais plus rien n'est modifiable ni cliquable. Le passage de HD à Full HD alourdit le fichier de %d %% pour un gain peu visible sur ce contenu." % (t_hd, t_fhd, plus),
    "Le .pptx (%s) reste le seul format modifiable." % t_pptx)

# ------------------------------------------------------------------ PROJET
rep("Explorez les sites, choisissez une thématique",
    "Jeu de données choisi : « Parcoursup 2025 — vœux de poursuite d'études et de réorientation dans l'enseignement supérieur et réponses des établissements », publié en données ouvertes par le ministère chargé de l'Enseignement supérieur (SIES) sur data.enseignementsup-recherche.gouv.fr, sous Licence Ouverte.",
    "Téléchargé le 12/10/2026 au format CSV, filtré sur la région Île-de-France : 2 691 formations, 8 départements. C'est la version officielle et complète du thème Parcoursup proposé par l'énoncé.")

rep("Définissez une problématique à votre jeu de données",
    "Problématique : en Île-de-France, dans quelles filières est-il le plus difficile d'obtenir une place ?",
    "Question complémentaire : la situation est-elle la même à Paris, en petite couronne et en grande couronne ?",
    "Indicateur retenu : le nombre de candidatures par place offerte.")

rep("TCD (Tableau croisé dynamique)",
    "Outils utilisés dans le classeur %s.xlsx :" % PROJET,
    "• formules : RECHERCHEV (zone géographique à partir du département), SIERREUR (candidatures ÷ places, admis ÷ places), SI imbriqués avec des seuils en références absolues (niveau de tension), NB.SI, SOMME.SI et SOMME.SI.ENS (tableau de synthèse) ;",
    "• tri à deux niveaux (Filière de A à Z, puis candidatures par place décroissant) et filtres automatiques ; filtre avancé avec extraction pour isoler les 87 formations de l'Université Paris Nanterre ;",
    "• sous-totaux par filière (Données > Sous-total) sur les places, les candidatures et les admis ;",
    "• deux tableaux croisés dynamiques avec un champ calculé « candidatures par place » : par filière, puis filière × zone ;",
    "• mise en forme conditionnelle : barres de données et couleur selon le niveau de tension.")

rep("présenter une synthèse chiffrée",
    "Synthèse (feuilles Synthèse et TCD) : 153 229 places pour 4 197 242 candidatures, soit 27,4 candidatures par place en moyenne.",
    "Par filière : PASS 81,1 ; IFSI 58,4 ; BUT 40,9 ; Licence L.AS 40,8 ; CPGE 38,6 ; BTS 28,9 ; école d'ingénieur 26,4 ; école de commerce 20,8 ; autre formation 20,2 ; licence 17,5 ; EFTS 10,4.",
    "Par zone : Paris 29,3 ; petite couronne 27,9 ; grande couronne 23,5.",
    "Les formules, le tableau croisé dynamique et les sous-totaux donnent les mêmes totaux : c'est mon contrôle de cohérence.")

rep("faire un graphique répondant à votre problématique",
    "Graphique en barres horizontales (feuille Synthèse), trié de la filière la plus demandée à la moins demandée.",
    "Réponse à la problématique : les filières de santé sont de loin les plus disputées (PASS, IFSI), alors que la licence, qui offre 4 places sur 10, est parmi les moins tendues ; Paris est plus tendu que la grande couronne.",
    "Limite : un candidat formule plusieurs vœux ; le ratio mesure donc l'attractivité d'une filière, pas la chance d'y être admis.")

rep("Déposez le fichier sur coursenligne",
    "Présentation : %s.pptx (le nom contient mon numéro d'étudiant), 6 diapositives :" % PROJET,
    "1. Titre — sujet des données, nom et prénom, année universitaire et filière ;",
    "2. Problématique et démarche ;",
    "3. Les données — capture du classeur Excel, description des données et des quatre calculs ajoutés ;",
    "4. Tri et filtre — capture du filtre avancé sur mon université ;",
    "5. Le tableau de synthèse — capture du tableau croisé dynamique et texte court ;",
    "6. Le graphique — figure générée à partir de la synthèse, avec son titre et la réponse à la problématique.",
    "Masque des diapositives personnalisé : couleurs et polices du thème, position et style du titre, pied de page (nom, numéro d'étudiant, sujet), numéro de diapositive, pastille de couleur, et trois dispositions renommées, dont une couverture à fond bleu nuit.",
    "Transition « Fondu » sur toutes les diapositives ; animations d'apparition calées sur la narration.",
    "Narration : un fichier audio de 27 à 35 secondes par diapositive (Insertion > Audio), lu automatiquement ; son texte figure dans les pages de notes.")

print("paragraphes vides retires en fin de document :", f.nettoyer_fin())
f.enregistrer(OUT)
