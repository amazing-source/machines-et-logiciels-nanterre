# Séance 5 — Documents multimédia

## À déposer

| Devoir | Statut | Fichiers |
|---|---|---|
| **DEVOIR 5 : PROJET PERSONNEL** | **obligatoire, noté** | `PROJET_MULTIMEDIA_NUMERO.xlsx` + `PROJET_MULTIMEDIA_NUMERO.pptx` **avec ta voix** |
| DEVOIR 5 : MULTIMEDIA (activité 1) | bonus | `SEANCE 5 - Reponses - NOM_Prenom.docx` + `SEANCE5_MULTIMEDIA.pptx` |

`NUMERO` = ton numéro d'étudiant : l'énoncé veut qu'il figure dans le nom des deux fichiers du projet. Le 12 octobre 2026, le devoir du projet affichait encore la date de l'an dernier : surveille la vraie.

> [!CAUTION]
> **🔴 À faire toi-même : la narration du projet.**
> Le devoir exige « une narration orale », de 10 à 40 secondes par diapositive. La présentation est livrée **sans son** : tant que tu n'as pas enregistré ta voix sur les 6 diapositives, le projet est incomplet.

## Ta narration

Dix minutes, dans PowerPoint, avec le micro de ton ordinateur ou d'un casque.

1. Ouvre `PROJET_MULTIMEDIA_….pptx`. Le texte à dire est dans les **notes** de chaque diapositive (*Affichage > Page de notes*) et dans [`narration/texte_de_la_narration.txt`](narration/texte_de_la_narration.txt). Adapte-le pour qu'il sonne comme toi.
2. *Diaporama > Enregistrer > À partir du début* : parle, passe à la diapositive suivante, et ainsi de suite jusqu'à la sixième.
3. Relance le diaporama pour t'écouter. Une diapositive ratée se refait seule : *Diaporama > Enregistrer > À partir de la diapositive actuelle*.
4. Enregistre le fichier : c'est celui-là que tu déposes.

Les animations sont réglées pour une trentaine de secondes par diapositive. Si tu parles moins longtemps, réduis les délais dans *Animations > Volet Animation*.

## Le projet du modèle

- **Données** : Parcoursup 2025, données ouvertes du ministère, 2 691 formations d'Île-de-France.
- **Question** : dans quelles filières est-il le plus difficile d'obtenir une place ? Est-ce pareil à Paris et en banlieue ?
- **Réponse** : 27,4 candidatures par place en moyenne ; 81,1 en PASS, 58,4 en IFSI, 17,5 en licence ; 29,3 à Paris contre 23,5 en grande couronne.

Tout ce que l'énoncé exige y est :

| Exigence | Dans le modèle |
|---|---|
| 5 à 6 diapositives, dont un titre avec nom, année et filière | 6 diapositives |
| Un masque défini par toi | masque personnalisé : couleurs, polices, pied de page, trois dispositions |
| Capture du fichier Excel + texte sur les données et les calculs | diapositive 3 |
| Capture d'un tableau de synthèse + texte court | diapositive 5 (tableau croisé dynamique) |
| Graphique final, avec un titre | diapositive 6 |
| Transitions et animations | fondu, apparitions échelonnées |
| Excel : formules, tri, filtre, sous-totaux, TCD, synthèse, graphique | les sept feuilles du classeur |
| **Narration** | **🔴 à enregistrer** |

Diapositive 4 : l'exemple cité est la licence MIASHS de Paris Nanterre. Remplace-le par ta formation si tu veux (feuille `Zoom Paris Nanterre` du classeur).

## Changer de question

Le devoir s'appelle « projet personnel ». Le même fichier de données répond à bien d'autres questions, et le tableau croisé se réorganise en glissant les champs :

| Ta question | Comment l'obtenir |
|---|---|
| Les formations sélectives sont-elles plus demandées ? | Feuille `TCD` : glisse « Sélectivité » de *Filtres* vers *Colonnes* |
| Quelles filières sont les plus disputées dans mon département ? | Feuille `TCD` : mets « Département » dans *Filtres* et choisis le tien |
| Quelles filières ne remplissent pas leurs places ? | Fais de « Taux de remplissage » la valeur du TCD (moyenne) |
| Où se situe ma formation, ou une autre université ? | Feuille `Zoom Paris Nanterre` : change le critère, relance *Données > Avancé* |
| Les dix formations les plus demandées de ma filière ? | Feuille `Données` : filtre sur la filière, tri décroissant sur « Candidatures par place » |

Ce qui est alors à refaire : le tableau de synthèse et le graphique dans Excel ; les captures d'écran (*Windows + Maj + S*, puis clic droit sur l'image de la diapositive > *Modifier l'image* : taille et animation sont conservées) ; le graphique de la diapositive 6 (*clic droit > Modifier les données*) ; les titres et les chiffres des diapositives 2 à 6 ; le texte de ta narration.

Tu peux aussi partir d'une autre région sur [le portail](https://data.enseignementsup-recherche.gouv.fr/explore/dataset/fr-esr-parcoursup/), ou d'un autre jeu de données (data.gouv.fr, SNCF, RATP…) : la démarche reste la même.

## Activité 1 (bonus)

`SEANCE5_MULTIMEDIA.pptx` est la présentation « Numérique & Environnement » de l'énoncé, mise en forme : 16 diapositives, thème Facette vert, pied de page, liens, SmartArt, graphique en secteurs, image sourcée, transition « Pousser », animations sur les diapositives 5 et 7. La feuille raconte chaque exercice.

L'énoncé laisse le thème, la transition et l'image « à votre convenance » : onglets *Création* et *Transitions*. Si tu changes de thème, repasse sur chaque diapositive pour vérifier que rien ne déborde. L'exercice 6 (exports `.ppsx`, `.pdf`, `.mp4`) n'est pas à déposer.

## Si tu remplis à la main

| Fichier | Où |
|---|---|
| Présentation du projet | diapositive 1 ; pied de page (*Insertion > En-tête/Pied > Appliquer partout*) ; ligne de source de la diapositive 6 ; notes de la diapositive 1 |
| Classeur du projet | feuille `Lisez-moi` |
| Présentation de l'activité 1 | diapositive 1 |
| Tous | *Fichier > Informations > Auteur*, et le nom du fichier |

## Les autres dossiers

`donnees/` : l'extrait d'origine en CSV et le dictionnaire des champs (Licence Ouverte 2.0) · `captures/` : les trois captures de la présentation · `images/` : l'ampoule et l'image d'empreinte carbone (CC0) · `narration/` : le texte à dire.
