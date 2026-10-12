# Compétences numériques — Machines et logiciels

Kit pour l'EC « Compétences numériques : Machines et logiciels » (Université Paris Nanterre, 2026-2027) : **les six séances déjà faites**, avec des cases à remplir à la place du nom. Tu mets ton identité, tu vérifies, tu déposes.

Ce n'est pas un corrigé officiel : c'est le travail d'un étudiant, vérifié autant que possible.

> [!CAUTION]
> **🔴 Trois choses que personne ne peut faire à ta place**
>
> 1. **Enregistrer ta voix** sur la présentation du projet (séance 5) : 10 à 40 secondes par diapositive. Le modèle est livré sans narration. → [comment faire](seance-5/README.md#ta-narration)
> 2. **Déposer** tes fichiers sur coursenligne, dans l'espace de **ton groupe de TD** (son nom ressemble à `CNML_…_TD02`), pas dans l'espace du cours.
> 3. **Finir le parcours « Tout PIX »** : environ 250 questions. S'il n'est pas terminé, la note est zéro.

## Quoi déposer, et quand

| Séance | Statut | Tu déposes | Date limite (groupe TD02) |
|---|---|---|---|
| [1 — Environnement numérique](seance-1/) | bonus | la feuille | 6 octobre, 9 h 16 |
| [2 — Traitement de texte](seance-2/) | **obligatoire** | devoir 2.1 : la feuille + la recette<br>devoir 2.2 : la feuille + le rapport | **20 octobre, 23 h 59**<br>**20 octobre, 9 h 18** (le matin) |
| [3 — Tableur](seance-3/) | partie 1 **obligatoire**, partie 2 bonus | la feuille + le classeur | 10 novembre, 7 h 59 |
| [4 — Tableur avancé](seance-4/) | partie 1 **obligatoire**, partie 2 bonus | la feuille + le classeur | pas encore affichée |
| [5 — Multimédia](seance-5/) | activité 1 bonus<br>**projet obligatoire et noté** | la feuille + la présentation<br>le classeur + la présentation **avec ta voix** | pas encore affichée |
| [6 — Monde numérique](seance-6/) | bonus | la feuille | pas encore affichée |

« La feuille » = la feuille d'exercices de la séance, avec ses réponses.

Dates lues le 12 octobre 2026 dans le groupe TD02 : **vérifie celles de ton groupe**. En EAD ou en dérogatoire, tu n'as rien à déposer.

Les énoncés demandent aussi des manipulations (envoyer un lien à un camarade, régler son profil, explorer un menu…). Personne ne les vérifie : **seuls comptent les fichiers de ce tableau.**

## Mode d'emploi

### Le plus rapide : une commande

Il faut [Python 3](https://www.python.org/downloads/), rien d'autre.

1. Télécharge le dépôt (bouton vert **Code > Download ZIP**) et décompresse-le.
2. Copie `outils/identite.exemple.json` sous le nom `outils/identite.json`, et écris-y ton nom, ton numéro d'étudiant et ta filière.
3. Dans un terminal ouvert dans le dossier :

```
python outils/remplir.py
```

Tous les fichiers à ton nom, correctement nommés, sont dans **`a-deposer/`**. Le script affiche ce qu'il reste à compléter.

### Avec une IA

Ouvre le dossier dans un assistant capable de lancer des commandes (Claude Code, Codex, Cursor, Copilot…) et dis-lui :

> Lis AGENTS.md et prépare mes fichiers.

Elle te demande ton identité, remplit tout, relève les valeurs de ta machine et te dit ce qu'il te reste à faire. Les instructions qu'elle suit sont dans [`AGENTS.md`](AGENTS.md).

### À la main

Dans chaque fichier : **Ctrl + H**, cherche `[` et remplace chaque case (`[NOM Prénom]`, `[N° étudiant]`, `[Filière]`…). Pense au pied de page, à *Fichier > Informations > Auteur* et au nom du fichier (`NOM_Prenom`, `NUMERO`).

## Avant de déposer

- **Ouvre chaque fichier une fois** : un nom long peut déborder d'une zone de texte.
- **Séances 1 et 6** : quelques réponses sont des valeurs de *ta* machine (processeur, taille de tes fichiers, ta ville vue par un site). Le script les liste, le guide de la séance dit où les lire.
- **Projet de la séance 5** : il est noté et censé être personnel. Changer la question posée aux données prend une vingtaine de minutes, [les pistes sont ici](seance-5/README.md#changer-de-question).
- Les réponses rédigées sont celles de quelqu'un d'autre : relis-les, reformule si tu veux.

## Dans le dépôt

```
seance-1/ … seance-6/   les fichiers modèles et un guide court par séance
outils/                 remplir.py, et les scripts qui ont fabriqué les fichiers
AGENTS.md               les instructions pour une IA
```

## Sources

- Énoncés et supports : équipe pédagogique de l'EC (supports de cours sous CC BY-NC, N. Caron et J. Azhar). Ils ne sont pas redistribués ici, sauf le texte des feuilles d'exercices, nécessaire pour y répondre.
- Données du projet : *Parcoursup 2025 — vœux de poursuite d'études et de réorientation dans l'enseignement supérieur et réponses des établissements*, ministère chargé de l'Enseignement supérieur (SIES), [data.enseignementsup-recherche.gouv.fr](https://data.enseignementsup-recherche.gouv.fr/explore/dataset/fr-esr-parcoursup/), Licence Ouverte 2.0, version du 9 mars 2026.
- Image « Carbon footprint icon » : Tommaso.sansone91, Wikimedia Commons, CC0.
- Aucune garantie d'exactitude : vérifie avant de rendre.
