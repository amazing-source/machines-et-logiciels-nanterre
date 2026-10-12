# Instructions pour un assistant IA

Tu aides un étudiant de l'Université Paris Nanterre à préparer ses rendus de l'EC « Compétences numériques : Machines et logiciels » à partir de ce kit.

**Ton objectif** : produire dans `a-deposer/` des fichiers à son nom, prêts à être déposés, sans une seule case entre crochets, puis lui dire clairement ce qu'il doit encore faire lui-même.

Parle-lui en français, simplement. Il n'a pas besoin de comprendre les scripts.

## Règles

1. **N'invente rien.** Identité, valeurs de sa machine, mesures : tu les demandes ou tu les relèves. Une case sans valeur sûre reste vide, et tu le signales.
2. **Ne modifie pas les modèles** des dossiers `seance-N/`. Tout se passe dans `a-deposer/`.
3. **N'écrase jamais un fichier de `a-deposer/`** : il contient peut-être la narration enregistrée par l'étudiant. Pas de `--ecraser` sans son accord explicite.
4. **Ne publie rien.** `outils/identite.json` et `a-deposer/` contiennent ses données personnelles ; git les ignore, laisse-les ainsi.
5. Trois choses ne sont pas à ta portée : enregistrer sa voix, déposer sur coursenligne, faire le parcours PIX. Ne prétends pas les avoir faites.

## Étape 1 — L'identité

Demande à l'étudiant, en une seule fois :

- son **NOM** et son **prénom** ;
- son **numéro d'étudiant** ;
- sa **filière et son année** (exemple : `L1 Économie-Gestion`).

Copie `outils/identite.exemple.json` vers `outils/identite.json` et remplis :

| Clé | Valeur | Exemple |
|---|---|---|
| `nom_prenom` | NOM en capitales, puis prénom | `DUPONT Marie` |
| `prenom_nom` | prénom, puis nom | `Marie Dupont` |
| `signature` | initiale du prénom, puis nom | `M. Dupont` |
| `numero` | numéro d'étudiant | `12345678` |
| `filiere` | filière et année | `L1 Économie-Gestion` |
| `fichier_numero` | le même numéro (il entre dans le nom des fichiers du projet) | `12345678` |

Garde `universite` et `annee` tels quels, sauf s'il te dit autre chose. Les clés `dnt_*` se remplissent à l'étape 3 ; les autres ne servent qu'aux scripts facultatifs.

## Étape 2 — Remplir

```
python outils/remplir.py
```

Python 3 suffit, aucune bibliothèque à installer. Le script copie chaque modèle dans `a-deposer/seance-N/`, remplace les cases d'identité partout (texte, pieds de page, notes, auteur du fichier, classeurs des graphiques), renomme les fichiers, puis affiche l'état de chaque fichier.

Après cette étape, les séances 2, 3, 4 et 5 sont complètes. Il reste des cases dans les feuilles des séances 1 et 6.

## Étape 3 — Les valeurs de sa machine

Elles sont listées dans `a-deposer/cases.json` : une entrée par case, avec la phrase qui l'entoure (`contexte`, la case visée entre `>>>` et `<<<`). Écris chaque valeur dans `"valeur"`, puis :

```
python outils/remplir.py --appliquer
```

- `null` : la case reste à remplir. `""` : la case est supprimée — c'est le réglage par défaut des remarques du modèle (« Vérifie… ») ; lis-les avant d'appliquer.
- Tu peux procéder en plusieurs passes : après chaque `--appliquer`, `cases.json` est régénéré avec ce qui reste.
- Ne modifie que `"valeur"`.

### Séance 1 (bonus) — 36 cases

| Cases | Où trouver la valeur |
|---|---|
| Système, processeur, mémoire, stockage | Sur la machine. Sous Windows : `Get-CimInstance Win32_OperatingSystem` (édition, version), `Win32_Processor` (nom), `Win32_ComputerSystem` (mémoire), `Get-PSDrive C` (espace libre et total) |
| `[ton nom d'utilisateur]` | le nom du dossier utilisateur (`$env:USERNAME`) |
| `[tonnom]`, `[ton graphique]`, `[nom de ton fichier]` | les noms de ses fichiers dans `Documents\MACHINES ET LOGICIELS` |
| `[largeur]`, `[hauteur]`, `[taille]` | les dimensions en pixels et la taille en octets de **ses** fichiers : les deux `.bmp`, le `.jpg`, le `.png`, leurs archives `.7z` et `.zip`, `recettefondant.docx` et `.txt` |
| Logiciel qui s'ouvre pour le `.bmp` et le `.jpg`, `[ta couleur]` | demande-lui |

Ces fichiers naissent des activités 4 et 5 de l'énoncé (un quart d'heure dans Paint et 7-Zip). S'ils n'existent pas : soit l'étudiant fait ces deux activités, soit il te demande de créer les fichiers — dans les deux cas, les tailles écrites sont celles de fichiers réels. Repère de contrôle : un bitmap 24 bits pèse `largeur × hauteur × 3 + 54` octets, à quelques octets de remplissage près.

### Séance 6 (bonus) — 4 cases

Ce que le site [donottrack-doc.com](https://donottrack-doc.com/fr/) affiche de lui (épisode 1) : sa ville et son pays, son fournisseur d'accès, son navigateur, son système. Demande-les-lui, ou relève-les avec son accord. Tu peux les écrire dans `identite.json` avant l'étape 2 (`dnt_ville`, `dnt_fai`, `dnt_navigateur`, `dnt_systeme`) ou dans `cases.json` après.

Les autres réponses de l'activité 6 sont de vraies mesures, faites le 12 octobre 2026 sur un autre poste (`seance-6/mesures/RESULTATS.md`). Pour les refaire sur le sien : `outils/README.md`, partie « Séance 6 ».

## Étape 4 — Contrôle

```
python outils/remplir.py --verifier
```

Le code de sortie vaut 0 quand il ne reste aucune case. Le script signale aussi la narration manquante du projet : c'est normal, elle revient à l'étudiant.

Si tu peux ouvrir les fichiers, regarde la diapositive 1 et le pied de page de `PROJET_MULTIMEDIA_….pptx` : un nom long peut déborder.

## Étape 5 — Ce que tu dis à l'étudiant

Termine par un message court :

1. **Où sont ses fichiers**, devoir par devoir (tableau ci-dessous).
2. **Ce qu'il lui reste à faire lui-même**, bien visible :
   - 🔴 **enregistrer sa narration** sur les 6 diapositives de `PROJET_MULTIMEDIA_….pptx` : *Diaporama > Enregistrer > À partir du début*, 10 à 40 secondes par diapositive. Le texte à dire est dans les notes de chaque diapositive (détails : `seance-5/README.md`) ;
   - 🔴 **déposer** chaque fichier sur coursenligne, dans l'espace de **son groupe de TD**, avant la date limite ;
   - 🔴 **finir le parcours « Tout PIX »** : s'il n'est pas terminé, la note est zéro ;
   - ouvrir chaque fichier une fois avant de le déposer.
3. **Les cases que tu as laissées vides**, s'il y en a, et pourquoi.

## Les devoirs

| Devoir sur coursenligne | Statut | Fichiers de `a-deposer/` |
|---|---|---|
| Devoir 1 : Environnement numérique | bonus | `seance-1/SEANCE 1 - Reponses - ….docx` |
| Devoir 2.1 : TTX court | obligatoire | `seance-2/SEANCE 2 - Reponses - ….docx` + `SEANCE2_ACTIVITE1_TTX.docx` |
| Devoir 2.2 : TTX long | obligatoire | `seance-2/SEANCE 2 - Reponses - ….docx` + `SEANCE2_ACTIVITE2_TTX.docx` |
| Séance 3, partie 1 | obligatoire | `seance-3/SEANCE 3 - Reponses - ….docx` + `…_TD_Excel.xlsx` |
| Séance 3, partie 2 | bonus | les deux mêmes fichiers |
| Séance 4, partie 1 | obligatoire | `seance-4/SEANCE 4 - Reponses - ….docx` + `…_TD_Excel_Approfondissement.xlsx` |
| Séance 4, partie 2 | bonus | les deux mêmes fichiers |
| Devoir 5 : Multimédia (activité 1) | bonus | `seance-5/SEANCE 5 - Reponses - ….docx` + `SEANCE5_MULTIMEDIA.pptx` |
| Devoir 5 : Projet personnel | obligatoire, noté | `seance-5/PROJET_MULTIMEDIA_….xlsx` + `PROJET_MULTIMEDIA_….pptx` (avec sa narration) |
| Devoir 6 : S'insérer dans un monde numérique | bonus | `seance-6/SEANCE 6 - Reponses - ….docx` |

Les dates limites sont dans le `README.md` de la racine ; elles valent pour le groupe TD02, l'étudiant doit vérifier celles de son groupe.

## S'il veut aller plus loin

- **Un projet à lui** (séance 5) : l'énoncé demande un sujet personnel. `seance-5/README.md`, partie « Changer de question », donne cinq questions faisables avec le même fichier de données, et la liste de ce qui est alors à refaire.
- **Régénérer les fichiers par script** (Windows et Microsoft Office en français) : `outils/README.md`. Les pièges du pilotage d'Office sont dans `outils/NOTES-TECHNIQUES.md`.
- **Ses propres mots** : les réponses rédigées des feuilles sont celles d'un autre étudiant. S'il te demande de les reformuler, garde le fond et les chiffres.
