# Outils

## `remplir.py` — mettre le kit à ton nom

Python 3 suffit, sur Windows, Mac ou Linux. Depuis la racine du dépôt :

```
python outils/remplir.py              # copie tout dans a-deposer/, à ton nom
python outils/remplir.py --appliquer  # reporte les valeurs écrites dans a-deposer/cases.json
python outils/remplir.py --verifier   # liste ce qui reste à compléter
```

1. Copie `identite.exemple.json` sous le nom `identite.json` et remplace `nom_prenom`, `numero` et `filiere`. Le reste est facultatif. Ce fichier est ignoré par git.
2. Lance `remplir.py`. Les modèles ne sont pas touchés ; un fichier déjà présent dans `a-deposer/` n'est jamais écrasé.
3. Les cases qui dépendent de ta machine (séances 1 et 6) sont listées dans `a-deposer/cases.json` : écris chaque valeur dans `"valeur"`, puis lance `--appliquer`. Ou remplis-les directement dans Word.

> [!CAUTION]
> 🔴 Le script ne remplace pas ta voix : la narration du projet reste à enregistrer dans PowerPoint ([séance 5](../seance-5/README.md#ta-narration)).

---

## Refaire les fichiers par script (facultatif)

Les fichiers des séances 5 et 6 ont été fabriqués par les scripts de ce dossier. Tu n'en as besoin que pour changer le projet en profondeur ou refaire les mesures.

**Il faut** : Windows 10 ou 11, Microsoft Office en français, PowerShell 5.1, Python 3 avec `pip install openpyxl lxml requests` ; pour les mesures de la séance 6, Node.js et Google Chrome.

Les scripts pilotent les vraies applications Office : **ferme Word, Excel et PowerPoint** et laisse la machine tranquille pendant qu'ils tournent. Dans PowerShell, depuis la racine du dépôt :

```powershell
Set-ExecutionPolicy -Scope Process Bypass     # autorise les scripts pour cette fenêtre
$racine = (Get-Location).Path
$numero = "12345678"                          # ton numéro d'étudiant
```

Dans `identite.json`, `dossier_rendu` est le dossier *MACHINES ET LOGICIELS* de tes *Documents*.

### Séance 5 — le projet

```powershell
$classeur = "$racine\seance-5\PROJET_MULTIMEDIA_$numero.xlsx"

# le classeur : données, formules, tri, filtres, sous-totaux, TCD, graphique, puis contrôle
.\outils\projet5_build.ps1 -Work "$env:TEMP\projet5" -Dst $classeur

# la présentation, sans son
.\outils\projet5_pptx_build.ps1 -Xlsx $classeur -Out "$racine\seance-5\PROJET_MULTIMEDIA_$numero.pptx" -SansVoix
```

Sans `-SansVoix`, la présentation reçoit une voix de synthèse Windows. Elle ne remplace pas la narration demandée : enregistre la tienne.

`projet5_pptx.ps1` relit dix-sept chiffres dans le classeur et **s'arrête si l'un d'eux ne correspond plus au texte des diapositives**. Si tu changes de question, c'est normal : réécris les textes dans le script, ou plus simplement à la main dans PowerPoint.

Captures d'écran d'Excel (seule la fenêtre d'Excel est photographiée) :

```powershell
.\outils\capture_excel.ps1 -Path $classeur -Sheet "Données" -Out "$racine\seance-5\captures\capture_donnees.png" -Zoom 120 -Select K2 -WidthPx 1900 -HeightPx 590
.\outils\capture_excel.ps1 -Path $classeur -Sheet "Zoom Paris Nanterre" -Out "$racine\seance-5\captures\capture_zoom_nanterre.png" -Zoom 110 -Select A6 -WidthPx 1560 -HeightPx 690
.\outils\capture_excel.ps1 -Path $classeur -Sheet "TCD" -Out "$racine\seance-5\captures\capture_tcd.png" -Zoom 130 -TopLeft A3 -Select A3 -WidthPx 990 -HeightPx 740
```

Retélécharger les données : `python outils\projet5_telecharge.py seance-5\donnees`.

### Séance 5 — l'activité 1

Télécharge d'abord sur coursenligne la présentation fournie (`NumériqueEnvironnement.pptx`).

```powershell
.\outils\seance5_activite1.ps1 -Src "C:\chemin\vers\NumériqueEnvironnement.pptx" -Images "$racine\seance-5\images" -Identite "$racine\outils\identite.json" -Out "$racine\seance-5\SEANCE5_MULTIMEDIA.pptx"

# exercice 6 : exports .ppsx, .pdf et .mp4
.\outils\seance5_exports.ps1 -Path "$racine\seance-5\SEANCE5_MULTIMEDIA.pptx" -OutDir "$racine\seance-5\exports"
```

### Les feuilles d'exercices des séances 5 et 6

Les scripts insèrent les réponses dans l'énoncé officiel, à télécharger sur coursenligne. Les réponses sont en clair dans `feuille5.py` et `feuille6.py` : c'est là qu'elles se réécrivent.

```powershell
.\outils\feuilles_build.ps1 -Seance 6 -Enonce "C:\chemin\vers\l-enonce-seance-6.docx" -Out "$racine\seance-6\SEANCE 6 - Reponses - NOM_Prenom.docx"
```

La feuille de la séance 5 cite la taille de tes exports : lance `seance5_exports.ps1` avant, et vérifie que `dossier_rendu` contient `SEANCE5_MULTIMEDIA.pptx` et le sous-dossier `SEANCE5_EXPORTS`.

### Séance 6 — refaire les mesures

```powershell
cd outils
npm install playwright-core          # pilote le Chrome déjà installé
node seance6_mesures.js cookies ..\seance-6\mesures     # cookies avant et après consentement
node seance6_mesures.js clavier ..\seance-6\mesures     # navigation au clavier
node seance6_mesures.js dnt ..\seance-6\mesures         # Do Not Track (ouvre une fenêtre, son coupé)
node seance6_mesures.js page ..\seance-6\mesures "https://www.websitecarbon.com/website/parisnanterre-fr/" carbone 15000
```

Chrome s'ouvre dans un profil vierge, sans aucun de tes comptes. Le site testé se change en tête des fonctions `cookies()` et `clavier()`.

### Les fichiers

| Fichier | Rôle |
|---|---|
| `remplir.py` | met le kit à ton nom |
| `identite.exemple.json` | modèle d'identité, à copier en `identite.json` |
| `bom.ps1` | réencode un `.ps1` en UTF-8 avec BOM — **à relancer après toute modification d'un script PowerShell** |
| `projet5_telecharge.py` | télécharge l'extrait Parcoursup |
| `projet5_build.ps1` | chaîne du classeur : `projet5_classeur.py`, `projet5_excel.ps1`, `projet5_verif.ps1` |
| `capture_excel.ps1` | capture de la fenêtre Excel |
| `projet5_pptx_build.ps1` | chaîne de la présentation : `projet5_voix.ps1`, `projet5_pptx.ps1`, `proprietes.py` |
| `projet5_textes.json` | texte de la narration |
| `seance5_activite1.ps1`, `seance5_exports.ps1` | activité 1 et ses exports |
| `feuilles_build.ps1`, `feuille5.py`, `feuille6.py`, `feuille_commun.py`, `verif_docx.ps1` | feuilles d'exercices |
| `seance6_mesures.js` | mesures dans le navigateur |
| `seances-2-a-4/` | scripts d'origine des séances 2 à 4 : chemins neutralisés, ils **ne s'exécutent pas tels quels** et sont là pour la méthode |
| `NOTES-TECHNIQUES.md` | les pièges du pilotage de Word, Excel et PowerPoint |
