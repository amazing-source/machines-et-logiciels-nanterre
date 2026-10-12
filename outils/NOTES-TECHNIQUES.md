# Notes techniques — piloter Word, Excel et PowerPoint par script

Relevé des pièges rencontrés en automatisant Office avec PowerShell (interface COM) et Python, sur Windows 11 et Office en français. Chacun a coûté du temps ; aucun n'est documenté clairement ailleurs.

## Avant tout

- **Ne tue jamais Word, Excel ou PowerPoint par leur nom** (`Stop-Process WINWORD`) : tu fermerais un document ouvert sans l'enregistrer. Un script doit créer sa propre instance et ne quitter que celle-là. PowerPoint n'a qu'une instance : ne le quitte que si `Presentations.Count` vaut 0.
- **Fichiers téléchargés** : ils portent une marque « Internet » et s'ouvrent en mode protégé, ce qui fige le pilotage sans message. `Unblock-File` avant toute chose.
- **Verrous `~$`** : un Office interrompu laisse un fichier `~$nom.docx` qui bloque l'ouverture suivante. À supprimer.
- **Le pilotage échoue parfois sans raison reproductible**, surtout si quelqu'un se sert de la machine en même temps. Prévois de relancer l'étape entière.

## PowerShell

- **Encodage** : PowerShell 5.1 lit un `.ps1` sans BOM comme de l'ANSI. Tous les accents cassent en silence et les recherches de texte ne trouvent plus rien. Enregistre en UTF-8 **avec BOM** (`bom.ps1`).
- **Nombres en argument** : `MaFonction 12.5` transmet le nombre dans une enveloppe qui provoque « Le cast spécifié n'est pas valide » dès qu'il atteint une propriété COM. Type les paramètres : `function F([single]$taille)`.
- **Tableaux à deux dimensions** : `$a.Value2 = $b.Value2` échoue. Copie par filtre avancé sans critère, ou cellule par cellule.
- `Workbook.BuiltinDocumentProperties` lève une exception. Pour l'auteur et le titre, réécris `docProps/core.xml` (`proprietes.py`).
- **Éditer un script depuis un heredoc Bash** mange un niveau d'antislash : `\n` devient un vrai saut de ligne, `C:\\Users` devient `C:\Users`. Écris les fichiers directement.

## Excel

- **Formats de nombre** : `.NumberFormat` suit la convention **française** dans les cellules — `"# ##0,00"`, `"0,0%"`. Un espace placé après les décimales divise la valeur par 1 000 : `# ##0,00" €"` est juste, `# ##0,00 "€"` est faux.
- **Mais dans un graphique**, `NumberFormat` suit la convention **anglaise** : `"0.0"`. Avec `"0,0"`, 81,1 s'affiche 81. Mets `NumberFormatLinked = $false` avant.
- **Mise en forme conditionnelle** : `FormatConditions.Add` exige ses trois arguments, avec un opérateur valide même pour une règle par formule ; la formule est en convention locale (`"=$J5<0,15"`).
- **Ne te fie pas à `.Text`** : il renvoie l'affichage (`####` si la colonne est étroite). Vérifie avec `.Value2`.
- **Formules** : dans le fichier et dans `Range.Formula`, les fonctions sont en anglais avec des virgules (`SUMIF`, `VLOOKUP`, `IFERROR`). Excel les affiche traduites.
- `AdvancedFilter` **désactive le filtre automatique** de la plage source : pose `AutoFilter` après.
- `Chart.SetSourceData` refuse une plage en plusieurs zones : crée la série avec `NewSeries()` puis `Values` et `XValues`.
- **Volets figés impossibles** : Excel rouvre avec la taille de sa dernière fenêtre. Après un graphique PowerPoint (mini-fenêtre de données), il ne reste qu'une ligne visible et `FreezePanes` échoue. Redonne une taille à la fenêtre : `$x.WindowState = -4143; $x.Width = 1000; $x.Height = 700`.
- **Traces personnelles** : un `.xlsx` enregistré par Excel contient le chemin complet du fichier (`x15ac:absPath`) et le nom de l'utilisateur qui a actualisé chaque tableau croisé (`refreshedBy`).

## Word

- `Document.Range()` désigne toujours le corps du document. Pour un pied de page, pars de `$footer.Range`.
- `$copie = $plage.FormattedText` est une **référence vivante**, pas une copie : colle d'abord, supprime l'original ensuite.
- **Numérotation des titres** : n'applique jamais un modèle de liste à tout le contenu, tous les paragraphes deviendraient des titres. Utilise `Style.LinkToListTemplate`.
- Titre 1, 2 et 3 sont basés sur Normal, pas les uns sur les autres.
- **Export PDF** : `ExportAsFixedFormat` se fige ; `SaveAs2($chemin, 17)` fonctionne, mais reste très lent sur un document à bordure de page et formes flottantes.
- **Insérer du texte dans un énoncé** : s'il contient des zones de texte imbriquées, un découpage par expressions régulières coupe au mauvais endroit. Travaille sur l'arbre XML (`feuille_commun.py`).

## PowerPoint

- `Chart.ChartData.Workbook` vaut parfois `$null` juste après `Activate()`, et ses feuilles ne répondent pas tout de suite : boucle avec attente.
- Sur ce classeur de données, affecter un nombre à `Value2` échoue : passe par `.Formula` avec du texte (`"0.3"`).
- Dans un **titre**, le retour chariot devient un saut de ligne au sein du même paragraphe : `Paragraphs(2)` ne désigne rien.
- Le texte d'un énoncé contient des espaces insécables : `TextRange.Find("a b")` peut ne rien trouver. Délimite la plage par son premier et son dernier mot.
- `Shapes.AddPicture` accepte un SVG mais sans image de repli : convertis en PNG.
- `Shape.ConvertTextToSmartArt($disposition)` fonctionne ; les niveaux de retrait des paragraphes deviennent ceux du SmartArt.
- **Animations calées sur une narration** : mets tous les effets « avec la précédente » et décale-les par un délai. Un effet « après la précédente » placé derrière la lecture audio attendrait la fin du son.
- **Contrôle visuel** : `Slide.Export($fichier, "PNG", $largeur, $hauteur)` rend une diapositive sans ouvrir de fenêtre.

## Captures d'écran

Ne capture jamais l'écran entier : tu photographierais les autres fenêtres. `capture_excel.ps1` utilise l'API `PrintWindow` sur la seule fenêtre d'Excel.

## Avant de partager un fichier Office

Un fichier Office en dit plus que son contenu visible. À contrôler à l'intérieur de l'archive (un `.docx`, `.xlsx` ou `.pptx` est un ZIP) :

- `docProps/core.xml` : auteur et dernier modificateur — y compris dans le classeur **incorporé à chaque graphique** d'une présentation ;
- `xl/workbook.xml` : chemin du fichier sur ton disque ;
- `xl/pivotCache/` : nom de l'utilisateur qui a actualisé le tableau croisé ;
- `docProps/thumbnail.jpeg`, s'il existe : miniature de la première page.

Le plus sûr est de chercher ton nom, ton numéro et ton nom d'utilisateur Windows dans tous les fichiers de l'archive, y compris les archives imbriquées.
