$ErrorActionPreference = "Stop"

$SRC = "C:\Users\utilisateur\Desktop\Cours\Environnement Numerique\seance-4\supports\Support_Exercices_Excel_Approfondissement_ML_Nanterre.xlsx"
$DST = "C:\Users\utilisateur\Desktop\Cours\Environnement Numerique\seance-4\NOM_Prenom_TD_Excel_Approfondissement.xlsx"
$MDP = "TD2026Excel"

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 800
Get-ChildItem (Split-Path $DST) -Force -Filter "~`$*" | Remove-Item -Force -ErrorAction SilentlyContinue
Copy-Item $SRC $DST -Force
Unblock-File $DST

$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
$x.ScreenUpdating = $false
$wb = $x.Workbooks.Open($DST)
"etape: classeur ouvert"

# ---------- constantes ----------
$xlAscending = 1; $xlDescending = 2; $xlYes = 1
$xlCellValue = 1; $xlExpression = 2; $xlGreater = 5
$xlValidateList = 3; $xlValidateWholeNumber = 1; $xlValidateDecimal = 2
$xlBetween = 1; $xlValidAlertStop = 1
$xlFilterCopy = 2
$xlDatabase = 1
$xlPie = 5; $xlColumnClustered = 51
$xlRowField = 1; $xlColumnField = 2; $xlPageField = 3; $xlDataField = 4
$xlSum = -4157; $xlCount = -4112
$xl3TrafficLights1 = 3
$xlSortOnValues = 0

# Deux pieges sur FormatConditions.Add en liaison tardive :
#  1. les trois arguments sont obligatoires, et l'operateur doit etre une
#     valeur LEGALE de XlFormatConditionOperator (1 a 8) meme s'il est ignore
#     pour une regle de type formule — 0, $null ou Missing echouent ;
#  2. la formule est interpretee en convention LOCALE (francaise), donc
#     separateur decimal virgule : "=$J5<0,15" et non "=$J5<0.15".
#     C'est l'inverse de Range.Formula, qui attend la convention anglaise.
function AjouterRegleFormule($plageCible, $formule, $couleurFond) {
  $r = $plageCible.FormatConditions.Add(2, 1, $formule)
  $r.Interior.Color = $couleurFond
  return $r
}

# =====================================================================
# EXERCICE 1 — tri, filtres, mise en forme conditionnelle
# =====================================================================
$ws = $wb.Worksheets("Exercice 1")
$ws.Activate()
$plage = $ws.Range("A4:J304")

# 1. tri sur deux niveaux : Region A-Z puis Montant decroissant
$M = [Type]::Missing
[void]$plage.Sort($ws.Range("D4"), $xlAscending, $ws.Range("I4"), $M, $xlDescending, $M, $xlAscending, $xlYes)
"  ex1: tri OK"

# 2. filtres automatiques + filtre Categorie = Informatique ET Montant > 500
if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false }
$plage.AutoFilter(6, "Informatique") | Out-Null
$plage.AutoFilter(9, ">500") | Out-Null
"  ex1: filtres OK"

# 3. mise en forme conditionnelle
$mont = $ws.Range("I5:I304")
$mont.FormatConditions.Delete()
$db = $mont.FormatConditions.AddDatabar()
try { $db.BarColor.Color = 15773696 } catch { "  (couleur barre par defaut)" }
"  ex1: barres de donnees OK"

$seuil = $mont.FormatConditions.Add($xlCellValue, $xlGreater, "2000")
$seuil.Interior.Color = 13551615
$seuil.Font.Color = -16383844
"  ex1: regle de seuil OK"

$marge = $ws.Range("J5:J304")
$marge.FormatConditions.Delete()
$icones = $marge.FormatConditions.AddIconSetCondition()
try { $icones.IconSet = $wb.IconSets.Item($xl3TrafficLights1) } catch { "  (jeu d'icones par defaut)" }
"  ex1: jeu d'icones OK"

# ligne entiere surlignee quand la marge est critique (< 15 %)
# pour une regle de type xlExpression l'operateur doit etre OMIS, pas 0
$tout = $ws.Range("A5:J304")
[void](AjouterRegleFormule $tout "=`$J5<0,15" 13561798)
"etape: exercice 1 OK"

# =====================================================================
# EXERCICE 2 — validation + graphiques
# =====================================================================
$ws = $wb.Worksheets("Exercice 2")
$ws.Activate()

# liste deroulante sur le Statut
$st = $ws.Range("I5:I204")
$st.Validation.Delete()
$st.Validation.Add($xlValidateList, $xlValidAlertStop, 1, "=`$L`$5:`$L`$8") | Out-Null
$st.Validation.IgnoreBlank = $true
$st.Validation.InCellDropdown = $true
"  ex2: liste deroulante OK"

# validation numerique sur la note de satisfaction
$nt = $ws.Range("J5:J204")
$nt.Validation.Delete()
$nt.Validation.Add($xlValidateWholeNumber, $xlValidAlertStop, $xlBetween, "1", "5") | Out-Null
"  ex2: validation numerique OK"
$nt.Validation.InputTitle = "Note satisfaction"
$nt.Validation.InputMessage = "Note entre 1 et 5"
$nt.Validation.ErrorTitle = "Valeur refusee"
$nt.Validation.ErrorMessage = "La note doit etre un entier compris entre 1 et 5."
$nt.Validation.ShowInput = $true
$nt.Validation.ShowError = $true

# graphique en secteurs : repartition du montant par categorie
$co = $ws.ChartObjects().Add(700, 220, 380, 260)
$ch = $co.Chart
$ch.SetSourceData($ws.Range("N5:O10"))
$ch.ChartType = $xlPie
$ch.HasTitle = $true
$ch.ChartTitle.Text = "Repartition du chiffre d'affaires par categorie"
$ch.ApplyDataLabels() | Out-Null
$ch.SeriesCollection(1).DataLabels().ShowPercentage = $true
$ch.SeriesCollection(1).DataLabels().ShowValue = $false
$ch.HasLegend = $true
"  ex2: graphique secteurs OK"

# histogramme : evolution mensuelle
# SetSourceData refuse une plage multi-zones : on prend Q5:S13 en entier
# puis on supprime la serie "Debut du mois", qui ne doit pas etre tracee.
$co2 = $ws.ChartObjects().Add(700, 500, 460, 260)
$ch2 = $co2.Chart
$ch2.SetSourceData($ws.Range("Q5:S13"))
$ch2.ChartType = $xlColumnClustered
for ($s = $ch2.SeriesCollection().Count; $s -ge 1; $s--) {
  if ($ch2.SeriesCollection($s).Name -like "*ebut*") { $ch2.SeriesCollection($s).Delete() }
}
$ch2.HasTitle = $true
$ch2.ChartTitle.Text = "Evolution du montant total des commandes par mois"
$ch2.HasLegend = $false
"etape: exercice 2 OK"

# =====================================================================
# EXERCICE 3 — tableaux croises dynamiques
# =====================================================================
$src3 = $wb.Worksheets("Exercice 3").Range("A4:I404")

$f1 = $wb.Worksheets.Add()
$f1.Name = "TCD Region x Categorie"
$pc1 = $wb.PivotCaches().Create($xlDatabase, $src3)
$pt1 = $pc1.CreatePivotTable($f1.Range("A3"), "TCD_Region")
$pt1.PivotFields("Region").Orientation = $xlRowField
$pt1.PivotFields("Categorie").Orientation = $xlColumnField
$pt1.PivotFields("Vendeur").Orientation = $xlPageField
$df1 = $pt1.AddDataField($pt1.PivotFields("Montant (EUR)"), "Somme du montant", $xlSum)
$df1.NumberFormat = "# ##0,00`" EUR`""
$pt1.RowGrand = $true
$pt1.ColumnGrand = $true
$f1.Range("A1").Value2 = "TCD 1 - Montant total par Region et par Categorie (filtre : Vendeur)"
$f1.Range("A1").Font.Bold = $true

$f2 = $wb.Worksheets.Add()
$f2.Name = "TCD Vendeur"
$pc2 = $wb.PivotCaches().Create($xlDatabase, $src3)
$pt2 = $pc2.CreatePivotTable($f2.Range("A3"), "TCD_Vendeur")
$pt2.PivotFields("Vendeur").Orientation = $xlRowField
[void]$pt2.AddDataField($pt2.PivotFields("ID Commande"), "Nombre de commandes", $xlCount)
$pt2.PivotFields("Vendeur").AutoSort($xlDescending, "Vendeur")
$f2.Range("A1").Value2 = "TCD 2 - Nombre de commandes par vendeur (tri alphabetique inverse)"
$f2.Range("A1").Font.Bold = $true
"etape: exercice 3 OK (TCD)"

# =====================================================================
# EXERCICE 4 — fonctions avancees, filtre avance, protection
# =====================================================================
$ws = $wb.Worksheets("Exercice 4")
$ws.Activate()
if ($ws.ProtectContents) { $ws.Unprotect($MDP) }

$ws.Range("G257").Formula = '=COUNTIFS($E$5:$E$254,"Cadre",$C$5:$C$254,"Commercial",$D$5:$D$254,"Paris")'
$ws.Range("G258").Formula = '=DSUM($A$4:$I$254,$G$4,$K$262:$L$263)'
$ws.Range("G259").Formula = '=DAVERAGE($A$4:$I$254,$G$4,$K$266:$K$267)'
$ws.Range("G258:G259").NumberFormat = "# ##0,00`" EUR`""

# affecter la formule a toute la plage : Excel ajuste les references
# relatives ligne par ligne, sans passer par le presse-papiers
$ws.Range("I5:I254").Formula = '=IFERROR(H5/(G5*12),"N/A")'
$ws.Range("I5:I254").NumberFormat = "0,0%"

# filtre avance : extraction vers A274
$ws.Range("A4:H254").AdvancedFilter($xlFilterCopy, $ws.Range("C270:G271"), $ws.Range("A274:H274"), $false) | Out-Null
"etape: exercice 4 formules + filtre avance OK"

# =====================================================================
# EXERCICE 5 — revision
# =====================================================================
$ws = $wb.Worksheets("Exercice 5")
$ws.Activate()
$pl5 = $ws.Range("A4:I354")

[void]$pl5.Sort($ws.Range("D4"), $xlAscending, $ws.Range("H4"), $M, $xlDescending, $M, $xlAscending, $xlYes)
"  ex5: tri OK"

$ws.Range("I5:I354").Formula = '=IFERROR(E5/F5,"Seuil non defini")'
$ws.Range("I5:I354").NumberFormat = "0,00"

if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false }
$pl5.AutoFilter(3, "Informatique") | Out-Null

$val = $ws.Range("H5:H354")
$val.FormatConditions.Delete()
$db5 = $val.FormatConditions.AddDatabar()
$db5.BarColor.Color = 10079487

$ratio = $ws.Range("I5:I354")
$ratio.FormatConditions.Delete()
$ic5 = $ratio.FormatConditions.AddIconSetCondition()
$ic5.IconSet = $wb.IconSets.Item($xl3TrafficLights1)

$tout5 = $ws.Range("A5:I354")
[void](AjouterRegleFormule $tout5 "=`$E5<`$F5" 13561798)

$ws.Range("G357").Formula = '=COUNTIFS($E$5:$E$354,0,$D$5:$D$354,"Paris")'

$f3 = $wb.Worksheets.Add()
$f3.Name = "TCD Stock"
$pc3 = $wb.PivotCaches().Create($xlDatabase, $ws.Range("A4:I354"))
$pt3 = $pc3.CreatePivotTable($f3.Range("A3"), "TCD_Stock")
$pt3.PivotFields("Entrepot").Orientation = $xlRowField
$pt3.PivotFields("Categorie").Orientation = $xlColumnField
$df3 = $pt3.AddDataField($pt3.PivotFields("Valeur stock (EUR)"), "Somme valeur stock", $xlSum)
$df3.NumberFormat = "# ##0,00`" EUR`""
$pt3.RowGrand = $true
$pt3.ColumnGrand = $true
$f3.Range("A1").Value2 = "TCD 3 - Valeur du stock par Entrepot et par Categorie"
$f3.Range("A1").Font.Bold = $true
"etape: exercice 5 OK"

# =====================================================================
# EXERCICE 6 — revision
# =====================================================================
$ws = $wb.Worksheets("Exercice 6")
$ws.Activate()
if ($ws.ProtectContents) { $ws.Unprotect($MDP) }

$sess = $ws.Range("H5:H304")
$sess.Validation.Delete()
$sess.Validation.Add($xlValidateList, $xlValidAlertStop, 1, "1,2") | Out-Null
$sess.Validation.InCellDropdown = $true

$notes = $ws.Range("E5:F304")
$notes.Validation.Delete()
$notes.Validation.Add($xlValidateDecimal, $xlValidAlertStop, $xlBetween, "0", "20") | Out-Null
$notes.Validation.InputTitle = "Note"
$notes.Validation.InputMessage = "Note comprise entre 0 et 20"
$notes.Validation.ErrorTitle = "Valeur refusee"
$notes.Validation.ErrorMessage = "La note doit etre comprise entre 0 et 20."
$notes.Validation.ShowInput = $true
$notes.Validation.ShowError = $true

$ws.Range("G5:G304").Formula = '=IFERROR(AVERAGE(E5:F5),"Absent")'
$ws.Range("G5:G304").NumberFormat = "0,00"

$ws.Range("G311").Formula = '=DAVERAGE($A$4:$H$304,$G$4,$A$307:$A$308)'
$ws.Range("G311").NumberFormat = "0,00"

$co3 = $ws.ChartObjects().Add(420, 20, 380, 250)
$ch3 = $co3.Chart
$ch3.SetSourceData($ws.Range("A314:B319"))
$ch3.ChartType = $xlPie
$ch3.HasTitle = $true
$ch3.ChartTitle.Text = "Repartition des effectifs par filiere"
$ch3.ApplyDataLabels() | Out-Null
$ch3.SeriesCollection(1).DataLabels().ShowPercentage = $true

$co4 = $ws.ChartObjects().Add(420, 290, 380, 250)
$ch4 = $co4.Chart
$ch4.SetSourceData($ws.Range("A314:C319"))
$ch4.ChartType = $xlColumnClustered
for ($s = $ch4.SeriesCollection().Count; $s -ge 1; $s--) {
  if ($ch4.SeriesCollection($s).Name -like "*ffectif*") { $ch4.SeriesCollection($s).Delete() }
}
$ch4.HasTitle = $true
$ch4.ChartTitle.Text = "Moyenne par filiere"
$ch4.HasLegend = $false
"etape: exercice 6 OK"

# =====================================================================
# recalcul, protection, enregistrement
# =====================================================================
$x.CalculateFullRebuild()
"--- verification (avant protection) ---"
$e4 = $wb.Worksheets("Exercice 4")
"G257 (NB.SI.ENS)  = " + $e4.Range("G257").Text
"G258 (BDSOMME)    = " + $e4.Range("G258").Text
"G259 (BDMOYENNE)  = " + $e4.Range("G259").Text
"I5   (SIERREUR)   = " + $e4.Range("I5").Text
"lignes extraites par le filtre avance : " + ($e4.Range("A275:A400").SpecialCells(2).Count)
$e5 = $wb.Worksheets("Exercice 5")
"Ex5 I5            = " + $e5.Range("I5").Text
"Ex5 G357          = " + $e5.Range("G357").Text
$e6 = $wb.Worksheets("Exercice 6")
"Ex6 G5            = " + $e6.Range("G5").Text
"Ex6 G311          = " + $e6.Range("G311").Text
"feuilles          : " + (($wb.Worksheets | ForEach-Object { $_.Name }) -join " | ")
$wb.Worksheets("Exercice 4").Protect($MDP)
$wb.Worksheets("Exercice 6").Protect($MDP)
$wb.Worksheets("Exercice 1").Activate()
$wb.Save()
"feuilles Exercice 4 et 6 protegees par mot de passe"
$wb.Close($true)
$x.Quit()
"TERMINE"












