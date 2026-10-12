# Seance 5, projet : traitements Excel poses sur le classeur produit par
# projet5_classeur.py — tri, filtres, mises en forme conditionnelles,
# sous-totaux, filtre avance, tableaux croises dynamiques, graphique.
# A reencoder en UTF-8 avec BOM avant execution (bom.ps1).
# Usage : projet5_excel.ps1 -Src etape1.xlsx -Dst final.xlsx
param(
    [Parameter(Mandatory = $true)][string]$Src,
    [Parameter(Mandatory = $true)][string]$Dst
)
$ErrorActionPreference = "Stop"

$xlAscending = 1; $xlDescending = 2; $xlYes = 1
$xlCellValue = 1; $xlEqual = 3
$xlFilterCopy = 2; $xlDatabase = 1
$xlRowField = 1; $xlColumnField = 2; $xlPageField = 3
$xlSum = -4157; $xlCount = -4112
$xlBarClustered = 57
$M = [Type]::Missing

function RGB($hex) {
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return $r + 256 * $g + 65536 * $b
}
$NAVY = RGB "12284C"; $BLUE = RGB "2A78D6"; $ORANGE = RGB "EB6834"
$INK = RGB "0B0B0B"; $INK2 = RGB "52514E"; $PALE = RGB "EAF1FB"; $PALEO = RGB "FCEFE8"
$WHITE = RGB "FFFFFF"

function Entete($plage, $couleur) {
    $plage.Font.Bold = $true
    $plage.Font.Color = $WHITE
    $plage.Interior.Color = $couleur
    $plage.HorizontalAlignment = -4108
    $plage.VerticalAlignment = -4108
    $plage.WrapText = $true
    $plage.RowHeight = 32
}

# Fige les volets sous la ligne indiquee. Le reglage passe par la fenetre
# active, qui n'est pas toujours prete dans une instance Excel invisible :
# on reessaie, et en dernier recours on continue sans (confort d'affichage).
function Figer($feuille, [int]$ligne) {
    for ($essai = 1; $essai -le 4; $essai++) {
        try {
            # a partir du 3e essai, Excel est rendu visible le temps du reglage
            if ($essai -ge 3) { $x.ScreenUpdating = $true; $x.Visible = $true; Start-Sleep -Milliseconds 800 }
            $feuille.Activate()
            $w = $x.ActiveWindow
            $w.FreezePanes = $false
            $w.SplitColumn = 0
            $w.SplitRow = $ligne
            $w.FreezePanes = $true
            if ($essai -ge 3) { $x.Visible = $false; $x.ScreenUpdating = $false }
            return
        }
        catch { Start-Sleep -Milliseconds 500 }
    }
    $x.Visible = $false; $x.ScreenUpdating = $false
    "  (volets non figes sur « $($feuille.Name) » : sans consequence sur le contenu)"
}

if (Test-Path $Dst) { Remove-Item $Dst -Force }
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
# Excel reprend la taille de sa derniere fenetre. Apres un graphique
# PowerPoint (mini-fenetre de donnees), elle ne montre plus qu'une ligne et
# les volets ne peuvent plus etre figes : on lui redonne une taille normale.
try { $x.WindowState = -4143; $x.Width = 1000; $x.Height = 700 } catch { }
$x.ScreenUpdating = $false
try {
    $wb = $x.Workbooks.Open($Src)
    $x.CalculateFullRebuild()
    "etape: classeur ouvert et recalcule"

    # =================================================================
    # DONNEES — tri a deux niveaux, filtres, mise en forme conditionnelle
    # =================================================================
    $wd = $wb.Worksheets.Item("Données")
    $wd.Activate()
    $last = $wd.Cells($wd.Rows.Count, 1).End(-4162).Row
    $plage = $wd.Range("A1:M$last")
    "  donnees: derniere ligne = $last"

    [void]$plage.Sort($wd.Range("D1"), $xlAscending, $wd.Range("K1"), $M, $xlDescending, $M, $xlAscending, $xlYes)
    "  donnees: tri Filiere (A-Z) puis Candidatures par place (decroissant) OK"

    $kk = $wd.Range("K2:K$last")
    $kk.FormatConditions.Delete()
    $db = $kk.FormatConditions.AddDatabar()
    $db.BarColor.Color = RGB "9EC5F4"
    try { $db.BarFillType = 0 } catch { }
    "  donnees: barres de donnees OK"

    # niveaux de tension : rampe ordinale d'une seule teinte, du clair au fonce
    $mm = $wd.Range("M2:M$last")
    $mm.FormatConditions.Delete()
    $niveaux = @(
        @("Faible", "EAF1FB", "0B0B0B"),
        @("Modérée", "B7D3F6", "0B0B0B"),
        @("Forte", "6DA7EC", "0B0B0B"),
        @("Très forte", "2A78D6", "FFFFFF")
    )
    foreach ($n in $niveaux) {
        $fc = $mm.FormatConditions.Add($xlCellValue, $xlEqual, ('="' + $n[0] + '"'))
        $fc.Interior.Color = RGB $n[1]
        $fc.Font.Color = RGB $n[2]
    }
    "  donnees: regles de couleur sur le niveau de tension OK"

    # =================================================================
    # SOUS-TOTAUX — Donnees > Sous-total, a chaque changement de filiere
    # =================================================================
    $ws = $wb.Worksheets.Add($M, $wb.Worksheets.Item($wb.Worksheets.Count))
    $ws.Name = "Sous-totaux"
    # copie en valeurs des 7 colonnes utiles, dans l'ordre voulu : un filtre
    # avance sans critere, dont la plage de destination porte les en-tetes
    $titres = @("Filière", "Zone", "Établissement", "Formation", "Places", "Candidatures", "Admis")
    for ($j = 0; $j -lt 7; $j++) { $ws.Cells(1, $j + 1).Value2 = $titres[$j] }
    $ws.Activate()
    [void]$wd.Range("A1:M$last").AdvancedFilter($xlFilterCopy, $M, $ws.Range("A1:G1"), $false)
    $ws.Range("A2:G$last").Interior.ColorIndex = -4142
    $ws.Range("A2:G$last").Font.Color = $INK
    $ws.Range("A2:G$last").FormatConditions.Delete()
    [void]$ws.Range("A1:G$last").Subtotal(1, $xlSum, @(5, 6, 7), $true, $false, 1)
    $lastS = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row
    $ws.Cells(1, 8).Value2 = "Candidatures par place"
    $ws.Range("H2:H$lastS").Formula = '=IFERROR(F2/E2,"")'
    Entete $ws.Range("A1:G1") $NAVY
    Entete $ws.Range("H1:H1") $ORANGE
    $ws.Range("E2:G$lastS").NumberFormat = "# ##0"
    $ws.Range("H2:H$lastS").NumberFormat = "0,0"
    $larg = @(22, 16, 30, 44, 9, 13, 9, 14)
    for ($j = 0; $j -lt 8; $j++) { $ws.Columns($j + 1).ColumnWidth = $larg[$j] }
    # lignes de sous-total en gras sur toute la largeur
    for ($r = 2; $r -le $lastS; $r++) {
        if ($ws.Cells($r, 5).HasFormula) {
            $ws.Range("A${r}:H${r}").Font.Bold = $true
            $ws.Range("A${r}:H${r}").Interior.Color = $PALE
        }
    }
    [void]$ws.Outline.ShowLevels(2)
    $ws.Activate()
    Figer $ws 1
    "etape: sous-totaux OK ($($lastS - 1) lignes, vue au niveau 2)"

    # =================================================================
    # ZOOM PARIS NANTERRE — filtre avance avec extraction, puis tri
    # =================================================================
    $wz = $wb.Worksheets.Add($M, $wb.Worksheets.Item($wb.Worksheets.Count))
    $wz.Name = "Zoom Paris Nanterre"
    $wz.Range("A1").Value2 = "Zoom : les formations de l'Université Paris Nanterre, de la plus demandée à la moins demandée"
    $wz.Range("A1").Font.Bold = $true; $wz.Range("A1").Font.Size = 14; $wz.Range("A1").Font.Color = $NAVY
    $wz.Range("A2").Value2 = "Obtenu par filtre avancé (Données > Avancé > Copier vers un autre emplacement) puis tri décroissant sur les candidatures par place."
    $wz.Range("A2").Font.Italic = $true; $wz.Range("A2").Font.Color = $INK2
    $wz.Range("A4").Value2 = "Zone de critères"
    $wz.Range("A4").Font.Bold = $true
    $wz.Range("A5").Value2 = "Établissement"
    $wz.Range("A6").Value2 = "*Paris Nanterre*"
    Entete $wz.Range("A5:A5") $NAVY
    $wz.Range("A6").Interior.Color = $PALEO
    $wz.Range("B6").Value2 = "← l'astérisque remplace n'importe quel texte : « contient Paris Nanterre »"
    $wz.Range("B6").Font.Italic = $true; $wz.Range("B6").Font.Color = $INK2

    $ext = @("Établissement", "Filière", "Formation", "Places", "Candidatures", "Admis",
        "Candidatures par place", "Taux de remplissage", "Niveau de tension")
    for ($j = 0; $j -lt $ext.Count; $j++) { $wz.Cells(8, $j + 1).Value2 = $ext[$j] }
    $wz.Activate()
    [void]$wd.Range("A1:M$last").AdvancedFilter($xlFilterCopy, $wz.Range("A5:A6"), $wz.Range("A8:I8"), $false)
    $lastZ = $wz.Cells($wz.Rows.Count, 1).End(-4162).Row
    $nz = $lastZ - 8
    [void]$wz.Range("A8:I$lastZ").Sort($wz.Range("G8"), $xlDescending, $M, $M, $xlAscending, $M, $xlAscending, $xlYes)
    Entete $wz.Range("A8:F8") $NAVY
    Entete $wz.Range("G8:I8") $ORANGE
    $wz.Range("A9:I$lastZ").Interior.ColorIndex = -4142
    $wz.Range("A9:I$lastZ").Font.Color = $INK
    $wz.Range("D9:F$lastZ").NumberFormat = "# ##0"
    $wz.Range("G9:G$lastZ").NumberFormat = "0,0"
    $wz.Range("H9:H$lastZ").NumberFormat = "0%"
    $wz.Range("I9:I$lastZ").HorizontalAlignment = -4108
    $wz.Range("A9:I$lastZ").Borders.Color = RGB "C3C2B7"
    $largZ = @(27, 14, 64, 8, 13, 8, 14, 13, 13)
    for ($j = 0; $j -lt 9; $j++) { $wz.Columns($j + 1).ColumnWidth = $largZ[$j] }
    $form = 0
    for ($r = 9; $r -le $lastZ; $r++) { if ($wz.Cells($r, 7).HasFormula) { $form++ } }
    # ma propre formation, mise en evidence
    $miashs = "Licence - Mathématiques et informatique appliquées aux sciences humaines et sociales"
    $rm = 0
    for ($r = 9; $r -le $lastZ; $r++) {
        if ($wz.Cells($r, 3).Value2 -eq $miashs -and $wz.Cells($r, 1).Value2 -eq "Université Paris Nanterre") { $rm = $r }
    }
    if ($rm -gt 0) {
        $wz.Range("A${rm}:I${rm}").Interior.Color = $PALEO
        $wz.Range("A${rm}:I${rm}").Font.Bold = $true
        $wz.Cells($rm, 10).Value2 = "← ma formation (rang $($rm - 8) sur $nz)"
        $wz.Cells($rm, 10).Font.Bold = $true; $wz.Cells($rm, 10).Font.Color = RGB "C8501E"
    }
    $kz = $wz.Range("G9:G$lastZ")
    $dbz = $kz.FormatConditions.AddDatabar()
    $dbz.BarColor.Color = RGB "9EC5F4"
    try { $dbz.BarFillType = 0 } catch { }
    Figer $wz 8
    "etape: zoom Paris Nanterre OK ($nz lignes extraites, $form formules restantes, MIASHS ligne $rm)"

    # Filtres automatiques de la feuille Donnees : a poser APRES les filtres
    # avances, qui desactivent le filtre automatique de la plage source.
    $wd.Activate()
    if ($wd.AutoFilterMode) { $wd.AutoFilterMode = $false }
    [void]$wd.Range("A1:M$last").AutoFilter(1)
    "etape: filtres automatiques de la feuille Donnees OK (actif = $($wd.AutoFilterMode))"

    # =================================================================
    # TCD — deux tableaux croises dynamiques sur le meme cache
    # =================================================================
    $wt = $wb.Worksheets.Add($M, $wb.Worksheets.Item("Synthèse"))
    $wt.Name = "TCD"
    $pc = $wb.PivotCaches().Create($xlDatabase, $wd.Range("A1:M$last"))

    $pt = $pc.CreatePivotTable($wt.Range("A6"), "TCD_Filiere")
    [void]$pt.CalculatedFields().Add("Ratio", "=Candidatures/Places", $true)
    $pt.PivotFields("Filière").Orientation = $xlRowField
    $pt.PivotFields("Sélectivité").Orientation = $xlPageField
    $d1 = $pt.AddDataField($pt.PivotFields("Formation"), "Nb de formations", $xlCount)
    $d2 = $pt.AddDataField($pt.PivotFields("Places"), "Somme des places", $xlSum)
    $d3 = $pt.AddDataField($pt.PivotFields("Candidatures"), "Somme des candidatures", $xlSum)
    $d4 = $pt.AddDataField($pt.PivotFields("Ratio"), "Cand. par place", $xlSum)
    $d1.NumberFormat = "# ##0"; $d2.NumberFormat = "# ##0"; $d3.NumberFormat = "# ##0"; $d4.NumberFormat = "0,0"
    $pt.PivotFields("Filière").AutoSort($xlDescending, "Cand. par place")
    $pt.RowGrand = $true; $pt.ColumnGrand = $true
    $pt.TableStyle2 = "PivotStyleMedium2"
    $pt.CompactLayoutRowHeader = "Filière"
    "  tcd 1 (par filiere) OK"

    $pt2 = $pc.CreatePivotTable($wt.Range("H6"), "TCD_Filiere_Zone")
    $pt2.PivotFields("Filière").Orientation = $xlRowField
    $pt2.PivotFields("Zone").Orientation = $xlColumnField
    $e1 = $pt2.AddDataField($pt2.PivotFields("Ratio"), "Candidatures par place ", $xlSum)
    $e1.NumberFormat = "0,0"
    $pz = $pt2.PivotFields("Zone")
    $pz.PivotItems("Paris").Position = 1
    $pz.PivotItems("Petite couronne").Position = 2
    $pz.PivotItems("Grande couronne").Position = 3
    $pt2.PivotFields("Filière").AutoSort($xlDescending, "Candidatures par place ")
    $pt2.RowGrand = $true; $pt2.ColumnGrand = $true
    $pt2.TableStyle2 = "PivotStyleMedium2"
    $pt2.CompactLayoutRowHeader = "Filière"
    $pt2.CompactLayoutColumnHeader = "Zone"
    "  tcd 2 (filiere x zone) OK"

    # sans cela, chaque actualisation du TCD recalcule les largeurs de colonnes
    $pt.HasAutoFormat = $false
    $pt2.HasAutoFormat = $false
    $wt.Range("A1").Value2 = "Tableaux croisés dynamiques — source : feuille Données (" + ($last - 1).ToString("N0", [Globalization.CultureInfo]::GetCultureInfo("fr-FR")) + " formations)"
    $wt.Range("A1").Font.Bold = $true; $wt.Range("A1").Font.Size = 14; $wt.Range("A1").Font.Color = $NAVY
    $wt.Range("A2").Value2 = "Champ calculé « Cand. par place » = Candidatures / Places (Analyse du tableau croisé dynamique > Champs, éléments et jeux > Champ calculé)."
    $wt.Range("A2").Font.Italic = $true; $wt.Range("A2").Font.Color = $INK2
    $wt.Range("A3").Value2 = "TCD 1 — par filière, de la plus demandée à la moins demandée"
    $wt.Range("A3").Font.Bold = $true
    $wt.Range("H3").Value2 = "TCD 2 — filière × zone géographique"
    $wt.Range("H3").Font.Bold = $true
    $largT = @(22, 17, 17, 23, 16, 3, 3, 24, 12, 17, 17, 14)
    for ($j = 0; $j -lt $largT.Count; $j++) { $wt.Columns($j + 1).ColumnWidth = $largT[$j] }
    "etape: TCD OK"

    # =================================================================
    # SYNTHESE — graphique en barres
    # =================================================================
    $wy = $wb.Worksheets.Item("Synthèse")
    $wy.Activate()
    $r0 = 5; $rT = 16
    $anc = $wy.Range("J4")
    $co = $wy.ChartObjects().Add($anc.Left, $anc.Top, 560, 360)
    $co.Name = "Graphique_filieres"
    $ch = $co.Chart
    $ch.ChartType = $xlBarClustered
    while ($ch.SeriesCollection().Count -gt 0) { $ch.SeriesCollection(1).Delete() }
    $s = $ch.SeriesCollection().NewSeries()
    $s.Name = "Candidatures par place"
    $s.Values = $wy.Range("F${r0}:F$($rT - 1)")
    $s.XValues = $wy.Range("A${r0}:A$($rT - 1)")
    $ch.HasLegend = $false
    $ch.HasTitle = $true
    $ch.ChartTitle.Text = "Candidatures par place selon la filière — Île-de-France, Parcoursup 2025"
    $ch.ChartTitle.Font.Size = 12; $ch.ChartTitle.Font.Bold = $true; $ch.ChartTitle.Font.Color = $INK
    $ax = $ch.Axes(1)
    $ax.ReversePlotOrder = $true
    $ax.TickLabels.Font.Size = 10; $ax.TickLabels.Font.Color = $INK
    $ax.Format.Line.ForeColor.RGB = RGB "C3C2B7"
    $ax.MajorTickMark = -4142
    $av = $ch.Axes(2)
    $av.HasMajorGridlines = $false
    $av.Delete()
    $ch.ChartGroups(1).GapWidth = 80
    $s.Format.Fill.ForeColor.RGB = $BLUE
    $s.HasDataLabels = $true
    $dl = $s.DataLabels()
    # dans un graphique, NumberFormat suit la convention ANGLAISE (contrairement
    # aux cellules, cf. HANDOFF 5.10) : "0,0" y afficherait 81 au lieu de 81,1
    $dl.NumberFormatLinked = $false
    $dl.NumberFormat = "0.0"
    $dl.Position = 2
    $dl.Font.Size = 10; $dl.Font.Color = $INK
    # mise en evidence de la licence : la filiere qui offre le plus de places
    for ($k = $r0; $k -lt $rT; $k++) {
        if ($wy.Cells($k, 1).Value2 -eq "Licence") { $s.Points($k - $r0 + 1).Format.Fill.ForeColor.RGB = $ORANGE }
    }
    $ch.ChartArea.Format.Line.Visible = 0
    "  etiquettes : " + (($s.Points(1).DataLabel.Text, $s.Points(2).DataLabel.Text, $s.Points(10).DataLabel.Text) -join " / ")
    $rn = $co.BottomRightCell.Row + 1
    $wy.Range("J$rn").Value2 = "En orange : la licence, qui offre à elle seule 4 places sur 10."
    $wy.Range("J$rn").Font.Italic = $true; $wy.Range("J$rn").Font.Color = $INK2
    "etape: graphique OK"

    # =================================================================
    # recalcul, controle, enregistrement
    # =================================================================
    $x.CalculateFullRebuild()
    $wb.Worksheets.Item("Données").Activate()
    $wd.Range("A2").Select() | Out-Null
    $wb.Worksheets.Item("Lisez-moi").Activate()
    $wb.SaveAs($Dst, 51)
    "etape: enregistre -> $Dst"
    "feuilles : " + (($wb.Worksheets | ForEach-Object { $_.Name }) -join " | ")
}
finally {
    if ($wb) { $wb.Close($false) }
    $x.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
"TERMINE"
