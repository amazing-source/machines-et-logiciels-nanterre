# Seance 5, projet : controle du classeur dans Excel (valeurs reelles via
# Value2, jamais via Text) et balayage des erreurs sur toutes les feuilles.
# Usage : projet5_verif.ps1 -Path classeur.xlsx
param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = "Stop"
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
    $wb = $x.Workbooks.Open($Path, 0, $true)
    $x.CalculateFullRebuild()

    "=== erreurs par feuille ==="
    foreach ($ws in $wb.Worksheets) {
        $n = 0
        foreach ($type in -4123, 2) {
            try { $n += $ws.UsedRange.SpecialCells($type, 16).Count } catch { }
        }
        "{0,-22} cellules utilisees={1,-7} erreurs={2}" -f $ws.Name, $ws.UsedRange.Count, $n
    }

    function Ligne($ws, $r, $c1, $c2) {
        $v = @()
        for ($c = $c1; $c -le $c2; $c++) {
            $z = $ws.Cells($r, $c).Value2
            if ($z -is [double]) { $v += [Math]::Round($z, 3) } else { $v += "$z" }
        }
        return ($v -join " | ")
    }

    $wy = $wb.Worksheets.Item("Synthèse")
    "=== Synthese ==="
    foreach ($r in 4..16) { Ligne $wy $r 1 8 }
    "--- zones ---"
    foreach ($r in 19..22) { Ligne $wy $r 1 6 }
    "--- Nanterre ---"
    foreach ($r in 26..28) { Ligne $wy $r 1 6 }

    $wt = $wb.Worksheets.Item("TCD")
    "=== TCD 1 ==="
    $t = $wb.Worksheets.Item("TCD").PivotTables("TCD_Filiere").TableRange2
    for ($r = $t.Row; $r -lt $t.Row + $t.Rows.Count; $r++) { Ligne $wt $r $t.Column ($t.Column + $t.Columns.Count - 1) }
    "=== TCD 2 ==="
    $t = $wb.Worksheets.Item("TCD").PivotTables("TCD_Filiere_Zone").TableRange2
    for ($r = $t.Row; $r -lt $t.Row + $t.Rows.Count; $r++) { Ligne $wt $r $t.Column ($t.Column + $t.Columns.Count - 1) }

    $wsst = $wb.Worksheets.Item("Sous-totaux")
    "=== Sous-totaux (lignes visibles) ==="
    $last = $wsst.Cells($wsst.Rows.Count, 1).End(-4162).Row
    for ($r = 1; $r -le $last; $r++) { if (-not $wsst.Rows($r).Hidden) { Ligne $wsst $r 1 8 } }

    $wz = $wb.Worksheets.Item("Zoom Paris Nanterre")
    "=== Zoom Paris Nanterre ==="
    foreach ($r in 5, 6, 8, 9, 10, 11, 12, 13, 49, 94, 95) { "[$r] " + (Ligne $wz $r 1 10) }

    $wd = $wb.Worksheets.Item("Données")
    "=== Donnees ==="
    "filtre automatique actif : " + $wd.AutoFilterMode + " | plage : " + $(if ($wd.AutoFilterMode) { $wd.AutoFilter.Range.Address() } else { "(aucune)" })
    foreach ($r in 1, 2, 3, 403, 404, 2692) { "[$r] " + (Ligne $wd $r 1 13) }
    "formule J2 : " + $wd.Range("J2").FormulaLocal
    "formule K2 : " + $wd.Range("K2").FormulaLocal
    "formule L2 : " + $wd.Range("L2").FormulaLocal
    "formule M2 : " + $wd.Range("M2").FormulaLocal
    "formule Synthese C5 : " + $wy.Range("C5").FormulaLocal
    "formule Synthese C28 : " + $wy.Range("C28").FormulaLocal
    "niveaux : " + ((("Faible", "Modérée", "Forte", "Très forte") | ForEach-Object { "$_=" + $x.WorksheetFunction.CountIf($wd.Range("M2:M2692"), $_) }) -join ", ")
    "zones   : " + ((("Paris", "Petite couronne", "Grande couronne") | ForEach-Object { "$_=" + $x.WorksheetFunction.CountIf($wd.Range("J2:J2692"), $_) }) -join ", ")
    "proprietes : auteur=" + $wb.Author + " | titre=" + $wb.Title
}
finally {
    if ($wb) { $wb.Close($false) }
    $x.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
