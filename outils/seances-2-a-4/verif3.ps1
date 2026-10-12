$ErrorActionPreference = "Stop"
$P = "C:\Users\utilisateur\Desktop\Cours\Environnement Numerique\seance-3\NOM_Prenom_TD_Excel.xlsx"

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 700

$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
$wb = $x.Workbooks.Open($P, 0, $false)
$x.CalculateFullRebuild()

function Cellules($feuille, $liste) {
  $ws = $wb.Worksheets($feuille)
  foreach ($c in $liste) {
    $v = $ws.Range($c).Text
    $f = $ws.Range($c).Formula
    "  {0,-5} = {1,-26} [{2}]" -f $c, $v, $f
  }
}

"=== Exercice 1 ==="
Cellules "Exercice 1" @("A5", "E5", "E14", "D15", "E15")
"=== Exercice 2 ==="
Cellules "Exercice 2" @("H6", "B9", "H9", "D15", "E15", "E17", "C21", "E23")
"=== Exercice 3 ==="
Cellules "Exercice 3" @("D5", "D6", "E5", "F5", "G6", "D14", "D15")
"=== Exercice 4 ==="
Cellules "Exercice 4" @("B15", "C15", "E15", "E20", "E21")
"=== Exercice 5 ==="
Cellules "Exercice 5" @("E15", "F15", "G15", "G16", "G17", "F22", "F23", "F24", "F25")
"=== Exercice 6 ==="
Cellules "Exercice 6" @("D5", "D9", "B10", "D10", "D11", "D12", "D13", "D14", "D15")

"=== erreurs eventuelles ==="
$nbErr = 0
foreach ($ws in $wb.Worksheets) {
  $used = $ws.UsedRange
  foreach ($cell in $used) {
    $v = $cell.Text
    if ($v -like "*#*" -and ($v -like "*#N/A*" -or $v -like "*#VALEUR*" -or $v -like "*#REF*" -or $v -like "*#NOM*" -or $v -like "*#DIV*" -or $v -like "*#NAME*" -or $v -like "*#VALUE*")) {
      "  ERREUR " + $ws.Name + "!" + $cell.Address(0,0) + " = " + $v
      $nbErr++
    }
  }
}
"  total erreurs : $nbErr"

$wb.Save()
$wb.Close($false)
$x.Quit()
"TERMINE"

