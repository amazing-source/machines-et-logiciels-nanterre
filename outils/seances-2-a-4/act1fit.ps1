$ErrorActionPreference = "Stop"
$M   = "C:\Users\utilisateur\Documents\MACHINES ET LOGICIELS"
$out = Join-Path $M "SEANCE2_ACTIVITE1_TTX.docx"
$pdf = Join-Path $M "SEANCE2_ACTIVITE1_TTX.pdf"

Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 600
Remove-Item (Join-Path $M "~`$*") -Force -ErrorAction SilentlyContinue

$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
$d = $w.Documents.Open($out, $false, $false)

function ParaIdx($doc, $prefix) {
  for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
    $t = ($doc.Paragraphs($i).Range.Text -replace "`r", "").Trim()
    if ($t.StartsWith($prefix)) { return $i }
  }
  return -1
}

# marges plus serrees
foreach ($s in $d.Sections) {
  $s.PageSetup.TopMargin = 34; $s.PageSetup.BottomMargin = 30
  $s.PageSetup.LeftMargin = 40; $s.PageSetup.RightMargin = 40
}

# titre un peu plus compact
$i = ParaIdx $d "Recette du cake"
$d.Paragraphs($i).Range.Font.Size = 20
$d.Paragraphs($i).SpaceAfter = 4
$d.Paragraphs($i).SpaceBefore = 0

$size = 10.5
$loops = 0
do {
  # corps de texte : tout sauf le titre
  foreach ($par in $d.Paragraphs) {
    $t = ($par.Range.Text -replace "`r", "").Trim()
    if ($t.StartsWith("Recette du cake")) { continue }
    if ($t.Length -eq 0) { continue }
    $par.Range.Font.Size = $size
    $par.LineSpacingRule = 0          # simple
    if ($par.SpaceAfter -gt 4) { $par.SpaceAfter = 4 }
    if ($par.SpaceBefore -gt 4) { $par.SpaceBefore = 4 }
  }
  # les 3 intertitres restent visibles
  foreach ($p in @("Ingr", "Instructions", "VARIANTE")) {
    $k = ParaIdx $d $p
    if ($k -gt 0) { $d.Paragraphs($k).Range.Font.Size = $size + 1.5 }
  }
  $d.Repaginate()
  $pages = $d.ComputeStatistics(2)
  "taille $size -> $pages page(s)"
  $size = $size - 0.5
  $loops++
} while ($pages -gt 1 -and $size -ge 8 -and $loops -lt 8)

"Pages finales : " + $d.ComputeStatistics(2)
$d.Save()
$d.ExportAsFixedFormat($pdf, 17)
"PDF exporte"
$d.Close($false)
$w.Quit()
