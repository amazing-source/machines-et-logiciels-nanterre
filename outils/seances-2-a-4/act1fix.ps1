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

# prefixes ASCII uniquement (PowerShell 5.1 lit ce fichier en ANSI)
function ParaIdx($doc, $prefix) {
  for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
    $t = ($doc.Paragraphs($i).Range.Text -replace "`r", "").Trim()
    if ($t.StartsWith($prefix)) { return $i }
  }
  return -1
}

# 1. tout remettre sur une colonne
$d.Content.PageSetup.TextColumns.SetCount(1)

# 2. enlever l'ancien saut de colonne s'il existe (caractere 14)
$f = $d.Content.Find
$f.ClearFormatting(); $f.Replacement.ClearFormatting()
$f.Text = "^n"; $f.Replacement.Text = ""
[void]$f.Execute([ref]"^n", [ref]$false, [ref]$false, [ref]$false, [ref]$false,
                 [ref]$false, [ref]$true, [ref]1, [ref]$false, [ref]"", [ref]2)

# 3. saut de section continu AVANT "Ingredients"
$i = ParaIdx $d "Ingr"
$d.Range($d.Paragraphs($i).Range.Start, $d.Paragraphs($i).Range.Start).Select()
$w.Selection.InsertBreak(3)    # wdSectionBreakContinuous

# 4. saut de section continu AVANT "VARIANTE"
$i = ParaIdx $d "VARIANTE"
$d.Range($d.Paragraphs($i).Range.Start, $d.Paragraphs($i).Range.Start).Select()
$w.Selection.InsertBreak(3)

"Sections apres decoupe : " + $d.Sections.Count

# 5. section 2 (ingredients + preparation) sur deux colonnes
$s2 = $d.Sections(2)
$s2.PageSetup.TextColumns.SetCount(2)
$s2.PageSetup.TextColumns.Spacing = 30
$s2.PageSetup.TextColumns.LineBetween = $true

# 6. saut de colonne avant "Instructions" pour separer les deux colonnes
$i = ParaIdx $d "Instructions"
$d.Range($d.Paragraphs($i).Range.Start, $d.Paragraphs($i).Range.Start).Select()
$w.Selection.InsertBreak(8)    # wdColumnBreak

# 7. mise en page identique sur toutes les sections + pied de page
foreach ($s in $d.Sections) {
  $s.PageSetup.PaperSize    = 7
  $s.PageSetup.Orientation  = 1
  $s.PageSetup.TopMargin    = 42
  $s.PageSetup.BottomMargin = 36
  $s.PageSetup.LeftMargin   = 50
  $s.PageSetup.RightMargin  = 50
  foreach ($b in @(-1, -2, -3, -4)) {
    $bd = $s.Borders($b); $bd.ArtStyle = 0; $bd.LineStyle = 1
    $bd.LineWidth = 12; $bd.Color = 2832064
  }
  $fr = $s.Footers(1)
  $fr.LinkToPrevious = $false
  $fr.Range.Text = "[NOM Prénom] - [N° étudiant]"
  $fr.Range.ParagraphFormat.Alignment = 1
  $fr.Range.Font.Size = 9
  $fr.Range.Font.Color = 7895160
}

$d.Repaginate()
"Pages : " + $d.ComputeStatistics(2)
"Colonnes s1/s2/s3 : " + $d.Sections(1).PageSetup.TextColumns.Count + "/" + $d.Sections(2).PageSetup.TextColumns.Count + "/" + $d.Sections(3).PageSetup.TextColumns.Count
$d.Save()
$d.Close($false)
$w.Quit()
"docx enregistre"
