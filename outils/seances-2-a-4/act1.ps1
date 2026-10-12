$ErrorActionPreference = "Stop"
function RGBv($r, $g, $b) { return $r + ($g * 256) + ($b * 65536) }

$M    = "C:\Users\utilisateur\Documents\MACHINES ET LOGICIELS"
$src  = Join-Path $M "recette cake citron pavot.docx"
$out  = Join-Path $M "SEANCE2_ACTIVITE1_TTX.docx"
$pdf  = Join-Path $M "SEANCE2_ACTIVITE1_TTX.pdf"
$img  = Join-Path $M "cake citron pavot.jpg"

Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 800
Remove-Item (Join-Path $M "~`$*") -Force -ErrorAction SilentlyContinue
Copy-Item $src $out -Force
Unblock-File $out

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try { $w.Options.PrintBackgrounds = $true } catch {}
"etape: Word demarre"

$d = $w.Documents.Open($out, $false, $false)
"etape: copie de travail ouverte"

# ---------- helper : index de paragraphe par debut de texte ----------
function ParaIdx($doc, $prefix) {
  for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
    $t = ($doc.Paragraphs($i).Range.Text -replace "`r", "").Trim()
    if ($t.StartsWith($prefix)) { return $i }
  }
  return -1
}
function ReplaceAll($doc, $from, $to) {
  $f = $doc.Content.Find
  $f.ClearFormatting(); $f.Replacement.ClearFormatting()
  $f.Text = $from; $f.Replacement.Text = $to
  $f.Forward = $true; $f.Wrap = 1; $f.MatchCase = $true
  $f.MatchWholeWord = $false; $f.MatchWildcards = $false
  [void]$f.Execute([ref]$from, [ref]$true, [ref]$false, [ref]$false, [ref]$false,
                   [ref]$false, [ref]$true, [ref]1, [ref]$false, [ref]$to, [ref]2)
}

# ================= 1. RECHERCHER / REMPLACER =================
ReplaceAll $d "pistache" "noisette"
ReplaceAll $d "Pistache" "Noisette"

# ================= 2. CORRECTION ORTHOGRAPHE / GRAMMAIRE =====
ReplaceAll $d "beure"        "beurre"
ReplaceAll $d "cuilères"     "cuillères"
ReplaceAll $d "Laisez"       "Laissez"
ReplaceAll $d "emprisoner"   "emprisonner"
ReplaceAll $d "acentuer"     "accentuer"
ReplaceAll $d "citron  pavot" "citron pavot"
ReplaceAll $d "on peut  rajouter" "on peut rajouter"
ReplaceAll $d ":5 min."      " : 5 min. "
ReplaceAll $d ":25-30 min."  " : 25-30 min. "
ReplaceAll $d "Préparation :15 mn" "Préparation : 15 mn"
ReplaceAll $d "  "           " "

# phrase dupliquee (fin de l'etape 4 repetee en paragraphe isole)
$iDup = ParaIdx $d "Surveillez bien : le dessus"
if ($iDup -gt 0) { $d.Paragraphs($iDup).Range.Delete() | Out-Null }

"etape: remplacements et corrections faits"

# ================= 3. MISE EN FORME PARAGRAPHES ==============
$iTitre = ParaIdx $d "Recette du cake"
$iPrep  = ParaIdx $d "Préparation"
$iCuis  = ParaIdx $d "Cuisson"
$iIngT  = ParaIdx $d "Ingrédients"
$iInsT  = ParaIdx $d "Instructions"
$iVar   = ParaIdx $d "VARIANTE"

# titre de la recette
$t = $d.Paragraphs($iTitre)
$t.Range.Font.Name = "Georgia"
$t.Range.Font.Size = 24
$t.Range.Font.Bold = $true
$t.Range.Font.Color = RGBv 192 57 43
$t.Range.Font.Shadow = $true
$t.Alignment = 1
$t.SpaceAfter = 10

# temps de preparation : couleur + cadrage a droite
$p = $d.Paragraphs($iPrep)
$p.Range.Font.Color = RGBv 197 90 17
$p.Range.Font.Bold = $true
$p.Range.Font.Italic = $true
$p.Alignment = 2
$p.SpaceAfter = 0
# reproduction de mise en forme (pinceau) sur le temps de cuisson
$c = $d.Paragraphs($iCuis)
$c.Range.Font.Color = RGBv 197 90 17
$c.Range.Font.Bold = $true
$c.Range.Font.Italic = $true
$c.Alignment = 2
$c.SpaceAfter = 10

# titres de sections : espace apres
foreach ($i in @($iIngT, $iInsT, $iVar)) {
  $q = $d.Paragraphs($i)
  $q.Range.Font.Bold = $true
  $q.Range.Font.Size = 13
  $q.Range.Font.Color = RGBv 192 57 43
  $q.SpaceAfter = 8
  $q.SpaceBefore = 6
}

# ================= 4. PUCES : INGREDIENTS ====================
$iIngT = ParaIdx $d "Ingrédients"
$iInsT = ParaIdx $d "Instructions"
$r = $d.Range($d.Paragraphs($iIngT + 1).Range.Start, $d.Paragraphs($iInsT - 1).Range.End)
$r.ListFormat.ApplyBulletDefault()
$r.ParagraphFormat.SpaceAfter = 0
$r.ParagraphFormat.LeftIndent = 28
$r.ParagraphFormat.RightIndent = 10

# bordure autour du bloc ingredients
foreach ($b in @(-1, -2, -3, -4)) {
  $bd = $r.Borders($b)
  $bd.LineStyle = 1
  $bd.LineWidth = 8
  $bd.Color = RGBv 197 90 17
}

# ================= 5. PUCES : ETAPES DE PREPARATION ==========
$iInsT = ParaIdx $d "Instructions"
$iVar  = ParaIdx $d "VARIANTE"
$r2 = $d.Range($d.Paragraphs($iInsT + 1).Range.Start, $d.Paragraphs($iVar - 1).Range.End)
$r2.ListFormat.ApplyBulletDefault()
$r2.ParagraphFormat.SpaceAfter = 4
$r2.ParagraphFormat.LeftIndent = 28
$r2.ParagraphFormat.Alignment = 3

# ================= 6. MISE EN PAGE : A4 PAYSAGE ==============
foreach ($s in $d.Sections) {
  $s.PageSetup.PaperSize    = 7    # A4
  $s.PageSetup.Orientation  = 1    # paysage
  $s.PageSetup.TopMargin    = 42
  $s.PageSetup.BottomMargin = 36
  $s.PageSetup.LeftMargin   = 50
  $s.PageSetup.RightMargin  = 50
}

# couleur d'arriere-plan de page
$d.Background.Fill.Visible = $true
$d.Background.Fill.Solid()
$d.Background.Fill.ForeColor.RGB = RGBv 253 246 227

# ================= 7. IMAGE APRES LE TITRE ===================
$iTitre = ParaIdx $d "Recette du cake"
$d.Paragraphs($iTitre).Range.InsertParagraphAfter()
$rImg = $d.Paragraphs($iTitre + 1).Range
$sh = $d.InlineShapes.AddPicture($img, $false, $true, $rImg)
$sh.LockAspectRatio = $true
$sh.Height = 142          # 5 cm = 141.7 pt
$d.Paragraphs($iTitre + 1).Alignment = 1
$d.Paragraphs($iTitre + 1).SpaceAfter = 6

"etape: image inseree"

# ================= 8. COLONNES ===============================
$iIngT = ParaIdx $d "Ingrédients"
$iVar  = ParaIdx $d "VARIANTE"
$rCol = $d.Range($d.Paragraphs($iIngT).Range.Start, $d.Paragraphs($iVar - 1).Range.End)
$rCol.Select()
$w.Selection.PageSetup.TextColumns.SetCount(2)
$w.Selection.PageSetup.TextColumns.Spacing = 30
$w.Selection.PageSetup.TextColumns.LineBetween = $true

# saut de colonne manuel avant "Instructions"
$iInsT = ParaIdx $d "Instructions"
$d.Range($d.Paragraphs($iInsT).Range.Start, $d.Paragraphs($iInsT).Range.Start).Select()
$w.Selection.InsertBreak(8)   # wdColumnBreak

# la mise en page doit rester paysage sur toutes les sections creees
foreach ($s in $d.Sections) {
  $s.PageSetup.PaperSize   = 7
  $s.PageSetup.Orientation = 1
}

# ================= 9. BORDURE DE PAGE ========================
foreach ($s in $d.Sections) {
  foreach ($b in @(-1, -2, -3, -4)) {
    $bd = $s.Borders($b)
    $bd.ArtStyle = 0
    $bd.LineStyle = 1
    $bd.LineWidth = 12
    $bd.Color = RGBv 192 57 43
  }
}

"etape: colonnes et bordure de page faites"

# ================= 10. ETOILES (notation) ====================
$ps = $d.Sections(1).PageSetup
$left0 = $ps.PageWidth - $ps.RightMargin - 5 * 22 - 4 * 4
for ($k = 0; $k -lt 5; $k++) {
  $st = $d.Shapes.AddShape(92, $left0 + $k * 26, $ps.TopMargin + 2, 22, 22)
  $st.Fill.ForeColor.RGB = RGBv 241 196 15
  $st.Line.ForeColor.RGB = RGBv 197 90 17
  $st.Line.Weight = 0.75
  $st.WrapFormat.Type = 3          # devant le texte
  $st.RelativeHorizontalPosition = 1
  $st.RelativeVerticalPosition = 1
}

# ================= 11. PIED DE PAGE ==========================
foreach ($s in $d.Sections) {
  $fr = $s.Footers(1)
  $fr.LinkToPrevious = $false
  $fr.Range.Text = "[NOM Prénom] - [N° étudiant]"
  $fr.Range.ParagraphFormat.Alignment = 1
  $fr.Range.Font.Size = 9
  $fr.Range.Font.Name = "Calibri"
  $fr.Range.Font.Color = RGBv 120 120 120
}

$d.Repaginate()
"Pages avant ajustement : " + $d.ComputeStatistics(2)

# ================= 12. CONTRAINTE UNE SEULE PAGE =============
$guard = 0
while ($d.ComputeStatistics(2) -gt 1 -and $guard -lt 6) {
  foreach ($par in $d.Paragraphs) {
    $f = $par.Range.Font
    if ($f.Size -gt 8) { $f.Size = $f.Size - 0.5 }
  }
  $d.Repaginate()
  $guard++
}
"Pages apres ajustement : " + $d.ComputeStatistics(2) + " (iterations: $guard)"

$d.Save()
$d.ExportAsFixedFormat($pdf, 17)
"Enregistre : $out"
"PDF        : $pdf"
$d.Close($false)
$w.Quit()

