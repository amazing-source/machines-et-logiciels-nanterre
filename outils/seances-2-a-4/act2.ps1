$ErrorActionPreference = "Stop"
function RGBv($r, $g, $b) { return $r + ($g * 256) + ($b * 65536) }

$M   = "C:\Users\utilisateur\Documents\MACHINES ET LOGICIELS"
$src = Join-Path $M "TEXTE LONG\TEXTE BRUT.docx"
$out = Join-Path $M "SEANCE2_ACTIVITE2_TTX.docx"

Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 700
Get-ChildItem $M -Force -Filter "~`$*" | Remove-Item -Force -ErrorAction SilentlyContinue
Copy-Item $src $out -Force
Unblock-File $out

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$w.UserName = "[NOM Prénom]"
$d = $w.Documents.Open($out, $false, $false)
$d.BuiltInDocumentProperties("Author") = "[NOM Prénom]"
$d.BuiltInDocumentProperties("Last author") = "[NOM Prénom]"

$ST_NORMAL = -1; $ST_H1 = -2; $ST_H2 = -3; $ST_H3 = -4; $ST_TITRE = -63

function ParaIdx($doc, $prefix) {
  for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
    $t = ($doc.Paragraphs($i).Range.Text -replace "`r", "").Trim()
    if ($t.StartsWith($prefix)) { return $i }
  }
  return -1
}
function ParaIdxFrom($doc, $prefix, $start) {
  for ($i = $start; $i -le $doc.Paragraphs.Count; $i++) {
    $t = ($doc.Paragraphs($i).Range.Text -replace "`r", "").Trim()
    if ($t.StartsWith($prefix)) { return $i }
  }
  return -1
}
function SetStyle($doc, $prefix, $styleId) {
  $i = ParaIdx $doc $prefix
  if ($i -gt 0) { $doc.Paragraphs($i).Range.Style = $doc.Styles($styleId) }
  else { "  ATTENTION introuvable: $prefix" }
}

# ---- EX3 : nettoyage des paragraphes vides (fait en premier) ----
for ($k = 0; $k -lt 5; $k++) {
  $f = $d.Content.Find
  $f.ClearFormatting(); $f.Replacement.ClearFormatting()
  [void]$f.Execute([ref]"^p^p", [ref]$false, [ref]$false, [ref]$false, [ref]$false,
                   [ref]$false, [ref]$true, [ref]1, [ref]$false, [ref]"^p", [ref]2)
}

# ---- EX1 Q1 : style par defaut du grand titre ----
"REPONSE EX1-Q1 : " + $d.Paragraphs(1).Style.NameLocal

# ---- EX1 Q2 : le titre recoit le style Titre, puis on modifie CE style ----
$d.Paragraphs(1).Range.Style = $d.Styles($ST_TITRE)
$stT = $d.Styles($ST_TITRE)
$stT.ParagraphFormat.Alignment = 1
$stT.Font.Bold = $true
$stT.Font.Size = 18
$stT.ParagraphFormat.SpaceBefore = 113.4

# ---- EX1 Q3 : style Normal ----
$n = $d.Styles($ST_NORMAL)
$n.Font.Name = "Arial"
$n.ParagraphFormat.Alignment = 3
$n.ParagraphFormat.LineSpacingRule = 5
$n.ParagraphFormat.LineSpacing = 13.8
$n.ParagraphFormat.SpaceAfter = 8

# ---- EX1 Q4/Q6/Q7 : application des styles de titres ----
SetStyle $d "Introduction" $ST_H1
SetStyle $d "COMPRENDRE LES M" $ST_H1
SetStyle $d "LES SECTEURS EN PREMI" $ST_H1
SetStyle $d "ACCOMPAGNER LE D" $ST_H1
SetStyle $d "Conclusion" $ST_H1
SetStyle $d ("D" + [char]0xC9 + "FINIR LES LIENS") $ST_H2
SetStyle $d ("La compl" + [char]0xE9 + "mentarit" + [char]0xE9) $ST_H3
SetStyle $d "...Et la confrontation" $ST_H3
SetStyle $d ("La sant" + [char]0xE9 + " et le diagnostic") $ST_H3
SetStyle $d ("L'" + [char]0xE9 + "ducation et la formation") $ST_H3
SetStyle $d ("Le d" + [char]0xE9 + "veloppement logiciel") $ST_H3

# ---- EX1 Q5 : saut de page avant chaque Titre 1 ----
$d.Styles($ST_H1).ParagraphFormat.PageBreakBefore = $true

# ---- EX1 Q8 : heritage, bleu fonce sur le style parent ----
$bleu = RGBv 31 78 121
$d.Styles($ST_H1).Font.Color = $bleu
"etape: styles de titres OK"

# ---- EX2 : style MiseEnValeur ----
try { $null = $d.Styles("MiseEnValeur") } catch { [void]$d.Styles.Add("MiseEnValeur", 1) }
$mv = $d.Styles("MiseEnValeur")
$mv.Font.Color = RGBv 89 89 89
$mv.Font.Size = 12
$mv.Font.Italic = $true
$mv.ParagraphFormat.LeftIndent = 18
$bmv = $mv.ParagraphFormat.Borders(-2)
$bmv.LineStyle = 1; $bmv.LineWidth = 12; $bmv.Color = RGBv 0 112 192

$iIntro = ParaIdx $d "Introduction"
$d.Paragraphs($iIntro + 1).Range.Style = $mv
$iConc = ParaIdx $d "Conclusion"
if ($iConc -gt 0 -and $iConc -lt $d.Paragraphs.Count) {
  $d.Paragraphs($iConc + 1).Range.Style = $mv
}
"etape: MiseEnValeur OK"

# ---- EX3 Q1 : y a-t-il du Titre 4 ? ----
$nbT4 = 0
foreach ($par in $d.Paragraphs) {
  if ($par.Style.NameLocal -like "Titre 4*" -or $par.Style.NameLocal -like "Heading 4*") { $nbT4++ }
}
"REPONSE EX3-Q1 : paragraphes Titre 4 = $nbT4"

# ---- EX4 : deplacer le bloc Education avant Sante ----
$iEdu   = ParaIdx $d ("L'" + [char]0xE9 + "ducation et la formation")
$iSante = ParaIdx $d ("La sant" + [char]0xE9 + " et le diagnostic")
if ($iEdu -gt 0 -and $iSante -gt 0 -and $iEdu -gt $iSante) {
  $fin = $iEdu
  for ($j = $iEdu + 1; $j -le $d.Paragraphs.Count; $j++) {
    $st = $d.Paragraphs($j).Style.NameLocal
    if ($st -like "Titre*" -or $st -like "Heading*") { break }
    $fin = $j
  }
  $nbBloc = $fin - $iEdu + 1
  $rBloc = $d.Range($d.Paragraphs($iEdu).Range.Start, $d.Paragraphs($fin).Range.End)
  # on COPIE d'abord (FormattedText est une plage vivante, invalide apres Delete)
  $cible = $d.Range($d.Paragraphs($iSante).Range.Start, $d.Paragraphs($iSante).Range.Start)
  $cible.FormattedText = $rBloc.FormattedText
  # puis on supprime l'original, qui est maintenant la 2e occurrence
  $iDup = ParaIdxFrom $d ("L'" + [char]0xE9 + "ducation et la formation") ($iSante + $nbBloc)
  if ($iDup -gt 0) {
    $rDup = $d.Range($d.Paragraphs($iDup).Range.Start, $d.Paragraphs($iDup + $nbBloc - 1).Range.End)
    [void]$rDup.Delete()
  }
  "etape: bloc Education deplace (" + $nbBloc + " paragraphes)"
}

# ---- EX4 : DEFINIR LES LIENS passe en Titre 1 ----
SetStyle $d ("D" + [char]0xC9 + "FINIR LES LIENS") $ST_H1

# ---- EX5 : numerotation multiniveau liee AUX STYLES ----
$lt = $d.ListTemplates.Add($true)
$lt.ListLevels(1).NumberFormat = "SECTION %1 :"
$lt.ListLevels(1).NumberStyle = 1
$lt.ListLevels(1).TrailingCharacter = 1
$lt.ListLevels(2).NumberFormat = "%1.%2"
$lt.ListLevels(2).NumberStyle = 0
$lt.ListLevels(3).NumberFormat = "%1.%3"
$lt.ListLevels(3).NumberStyle = 0
$lt.ListLevels(3).ResetOnHigher = 1
$d.Styles($ST_H1).LinkToListTemplate($lt, 1)
$d.Styles($ST_H2).LinkToListTemplate($lt, 2)
$d.Styles($ST_H3).LinkToListTemplate($lt, 3)
"etape: numerotation liee aux styles"

# ---- EX5 : page SOMMAIRE ----
# saut de page apres le sous-titre (paragraphe 2)
$rEnd = $d.Paragraphs(2).Range
$rEnd.InsertParagraphAfter()
$pNew = $d.Paragraphs(3)
$pNew.Range.Style = $d.Styles($ST_NORMAL)
$d.Range($pNew.Range.Start, $pNew.Range.Start).InsertBreak(7)

# apres le saut, le paragraphe vide est le 4e
$pSom = $d.Paragraphs(4)
$pSom.Range.Style = $d.Styles($ST_NORMAL)
$pSom.Range.InsertBefore("SOMMAIRE")
$pSom.Range.Font.Size = 20
$pSom.Range.Font.Bold = $true
$pSom.Range.Font.Color = $bleu
$pSom.Alignment = 1
$pSom.SpaceAfter = 18
$pSom.Range.InsertParagraphAfter()

$rToc = $d.Paragraphs(5).Range
$rToc.Style = $d.Styles($ST_NORMAL)
[void]$d.TablesOfContents.Add($rToc, $true, 1, 2, $false, "", $true, $true)
"etape: sommaire insere"

# ---- EX5 : renommer Conclusion en Conclusion generale ----
$iConc = ParaIdx $d "Conclusion"
if ($iConc -gt 0) {
  $rc = $d.Paragraphs($iConc).Range
  [void]$rc.MoveEnd(1, -1)
  $rc.Text = "Conclusion g" + [char]0xE9 + "n" + [char]0xE9 + "rale"
}

# ---- EX6 : pied de page (auteur + Page X / Y) ----
foreach ($s in $d.Sections) {
  $fr = $s.Footers(1)
  $fr.LinkToPrevious = $false
  $fr.Range.Text = ""
  [void]$d.Fields.Add($fr.Range, 17)              # AUTHOR
  $fr.Range.InsertAfter([char]9 + "Page ")
  [void]$d.Fields.Add($d.Range($fr.Range.End - 1, $fr.Range.End - 1), 33)   # PAGE
  $fr.Range.InsertAfter(" / ")
  [void]$d.Fields.Add($d.Range($fr.Range.End - 1, $fr.Range.End - 1), 26)   # NUMPAGES
  $fr.Range.Font.Size = 9
  $fr.Range.Font.Name = "Arial"
}
"etape: pied de page OK"

# ---- EX7 : index lexical ----
$termes = @("intelligence artificielle", "machine learning", "Big Data",
            "soft skills", "skills gap", ("cybers" + [char]0xE9 + "curit" + [char]0xE9))
# on ne marque qu'a partir du corps du texte, jamais dans le titre ni le sommaire
$iDebutCorps = ParaIdx $d "Introduction"
foreach ($terme in $termes) {
  $zone = $d.Range($d.Paragraphs($iDebutCorps).Range.Start, $d.Content.End - 1)
  $f = $zone.Find
  $f.ClearFormatting(); $f.Replacement.ClearFormatting()
  $f.Text = $terme; $f.Forward = $true; $f.Wrap = 0; $f.MatchCase = $false
  if ($f.Execute()) { [void]$d.Indexes.MarkEntry($f.Parent, $terme) }
}
$fin = $d.Range($d.Content.End - 1, $d.Content.End - 1)
$fin.InsertParagraphAfter()
$pIdx = $d.Paragraphs($d.Paragraphs.Count)
$pIdx.Range.Style = $d.Styles($ST_H1)
$pIdx.Range.InsertBefore("Index des notions")
$pIdx.Range.InsertParagraphAfter()
$rIdx = $d.Paragraphs($d.Paragraphs.Count).Range
$rIdx.Style = $d.Styles($ST_NORMAL)
[void]$d.Indexes.Add($rIdx)
"etape: index genere"

$d.Fields.Update()
$d.TablesOfContents(1).Update()
$d.Indexes(1).Update()
$d.Repaginate()
"Auteur : " + $d.BuiltInDocumentProperties("Author").Value
"Pages : " + $d.ComputeStatistics(2) + " | Mots : " + $d.ComputeStatistics(0) + " | Paragraphes : " + $d.Paragraphs.Count
$d.Save()
$d.Close($false)
$w.Quit()
"TERMINE"


