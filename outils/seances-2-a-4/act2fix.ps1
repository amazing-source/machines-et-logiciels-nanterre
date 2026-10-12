$ErrorActionPreference = "Stop"
$M = "C:\Users\utilisateur\Documents\MACHINES ET LOGICIELS"
$P = Join-Path $M "SEANCE2_ACTIVITE2_TTX.docx"

Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 600
Get-ChildItem $M -Force -Filter "~`$*" | Remove-Item -Force -ErrorAction SilentlyContinue

$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
$w.UserName = "[NOM Prénom]"
$d = $w.Documents.Open($P, $false, $false)

# 1. supprimer les champs parasites du corps (PAGE/NUMPAGES mal places)
# les champs sont indexes : on supprime un par un en rescannant a chaque fois
$nbSupprimes = 0
$encore = $true
while ($encore) {
  $encore = $false
  for ($i = 1; $i -le $d.Content.Fields.Count; $i++) {
    $t = $d.Content.Fields($i).Type
    if ($t -eq 33 -or $t -eq 26 -or $t -eq 3) {
      $d.Content.Fields($i).Delete()
      $nbSupprimes++
      $encore = $true
      break
    }
  }
}
"champs parasites supprimes : $nbSupprimes"

# 2. reconstruire le pied de page dans SA PROPRE histoire
foreach ($s in $d.Sections) {
  $fr = $s.Footers(1)
  $fr.LinkToPrevious = $false
  $fr.Range.Text = ""

  $largeur = $s.PageSetup.PageWidth - $s.PageSetup.LeftMargin - $s.PageSetup.RightMargin
  $fr.Range.ParagraphFormat.TabStops.ClearAll()
  [void]$fr.Range.ParagraphFormat.TabStops.Add($largeur, 2)   # taquet droit

  # champ AUTEUR au debut
  $r = $fr.Range
  $r.SetRange($fr.Range.Start, $fr.Range.Start)
  [void]$d.Fields.Add($r, 17)

  # texte intermediaire
  $fr.Range.InsertAfter([char]9 + "Page ")

  # champ PAGE : plage construite DANS l'histoire du pied de page
  $r = $fr.Range
  $r.SetRange($fr.Range.End - 1, $fr.Range.End - 1)
  [void]$d.Fields.Add($r, 33)

  $fr.Range.InsertAfter(" / ")

  $r = $fr.Range
  $r.SetRange($fr.Range.End - 1, $fr.Range.End - 1)
  [void]$d.Fields.Add($r, 26)

  $fr.Range.Font.Size = 9
  $fr.Range.Font.Name = "Arial"
}

# 3. numerotation de niveau 3 : repart a 1 dans chaque section
$lt = $d.Styles(-4).ListTemplate
if ($lt -ne $null) {
  $lt.ListLevels(3).NumberFormat = "%3."
  $lt.ListLevels(3).NumberStyle = 0
  $lt.ListLevels(3).ResetOnHigher = 1
}

[void]$d.Fields.Update()
$d.TablesOfContents(1).Update()
$d.Repaginate()

"--- verification ---"
"champs restants dans le corps : " + $d.Content.Fields.Count
"pied de page : [" + ($d.Sections(1).Footers(1).Range.Text -replace "`r", "") + "]"
"titre        : [" + ($d.Paragraphs(1).Range.Text -replace "`r", "") + "]"
"pages        : " + $d.ComputeStatistics(2)
$d.Save(); $d.Close($false); $w.Quit()
"TERMINE"


