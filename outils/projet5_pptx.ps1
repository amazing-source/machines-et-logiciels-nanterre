# Seance 5, projet : construction de la presentation dans PowerPoint.
# Masque personnalise (couleurs, polices, pied de page, dispositions), six
# diapositives, graphique natif alimente par le classeur, animations,
# transitions, notes et narration audio.
# A reencoder en UTF-8 avec BOM avant execution (bom.ps1).
param(
    [Parameter(Mandatory = $true)][string]$Xlsx,
    [Parameter(Mandatory = $true)][string]$Captures,
    [Parameter(Mandatory = $true)][string]$Narration,
    [Parameter(Mandatory = $true)][string]$Textes,
    [Parameter(Mandatory = $true)][string]$Identite,
    [Parameter(Mandatory = $true)][string]$Out,
    # variante sans audio ni minutage, pour enregistrer sa propre voix
    [switch]$SansVoix
)
$ErrorActionPreference = "Stop"

function RGB($hex) {
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return $r + 256 * $g + 65536 * $b
}
$NAVY = "12284C"; $NAVY2 = "1D3B6B"; $BLUE = "2A78D6"; $ORANGE = "EB6834"; $ORANGED = "B8461A"
$INK = "1A1A19"; $INK2 = "52514E"; $PALE = "F1F5FB"; $PALEO = "FCEFE8"; $LINE = "D9DEE7"
$WHITE = "FFFFFF"; $ICE = "CDE2FB"; $ICE2 = "9EC5F4"

# ---------------------------------------------------------------------------
# 1. Chiffres : lus dans le classeur, puis compares a ceux cites dans le texte
# ---------------------------------------------------------------------------
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
    $wb = $x.Workbooks.Open($Xlsx, 0, $true)
    $wy = $wb.Worksheets.Item("Synthèse")
    $noms = @(); $vals = @()
    for ($r = 5; $r -le 15; $r++) { $noms += [string]$wy.Cells($r, 1).Value2; $vals += [double]$wy.Cells($r, 6).Value2 }
    $chk = [ordered]@{
        "formations IdF"       = @([double]$wy.Range("B16").Value2, 2691)
        "places IdF"           = @([double]$wy.Range("C16").Value2, 153229)
        "candidatures IdF"     = @([double]$wy.Range("D16").Value2, 4197242)
        "cand/place IdF"       = @([Math]::Round([double]$wy.Range("F16").Value2, 1), 27.4)
        "PASS"                 = @([Math]::Round($vals[0], 1), 81.1)
        "IFSI"                 = @([Math]::Round($vals[1], 1), 58.4)
        "Licence"              = @([Math]::Round($vals[9], 1), 17.5)
        "EFTS"                 = @([Math]::Round($vals[10], 1), 10.4)
        "part places licence"  = @([Math]::Round([double]$wy.Range("H14").Value2, 2), 0.41)
        "Paris"                = @([Math]::Round([double]$wy.Range("F20").Value2, 1), 29.3)
        "Petite couronne"      = @([Math]::Round([double]$wy.Range("F21").Value2, 1), 27.9)
        "Grande couronne"      = @([Math]::Round([double]$wy.Range("F22").Value2, 1), 23.5)
        "Nanterre formations"  = @([double]$wy.Range("B27").Value2, 87)
        "Nanterre cand/place"  = @([Math]::Round([double]$wy.Range("F27").Value2, 1), 19.8)
        "MIASHS cand/place"    = @([Math]::Round([double]$wy.Range("F28").Value2, 1), 19.2)
    }
    $wz = $wb.Worksheets.Item("Zoom Paris Nanterre")
    $chk["Psychologie L.AS"] = @([Math]::Round([double]$wz.Range("G9").Value2, 0), 171)
    $rangM = 0
    for ($r = 9; $r -le 95; $r++) { if ([string]$wz.Cells($r, 10).Value2 -like "*ma formation*") { $rangM = $r - 8 } }
    $chk["rang MIASHS"] = @($rangM, 41)
    $wb.Close($false)
}
finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
$ko = 0
foreach ($k in $chk.Keys) { if ($chk[$k][0] -ne $chk[$k][1]) { $ko++; "ECART sur $k : classeur=$($chk[$k][0]) texte=$($chk[$k][1])" } }
if ($ko -gt 0) { throw "Les chiffres du texte ne correspondent plus au classeur ($ko ecart(s))." }
"etape: $($chk.Count) chiffres du texte confrontes au classeur, aucun ecart"

$J = Get-Content -Raw -Encoding UTF8 $Textes | ConvertFrom-Json
$ID = Get-Content -Raw -Encoding UTF8 $Identite | ConvertFrom-Json
function Remplir([string]$t) {
    foreach ($p in $ID.PSObject.Properties) { $t = $t.Replace('{' + $p.Name + '}', [string]$p.Value) }
    return $t
}
# duree de chaque narration (inutile, et fichiers peut-etre absents, sans voix)
$durees = @()
for ($i = 1; $i -le 6; $i++) {
    if ($SansVoix) { $durees += 0 }
    else { $durees += ((Get-Item (Join-Path $Narration "narration_diapo$i.wav")).Length - 44) / 44100.0 }
}
# dans le modele a trous, le numero est deja la mention « [N° étudiant] »
$ligneNumero = $(if ("$($ID.numero)".StartsWith("[")) { "$($ID.numero)" } else { "N° étudiant $($ID.numero)" })

# ---------------------------------------------------------------------------
# 2. PowerPoint
# ---------------------------------------------------------------------------
if (Test-Path $Out) { Remove-Item $Out -Force }
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add(-1)
try { $ppt.WindowState = 2 } catch { }

# Typographie francaise : espace insecable devant ? : ; ! et dans les
# guillemets, et entre les tranches de trois chiffres (2 691 ne se coupe pas).
function Typo([string]$t) {
    $n = [string][char]0x00A0
    $t = $t -replace ' ([?:;!»])', ($n + '$1')
    $t = $t -replace '« ', ('«' + $n)
    $t = $t -replace '(\d) (\d{3})', ('$1' + $n + '$2')
    return $t
}
function Txt($sl, [string]$text, [single]$l, [single]$t, [single]$w, [single]$h, [single]$size, [string]$color, [bool]$bold = $false, [int]$align = 1) {
    $s = $sl.Shapes.AddTextbox(1, $l, $t, $w, $h)
    $tf = $s.TextFrame
    $tf.WordWrap = -1; $tf.AutoSize = 0
    $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $tf.VerticalAnchor = 1
    $s.Height = $h
    $r = $tf.TextRange
    $r.Text = (Typo $text)
    $r.Font.Size = $size
    $r.Font.Color.RGB = RGB $color
    $r.Font.Bold = $(if ($bold) { -1 } else { 0 })
    $r.ParagraphFormat.Alignment = $align
    $r.ParagraphFormat.SpaceWithin = 1.0
    return $s
}
# texte a mises en forme multiples : tableau de (texte, taille, couleur, gras)
function Runs($sl, $runs, [single]$l, [single]$t, [single]$w, [single]$h, [single]$spaceAfter = 0) {
    $s = Txt $sl "" $l $t $w $h 12 $INK
    $r = $s.TextFrame.TextRange
    foreach ($u in $runs) {
        $a = $r.InsertAfter((Typo $u[0]))
        $a.Font.Size = [single]$u[1]
        $a.Font.Color.RGB = RGB $u[2]
        $a.Font.Bold = $(if ($u[3]) { -1 } else { 0 })
    }
    if ($spaceAfter -gt 0) { $r.ParagraphFormat.SpaceAfter = $spaceAfter }
    return $s
}
function Card($sl, [single]$l, [single]$t, [single]$w, [single]$h, [string]$fill, $line = $null, [single]$radius = 0.06) {
    $s = $sl.Shapes.AddShape(5, $l, $t, $w, $h)
    $s.Adjustments.Item(1) = $radius
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB = RGB $fill
    if ($line) { $s.Line.Visible = -1; $s.Line.ForeColor.RGB = RGB $line; $s.Line.Weight = 0.75 } else { $s.Line.Visible = 0 }
    $s.Shadow.Visible = 0
    $s.TextFrame.TextRange.Text = ""
    return $s
}
function Label($sl, [string]$text, [single]$l, [single]$t, [single]$w, [string]$color) {
    $s = Txt $sl $text $l $t $w 16 11 $color $true
    $s.TextFrame2.TextRange.Font.Spacing = [single]0.8
    return $s
}
function Pic($sl, [string]$file, [single]$l, [single]$t, [single]$w, [single]$h, [string]$alt) {
    $p = $sl.Shapes.AddPicture($file, 0, -1, $l, $t, $w, $h)
    $p.LockAspectRatio = -1
    $p.AlternativeText = $alt
    $p.Line.Visible = -1; $p.Line.ForeColor.RGB = RGB $LINE; $p.Line.Weight = 0.75
    $p.Shadow.Visible = -1; $p.Shadow.Blur = 14; $p.Shadow.OffsetX = 0; $p.Shadow.OffsetY = 4
    $p.Shadow.ForeColor.RGB = 0; $p.Shadow.Transparency = 0.78
    return $p
}
function Bulle($sl, [int]$n, [single]$l, [single]$t, [single]$d, [string]$fill) {
    $s = $sl.Shapes.AddShape(9, $l, $t, $d, $d)
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB = RGB $fill; $s.Line.Visible = 0; $s.Shadow.Visible = 0
    $tf = $s.TextFrame
    $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $tf.VerticalAnchor = 3; $tf.WordWrap = 0
    $tf.TextRange.Text = "$n"
    $tf.TextRange.Font.Size = [single][Math]::Round($d * 0.45)
    $tf.TextRange.Font.Bold = -1
    $tf.TextRange.Font.Color.RGB = RGB $WHITE
    $tf.TextRange.ParagraphFormat.Alignment = 2
    return $s
}
function Grouper($sl, $formes, $nom) {
    $noms = @(); foreach ($f in $formes) { $noms += $f.Name }
    $g = $sl.Shapes.Range($noms).Group()
    $g.Name = $nom
    return $g
}
# toutes les animations partent « avec la precedente », decalees par un delai
# compte depuis l'entree de la diapositive : elles suivent ainsi la narration
function Anim($sl, $forme, [single]$delai, [int]$effet = 10, [single]$duree = 0.6) {
    $e = $sl.TimeLine.MainSequence.AddEffect($forme, $effet, 0, 2)
    $e.Timing.TriggerDelayTime = $delai
    $e.Timing.Duration = $duree
    return $e
}

try {
    "etape: PowerPoint ouvert, diapositive {0} x {1} pt" -f $pres.PageSetup.SlideWidth, $pres.PageSetup.SlideHeight

    # ======================================================================
    # MASQUE DES DIAPOSITIVES
    # ======================================================================
    $m = $pres.SlideMaster
    $m.Name = "Masque Parcoursup"
    $cs = $m.Theme.ThemeColorScheme
    $couleurs = @($INK, $WHITE, $NAVY, $PALE, $BLUE, $ORANGE, "1BAF7A", "EDA100", "6250D6", "898781", $BLUE, "6250D6")
    for ($i = 1; $i -le 12; $i++) { $cs.Colors($i).RGB = RGB $couleurs[$i - 1] }
    $m.Theme.ThemeFontScheme.MajorFont.Item(1).Name = "Calibri"
    $m.Theme.ThemeFontScheme.MinorFont.Item(1).Name = "Calibri"

    foreach ($sh in $m.Shapes) {
        if ($sh.Type -ne 14) { continue }
        switch ($sh.PlaceholderFormat.Type) {
            1 {
                $sh.Left = 44; $sh.Top = 26; $sh.Width = 820; $sh.Height = 62
                $sh.TextFrame.VerticalAnchor = 3
                $sh.TextFrame.MarginLeft = 0; $sh.TextFrame.MarginRight = 0
                $f = $sh.TextFrame.TextRange.Font
                $f.Name = "Calibri"; $f.Size = 28; $f.Bold = -1; $f.Color.RGB = RGB $NAVY
                $sh.TextFrame.TextRange.ParagraphFormat.Alignment = 1
            }
            2 {
                $sh.Left = 44; $sh.Top = 104; $sh.Width = 872; $sh.Height = 382
                $sh.TextFrame.TextRange.Font.Color.RGB = RGB $INK
            }
            15 {
                $sh.Left = 60; $sh.Top = 506; $sh.Width = 640; $sh.Height = 20
                $sh.TextFrame.MarginLeft = 0
                $sh.TextFrame.TextRange.Font.Size = 10.5; $sh.TextFrame.TextRange.Font.Color.RGB = RGB $INK2
                $sh.TextFrame.TextRange.ParagraphFormat.Alignment = 1
            }
            13 {
                $sh.Left = 846; $sh.Top = 506; $sh.Width = 70; $sh.Height = 20
                $sh.TextFrame.MarginRight = 0
                $sh.TextFrame.TextRange.Font.Size = 11; $sh.TextFrame.TextRange.Font.Bold = -1
                $sh.TextFrame.TextRange.Font.Color.RGB = RGB $NAVY
                $sh.TextFrame.TextRange.ParagraphFormat.Alignment = 3
            }
            16 {
                $sh.Left = 706; $sh.Top = 506; $sh.Width = 130; $sh.Height = 20
                $sh.TextFrame.TextRange.Font.Size = 10.5; $sh.TextFrame.TextRange.Font.Color.RGB = RGB $INK2
                $sh.TextFrame.TextRange.ParagraphFormat.Alignment = 3
            }
        }
    }
    # repere graphique du masque : une pastille devant le pied de page
    $pastille = $m.Shapes.AddShape(5, 44, 511, 9, 9)
    $pastille.Adjustments.Item(1) = 0.3
    $pastille.Fill.Solid(); $pastille.Fill.ForeColor.RGB = RGB $ORANGE; $pastille.Line.Visible = 0
    $pastille.Name = "Pastille du masque"

    $layCouv = $m.CustomLayouts.Item(1)
    $layCont = $m.CustomLayouts.Item(2)
    $laySeul = $m.CustomLayouts.Item(6)
    "  dispositions d'origine : " + ((1..$m.CustomLayouts.Count | ForEach-Object { $m.CustomLayouts.Item($_).Name }) -join " | ")
    for ($i = $m.CustomLayouts.Count; $i -ge 1; $i--) { if ($i -notin 1, 2, 6) { $m.CustomLayouts.Item($i).Delete() } }
    $layCouv.Name = "Couverture (fond bleu nuit)"
    $layCont.Name = "Titre et contenu"
    $laySeul.Name = "Titre seul (contenu libre)"

    # disposition de couverture : fond bleu nuit, textes blancs alignes a gauche
    $layCouv.FollowMasterBackground = 0
    $layCouv.Background.Fill.Solid()
    $layCouv.Background.Fill.ForeColor.RGB = RGB $NAVY
    $layCouv.DisplayMasterShapes = 0
    foreach ($sh in $layCouv.Shapes) {
        if ($sh.Type -ne 14) { continue }
        switch ($sh.PlaceholderFormat.Type) {
            3 {
                $sh.Left = 60; $sh.Top = 100; $sh.Width = 800; $sh.Height = 140
                $sh.TextFrame.VerticalAnchor = 1; $sh.TextFrame.MarginLeft = 0
                $f = $sh.TextFrame.TextRange.Font
                $f.Size = 42; $f.Bold = -1; $f.Color.RGB = RGB $WHITE
                $sh.TextFrame.TextRange.ParagraphFormat.Alignment = 1
            }
            4 {
                $sh.Left = 60; $sh.Top = 236; $sh.Width = 800; $sh.Height = 96
                $sh.TextFrame.VerticalAnchor = 1; $sh.TextFrame.MarginLeft = 0
                $f = $sh.TextFrame.TextRange.Font
                $f.Size = 24; $f.Color.RGB = RGB $WHITE
                $sh.TextFrame.TextRange.ParagraphFormat.Alignment = 1
            }
        }
    }
    "etape: masque OK (" + $m.CustomLayouts.Count + " dispositions : " + ((1..$m.CustomLayouts.Count | ForEach-Object { $m.CustomLayouts.Item($_).Name }) -join " | ") + ")"

    $pied = "$($ID.nom_prenom) · $($ID.numero)  —  Parcoursup 2025 en Île-de-France  ·  Année universitaire $($ID.annee)"

    # ======================================================================
    # DIAPOSITIVE 1 — couverture
    # ======================================================================
    $s1 = $pres.Slides.AddSlide(1, $layCouv)
    [void](Label $s1 "PROJET PERSONNEL  ·  SÉANCE 5 — DOCUMENTS MULTIMÉDIA" 60 66 800 $ICE2)
    $t = $s1.Shapes.Title.TextFrame.TextRange
    $l1 = "Parcoursup 2025 en Île-de-France"
    $l2 = Typo "Où la concurrence est-elle la plus forte ?"
    $t.Text = $l1 + "`r" + $l2
    # dans un titre, le retour chariot devient un saut de ligne a l'interieur du
    # meme paragraphe : on vise donc la seconde ligne par ses caracteres
    $c2 = $t.Characters($l1.Length + 2, $l2.Length)
    $c2.Font.Size = 26; $c2.Font.Bold = 0; $c2.Font.Color.RGB = RGB "F6A27C"
    $st = $s1.Shapes.Placeholders.Item(2).TextFrame.TextRange
    $st.Text = "$($ID.nom_prenom)`rAnnée universitaire $($ID.annee)  ·  $($ID.filiere), $($ID.universite)`r$ligneNumero"
    $st.ParagraphFormat.Bullet.Visible = 0
    $st.Paragraphs(1).Font.Size = 24; $st.Paragraphs(1).Font.Bold = -1
    $st.Paragraphs(2).Font.Size = 15; $st.Paragraphs(2).Font.Color.RGB = RGB $ICE
    $st.Paragraphs(3).Font.Size = 13; $st.Paragraphs(3).Font.Color.RGB = RGB $ICE2
    $st.Paragraphs(1).ParagraphFormat.SpaceAfter = 4

    $tuiles = @(@("2 691", "formations en Île-de-France"), @("153 229", "places offertes"), @("4,2 millions", "de candidatures"))
    $g1 = @()
    for ($k = 0; $k -lt 3; $k++) {
        $xx = 60 + $k * 270
        $c = Card $s1 $xx 376 250 90 $NAVY2 $null 0.10
        $a = Txt $s1 $tuiles[$k][0] ($xx + 20) 386 210 42 32 $WHITE $true
        $b = Txt $s1 $tuiles[$k][1] ($xx + 20) 430 210 20 12.5 $ICE
        $g1 += (Grouper $s1 @($c, $a, $b) "Tuile $($k + 1)")
    }
    [void](Txt $s1 "Données ouvertes : Parcoursup 2025, ministère chargé de l'Enseignement supérieur (SIES) — Licence Ouverte" 60 506 800 18 10.5 $ICE2)
    "etape: diapositive 1 OK"

    # ======================================================================
    # DIAPOSITIVE 2 — problematique et demarche
    # ======================================================================
    $s2 = $pres.Slides.AddSlide(2, $laySeul)
    $s2.Shapes.Title.TextFrame.TextRange.Text = Typo "Problématique et démarche"
    [void](Card $s2 44 104 416 382 $PALE)
    [void](Label $s2 "LA QUESTION" 68 126 368 $BLUE)
    [void](Txt $s2 ("En Île-de-France, dans quelles" + [char]11 + "filières est-il le plus difficile" + [char]11 + "d'obtenir une place ?") 68 150 368 110 22 $NAVY $true)
    $q2 = Txt $s2 "Et la situation est-elle la même à Paris, en petite couronne et en grande couronne ?" 68 252 368 70 16 $INK2
    $ci = Card $s2 68 372 368 90 $WHITE $LINE 0.10
    $li = Label $s2 "L'INDICATEUR" 86 386 332 $ORANGED
    $ti = Runs $s2 @(@("Candidatures par place`r", 18, $NAVY, $true), @("= candidatures reçues ÷ places offertes", 13.5, $INK2, $false)) 86 406 332 46
    $gInd = Grouper $s2 @($ci, $li, $ti) "Indicateur"

    [void](Label $s2 "LA DÉMARCHE" 500 126 416 $BLUE)
    $etapes = @(
        @("Collecter", "Données ouvertes Parcoursup 2025, sur le portail du ministère chargé de l'Enseignement supérieur (Licence Ouverte)."),
        @("Calculer", "Quatre colonnes ajoutées par formule : RECHERCHEV, SIERREUR et SI imbriqués."),
        @("Synthétiser", "Tri, filtres, sous-totaux et tableaux croisés dynamiques."),
        @("Visualiser", "Un graphique en barres qui répond à la question.")
    )
    $gE = @()
    for ($k = 0; $k -lt 4; $k++) {
        $yy = 154 + $k * 84
        $b = Bulle $s2 ($k + 1) 500 $yy 36 $NAVY
        $a = Txt $s2 $etapes[$k][0] 552 ($yy - 2) 364 22 17 $NAVY $true
        $d = Txt $s2 $etapes[$k][1] 552 ($yy + 21) 364 40 13 $INK2
        $gE += (Grouper $s2 @($b, $a, $d) "Etape $($k + 1)")
    }
    "etape: diapositive 2 OK"

    # ======================================================================
    # DIAPOSITIVE 3 — les donnees
    # ======================================================================
    $s3 = $pres.Slides.AddSlide(3, $laySeul)
    $s3.Shapes.Title.TextFrame.TextRange.Text = "Les données"
    $p3 = Pic $s3 (Join-Path $Captures "capture_donnees.png") 44 100 872 184.9 "Capture d'écran de la feuille Données du classeur Excel : une ligne par formation, en-têtes bleus pour les colonnes d'origine, en-têtes orange pour les quatre colonnes calculées."
    $k1 = $s3.Shapes.AddShape(1, 44, 297, 10, 10); $k1.Fill.ForeColor.RGB = RGB $NAVY; $k1.Line.Visible = 0
    $k1t = Txt $s3 "Colonnes d'origine (portail open data)" 60 293 280 16 11.5 $INK2
    $k2 = $s3.Shapes.AddShape(1, 330, 297, 10, 10); $k2.Fill.ForeColor.RGB = RGB "C8501E"; $k2.Line.Visible = 0
    $k2t = Txt $s3 "Colonnes calculées par mes formules" 346 293 300 16 11.5 $INK2
    $gLeg = Grouper $s3 @($k1, $k1t, $k2, $k2t) "Legende des colonnes"

    $c31 = Card $s3 44 322 428 164 $PALE
    $l31 = Label $s3 "LE JEU DE DONNÉES" 64 337 388 $BLUE
    $t31 = Txt $s3 "2 691 formations d'Île-de-France proposées sur Parcoursup en 2025. Une ligne = une formation dans un établissement : ses places, ses candidatures, ses admis." 64 358 388 76 14 $INK
    $u31 = Txt $s3 "Source : data.enseignementsup-recherche.gouv.fr — Licence Ouverte, téléchargé le 12/10/2026." 64 446 388 30 10.5 $INK2
    $g31 = Grouper $s3 @($c31, $l31, $t31, $u31) "Description des donnees"

    $c32 = Card $s3 488 322 428 164 $PALEO
    $l32 = Label $s3 "MES QUATRE CALCULS" 508 337 388 $ORANGED
    $t32 = Runs $s3 @(
        @("Zone", 13, $INK, $true), @(" — RECHERCHEV sur le code du département`r", 13, $INK2, $false),
        @("Candidatures par place", 13, $INK, $true), @(" — candidatures ÷ places (SIERREUR)`r", 13, $INK2, $false),
        @("Taux de remplissage", 13, $INK, $true), @(" — admis ÷ places`r", 13, $INK2, $false),
        @("Niveau de tension", 13, $INK, $true), @(" — SI imbriqués, de Faible à Très forte", 13, $INK2, $false)
    ) 508 360 388 116 7
    $g32 = Grouper $s3 @($c32, $l32, $t32) "Mes calculs"
    "etape: diapositive 3 OK"

    # ======================================================================
    # DIAPOSITIVE 4 — tri et filtre
    # ======================================================================
    $s4 = $pres.Slides.AddSlide(4, $laySeul)
    $s4.Shapes.Title.TextFrame.TextRange.Text = Typo "Tri et filtre : le cas de mon université"
    $t4 = @(
        @("87 formations", "extraites par filtre avancé : celles de l'Université Paris Nanterre", $PALE),
        @("19,8", "candidatures par place en moyenne, contre 27,4 dans toute l'Île-de-France", $PALE),
        @("41e sur 87", "le rang de ma licence, MIASHS, avec 19,2 candidatures par place", $PALEO)
    )
    $g4 = @()
    for ($k = 0; $k -lt 3; $k++) {
        $xx = 44 + $k * 296
        $c = Card $s4 $xx 102 280 90 $t4[$k][2] $null 0.10
        $a = Txt $s4 $t4[$k][0] ($xx + 18) 108 244 38 28 $NAVY $true
        $b = Txt $s4 $t4[$k][1] ($xx + 18) 148 244 36 12 $INK2
        $g4 += (Grouper $s4 @($c, $a, $b) "Chiffre $($k + 1)")
    }
    $p4 = Pic $s4 (Join-Path $Captures "capture_zoom_nanterre.png") 54 206 852 276.1 "Capture d'écran de la feuille Zoom Paris Nanterre : zone de critères du filtre avancé, puis les formations de l'université triées de la plus demandée à la moins demandée."
    "etape: diapositive 4 OK"

    # ======================================================================
    # DIAPOSITIVE 5 — tableau de synthese
    # ======================================================================
    $s5 = $pres.Slides.AddSlide(5, $laySeul)
    $s5.Shapes.Title.TextFrame.TextRange.Text = "Le tableau de synthèse"
    $p5 = Pic $s5 (Join-Path $Captures "capture_tcd.png") 44 104 520 293.6 "Capture d'écran du tableau croisé dynamique : pour chaque filière, nombre de formations, somme des places, somme des candidatures et candidatures par place."
    $cap5 = Txt $s5 "Feuille « TCD » : tableau croisé dynamique par filière, trié de la plus demandée à la moins demandée. Le champ calculé « Cand. par place » divise les candidatures par les places." 44 408 520 48 11.5 $INK2

    $l51 = Label $s5 "CE QUE DIT LE TABLEAU" 588 108 328 $BLUE
    $t51 = Runs $s5 @(
        @("4,2 millions de candidatures pour 153 229 places : en moyenne ", 14, $INK, $false),
        @("27,4 candidatures par place", 14, $NAVY, $true),
        @(". Mais l'écart est énorme : de 10 dans les écoles de travail social (EFTS) à 81 en PASS.", 14, $INK, $false)
    ) 588 130 328 112
    $g51 = Grouper $s5 @($l51, $t51) "Lecture du tableau"

    $l52 = Label $s5 "ET SELON LA ZONE ? (SECOND TCD)" 588 232 328 $BLUE
    $z5 = @(@("29,3", "Paris"), @("27,9", "Petite couronne"), @("23,5", "Grande couronne"))
    $g5z = @()
    for ($k = 0; $k -lt 3; $k++) {
        $xx = 588 + $k * 113.5
        $c = Card $s5 $xx 254 101 76 $PALE $null 0.10
        $a = Txt $s5 $z5[$k][0] ($xx + 6) 262 89 30 23 $NAVY $true 2
        $b = Txt $s5 $z5[$k][1] ($xx + 4) 298 93 28 10.5 $INK2 $false 2
        $g5z += (Grouper $s5 @($c, $a, $b) "Zone $($k + 1)")
    }
    $n5 = Card $s5 588 348 328 62 $PALEO $null 0.10
    $n5t = Runs $s5 @(@("Contrôle : ", 11.5, $ORANGED, $true), @("les sous-totaux et les formules NB.SI / SOMME.SI redonnent exactement les mêmes totaux.", 11.5, $INK, $false)) 602 358 300 44
    $g5n = Grouper $s5 @($n5, $n5t) "Controle"
    "etape: diapositive 5 OK"

    # ======================================================================
    # DIAPOSITIVE 6 — graphique
    # ======================================================================
    $s6 = $pres.Slides.AddSlide(6, $laySeul)
    $s6.Shapes.Title.TextFrame.TextRange.Text = Typo "Réponse : les filières de santé sont les plus disputées"
    $gs = $s6.Shapes.AddChart2(-1, 57, 44, 98, 590, 388, $true)
    $gs.Name = "Graphique des filieres"
    $gs.AlternativeText = "Graphique en barres : candidatures par place selon la filière en Île-de-France. PASS 81,1 ; IFSI 58,4 ; BUT 40,9 ; Licence L.AS 40,8 ; CPGE 38,6 ; BTS 28,9 ; école d'ingénieur 26,4 ; école de commerce 20,8 ; autre formation 20,2 ; licence 17,5 ; EFTS 10,4."
    $ch = $gs.Chart
    # Le classeur de donnees du graphique s'ouvre dans un Excel a part : il
    # n'est pas toujours pret tout de suite, d'ou l'attente et les relances.
    # Le classeur peut exister avant que ses feuilles soient accessibles : on
    # n'avance que lorsque la premiere feuille repond.
    $cwb = $null; $cws = $null
    for ($essai = 1; $essai -le 30 -and -not $cws; $essai++) {
        try {
            if ($essai % 6 -eq 1) { $ch.ChartData.Activate() }
            $cwb = $ch.ChartData.Workbook
            if ($cwb) { $cws = $cwb.Worksheets.Item(1); [void]$cws.Name }
        }
        catch { $cws = $null }
        if (-not $cws) { Start-Sleep -Milliseconds 700 }
    }
    if (-not $cws) { throw "Le classeur de donnees du graphique ne s'est pas ouvert." }
    $cws.Range("A1:F30").ClearContents()
    # sur le classeur de donnees d'un graphique PowerPoint, l'affectation d'un
    # nombre a Value2 echoue depuis PowerShell : on passe par Formula, qui
    # attend du texte en notation anglaise (point decimal)
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $cws.Range("A1").Formula = "Filière"
    $cws.Range("B1").Formula = "Candidatures par place"
    for ($k = 0; $k -lt $noms.Count; $k++) {
        $cws.Cells($k + 2, 1).Formula = [string]$noms[$k]
        $cws.Cells($k + 2, 2).Formula = $vals[$k].ToString("R", $inv)
    }
    $fin = $noms.Count + 1
    try { $cws.ListObjects.Item(1).Resize($cws.Range("A1:B$fin")) } catch { }
    $ch.SetSourceData("='" + $cws.Name + "'!`$A`$1:`$B`$$fin")
    $cwb.Close()
    $ch.HasLegend = $false
    $ch.HasTitle = $true
    $ch.ChartTitle.Text = "Candidatures par place selon la filière — Île-de-France, 2025"
    $ch.ChartTitle.Font.Size = 14; $ch.ChartTitle.Font.Bold = $true; $ch.ChartTitle.Font.Color = RGB $INK
    $ax = $ch.Axes(1)
    $ax.ReversePlotOrder = $true
    $ax.TickLabels.Font.Size = 13; $ax.TickLabels.Font.Color = RGB $INK
    $ax.Format.Line.ForeColor.RGB = RGB "C3C2B7"
    $ax.MajorTickMark = -4142
    $av = $ch.Axes(2)
    $av.HasMajorGridlines = $false
    $av.Delete()
    $ch.ChartGroups(1).GapWidth = 75
    $se = $ch.SeriesCollection(1)
    $se.Format.Fill.ForeColor.RGB = RGB $BLUE
    $se.HasDataLabels = $true
    $dl = $se.DataLabels()
    $dl.NumberFormatLinked = $false
    $dl.NumberFormat = "0.0"
    $dl.Position = 2
    $dl.Font.Size = 13; $dl.Font.Bold = $true; $dl.Font.Color = RGB $INK
    for ($k = 0; $k -lt $noms.Count; $k++) { if ($noms[$k] -eq "Licence") { $se.Points($k + 1).Format.Fill.ForeColor.RGB = RGB $ORANGE } }
    $ch.ChartArea.Format.Fill.Visible = 0
    $ch.ChartArea.Format.Line.Visible = 0
    $ch.PlotArea.Format.Fill.Visible = 0
    "  graphique : etiquettes " + (($se.Points(1).DataLabel.Text, $se.Points(2).DataLabel.Text, $se.Points(10).DataLabel.Text) -join " / ")

    [void](Label $s6 "À RETENIR" 658 108 258 $BLUE)
    $ret = @(
        @("La santé loin devant. ", "81 candidatures par place en PASS, 58 en IFSI."),
        @("La licence, peu tendue. ", "Elle offre 4 places sur 10 et ne reçoit que 17,5 candidatures par place."),
        @("Paris devant la banlieue. ", "29,3 à Paris contre 23,5 en grande couronne.")
    )
    $g6 = @()
    for ($k = 0; $k -lt 3; $k++) {
        $yy = 134 + $k * 80
        $b = Bulle $s6 ($k + 1) 658 $yy 26 $(if ($k -eq 1) { "C8501E" } else { $NAVY })
        $a = Runs $s6 @(@($ret[$k][0], 13.5, $NAVY, $true), @($ret[$k][1], 13.5, $INK, $false)) 694 ($yy + 1) 222 70
        $g6 += (Grouper $s6 @($b, $a) "A retenir $($k + 1)")
    }
    $c6 = Card $s6 658 378 258 66 $PALEO $null 0.10
    $c6t = Runs $s6 @(@("À nuancer : ", 11.5, $ORANGED, $true), @("un candidat formule plusieurs vœux. Le ratio mesure l'attractivité, pas la chance d'être admis.", 11.5, $INK, $false)) 670 386 234 54
    $g6n = Grouper $s6 @($c6, $c6t) "Nuance"
    [void](Txt $s6 ("Source : Parcoursup 2025, MESR-SIES (données ouvertes). Calculs : " + $ID.signature + ".") 658 454 258 30 9.5 $INK2)
    "etape: diapositive 6 OK"

    # ======================================================================
    # PIEDS DE PAGE, NOTES, NARRATION, TRANSITIONS
    # ======================================================================
    for ($i = 1; $i -le 6; $i++) {
        $s = $pres.Slides.Item($i)
        $hf = $s.HeadersFooters
        if ($i -gt 1) {
            $hf.Footer.Visible = -1; $hf.Footer.Text = $pied
            $hf.SlideNumber.Visible = -1
        }
        else { $hf.Footer.Visible = 0; $hf.SlideNumber.Visible = 0 }
        $hf.DateAndTime.Visible = 0
        $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = (Remplir $J.narration[$i - 1].notes)

        if (-not $SansVoix) {
            $au = $s.Shapes.AddMediaObject2((Join-Path $Narration "narration_diapo$i.wav"), 0, -1, 886, 28, 28, 28)
            $au.Name = "Narration"
            $au.AlternativeText = "Narration audio de la diapositive $i"
            $e = $s.TimeLine.MainSequence.AddEffect($au, 83, 0, 2)
            $e.MoveTo(1)
            $au.AnimationSettings.PlaySettings.HideWhileNotPlaying = 0
        }

        $tr = $s.SlideShowTransition
        $tr.EntryEffect = 3849
        $tr.Duration = 0.8
        $tr.AdvanceOnClick = -1
        if ($SansVoix) { $tr.AdvanceOnTime = 0 }
        else {
            $tr.AdvanceOnTime = -1
            $tr.AdvanceTime = [Math]::Round($durees[$i - 1] + 1.5, 1)
        }
    }
    "etape: pieds de page, notes, narration et transitions OK"

    # ======================================================================
    # ANIMATIONS (calees sur la narration)
    # ======================================================================
    [void](Anim $s1 $g1[0] 14.0 42); [void](Anim $s1 $g1[1] 17.0 42); [void](Anim $s1 $g1[2] 20.0 42)

    [void](Anim $s2 $q2 7.0); [void](Anim $s2 $gInd 11.5 42)
    [void](Anim $s2 $gE[0] 18.5); [void](Anim $s2 $gE[1] 21.5); [void](Anim $s2 $gE[2] 23.5); [void](Anim $s2 $gE[3] 25.5)

    [void](Anim $s3 $p3 0.4); [void](Anim $s3 $gLeg 1.2); [void](Anim $s3 $g31 2.5 42); [void](Anim $s3 $g32 12.0 42)

    [void](Anim $s4 $g4[0] 3.5 42); [void](Anim $s4 $p4 8.0); [void](Anim $s4 $g4[1] 13.0 42); [void](Anim $s4 $g4[2] 19.0 42)

    [void](Anim $s5 $p5 0.4); [void](Anim $s5 $cap5 1.2); [void](Anim $s5 $g51 5.0)
    [void](Anim $s5 $l52 13.5); [void](Anim $s5 $g5z[0] 15.5 42); [void](Anim $s5 $g5z[1] 18.0 42); [void](Anim $s5 $g5z[2] 20.5 42)
    [void](Anim $s5 $g5n 25.0)

    $ew = Anim $s6 $gs 0.4 22 1.6
    try { $ew.EffectParameters.Direction = 4 } catch { }
    [void](Anim $s6 $g6[0] 4.0); [void](Anim $s6 $g6[1] 13.0); [void](Anim $s6 $g6[2] 19.0); [void](Anim $s6 $g6n 23.5 42)
    "etape: animations OK (" + ((1..6 | ForEach-Object { $pres.Slides.Item($_).TimeLine.MainSequence.Count }) -join ", ") + " effets par diapositive)"

    $pres.SlideShowSettings.AdvanceMode = $(if ($SansVoix) { 1 } else { 2 })
    try {
        $pres.BuiltInDocumentProperties.Item("Title").Value = "Parcoursup 2025 en Île-de-France : où la concurrence est-elle la plus forte ?"
        $pres.BuiltInDocumentProperties.Item("Author").Value = $ID.nom_prenom
        $pres.BuiltInDocumentProperties.Item("Subject").Value = "Projet multimédia — $($ID.numero)"
    }
    catch { "  (proprietes du document non modifiables par COM : corrigees ensuite)" }

    $pres.SaveAs($Out)
    "etape: enregistre -> $Out"
}
catch {
    "ERREUR ligne " + $_.InvocationInfo.ScriptLineNumber + " : " + $_.Exception.Message
    $_.ScriptStackTrace
    throw
}
finally {
    if ($pres) { $pres.Close() }
    if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
}
"TERMINE"
