# Seance 5, activite 1 : « Numerique & Environnement » mis en forme dans
# PowerPoint, exercice par exercice (theme, masque, liens, SmartArt, icone,
# zone de texte, graphique en secteurs, nouvelle diapositive, transitions,
# animations). A reencoder en UTF-8 avec BOM avant execution (bom.ps1).
param(
    [Parameter(Mandatory = $true)][string]$Src,
    [Parameter(Mandatory = $true)][string]$Images,
    [Parameter(Mandatory = $true)][string]$Identite,
    [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = "Stop"
$THEMES = "C:\Program Files\Microsoft Office\root\Document Themes 16"
$NB = [string][char]0x00A0
$ID = Get-Content -Raw -Encoding UTF8 $Identite | ConvertFrom-Json

function RGB($hex) {
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return $r + 256 * $g + 65536 * $b
}
function Typo([string]$t) {
    $t = $t -replace ' ([?:;!»%])', ($NB + '$1')
    $t = $t -replace '« ', ('«' + $NB)
    $t = $t -replace '(\d) (\d{3})', ('$1' + $NB + '$2')
    return $t
}
function Forme($sl, [string]$nom) {
    foreach ($sh in $sl.Shapes) { if ($sh.Name -eq $nom) { return $sh } }
    throw "Forme introuvable sur la diapositive $($sl.SlideIndex) : $nom"
}
function Placer($sh, [single]$l, [single]$t, [single]$w, [single]$h) {
    $sh.Left = $l; $sh.Top = $t; $sh.Width = $w; $sh.Height = $h
}
# ecrit des paragraphes (niveau, texte) dans une zone de texte
function Paragraphes($sh, $items, [single]$t1, [single]$t2) {
    $tr = $sh.TextFrame.TextRange
    $tr.Text = (($items | ForEach-Object { Typo $_[1] }) -join "`r")
    for ($i = 0; $i -lt $items.Count; $i++) {
        $p = $tr.Paragraphs($i + 1)
        $p.IndentLevel = $items[$i][0]
        $p.Font.Size = $(if ($items[$i][0] -eq 1) { $t1 } else { $t2 })
        $p.Font.Bold = 0
    }
}
function SmartArtDe($sl) {
    foreach ($sh in $sl.Shapes) { if ($sh.HasSmartArt) { return $sh } }
    throw "Aucun SmartArt sur la diapositive $($sl.SlideIndex)"
}

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add(-1)
try { $ppt.WindowState = 2 } catch { }
if (Test-Path $Out) { Remove-Item $Out -Force }

try {
    # ======================================================================
    # EXERCICE 1 — nouvelle presentation, modele, variante, couleurs, masque
    # ======================================================================
    $thmx = "$THEMES\Facet.thmx"
    $tv = $ppt.OpenThemeFile($thmx)
    $pres.ApplyTemplate2($thmx, $tv.ThemeVariants.Item(1).Id)
    [void]$pres.Slides.InsertFromFile($Src, 0)
    "ex1: modele Facet applique, " + $pres.Slides.Count + " diapositives importees"

    # masque : jeu de couleurs « Vert » et style d'arriere-plan
    $m = $pres.SlideMaster
    $m.Theme.ThemeColorScheme.Load("$THEMES\Theme Colors\Green.xml")
    $m.BackgroundStyle = 1
    "ex1: masque — couleurs du theme « Vert », style d'arriere-plan " + $m.BackgroundStyle

    foreach ($cible in @($m) + @(1..$m.CustomLayouts.Count | ForEach-Object { $m.CustomLayouts.Item($_) })) {
        foreach ($sh in $cible.Shapes) {
            if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -in 13, 15) { $sh.TextFrame.TextRange.Font.Size = 11 }
        }
    }
    "ex1: masque — pied de page et numero de diapositive passes en corps 11"

    $L = @{}
    for ($i = 1; $i -le $m.CustomLayouts.Count; $i++) { $L[$m.CustomLayouts.Item($i).Name] = $m.CustomLayouts.Item($i) }
    $layCont = $L["Titre et contenu"]

    # les titres de 11 a 14 sont de simples zones de texte dans l'original :
    # on retablit le vrai titre de la disposition, puis on y reporte le texte
    foreach ($n in 11, 12, 13, 14) {
        $s = $pres.Slides.Item($n)
        $faux = Forme $s "Titre 1"
        $s.CustomLayout = $layCont
        $vrai = $null
        foreach ($sh in $s.Shapes) { if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -eq 1) { $vrai = $sh } }
        if (-not $vrai) { $s.Layout = 2; foreach ($sh in $s.Shapes) { if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -eq 1) { $vrai = $sh } } }
        if ($vrai) { $vrai.TextFrame.TextRange.Text = "x"; $faux.Delete() } else { $faux.Name = "Title 1" }
    }
    "ex1: titres des diapositives 11 a 14 replaces dans l'espace reserve du titre"

    # titres : meme emplacement et meme taille partout
    $titres = @{
        2 = "Introduction"; 3 = "Introduction"
        4 = "Protection de l'environnement : le dilemme"; 5 = "Protection de l'environnement : le dilemme"
        6 = "Impact du numérique sur l'environnement"; 7 = "Impact du numérique sur l'environnement"
        8 = "Impact du numérique sur l'environnement"; 9 = "Impact du numérique sur l'environnement"
        10 = "Quelques éco-gestes pour un numérique responsable"; 11 = "Quelques éco-gestes pour un numérique responsable"
        12 = "Quelques éco-gestes pour un numérique responsable"; 13 = "Quelques éco-gestes pour un numérique responsable"
        14 = "Quelques éco-gestes pour un numérique responsable"; 15 = "Sources et références"
    }
    foreach ($n in $titres.Keys) {
        $s = $pres.Slides.Item($n)
        $t = $s.Shapes.Title
        Placer $t 53 40 677 86
        $t.TextFrame.TextRange.Text = Typo $titres[$n]
        $t.TextFrame.TextRange.Font.Size = 28
        $t.TextFrame.VerticalAnchor = 3
    }
    $t1 = $pres.Slides.Item(1).Shapes.Title.TextFrame.TextRange
    $t1.Text = Typo "Compétences numériques : protéger la santé, le bien-être et l'environnement"
    $t1.Font.Size = 36
    # identite de l'auteur de la mise en forme, sous le sous-titre
    $st1 = (Forme $pres.Slides.Item(1) "Subtitle 2").TextFrame.TextRange
    $st1.Text = "Vers un numérique responsable`rMise en forme : $($ID.nom_prenom) — $($ID.numero)"
    $st1.Paragraphs(2).Font.Size = 14
    "ex1: titres harmonises"

    # ======================================================================
    # EXERCICE 2 — objets, liens et illustrations
    # ======================================================================
    # 2.2 liens hypertextes (diapositive 2)
    $s = $pres.Slides.Item(2)
    $c = Forme $s "Content Placeholder 2"
    Placer $c 53 140 677 340
    $tr = $c.TextFrame.TextRange
    $tr.Paragraphs(1).Font.Size = 18; $tr.Paragraphs(2).Font.Size = 18
    $tr.Paragraphs(3).Font.Size = 15; $tr.Paragraphs(4).Font.Size = 15
    $f = $tr.Find("Pix")
    $f.ActionSettings(1).Hyperlink.Address = "https://pix.fr/"
    $f.ActionSettings(1).Hyperlink.ScreenTip = "Site de Pix"
    # « Protection et sécurité » : les espaces de l'original ne sont pas tous des
    # espaces ordinaires, on delimite donc la plage par son premier et son
    # dernier mot plutot que par une recherche de la chaine entiere
    $a = $tr.Find("Protection")
    $b = $tr.Find("curit", $a.Start)
    $fin = $b.Start + $b.Length
    while ($tr.Characters($fin, 1).Text -match '[\p{L}\p{M}]') { $fin++ }
    $f = $tr.Characters($a.Start, $fin - $a.Start)
    "ex2: plage du second lien = [" + $f.Text + "]"
    $f.ActionSettings(1).Hyperlink.Address = "https://pix.fr/competences"
    $f.ActionSettings(1).Hyperlink.ScreenTip = "Les compétences évaluées par Pix"
    "ex2: liens hypertextes poses sur « Pix » et « Protection et sécurité » (" + $s.Hyperlinks.Count + " liens)"

    # 2.3 SmartArt « Liste de blocs simple » (diapositive 3)
    $s = $pres.Slides.Item(3)
    $c = Forme $s "Content Placeholder 2"
    $themes = @("Ergonomie du poste de travail", "Communication sans fil et ondes", "Impact environnemental",
        "Accessibilité", "Vie connectée", "Capteurs", "Intelligence artificielle et robots", "Santé",
        "Vie privée et confidentialité")
    $c.TextFrame.TextRange.Text = "Thématiques associées"
    $c.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0
    $c.TextFrame.TextRange.Font.Size = 20; $c.TextFrame.TextRange.Font.Bold = -1
    Placer $c 53 132 677 36
    $layBloc = $null; $layVert = $null
    for ($i = 1; $i -le $ppt.SmartArtLayouts.Count; $i++) {
        # ne pas nommer cette variable $id : PowerShell ignore la casse et elle
        # ecraserait $ID, l'identite lue en debut de script
        $refDispo = $ppt.SmartArtLayouts.Item($i).Id
        if ($refDispo -like "*/layout/default") { $layBloc = $ppt.SmartArtLayouts.Item($i) }
        if ($refDispo -like "*/layout/vList5") { $layVert = $ppt.SmartArtLayouts.Item($i) }
    }
    $sa = $s.Shapes.AddSmartArt($layBloc, 53, 176, 677, 300)
    $sa.Name = "SmartArt thematiques"
    $nodes = $sa.SmartArt.AllNodes
    while ($nodes.Count -lt $themes.Count) { [void]$nodes.Add() }
    while ($nodes.Count -gt $themes.Count) { $nodes.Item($nodes.Count).Delete() }
    for ($i = 0; $i -lt $themes.Count; $i++) { $nodes.Item($i + 1).TextFrame2.TextRange.Text = $themes[$i] }
    $sa.AlternativeText = "Liste en blocs des neuf thématiques associées : " + ($themes -join ", ") + "."
    "ex2: SmartArt « " + $layBloc.Name + " » : " + $sa.SmartArt.AllNodes.Count + " blocs"

    # 2.4 icone ampoule devant « Le saviez-vous ? » + 3.1 image (diapositive 8)
    $s = $pres.Slides.Item(8)
    $c = Forme $s "Content Placeholder 2"
    $tr = $c.TextFrame.TextRange
    $tr.Text = (Typo "Le saviez-vous ?") + "`r" + (Typo "Pour calculer l'empreinte carbone d'un appareil, on compte toute l'énergie dépensée !")
    $tr.ParagraphFormat.Bullet.Visible = 0
    $tr.Paragraphs(1).Font.Size = 24; $tr.Paragraphs(1).Font.Bold = -1
    $tr.Paragraphs(2).Font.Size = 17
    $tr.Paragraphs(1).ParagraphFormat.SpaceAfter = 8
    Placer $c 100 136 420 110
    $amp = $s.Shapes.AddPicture("$Images\ampoule.png", 0, -1, 53, 136, 38, 38)
    $amp.Name = "Icone ampoule"
    $amp.AlternativeText = "Icône : ampoule"
    $ex = Forme $s "TextBox 3"
    $ex.TextFrame.TextRange.Text = "Exemple`r" + (Typo "Pour un smartphone, on analyse tout son « cycle de vie » : extraction des matières premières, fabrication, utilisation et recyclage. L'énergie consommée par un portable, ce n'est donc pas seulement le courant qui sort de la prise quand on le recharge, mais beaucoup, beaucoup plus !")
    $ex.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0
    $ex.TextFrame.TextRange.Font.Size = 15
    $ex.TextFrame.TextRange.Paragraphs(1).Font.Bold = -1
    $ex.TextFrame.WordWrap = -1; $ex.TextFrame.AutoSize = 0
    Placer $ex 53 262 450 150
    $ex.ShapeStyle = 23
    $ex.TextFrame.MarginLeft = 12; $ex.TextFrame.MarginRight = 12; $ex.TextFrame.MarginTop = 10; $ex.TextFrame.MarginBottom = 10
    $ex.TextFrame.VerticalAnchor = 3
    Placer $ex 53 262 450 150
    $emp = $s.Shapes.AddPicture("$Images\empreinte_carbone_commons_CC0.png", 0, -1, 548, 150, 170, 215.9)
    $emp.LockAspectRatio = -1
    $emp.Name = "Image empreinte carbone"
    $emp.AlternativeText = "Pictogramme d'une empreinte de pied portant la mention CO2, symbole de l'empreinte carbone."
    $src8 = $s.Shapes.AddTextbox(1, 520, 372, 226, 40)
    $src8.TextFrame.TextRange.Text = "Image : « Carbon footprint icon », Tommaso.sansone91, Wikimedia Commons (CC0)"
    $src8.TextFrame.TextRange.Font.Size = 9.5; $src8.TextFrame.TextRange.Font.Italic = -1
    $src8.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $src8.Name = "Source de l'image"
    "ex2: icone ampoule et image d'empreinte carbone inserees (diapositive 8)"

    # 2.5 zone de texte avec style rapide (diapositive 5)
    $s = $pres.Slides.Item(5)
    $im = Forme $s "Picture 3"
    $im.LockAspectRatio = -1
    $im.Width = 420; $im.Left = 53; $im.Top = 148
    $zt = $s.Shapes.AddTextbox(1, 492, 172, 238, 240)
    $zt.Name = "Zone de texte GES"
    $zt.TextFrame.WordWrap = -1
    $zt.TextFrame.TextRange.Text = Typo "L'émission des GES des appareils a doublé par rapport à 2012 et ne cesse de croître. Il peut encore doubler en 2025 si on n'applique pas les éco-gestes pour un numérique responsable"
    $zt.TextFrame.TextRange.Font.Size = 16
    $zt.ShapeStyle = 30
    $zt.TextFrame.MarginLeft = 14; $zt.TextFrame.MarginRight = 14; $zt.TextFrame.MarginTop = 12; $zt.TextFrame.MarginBottom = 12
    $zt.TextFrame.VerticalAnchor = 3
    $zt.TextFrame.AutoSize = 0
    Placer $zt 492 172 238 240
    $im.AlternativeText = "Graphique : part du numérique dans les émissions mondiales de gaz à effet de serre, de 2013 à 2025, selon plusieurs scénarios (The Shift Project, 2018)."
    "ex2: zone de texte inseree a cote de la figure, style rapide " + $zt.ShapeStyle

    # 2.6 graphique en secteurs (diapositive 9)
    $s = $pres.Slides.Item(9)
    foreach ($sh in @($s.Shapes)) { if ($sh.Type -eq 3) { $sh.Delete() } }
    $c = Forme $s "Content Placeholder 5"
    $c.TextFrame.TextRange.Text = Typo "Le numérique absorbe 10 % de la consommation mondiale d'électricité, un chiffre qui grimpe de 10 % par an :"
    $c.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0
    $c.TextFrame.TextRange.Font.Size = 18
    Placer $c 53 132 677 58
    $gs = $s.Shapes.AddChart2(-1, 5, 53, 196, 677, 290, $true)
    $gs.Name = "Graphique en secteurs"
    $ch = $gs.Chart
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
    $cws.Range("B1").Formula = "Part de consommation"
    $parts = @(@("Ordinateurs, smartphones et objets connectés", "0.3"), @("Data-centers qui hébergent nos données", "0.3"), @("Réseaux sur lesquels circulent les données", "0.4"))
    for ($k = 0; $k -lt 3; $k++) { $cws.Cells($k + 2, 1).Formula = $parts[$k][0]; $cws.Cells($k + 2, 2).Formula = $parts[$k][1] }
    try { $cws.ListObjects.Item(1).Resize($cws.Range("A1:B4")) } catch { }
    $ch.SetSourceData("='" + $cws.Name + "'!`$A`$1:`$B`$4")
    $cwb.Close()
    $ch.HasTitle = $true
    $ch.ChartTitle.Text = "Part de la consommation électrique du numérique"
    $ch.ChartTitle.Font.Size = 16; $ch.ChartTitle.Font.Bold = $true
    $ch.HasLegend = $true
    $ch.Legend.Position = -4152
    $ch.Legend.Font.Size = 15
    $se = $ch.SeriesCollection(1)
    $se.HasDataLabels = $true
    $dl = $se.DataLabels()
    $dl.ShowPercentage = $true; $dl.ShowValue = $false; $dl.ShowCategoryName = $false
    $dl.Position = 3
    $dl.Font.Size = 16; $dl.Font.Bold = $true; $dl.Font.Color = RGB "FFFFFF"
    $se.Points(3).DataLabel.Font.Color = RGB "1A1A19"
    $gs.AlternativeText = "Graphique en secteurs : 30 % pour les ordinateurs, smartphones et objets connectés ; 30 % pour les data-centers ; 40 % pour les réseaux."
    "ex2: graphique en secteurs cree (30 / 30 / 40)"

    # 2.7 SmartArt « Liste de blocs verticale » (diapositives 10 a 14)
    $listes = @{
        10 = @(
            @(1, "Allonger la durée de vie de son matériel"),
            @(2, "Privilégier le matériel reconditionné ou d'occasion, ou faire le choix d'appareils économes, pouvant être réparés"),
            @(2, "Éviter de multiplier les appareils"),
            @(2, "Bien entretenir son matériel avec des protections anti-virus pour éviter les pannes"),
            @(2, "Recycler son matériel informatique"))
        11 = @(
            @(1, "Aller vers la sobriété"),
            @(2, "Brancher ses équipements sur une multiprise à interrupteur et l'éteindre lorsqu'ils ne sont pas utilisés"),
            @(2, "Activer le « mode économie d'énergie » sur son ordinateur, ou le « mode avion » sur son smartphone"),
            @(2, "Fermer ou supprimer les applications inutilisées sur son smartphone"),
            @(2, "Désactiver le wifi de la box internet lorsqu'il n'est pas utilisé"))
        12 = @(
            @(1, "Mieux naviguer"),
            @(2, "Utiliser des mots-clés pour ses recherches"),
            @(2, "Consulter son historique et enregistrer ses sites préférés dans les favoris pour y accéder plus rapidement"),
            @(2, "Maîtriser ses usages : 10 minutes de vidéo en haute définition sur smartphone équivalent à un four de 2 000 W pendant 5 minutes. Mieux vaut télécharger."),
            @(1, "Limiter son temps d'écran"),
            @(2, "C'est bon pour la planète et pour soi !"))
        13 = @(
            @(1, "Stockage des données"),
            @(2, "Le stockage peut se faire en local, sur la mémoire de vos appareils, ou en externe grâce aux Clouds : ces espaces sont souvent gratuits mais très énergivores"),
            @(2, "Trier, classer et supprimer ses données régulièrement"),
            @(2, "Privilégier le stockage de ses données en local est plus éco-responsable"))
        14 = @(
            @(1, "Nettoyer sa boîte mail"),
            @(2, "Alléger les mails envoyés : fichiers compressés ou en basse définition, pièces jointes supprimées dans les réponses, liens hypertextes pour transmettre un document"),
            @(2, "Conserver uniquement les mails importants et nettoyer régulièrement sa boîte mail"),
            @(2, "Installer un anti-spam et se désinscrire des newsletters ou publicités que l'on ne lit pas"))
    }
    foreach ($n in 10, 11, 12, 13, 14) {
        $s = $pres.Slides.Item($n)
        $c = Forme $s "Content Placeholder 5"
        Placer $c 53 140 677 340
        Paragraphes $c $listes[$n] 18 14
        $c.ConvertTextToSmartArt($layVert)
        $sa = SmartArtDe $s
        $blocs = @($listes[$n] | Where-Object { $_[0] -eq 1 }).Count
        if ($blocs -eq 1) { Placer $sa 53 158 677 262 } else { Placer $sa 53 140 677 336 }
        for ($k = 1; $k -le $sa.SmartArt.AllNodes.Count; $k++) {
            $nd = $sa.SmartArt.AllNodes.Item($k)
            $nd.TextFrame2.TextRange.Font.Size = $(if ($nd.Level -eq 1) { [single]24 } else { [single]15 })
        }
        $sa.Name = "SmartArt eco-gestes"
        $sa.AlternativeText = "Liste verticale en blocs : " + (($listes[$n] | ForEach-Object { $_[1] }) -join " ; ")
    }
    "ex2: texte des diapositives 10 a 14 converti en SmartArt « " + $layVert.Name + " »"

    # ======================================================================
    # mise en forme des autres diapositives
    # ======================================================================
    $s = $pres.Slides.Item(4)
    $c = Forme $s "Content Placeholder 2"
    Placer $c 53 140 677 340
    $c.TextFrame.TextRange.Font.Size = 19

    $s = $pres.Slides.Item(6)
    $c = Forme $s "Content Placeholder 2"
    $c.TextFrame.TextRange.Font.Size = 18
    $c.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0
    Placer $c 53 136 677 104
    $ex = Forme $s "TextBox 3"
    $ex.TextFrame.TextRange.Text = "Exemples`r" + (Typo "L'envoi d'un mail d'un mégaoctet (1 Mo) correspond à l'utilisation d'une vieille ampoule de 60 watts pendant 25 minutes, soit l'équivalent de 20 g de CO2.") + "`r" + (Typo "Une heure de vidéo regardée en ligne consomme autant d'énergie qu'un réfrigérateur pendant une année entière.")
    $ex.TextFrame.TextRange.Font.Size = 15
    $ex.TextFrame.TextRange.Paragraphs(1).Font.Bold = -1
    $ex.TextFrame.TextRange.Paragraphs(1).ParagraphFormat.Bullet.Visible = 0
    foreach ($k in 2, 3) {
        $p = $ex.TextFrame.TextRange.Paragraphs($k)
        $p.ParagraphFormat.Bullet.Visible = -1; $p.ParagraphFormat.Bullet.Type = 1; $p.ParagraphFormat.Bullet.Character = 8226
        $p.ParagraphFormat.SpaceBefore = 6
    }
    $ex.TextFrame.WordWrap = -1; $ex.TextFrame.AutoSize = 0
    $ex.ShapeStyle = 23
    $ex.TextFrame.MarginLeft = 14; $ex.TextFrame.MarginRight = 14; $ex.TextFrame.MarginTop = 10; $ex.TextFrame.MarginBottom = 10
    $ex.TextFrame.VerticalAnchor = 3
    Placer $ex 53 262 677 150

    $s = $pres.Slides.Item(7)
    $c = Forme $s "Content Placeholder 5"
    $c.TextFrame.TextRange.Font.Size = 17
    $c.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0
    Placer $c 53 140 236 300
    $im = Forme $s "Picture 2"
    $im.LockAspectRatio = -1
    $im.Width = 420; $im.Left = 310; $im.Top = 144
    $im.AlternativeText = "Infographie ADEME : émissions de carbone d'un ordinateur à chaque étape de son cycle de vie ; la production des matières premières et la fabrication représentent la plus grande part."
    $cap = Forme $s "TextBox 8"
    $cap.TextFrame.WordWrap = -1; $cap.TextFrame.AutoSize = 0
    $cap.TextFrame.TextRange.Font.Size = 11.5
    $cap.TextFrame.TextRange.Font.Italic = -1
    Placer $cap 310 430 420 54

    $s = $pres.Slides.Item(15)
    $c = Forme $s "Content Placeholder 2"
    Placer $c 53 140 677 330
    $c.TextFrame.TextRange.Font.Size = 17
    "mise en forme des diapositives 4, 6, 7 et 15 OK"

    # ======================================================================
    # EXERCICE 3 — nouvelle diapositive « Contenu avec legende » apres la 10
    # ======================================================================
    $layLeg = $L["Contenu avec légende"]
    $ns = $pres.Slides.AddSlide(11, $layLeg)
    $tit = $null; $txt = $null; $obj = $null
    foreach ($sh in $ns.Shapes) {
        if ($sh.Type -ne 14) { continue }
        switch ($sh.PlaceholderFormat.Type) { 1 { $tit = $sh } 2 { $txt = $sh } 7 { $obj = $sh } }
    }
    "ex3: disposition « " + $ns.CustomLayout.Name + " » : " + (($ns.Shapes | ForEach-Object { $_.Name + "(" + $(if ($_.Type -eq 14) { $_.PlaceholderFormat.Type } else { "-" }) + ") [" + [int]$_.Left + "," + [int]$_.Top + " " + [int]$_.Width + "x" + [int]$_.Height + "]" }) -join " | ")
    $tit.TextFrame.TextRange.Text = Typo "L'empreinte carbone : la trace de nos usages"
    $txt.TextFrame.TextRange.Text = (Typo "Allonger la durée de vie d'un appareil, c'est amortir sur plus d'années le carbone émis pour le fabriquer : la fabrication représente la plus grosse part de son empreinte.") + "`r" + (Typo "Source de l'image : « Carbon footprint icon », Tommaso.sansone91, Wikimedia Commons, licence CC0 (domaine public).")
    $tit.TextFrame.TextRange.Font.Size = 24
    $txt.TextFrame.TextRange.Font.Size = 14
    $txt.TextFrame.TextRange.Paragraphs(2).Font.Size = 11.5
    $txt.TextFrame.TextRange.Paragraphs(2).Font.Italic = -1
    $ol = $obj.Left; $ot = $obj.Top; $ow = $obj.Width; $oh = $obj.Height
    $obj.Delete()
    $ratio = 886.0 / 1125.0
    $ph = 320; $pw = $ph * $ratio
    $pic = $ns.Shapes.AddPicture("$Images\empreinte_carbone_commons_CC0.png", 0, -1, ($ol + ($ow - $pw) / 2), ($ot + ($oh - $ph) / 2), $pw, $ph)
    $pic.LockAspectRatio = -1
    $pic.Name = "Image empreinte carbone"
    $pic.AlternativeText = "Pictogramme d'une empreinte de pied portant la mention CO2, symbole de l'empreinte carbone."
    "ex3: nouvelle diapositive 11 creee, image {0:N0} x {1:N0} pt (proportions conservees)" -f $pic.Width, $pic.Height

    # ======================================================================
    # EXERCICE 2.1 — numero de diapositive et pied de page, appliques partout
    # ======================================================================
    $annee = "Année universitaire $($ID.annee)"
    # PowerPoint repond parfois « Failed » juste apres un graphique ou un
    # SmartArt (il est encore occupe) : l'operation etant sans effet de bord,
    # on la reprend simplement.
    for ($essai = 1; $essai -le 6; $essai++) {
        try {
            foreach ($s in $pres.Slides) {
                $hf = $s.HeadersFooters
                $hf.Footer.Visible = -1; $hf.Footer.Text = $annee
                $hf.SlideNumber.Visible = -1
                $hf.DateAndTime.Visible = 0
                foreach ($sh in $s.Shapes) {
                    if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -in 13, 15) { $sh.TextFrame.TextRange.Font.Size = 11 }
                }
            }
            break
        }
        catch { if ($essai -eq 6) { throw }; Start-Sleep -Seconds 2 }
    }
    "ex2: pied de page « $annee » et numero sur les " + $pres.Slides.Count + " diapositives"

    # ======================================================================
    # EXERCICE 4 — transition appliquee a toutes les diapositives
    # ======================================================================
    foreach ($s in $pres.Slides) {
        $tr = $s.SlideShowTransition
        $tr.EntryEffect = 3853
        $tr.Duration = 1.0
        $tr.AdvanceOnClick = -1
        $tr.AdvanceOnTime = 0
    }
    "ex4: transition (EntryEffect " + $pres.Slides.Item(2).SlideShowTransition.EntryEffect + ", 1 s) appliquee partout"

    # ======================================================================
    # EXERCICE 5 — animations des diapositives 5 et 7
    # ======================================================================
    $s = $pres.Slides.Item(5)
    $e1 = $s.TimeLine.MainSequence.AddEffect((Forme $s "Picture 3"), 10, 0, 1)
    $e2 = $s.TimeLine.MainSequence.AddEffect((Forme $s "Zone de texte GES"), 10, 0, 3)
    $e2.Timing.TriggerDelayTime = 4
    $s = $pres.Slides.Item(7)
    $e3 = $s.TimeLine.MainSequence.AddEffect((Forme $s "Content Placeholder 5"), 10, 0, 1)
    $e4 = $s.TimeLine.MainSequence.AddEffect((Forme $s "TextBox 8"), 10, 0, 3)
    $e4.Timing.TriggerDelayTime = 4
    "ex5: animations — diapo 5 : {0} effets ; diapo 7 : {1} effets (2e effet : apres la precedente, delai {2} s)" -f $pres.Slides.Item(5).TimeLine.MainSequence.Count, $pres.Slides.Item(7).TimeLine.MainSequence.Count, $e4.Timing.TriggerDelayTime

    $pres.SaveAs($Out)
    "enregistre -> $Out (" + $pres.Slides.Count + " diapositives)"
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

