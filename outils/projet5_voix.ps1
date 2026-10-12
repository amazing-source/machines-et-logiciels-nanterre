# Seance 5, projet : fichiers audio de narration, un par diapositive, generes
# avec la voix de synthese francaise de Windows (Microsoft Hortense).
# C'est une voix de SYNTHESE : elle tient lieu de narration en attendant que
# l'etudiant enregistre la sienne (Insertion > Audio > Enregistrer l'audio).
# Usage : projet5_voix.ps1 -Textes projet5_textes.json -Identite identite.json -OutDir dossier
param(
    [Parameter(Mandatory = $true)][string]$Textes,
    [Parameter(Mandatory = $true)][string]$Identite,
    [Parameter(Mandatory = $true)][string]$OutDir
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Speech
New-Item -ItemType Directory -Force $OutDir | Out-Null
$j = Get-Content -Raw -Encoding UTF8 $Textes | ConvertFrom-Json
$ID = Get-Content -Raw -Encoding UTF8 $Identite | ConvertFrom-Json
# les textes contiennent des champs {nom} remplis depuis le fichier d'identite
function Remplir([string]$t) {
    foreach ($p in $ID.PSObject.Properties) { $t = $t.Replace('{' + $p.Name + '}', [string]$p.Value) }
    return $t
}

$fmt = New-Object System.Speech.AudioFormat.SpeechAudioFormatInfo(22050, [System.Speech.AudioFormat.AudioBitsPerSample]::Sixteen, [System.Speech.AudioFormat.AudioChannel]::Mono)
$syn = New-Object System.Speech.Synthesis.SpeechSynthesizer
$syn.SelectVoice("Microsoft Hortense Desktop")
$syn.Rate = 0
$syn.Volume = 100

$txt = @()
$i = 0
foreach ($n in $j.narration) {
    $i++
    $wav = Join-Path $OutDir ("narration_diapo{0}.wav" -f $i)
    $syn.SetOutputToWaveFile($wav, $fmt)
    $syn.Speak((Remplir $n.tts))
    $syn.SetOutputToNull()
    $sec = ((Get-Item $wav).Length - 44) / (22050 * 2)
    $mots = ((Remplir $n.notes) -split '\s+').Count
    "{0} : {1,5:N1} s | {2} mots | {3:N0} Ko" -f (Split-Path $wav -Leaf), $sec, $mots, ((Get-Item $wav).Length / 1KB)
    $txt += $n.titre
    $txt += (Remplir $n.notes)
    $txt += ""
}
$syn.Dispose()
[IO.File]::WriteAllLines((Join-Path $OutDir "texte_de_la_narration.txt"), $txt, (New-Object Text.UTF8Encoding($true)))
