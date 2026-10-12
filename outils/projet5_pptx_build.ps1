# Seance 5, projet : chaine de construction de la presentation.
#   1. projet5_voix.ps1  : fichiers audio de narration (sauf -SansVoix)
#   2. projet5_pptx.ps1  : presentation PowerPoint
#   3. proprietes.py     : auteur et titre du document
# Usage : projet5_pptx_build.ps1 -Xlsx <classeur> -Out <presentation.pptx>
#         [-Identite identite.json] [-Captures dossier] [-Narration dossier] [-SansVoix]
param(
    [Parameter(Mandatory = $true)][string]$Xlsx,
    [Parameter(Mandatory = $true)][string]$Out,
    [string]$Identite = "",
    [string]$Captures = "",
    [string]$Narration = "",
    [string]$Log = "$env:TEMP\projet5_pptx.log",
    [switch]$SansVoix
)
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
if (-not $Identite) { $Identite = "$here\identite.json" }
if (-not $Captures) { $Captures = "$root\seance-5\captures" }
if (-not $Narration) { $Narration = "$root\seance-5\narration" }
$ID = Get-Content -Raw -Encoding UTF8 $Identite | ConvertFrom-Json
& "$here\bom.ps1" "$here\projet5_pptx.ps1" "$here\projet5_voix.ps1"
# l'audio depend du nom et de la filiere : il est regenere a chaque fois
# (sauf en version sans voix, qui ne s'en sert que pour les durees)
if (-not $SansVoix) {
    & "$here\projet5_voix.ps1" -Textes "$here\projet5_textes.json" -Identite $Identite -OutDir $Narration
}
try {
    & "$here\projet5_pptx.ps1" -Xlsx $Xlsx -Captures $Captures -Narration $Narration -Textes "$here\projet5_textes.json" -Identite $Identite -Out $Out -SansVoix:$SansVoix *> $Log
}
catch { }
Get-Content $Log -Encoding Unicode | Select-Object -Last 9
if (Test-Path $Out) {
    $env:PYTHONIOENCODING = "utf-8"
    python -I "$here\proprietes.py" $Out $ID.nom_prenom "Parcoursup 2025 en Île-de-France : où la concurrence est-elle la plus forte ?" ("Projet multimédia — " + $ID.numero)
    "taille : " + (Get-Item $Out).Length + " octets"
}
