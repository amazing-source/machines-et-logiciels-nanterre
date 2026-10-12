# Seance 5, projet : chaine complete de construction du classeur.
#   1. projet5_classeur.py  : donnees, formules, synthese (openpyxl)
#   2. projet5_excel.ps1    : tri, filtres, sous-totaux, TCD, graphique (Excel)
#   3. projet5_verif.ps1    : controle des valeurs et des erreurs
# Usage : projet5_build.ps1 -Work <dossier de travail> -Dst <classeur final> [-Identite identite.json]
param(
    [Parameter(Mandatory = $true)][string]$Work,
    [Parameter(Mandatory = $true)][string]$Dst,
    [string]$Identite = ""
)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
if (-not $Identite) { $Identite = "$here\identite.json" }
New-Item -ItemType Directory -Force $Work | Out-Null
$env:PYTHONIOENCODING = "utf-8"

python -I "$here\projet5_classeur.py" "$root\seance-5\donnees\fr-esr-parcoursup_2025_ile-de-france_brut.csv" "$Work\etape1.xlsx" $Identite
if ($LASTEXITCODE -ne 0) { throw "projet5_classeur.py a echoue" }

& "$here\bom.ps1" "$here\projet5_excel.ps1" "$here\projet5_verif.ps1" "$here\capture_excel.ps1"
& "$here\projet5_excel.ps1" -Src "$Work\etape1.xlsx" -Dst $Dst
& "$here\projet5_verif.ps1" -Path $Dst | Out-File "$Work\verif.txt" -Encoding utf8
Get-Content "$Work\verif.txt" -Encoding UTF8 | Select-Object -First 8
"taille : " + (Get-Item $Dst).Length + " octets"
