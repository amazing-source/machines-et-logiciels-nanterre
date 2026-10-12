# Genere une feuille de reponses (seance 5 ou 6), verifie qu'elle s'ouvre
# dans Word et, si demande, en tire un PDF de controle.
# Usage : feuilles_build.ps1 -Seance 6 -Enonce <enonce.docx> -Out <sortie.docx>
#         [-Identite identite.json] [-PdfDir dossier]
param(
    [Parameter(Mandatory = $true)][int]$Seance,
    [Parameter(Mandatory = $true)][string]$Enonce,
    [Parameter(Mandatory = $true)][string]$Out,
    [string]$Identite = "",
    [string]$PdfDir = ""
)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Identite) { $Identite = "$here\identite.json" }
$env:PYTHONIOENCODING = "utf-8"
Push-Location $here
try { python "$here\feuille$Seance.py" $Enonce $Out $Identite } finally { Pop-Location }
if ($LASTEXITCODE -ne 0) { throw "feuille$Seance.py a echoue" }
& "$here\bom.ps1" "$here\verif_docx.ps1"
if ($PdfDir) {
    New-Item -ItemType Directory -Force $PdfDir | Out-Null
    & "$here\verif_docx.ps1" -Path $Out -Pdf (Join-Path $PdfDir "s$Seance.pdf")
}
else { & "$here\verif_docx.ps1" -Path $Out }
