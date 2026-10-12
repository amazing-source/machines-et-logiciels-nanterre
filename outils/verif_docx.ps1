# Verification d'un document Word : il doit s'ouvrir dans Word, on releve
# son nombre de pages et, si demande, on en tire un PDF de controle.
# Usage : verif_docx.ps1 -Path doc.docx [-Pdf sortie.pdf]
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$Pdf = ""
)
$ErrorActionPreference = "Stop"
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $d = $w.Documents.Open($Path, $false, $true)
    "{0} : {1} page(s), {2} mots, {3} paragraphes" -f (Split-Path $Path -Leaf), $d.ComputeStatistics(2), $d.ComputeStatistics(0), $d.Paragraphs.Count
    if ($Pdf) {
        # SaveAs2 format 17 (PDF) ; ExportAsFixedFormat se fige sur cette machine
        $d.SaveAs2($Pdf, 17)
        "pdf de controle : $Pdf"
    }
    $d.Close($false)
}
finally {
    $w.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
}
