# Seance 5, activite 1, exercice 6 : exports de la presentation dans un
# sous-dossier — diaporama (.ppsx), PDF, puis videos HD et Full HD (.mp4).
# Usage : seance5_exports.ps1 -Path SEANCE5_MULTIMEDIA.pptx -OutDir dossier [-SansVideo]
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [switch]$SansVideo
)
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force $OutDir | Out-Null
$base = [IO.Path]::GetFileNameWithoutExtension($Path)
$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($Path, -1, 0, 0)
try {
    $ppsx = Join-Path $OutDir "$base.ppsx"
    $pres.SaveCopyAs($ppsx, 28)
    "ppsx : {0:N0} octets" -f (Get-Item $ppsx).Length

    $pdf = Join-Path $OutDir "$base.pdf"
    $pres.SaveCopyAs($pdf, 32)
    "pdf  : {0:N0} octets" -f (Get-Item $pdf).Length

    if (-not $SansVideo) {
        foreach ($v in @(@(720, "HD_720p"), @(1080, "FullHD_1080p"))) {
            $mp4 = Join-Path $OutDir ("{0}_{1}.mp4" -f $base, $v[1])
            if (Test-Path $mp4) { Remove-Item $mp4 -Force }
            $t0 = Get-Date
            $pres.CreateVideo($mp4, $true, 5, $v[0], 30, 85)
            do { Start-Sleep -Seconds 3; $st = $pres.CreateVideoStatus } while ($st -eq 1 -or $st -eq 2)
            "mp4 {0} : statut {1} | {2:N0} octets | {3:N0} s" -f $v[1], $st, $(if (Test-Path $mp4) { (Get-Item $mp4).Length } else { 0 }), ((Get-Date) - $t0).TotalSeconds
        }
    }
}
finally {
    $pres.Close()
    if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
}
"EXPORTS TERMINES"
