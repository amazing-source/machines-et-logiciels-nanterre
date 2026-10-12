# Re-encode un script en UTF-8 AVEC BOM (PowerShell 5.1 lit sinon le fichier
# en ANSI et casse tous les accents en silence). Idempotent.
# Usage : bom.ps1 chemin1 [chemin2 ...]
foreach ($p in $args) {
    $full = (Resolve-Path $p).Path
    $t = [IO.File]::ReadAllText($full, (New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllText($full, $t, (New-Object Text.UTF8Encoding($true)))
}
