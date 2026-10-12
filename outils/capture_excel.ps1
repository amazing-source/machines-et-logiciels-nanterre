# Capture d'ecran d'une feuille Excel : capture de la SEULE fenetre Excel
# (API PrintWindow), jamais de l'ecran entier — les autres fenetres de
# l'utilisateur ne sont donc pas photographiees.
# Usage : capture_excel.ps1 -Path classeur.xlsx -Sheet "Données" -Out img.png
#         [-Zoom 130] [-TopLeft A1] [-Select K2] [-WidthPx 1800] [-HeightPx 900]
#         [-Crop grid|window]
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$Sheet,
    [Parameter(Mandatory = $true)][string]$Out,
    [int]$Zoom = 100,
    [string]$TopLeft = "A1",
    [string]$Select = "",
    [int]$WidthPx = 1800,
    [int]$HeightPx = 900,
    [string]$Crop = "grid"
)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Cap {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint f);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int hgt, bool repaint);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
}
"@
[void][Cap]::SetProcessDPIAware()

$x = New-Object -ComObject Excel.Application
$x.DisplayAlerts = $false
try {
    $wb = $x.Workbooks.Open($Path, 0, $true)
    $x.Visible = $true
    $x.WindowState = -4143
    $h = [IntPtr]$x.Hwnd
    $dpi = [Cap]::GetDpiForWindow($h)
    $k = $dpi / 96.0
    [void][Cap]::MoveWindow($h, 8, 8, [int]($WidthPx * $k), [int]($HeightPx * $k), $true)
    # le volet « Champs de tableau croise dynamique » mangerait la largeur
    try { $wb.ShowPivotTableFieldList = $false } catch { }
    $ws = $wb.Worksheets.Item($Sheet)
    $ws.Activate()
    $win = $x.ActiveWindow
    $win.Zoom = $Zoom
    $x.Goto($ws.Range($TopLeft), $true)
    if ($Select) { $ws.Range($Select).Select() | Out-Null } else { $ws.Range($TopLeft).Select() | Out-Null }
    try { $wb.ShowPivotTableFieldList = $false } catch { }
    $x.ScreenUpdating = $true
    Start-Sleep -Milliseconds 1500

    $r = New-Object Cap+RECT
    [void][Cap]::GetWindowRect($h, [ref]$r)
    $w = $r.R - $r.L; $hg = $r.B - $r.T
    $bmp = New-Object System.Drawing.Bitmap($w, $hg)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $dc = $g.GetHdc()
    $ok = [Cap]::PrintWindow($h, $dc, 2)
    $g.ReleaseHdc($dc)
    $g.Dispose()

    # recadrage : de la barre de formule jusqu'au bas de la fenetre
    $top = 0
    if ($Crop -eq "grid") {
        # Panes(1) = volet superieur gauche : son point 0 est le haut de la
        # ligne 1, meme quand des volets sont figes. Au-dessus : la bande des
        # lettres de colonnes (qui suit le zoom) puis la barre de formule.
        $p1 = $win.Panes(1)
        $gy = $p1.PointsToScreenPixelsY($p1.VisibleRange.Top) - $r.T
        $top = [Math]::Max(0, [int]($gy - (21 * $Zoom / 100 + 36) * $k))
    }
    $bord = [int](7 * $k)
    $rect = New-Object System.Drawing.Rectangle($bord, $top, ($w - 2 * $bord), ($hg - $top - $bord))
    $img = $bmp.Clone($rect, $bmp.PixelFormat)
    $img.Save($Out, [System.Drawing.Imaging.ImageFormat]::Png)
    "capture: ok=$ok dpi=$dpi fenetre=${w}x${hg} image=$($img.Width)x$($img.Height) visible=$($win.VisibleRange.Address($false,$false)) -> $Out"
    $img.Dispose(); $bmp.Dispose()
}
finally {
    if ($wb) { $wb.Close($false) }
    $x.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
}
