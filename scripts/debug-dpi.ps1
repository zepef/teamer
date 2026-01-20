# Debug DPI and screen info
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$screen = [System.Windows.Forms.Screen]::PrimaryScreen
Write-Host "Primary Screen Info:" -ForegroundColor Cyan
Write-Host "  Bounds: $($screen.Bounds.X),$($screen.Bounds.Y) - $($screen.Bounds.Width)x$($screen.Bounds.Height)"
Write-Host "  WorkingArea: $($screen.WorkingArea.X),$($screen.WorkingArea.Y) - $($screen.WorkingArea.Width)x$($screen.WorkingArea.Height)"

# Get DPI
Add-Type @"
using System;
using System.Runtime.InteropServices;

public class DpiHelper {
    [DllImport("user32.dll")]
    public static extern int GetDpiForSystem();

    [DllImport("shcore.dll")]
    public static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

    [DllImport("user32.dll")]
    public static extern IntPtr MonitorFromPoint(System.Drawing.Point pt, uint dwFlags);
}
"@ -ErrorAction SilentlyContinue

try {
    $systemDpi = [DpiHelper]::GetDpiForSystem()
    $scaleFactor = $systemDpi / 96.0
    Write-Host ""
    Write-Host "DPI Info:" -ForegroundColor Cyan
    Write-Host "  System DPI: $systemDpi"
    Write-Host "  Scale Factor: $scaleFactor (${scaleFactor}x)"
    Write-Host "  100% = 96 DPI, 125% = 120 DPI, 150% = 144 DPI"
}
catch {
    Write-Host "Could not get DPI info: $_" -ForegroundColor Yellow
}

# Check actual window positions
Write-Host ""
Write-Host "Current Window Positions:" -ForegroundColor Cyan

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class WinRect {
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("dwmapi.dll")]
    public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;
}
"@ -ErrorAction SilentlyContinue

$processes = @("WindowsTerminal", "EXCEL", "Notepad")
foreach ($procName in $processes) {
    $proc = Get-Process -Name $procName -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | Select-Object -First 1
    if ($proc) {
        $windowRect = New-Object WinRect+RECT
        $extendedRect = New-Object WinRect+RECT
        [WinRect]::GetWindowRect($proc.MainWindowHandle, [ref]$windowRect) | Out-Null
        [WinRect]::DwmGetWindowAttribute($proc.MainWindowHandle, [WinRect]::DWMWA_EXTENDED_FRAME_BOUNDS, [ref]$extendedRect, [System.Runtime.InteropServices.Marshal]::SizeOf([type][WinRect+RECT])) | Out-Null

        $borderLeft = $extendedRect.Left - $windowRect.Left
        $borderTop = $extendedRect.Top - $windowRect.Top
        $borderRight = $windowRect.Right - $extendedRect.Right
        $borderBottom = $windowRect.Bottom - $extendedRect.Bottom

        Write-Host "  $procName :" -ForegroundColor Yellow
        Write-Host "    Window Rect: Left=$($windowRect.Left) Top=$($windowRect.Top) Right=$($windowRect.Right) Bottom=$($windowRect.Bottom)"
        Write-Host "    Visible (Extended): Left=$($extendedRect.Left) Top=$($extendedRect.Top) Right=$($extendedRect.Right) Bottom=$($extendedRect.Bottom)"
        Write-Host "    Invisible Borders: Left=$borderLeft Top=$borderTop Right=$borderRight Bottom=$borderBottom"
    }
}
