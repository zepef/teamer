<#
.SYNOPSIS
    Captures current system configuration (monitors, virtual desktops, windows) and saves to JSON.

.DESCRIPTION
    This script discovers the current state of:
    - Monitor configuration (resolution, position, DPI, refresh rate)
    - Virtual desktops (count, names, current desktop)
    - Windows per desktop (position, size, state, pinned status)

    Output is saved to a timestamped JSON config file.

.PARAMETER OutputPath
    Directory to save the config file. Defaults to E:\Projects\teamer\config\

.EXAMPLE
    .\Get-SystemConfig.ps1

.EXAMPLE
    .\Get-SystemConfig.ps1 -OutputPath "C:\MyConfigs"

.NOTES
    Requires: Windows 10 2004+ or Windows 11
    Auto-installs PSVirtualDesktop module if not present
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath = "E:\Projects\teamer\config"
)

#region Win32 API Definitions
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;

public class Win32Window {
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsZoomed(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    public const int GWL_STYLE = -16;
    public const int GWL_EXSTYLE = -20;
    public const uint WS_VISIBLE = 0x10000000;
    public const uint WS_CAPTION = 0x00C00000;
    public const uint WS_EX_TOOLWINDOW = 0x00000080;
    public const uint WS_EX_APPWINDOW = 0x00040000;

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public static string GetWindowTitle(IntPtr hWnd) {
        int length = GetWindowTextLength(hWnd);
        if (length == 0) return "";
        StringBuilder sb = new StringBuilder(length + 1);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    public static string GetWindowClassName(IntPtr hWnd) {
        StringBuilder sb = new StringBuilder(256);
        GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    public static List<IntPtr> GetAllWindows() {
        List<IntPtr> windows = new List<IntPtr>();
        EnumWindows((hWnd, lParam) => {
            windows.Add(hWnd);
            return true;
        }, IntPtr.Zero);
        return windows;
    }

    public static bool IsAltTabWindow(IntPtr hWnd) {
        if (!IsWindowVisible(hWnd)) return false;

        int style = GetWindowLong(hWnd, GWL_STYLE);
        int exStyle = GetWindowLong(hWnd, GWL_EXSTYLE);

        // Skip tool windows unless they have app window style
        if ((exStyle & WS_EX_TOOLWINDOW) != 0 && (exStyle & WS_EX_APPWINDOW) == 0)
            return false;

        // Must have a caption (title bar)
        if ((style & WS_CAPTION) != WS_CAPTION)
            return false;

        // Must have a title
        if (GetWindowTextLength(hWnd) == 0)
            return false;

        return true;
    }
}
"@
#endregion

#region Helper Functions

function Ensure-Dependencies {
    <#
    .SYNOPSIS
        Ensures PSVirtualDesktop module is installed
    #>
    Write-Host "Checking dependencies..." -ForegroundColor Cyan

    # Check if module is already loaded
    $loadedModule = Get-Module -Name VirtualDesktop -ErrorAction SilentlyContinue

    if ($loadedModule) {
        Write-Host "PSVirtualDesktop module already loaded." -ForegroundColor Green
        return
    }

    # Check if module is available
    $availableModule = Get-Module -ListAvailable -Name VirtualDesktop -ErrorAction SilentlyContinue

    if (-not $availableModule) {
        Write-Host "PSVirtualDesktop module not found. Installing..." -ForegroundColor Yellow
        try {
            # Install module
            Install-Module -Name VirtualDesktop -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop

            # Refresh module paths to pick up newly installed module (including OneDrive paths)
            $userPath = [System.Environment]::GetEnvironmentVariable("PSModulePath", "User")
            $machinePath = [System.Environment]::GetEnvironmentVariable("PSModulePath", "Machine")
            $oneDrivePath = "$env:USERPROFILE\OneDrive\Documents\WindowsPowerShell\Modules"

            $env:PSModulePath = @($userPath, $machinePath, $oneDrivePath) -join ";"

            Write-Host "PSVirtualDesktop installed successfully." -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install PSVirtualDesktop module: $_"
            throw
        }
    }

    # Try to import the module
    try {
        Import-Module VirtualDesktop -Force -ErrorAction Stop
        Write-Host "PSVirtualDesktop module loaded." -ForegroundColor Green
    }
    catch {
        # Try finding and importing by path (including OneDrive-synced locations)
        $modulePaths = @(
            "$env:USERPROFILE\Documents\PowerShell\Modules\VirtualDesktop",
            "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\VirtualDesktop",
            "$env:USERPROFILE\OneDrive\Documents\PowerShell\Modules\VirtualDesktop",
            "$env:USERPROFILE\OneDrive\Documents\WindowsPowerShell\Modules\VirtualDesktop",
            "$env:ProgramFiles\PowerShell\Modules\VirtualDesktop",
            "$env:ProgramFiles\WindowsPowerShell\Modules\VirtualDesktop"
        )

        $imported = $false
        foreach ($path in $modulePaths) {
            if (Test-Path $path) {
                try {
                    Import-Module $path -Force -ErrorAction Stop
                    Write-Host "PSVirtualDesktop module loaded from: $path" -ForegroundColor Green
                    $imported = $true
                    break
                }
                catch {
                    Write-Warning "Failed to import from $path"
                }
            }
        }

        if (-not $imported) {
            Write-Error "Could not load VirtualDesktop module. Please restart PowerShell and try again."
            throw "Module import failed"
        }
    }

    # Verify the module is working
    try {
        $null = Get-Command Get-DesktopCount -ErrorAction Stop
        Write-Host "PSVirtualDesktop commands verified." -ForegroundColor Green
    }
    catch {
        Write-Error "VirtualDesktop module loaded but commands not available. Please restart PowerShell."
        throw
    }
}

function Get-MonitorConfig {
    <#
    .SYNOPSIS
        Retrieves monitor configuration via CIM/WMI
    #>
    Write-Host "Querying monitor configuration..." -ForegroundColor Cyan

    $monitors = @()

    # Get video controller info for refresh rates
    $videoControllers = Get-CimInstance -ClassName Win32_VideoController

    # Get monitor info
    $monitorInfo = Get-CimInstance -ClassName Win32_DesktopMonitor

    # Get detailed display config using WMI
    $displayConfigs = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams -ErrorAction SilentlyContinue

    # Primary method: Use .NET Screen class for accurate info
    Add-Type -AssemblyName System.Windows.Forms
    $screens = [System.Windows.Forms.Screen]::AllScreens

    $index = 0
    foreach ($screen in $screens) {
        $monitor = [ordered]@{
            deviceName   = $screen.DeviceName
            isPrimary    = $screen.Primary
            resolution   = [ordered]@{
                width  = $screen.Bounds.Width
                height = $screen.Bounds.Height
            }
            position     = [ordered]@{
                x = $screen.Bounds.X
                y = $screen.Bounds.Y
            }
            workingArea  = [ordered]@{
                x      = $screen.WorkingArea.X
                y      = $screen.WorkingArea.Y
                width  = $screen.WorkingArea.Width
                height = $screen.WorkingArea.Height
            }
            bitsPerPixel = $screen.BitsPerPixel
            refreshRate  = 60  # Default, will try to get actual
            dpiScale     = 100 # Default, will try to get actual
        }

        # Try to get refresh rate from video controller
        if ($videoControllers) {
            $vc = $videoControllers | Select-Object -First 1
            if ($vc.CurrentRefreshRate) {
                $monitor.refreshRate = $vc.CurrentRefreshRate
            }
        }

        # Try to get DPI scale
        try {
            # Use registry for DPI info per monitor
            $dpiKey = "HKCU:\Control Panel\Desktop\WindowMetrics"
            if (Test-Path $dpiKey) {
                $appliedDpi = (Get-ItemProperty -Path $dpiKey -Name AppliedDPI -ErrorAction SilentlyContinue).AppliedDPI
                if ($appliedDpi) {
                    $monitor.dpiScale = [math]::Round(($appliedDpi / 96) * 100)
                }
            }
        }
        catch {
            # Keep default
        }

        $monitors += $monitor
        $index++
    }

    Write-Host "Found $($monitors.Count) monitor(s)." -ForegroundColor Green
    return $monitors
}

function Get-DesktopConfig {
    <#
    .SYNOPSIS
        Retrieves virtual desktop configuration using PSVirtualDesktop
    #>
    Write-Host "Querying virtual desktop configuration..." -ForegroundColor Cyan

    $desktopCount = Get-DesktopCount
    $currentDesktop = Get-CurrentDesktop
    $currentIndex = Get-DesktopIndex -Desktop $currentDesktop

    $desktops = @()
    for ($i = 0; $i -lt $desktopCount; $i++) {
        $desktop = Get-Desktop -Index $i
        $desktopName = ""

        try {
            $desktopName = Get-DesktopName -Desktop $desktop
        }
        catch {
            $desktopName = "Desktop $($i + 1)"
        }

        if ([string]::IsNullOrWhiteSpace($desktopName)) {
            $desktopName = "Desktop $($i + 1)"
        }

        $desktops += [ordered]@{
            index = $i
            name  = $desktopName
            id    = $desktop.ToString()
        }
    }

    $config = [ordered]@{
        count        = $desktopCount
        currentIndex = $currentIndex
        desktops     = $desktops
    }

    Write-Host "Found $desktopCount virtual desktop(s). Current: $currentIndex" -ForegroundColor Green
    return $config
}

function Get-WindowConfig {
    <#
    .SYNOPSIS
        Enumerates all visible windows with position, size, and state
    #>
    Write-Host "Enumerating windows..." -ForegroundColor Cyan

    $windows = @()
    $allHandles = [Win32Window]::GetAllWindows()
    $processCache = @{}

    foreach ($hwnd in $allHandles) {
        # Filter to only "real" windows (alt-tab worthy)
        if (-not [Win32Window]::IsAltTabWindow($hwnd)) {
            continue
        }

        $title = [Win32Window]::GetWindowTitle($hwnd)
        $className = [Win32Window]::GetWindowClassName($hwnd)

        # Skip certain system windows
        $skipClasses = @('Progman', 'WorkerW', 'Shell_TrayWnd', 'Shell_SecondaryTrayWnd', 'NotifyIconOverflowWindow')
        if ($className -in $skipClasses) {
            continue
        }

        # Get process info
        $processId = 0
        [Win32Window]::GetWindowThreadProcessId($hwnd, [ref]$processId) | Out-Null

        $processName = ""
        if ($processCache.ContainsKey($processId)) {
            $processName = $processCache[$processId]
        }
        else {
            try {
                $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
                $processName = $process.ProcessName
                $processCache[$processId] = $processName
            }
            catch {
                $processName = "Unknown"
            }
        }

        # Get window rect
        $rect = New-Object Win32Window+RECT
        [Win32Window]::GetWindowRect($hwnd, [ref]$rect) | Out-Null

        # Determine window state
        $state = "Normal"
        if ([Win32Window]::IsIconic($hwnd)) {
            $state = "Minimized"
        }
        elseif ([Win32Window]::IsZoomed($hwnd)) {
            $state = "Maximized"
        }

        # Get desktop assignment
        $desktopIndex = -1
        $isPinned = $false

        try {
            $desktop = Get-DesktopFromWindow -Hwnd $hwnd.ToInt64()
            if ($desktop) {
                $desktopIndex = Get-DesktopIndex -Desktop $desktop
            }
        }
        catch {
            # Window might not be on any desktop (pinned or system)
        }

        # Check if pinned
        try {
            $isPinned = Test-WindowPinned -Hwnd $hwnd.ToInt64()
        }
        catch {
            # Pinned check failed, assume not pinned
        }

        $window = [ordered]@{
            hwnd         = $hwnd.ToInt64()
            processName  = $processName
            processId    = $processId
            title        = $title
            className    = $className
            desktopIndex = $desktopIndex
            position     = [ordered]@{
                x = $rect.Left
                y = $rect.Top
            }
            size         = [ordered]@{
                width  = $rect.Right - $rect.Left
                height = $rect.Bottom - $rect.Top
            }
            state        = $state
            isPinned     = $isPinned
        }

        $windows += $window
    }

    Write-Host "Found $($windows.Count) window(s)." -ForegroundColor Green
    return $windows
}

function Get-PinnedItems {
    <#
    .SYNOPSIS
        Gets lists of pinned windows and applications
    #>
    Write-Host "Checking pinned items..." -ForegroundColor Cyan

    $pinnedWindows = @()
    $pinnedApplications = @()

    $allHandles = [Win32Window]::GetAllWindows()

    foreach ($hwnd in $allHandles) {
        if (-not [Win32Window]::IsAltTabWindow($hwnd)) {
            continue
        }

        $hwndInt = $hwnd.ToInt64()

        try {
            if (Test-WindowPinned -Hwnd $hwndInt) {
                $title = [Win32Window]::GetWindowTitle($hwnd)
                $pinnedWindows += [ordered]@{
                    hwnd  = $hwndInt
                    title = $title
                }
            }
        }
        catch {
            # Skip
        }

        try {
            if (Test-ApplicationPinned -Hwnd $hwndInt) {
                $title = [Win32Window]::GetWindowTitle($hwnd)
                $processId = 0
                [Win32Window]::GetWindowThreadProcessId($hwnd, [ref]$processId) | Out-Null

                $pinnedApplications += [ordered]@{
                    hwnd      = $hwndInt
                    title     = $title
                    processId = $processId
                }
            }
        }
        catch {
            # Skip
        }
    }

    Write-Host "Found $($pinnedWindows.Count) pinned window(s), $($pinnedApplications.Count) pinned application(s)." -ForegroundColor Green

    return @{
        windows      = $pinnedWindows
        applications = $pinnedApplications
    }
}

function Export-Config {
    <#
    .SYNOPSIS
        Combines all config data and exports to JSON
    #>
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,

        [Parameter(Mandatory)]
        [string]$OutputPath
    )

    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    # Generate timestamp
    $timestamp = Get-Date -Format "yyyy-MM-ddTHH-mm-ss"
    $filename = "system-config-$timestamp.json"
    $fullPath = Join-Path $OutputPath $filename

    # Add timestamp to config
    $Config.timestamp = (Get-Date -Format "o")

    # Convert to JSON and save
    $json = $Config | ConvertTo-Json -Depth 10
    $json | Out-File -FilePath $fullPath -Encoding UTF8

    Write-Host "`nConfiguration saved to: $fullPath" -ForegroundColor Green
    return $fullPath
}

#endregion

#region Main Execution

Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "  System Configuration Discovery Tool  " -ForegroundColor Magenta
Write-Host "========================================`n" -ForegroundColor Magenta

try {
    # Ensure dependencies
    Ensure-Dependencies

    Write-Host ""

    # Gather all configuration (wrap in @() to ensure arrays stay arrays)
    $monitors = @(Get-MonitorConfig)
    $virtualDesktops = Get-DesktopConfig
    $windows = @(Get-WindowConfig)
    $pinnedItems = Get-PinnedItems

    # Build final config object
    $config = [ordered]@{
        timestamp          = $null  # Will be set in Export-Config
        monitors           = $monitors
        virtualDesktops    = $virtualDesktops
        windows            = $windows
        pinnedWindows      = @($pinnedItems.windows)
        pinnedApplications = @($pinnedItems.applications)
    }

    # Export to file
    $outputFile = Export-Config -Config $config -OutputPath $OutputPath

    # Summary - use @().Count to ensure proper array counting
    $monitorCount = @($monitors).Count
    $windowCount = @($windows).Count
    $pinnedWinCount = @($pinnedItems.windows).Count
    $pinnedAppCount = @($pinnedItems.applications).Count

    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "            Summary                    " -ForegroundColor Magenta
    Write-Host "========================================" -ForegroundColor Magenta
    Write-Host "Monitors:           $monitorCount" -ForegroundColor White
    Write-Host "Virtual Desktops:   $($virtualDesktops.count)" -ForegroundColor White
    Write-Host "Windows:            $windowCount" -ForegroundColor White
    Write-Host "Pinned Windows:     $pinnedWinCount" -ForegroundColor White
    Write-Host "Pinned Apps:        $pinnedAppCount" -ForegroundColor White
    Write-Host "========================================`n" -ForegroundColor Magenta

    # Return the config for pipeline use
    return $config
}
catch {
    Write-Error "Failed to capture system configuration: $_"
    throw
}

#endregion
