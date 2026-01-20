<#
.SYNOPSIS
    Teamer - CRUD operations for Windows Virtual Desktop management.

.DESCRIPTION
    Provides Create, Read, Update, Delete operations for virtual desktops and windows.

    SAFETY: All desktops existing at module load time are PROTECTED.
    Only newly created desktops (after loading) can be modified/deleted.

.NOTES
    Requires: Windows 10 2004+ or Windows 11
    Requires: PSVirtualDesktop module (auto-installed if missing)
#>

#region Constants

$script:PROJECT_ROOT = "E:\Projects\teamer"

# Protected desktops by NAME (more reliable than index which can shift)
# These desktops cannot be removed or have windows closed by Teamer
$script:PROTECTED_DESKTOP_NAMES = @(
    "Main",
    "Code",
    "Desktop 1",
    "Desktop 2",
    "Desktop 3"
)

# Legacy: count-based protection (kept for backward compatibility)
$script:PROTECTED_DESKTOP_COUNT = 0  # Will be set during initialization

#endregion

#region Win32 API Definitions

# Only add type if not already loaded
if (-not ([System.Management.Automation.PSTypeName]'TeamerWin32').Type) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Generic;

public class TeamerWin32 {
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

    [DllImport("user32.dll")]
    public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    public const int GWL_STYLE = -16;
    public const int GWL_EXSTYLE = -20;
    public const uint WS_VISIBLE = 0x10000000;
    public const uint WS_CAPTION = 0x00C00000;
    public const uint WS_EX_TOOLWINDOW = 0x00000080;
    public const uint WS_EX_APPWINDOW = 0x00040000;
    public const uint MONITOR_DEFAULTTONEAREST = 2;
    public const uint WM_CLOSE = 0x0010;

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct MONITORINFOEX {
        public int cbSize;
        public RECT rcMonitor;
        public RECT rcWork;
        public uint dwFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string szDevice;
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

        if ((exStyle & WS_EX_TOOLWINDOW) != 0 && (exStyle & WS_EX_APPWINDOW) == 0)
            return false;

        if ((style & WS_CAPTION) != WS_CAPTION)
            return false;

        if (GetWindowTextLength(hWnd) == 0)
            return false;

        return true;
    }

    public static string GetMonitorDeviceName(IntPtr hWnd) {
        IntPtr hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST);
        MONITORINFOEX mi = new MONITORINFOEX();
        mi.cbSize = Marshal.SizeOf(typeof(MONITORINFOEX));
        if (GetMonitorInfo(hMonitor, ref mi)) {
            return mi.szDevice;
        }
        return "Unknown";
    }

    public static bool CloseWindow(IntPtr hWnd) {
        return PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
    }
}
"@
}

#endregion

#region Module Initialization

function Initialize-TeamerModule {
    <#
    .SYNOPSIS
        Ensures PSVirtualDesktop module is loaded
    #>
    $loadedModule = Get-Module -Name VirtualDesktop -ErrorAction SilentlyContinue
    if ($loadedModule) { return }

    $availableModule = Get-Module -ListAvailable -Name VirtualDesktop -ErrorAction SilentlyContinue
    if (-not $availableModule) {
        Write-Host "Installing PSVirtualDesktop module..." -ForegroundColor Yellow
        Install-Module -Name VirtualDesktop -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop

        $userPath = [System.Environment]::GetEnvironmentVariable("PSModulePath", "User")
        $machinePath = [System.Environment]::GetEnvironmentVariable("PSModulePath", "Machine")
        $oneDrivePath = "$env:USERPROFILE\OneDrive\Documents\WindowsPowerShell\Modules"
        $env:PSModulePath = @($userPath, $machinePath, $oneDrivePath) -join ";"
    }

    try {
        Import-Module VirtualDesktop -Force -ErrorAction Stop
    }
    catch {
        $modulePaths = @(
            "$env:USERPROFILE\Documents\PowerShell\Modules\VirtualDesktop",
            "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\VirtualDesktop",
            "$env:USERPROFILE\OneDrive\Documents\PowerShell\Modules\VirtualDesktop",
            "$env:USERPROFILE\OneDrive\Documents\WindowsPowerShell\Modules\VirtualDesktop"
        )

        foreach ($path in $modulePaths) {
            if (Test-Path $path) {
                try {
                    Import-Module $path -Force -ErrorAction Stop
                    return
                }
                catch { }
            }
        }
        throw "Could not load VirtualDesktop module"
    }
}

# Initialize on module load
Initialize-TeamerModule

# Capture current desktop count as protection boundary
$script:PROTECTED_DESKTOP_COUNT = Get-DesktopCount

#endregion

#region Safety Functions

function Test-DesktopProtected {
    <#
    .SYNOPSIS
        Checks if a desktop is protected by NAME or by index
    .PARAMETER Index
        The 0-based desktop index
    .PARAMETER Name
        The desktop name (optional, will be looked up if not provided)
    .OUTPUTS
        Boolean - $true if protected, $false if safe to modify
    #>
    param(
        [Parameter(Mandatory)]
        [int]$Index,

        [Parameter()]
        [string]$Name
    )

    # If name not provided, look it up
    if (-not $Name) {
        try {
            $desktop = Get-Desktop -Index $Index
            $Name = $desktop.Name
        }
        catch {
            $Name = ""
        }
    }

    # Check if name is in protected list
    if ($Name -and $script:PROTECTED_DESKTOP_NAMES -contains $Name) {
        return $true
    }

    # Legacy fallback: check by index count
    return $Index -lt $script:PROTECTED_DESKTOP_COUNT
}

function Assert-SafeDesktopOperation {
    <#
    .SYNOPSIS
        Throws an error if attempting to modify a protected desktop
    .PARAMETER Index
        The 0-based desktop index
    .PARAMETER Name
        The desktop name (optional)
    .PARAMETER Operation
        Description of the operation being attempted
    #>
    param(
        [Parameter(Mandatory)]
        [int]$Index,

        [Parameter()]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$Operation
    )

    # Get name if not provided
    if (-not $Name) {
        try {
            $desktop = Get-Desktop -Index $Index
            $Name = $desktop.Name
        }
        catch {
            $Name = "Desktop $($Index + 1)"
        }
    }

    if (Test-DesktopProtected -Index $Index -Name $Name) {
        throw "BLOCKED: Cannot $Operation on protected desktop '$Name'. Protected desktops: $($script:PROTECTED_DESKTOP_NAMES -join ', ')"
    }
}

function Get-WindowDesktopIndex {
    <#
    .SYNOPSIS
        Gets the desktop index for a window handle
    #>
    param([int64]$Hwnd)

    try {
        $desktop = Get-DesktopFromWindow -Hwnd $Hwnd
        if ($desktop) {
            return Get-DesktopIndex -Desktop $desktop
        }
    }
    catch { }
    return -1
}

#endregion

#region Shell Detection Helpers

function Get-ShellType {
    <#
    .SYNOPSIS
        Detects shell type from window title
    #>
    param([string]$Title)

    # WSL/Linux detection
    if ($Title -match '(?i)(ubuntu|debian|kali|fedora|opensuse|arch|manjaro|wsl)') {
        return @{ shell = 'wsl'; distro = $Matches[1]; label = "WSL ($($Matches[1]))" }
    }
    if ($Title -match '^[a-z][a-z0-9_-]*:\s*[~/]' -or $Title -match '^[a-z][a-z0-9_-]*@[a-z]') {
        return @{ shell = 'wsl'; distro = 'Linux'; label = 'WSL (Linux)' }
    }

    # PowerShell detection
    if ($Title -match '(?i)(powershell|pwsh|PS\s+[A-Z]:\\)') {
        if ($Title -match '(?i)pwsh|PowerShell\s*7|Preview') {
            return @{ shell = 'powershell-core'; label = 'PowerShell Core' }
        }
        return @{ shell = 'powershell'; label = 'PowerShell' }
    }

    # Git Bash
    if ($Title -match '(?i)(mingw|git\s*bash|bash.*git)') {
        return @{ shell = 'git-bash'; label = 'Git Bash' }
    }

    # CMD
    if ($Title -match '(?i)(command\s*prompt|cmd\.exe|^[A-Z]:\\.*>)') {
        return @{ shell = 'cmd'; label = 'Command Prompt' }
    }

    # Generic shell
    if ($Title -match '(?i)(bash|zsh|fish|sh\s*-)') {
        return @{ shell = 'shell'; label = 'Shell' }
    }

    return @{ shell = 'unknown'; label = 'Terminal' }
}

function Get-WindowClassification {
    <#
    .SYNOPSIS
        Classifies a window as terminal or application
    #>
    param(
        [string]$ProcessName,
        [string]$Title
    )

    $terminalHosts = @{
        'WindowsTerminal' = 'Windows Terminal'
        'ConEmu64'        = 'ConEmu'
        'ConEmuC64'       = 'ConEmu'
        'Hyper'           = 'Hyper'
        'Alacritty'       = 'Alacritty'
        'wezterm-gui'     = 'WezTerm'
    }

    $directShells = @{
        'powershell' = @{ shell = 'powershell'; label = 'PowerShell' }
        'pwsh'       = @{ shell = 'powershell-core'; label = 'PowerShell Core' }
        'cmd'        = @{ shell = 'cmd'; label = 'Command Prompt' }
        'wsl'        = @{ shell = 'wsl'; label = 'WSL' }
        'wslhost'    = @{ shell = 'wsl'; label = 'WSL' }
        'mintty'     = @{ shell = 'git-bash'; label = 'Git Bash' }
    }

    if ($terminalHosts.ContainsKey($ProcessName)) {
        $shellInfo = Get-ShellType -Title $Title
        return @{
            type        = 'terminal'
            terminalApp = $terminalHosts[$ProcessName]
            shell       = $shellInfo.shell
            shellLabel  = $shellInfo.label
            distro      = $shellInfo.distro
        }
    }

    if ($directShells.ContainsKey($ProcessName)) {
        $info = $directShells[$ProcessName]
        return @{
            type        = 'terminal'
            terminalApp = $null
            shell       = $info.shell
            shellLabel  = $info.label
        }
    }

    return @{ type = 'application' }
}

#endregion

#region CREATE Functions

function New-TeamerDesktop {
    <#
    .SYNOPSIS
        Creates a new virtual desktop
    .DESCRIPTION
        Creates a new virtual desktop. New desktops are always created at index 3+,
        so they are safe to manage.
    .PARAMETER Name
        Optional name for the new desktop
    .EXAMPLE
        New-TeamerDesktop -Name "Project-Alpha"
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Name
    )

    try {
        $newDesktop = New-Desktop
        $newIndex = Get-DesktopIndex -Desktop $newDesktop

        if ($Name) {
            Set-DesktopName -Desktop $newDesktop -Name $Name
        }

        $displayName = if ($Name) { $Name } else { "Desktop $($newIndex + 1)" }

        Write-Host "Created new desktop: $displayName (index $newIndex)" -ForegroundColor Green

        return @{
            Success = $true
            Index   = $newIndex
            Name    = $displayName
            Desktop = $newDesktop
        }
    }
    catch {
        Write-Error "Failed to create desktop: $_"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function New-TeamerTerminal {
    <#
    .SYNOPSIS
        Launches a new terminal on a specified desktop
    .DESCRIPTION
        Opens a terminal with the specified shell type on the target desktop.
        Only works on Desktop 3+ (index 2+).
    .PARAMETER DesktopIndex
        The 0-based index of the target desktop (must be 2+)
    .PARAMETER Shell
        Shell type: ps (PowerShell), wsl (WSL), cmd (Command Prompt)
    .EXAMPLE
        New-TeamerTerminal -DesktopIndex 2 -Shell wsl
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$DesktopIndex,

        [Parameter()]
        [ValidateSet('ps', 'wsl', 'cmd', 'pwsh')]
        [string]$Shell = 'ps'
    )

    # Safety check
    Assert-SafeDesktopOperation -Index $DesktopIndex -Operation "launch terminal"

    # Get target desktop
    $targetDesktop = Get-Desktop -Index $DesktopIndex
    if (-not $targetDesktop) {
        throw "Desktop index $DesktopIndex does not exist"
    }

    # Switch to target desktop first
    $currentDesktop = Get-CurrentDesktop
    $currentIndex = Get-DesktopIndex -Desktop $currentDesktop

    if ($currentIndex -ne $DesktopIndex) {
        Switch-Desktop -Desktop $targetDesktop
        Start-Sleep -Milliseconds 300
    }

    # Launch terminal
    $process = $null
    switch ($Shell) {
        'ps' {
            $process = Start-Process -FilePath "wt.exe" -ArgumentList "-p", "Windows PowerShell" -PassThru
        }
        'pwsh' {
            $process = Start-Process -FilePath "wt.exe" -ArgumentList "-p", "PowerShell" -PassThru
        }
        'wsl' {
            $process = Start-Process -FilePath "wt.exe" -ArgumentList "-p", "Ubuntu" -PassThru
        }
        'cmd' {
            $process = Start-Process -FilePath "wt.exe" -ArgumentList "-p", "Command Prompt" -PassThru
        }
    }

    Start-Sleep -Milliseconds 500

    Write-Host "Launched $Shell terminal on Desktop $($DesktopIndex + 1)" -ForegroundColor Green

    return @{
        Success      = $true
        DesktopIndex = $DesktopIndex
        Shell        = $Shell
        ProcessId    = $process.Id
    }
}

#endregion

#region READ Functions

function Get-TeamerTree {
    <#
    .SYNOPSIS
        Builds and returns the system tree structure
    .DESCRIPTION
        Returns a hierarchical tree: Screen -> Desktop -> Window
        with all window and shell information
    .OUTPUTS
        Array of screen objects containing desktop and window hierarchies
    #>
    [CmdletBinding()]
    param()

    # Get screens
    Add-Type -AssemblyName System.Windows.Forms
    $screens = [System.Windows.Forms.Screen]::AllScreens
    $videoControllers = Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue

    # Get desktops
    $desktopCount = Get-DesktopCount
    $currentDesktop = Get-CurrentDesktop
    $currentIndex = Get-DesktopIndex -Desktop $currentDesktop

    $desktops = @()
    for ($i = 0; $i -lt $desktopCount; $i++) {
        $desktop = Get-Desktop -Index $i
        $desktopName = ""
        try { $desktopName = Get-DesktopName -Desktop $desktop } catch { }
        if ([string]::IsNullOrWhiteSpace($desktopName)) { $desktopName = "Desktop $($i + 1)" }

        $desktops += @{
            index      = $i
            name       = $desktopName
            isCurrent  = ($i -eq $currentIndex)
            isProtected = (Test-DesktopProtected -Index $i)
        }
    }

    # Get windows
    $windows = @()
    $allHandles = [TeamerWin32]::GetAllWindows()
    $processCache = @{}
    $skipClasses = @('Progman', 'WorkerW', 'Shell_TrayWnd', 'Shell_SecondaryTrayWnd', 'NotifyIconOverflowWindow')

    foreach ($hwnd in $allHandles) {
        if (-not [TeamerWin32]::IsAltTabWindow($hwnd)) { continue }

        $className = [TeamerWin32]::GetWindowClassName($hwnd)
        if ($className -in $skipClasses) { continue }

        $title = [TeamerWin32]::GetWindowTitle($hwnd)

        $processId = 0
        [TeamerWin32]::GetWindowThreadProcessId($hwnd, [ref]$processId) | Out-Null

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
            catch { $processName = "Unknown" }
        }

        $state = "Normal"
        if ([TeamerWin32]::IsIconic($hwnd)) { $state = "Minimized" }
        elseif ([TeamerWin32]::IsZoomed($hwnd)) { $state = "Maximized" }

        $monitorDevice = [TeamerWin32]::GetMonitorDeviceName($hwnd)
        $desktopIndex = Get-WindowDesktopIndex -Hwnd $hwnd.ToInt64()

        $isPinned = $false
        try { $isPinned = Test-WindowPinned -Hwnd $hwnd.ToInt64() } catch { }

        $classification = Get-WindowClassification -ProcessName $processName -Title $title

        $window = @{
            hwnd          = $hwnd.ToInt64()
            processName   = $processName
            processId     = $processId
            title         = $title
            monitorDevice = $monitorDevice
            desktopIndex  = $desktopIndex
            state         = $state
            isPinned      = $isPinned
            type          = $classification.type
        }

        if ($classification.type -eq 'terminal') {
            $window.terminalApp = $classification.terminalApp
            $window.shell = $classification.shell
            $window.shellLabel = $classification.shellLabel
            if ($classification.distro) { $window.distro = $classification.distro }
        }

        $windows += $window
    }

    # Build tree
    $tree = [System.Collections.ArrayList]@()
    $screenIndex = 1

    foreach ($screen in $screens) {
        $screenNode = [ordered]@{
            number      = $screenIndex
            deviceName  = $screen.DeviceName
            isPrimary   = $screen.Primary
            resolution  = @{ width = $screen.Bounds.Width; height = $screen.Bounds.Height }
            refreshRate = if ($videoControllers) { $videoControllers[0].CurrentRefreshRate } else { 60 }
            desktops    = [System.Collections.ArrayList]@()
        }

        foreach ($desktop in $desktops) {
            $desktopWindows = @($windows | Where-Object {
                ($_.desktopIndex -eq $desktop.index -or $_.isPinned) -and
                $_.monitorDevice -eq $screen.DeviceName
            })

            $desktopNode = [ordered]@{
                index       = $desktop.index
                name        = $desktop.name
                isCurrent   = $desktop.isCurrent
                isProtected = $desktop.isProtected
                windows     = [System.Collections.ArrayList]@()
            }

            foreach ($win in $desktopWindows) {
                [void]$desktopNode.windows.Add($win)
            }

            [void]$screenNode.desktops.Add($desktopNode)
        }

        [void]$tree.Add($screenNode)
        $screenIndex++
    }

    # Return as proper array (comma prefix forces array even with single item)
    return ,$tree.ToArray()
}

function Show-TeamerTree {
    <#
    .SYNOPSIS
        Displays the system tree in a formatted output
    .PARAMETER Filter
        Filter: 'terminals' for only terminals, 'apps' for only applications
    .PARAMETER DesktopIndex
        Show only a specific desktop index (0-based)
    .PARAMETER Manageable
        Show only manageable desktops (index 2+)
    .EXAMPLE
        Show-TeamerTree
    .EXAMPLE
        Show-TeamerTree -Filter terminals
    .EXAMPLE
        Show-TeamerTree -Manageable
    .EXAMPLE
        Show-TeamerTree -DesktopIndex 0
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('terminals', 'apps', 'all')]
        [string]$Filter = 'all',

        [Parameter()]
        [int]$DesktopIndex = -1,

        [Parameter()]
        [switch]$Manageable
    )

    $tree = @(Get-TeamerTree)

    Write-Host ""
    Write-Host ("=" * 70) -ForegroundColor Magenta
    Write-Host "  TEAMER - System Tree" -ForegroundColor Magenta
    Write-Host ("=" * 70) -ForegroundColor Magenta

    foreach ($screen in @($tree)) {
        $primaryMark = if ($screen.isPrimary) { " (Primary)" } else { "" }
        Write-Host "`n[Screen $($screen.number)]$primaryMark $($screen.resolution.width)x$($screen.resolution.height) @ $($screen.refreshRate)Hz" -ForegroundColor Yellow

        foreach ($dsk in @($screen.desktops)) {
            # Filter by desktop index
            if ($DesktopIndex -ge 0 -and $dsk.index -ne $DesktopIndex) { continue }

            # Filter manageable only
            if ($Manageable -and $dsk.isProtected) { continue }

            $protectedMark = if ($dsk.isProtected) { " *PROTECTED*" } else { "" }
            $currentMark = if ($dsk.isCurrent) { " *ACTIVE*" } else { "" }

            $desktopColor = if ($dsk.isProtected) { "DarkCyan" } else { "Cyan" }

            Write-Host "`n  [Desktop $($dsk.index + 1)] `"$($dsk.name)`"$protectedMark$currentMark" -ForegroundColor $desktopColor

            $filteredWindows = @($dsk.windows)
            if ($Filter -eq 'terminals') {
                $filteredWindows = @($dsk.windows | Where-Object { $_.type -eq 'terminal' })
            }
            elseif ($Filter -eq 'apps') {
                $filteredWindows = @($dsk.windows | Where-Object { $_.type -eq 'application' })
            }

            if ($filteredWindows.Count -eq 0) {
                Write-Host "    (no matching windows)" -ForegroundColor DarkGray
                continue
            }

            foreach ($win in @($filteredWindows)) {
                $pinnedMark = if ($win.isPinned) { " [+]" } else { "" }
                $stateIcon = switch ($win.state) {
                    "Minimized" { " [-]" }
                    "Maximized" { " [+]" }
                    default { "" }
                }

                $displayTitle = if ($win.title.Length -gt 50) {
                    $win.title.Substring(0, 47) + "..."
                } else { $win.title }

                if ($win.type -eq 'terminal') {
                    $shellIcon = switch ($win.shell) {
                        'powershell'      { "[PS]" }
                        'powershell-core' { "[PS7]" }
                        'wsl'             { "[WSL]" }
                        'cmd'             { "[CMD]" }
                        'git-bash'        { "[GIT]" }
                        default           { "[T]" }
                    }

                    $hostName = if ($win.terminalApp) { $win.terminalApp } else { $win.processName }
                    $shellName = if ($win.shellLabel) { $win.shellLabel } else { "Terminal" }

                    Write-Host "    $shellIcon " -ForegroundColor Green -NoNewline
                    Write-Host "$hostName" -ForegroundColor White -NoNewline
                    Write-Host " > " -ForegroundColor DarkGray -NoNewline
                    Write-Host "$shellName" -ForegroundColor Cyan -NoNewline
                    Write-Host "$stateIcon$pinnedMark" -ForegroundColor White
                    Write-Host "        $displayTitle" -ForegroundColor Gray
                }
                else {
                    Write-Host "    [A] " -ForegroundColor White -NoNewline
                    Write-Host "$($win.processName)$stateIcon$pinnedMark" -ForegroundColor White
                    Write-Host "        $displayTitle" -ForegroundColor Gray
                }
            }
        }
    }

    Write-Host ""
    Write-Host ("=" * 70) -ForegroundColor Magenta
    Write-Host "Legend: [PS]=PowerShell [PS7]=PS Core [WSL]=WSL [CMD]=Cmd [GIT]=Git Bash [A]=App" -ForegroundColor DarkGray
    Write-Host "        *PROTECTED*=Cannot modify  [-]=Min [+]=Max" -ForegroundColor DarkGray
    Write-Host ("=" * 70) -ForegroundColor Magenta
}

#endregion

#region UPDATE Functions

function Rename-TeamerDesktop {
    <#
    .SYNOPSIS
        Renames a virtual desktop
    .DESCRIPTION
        Renames a desktop. Only works on Desktop 3+ (index 2+).
    .PARAMETER Index
        The 0-based desktop index (must be 2+)
    .PARAMETER Name
        The new name for the desktop
    .EXAMPLE
        Rename-TeamerDesktop -Index 2 -Name "Project-Beta"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$Index,

        [Parameter(Mandatory)]
        [string]$Name
    )

    Assert-SafeDesktopOperation -Index $Index -Operation "rename"

    try {
        $desktop = Get-Desktop -Index $Index
        if (-not $desktop) {
            throw "Desktop index $Index does not exist"
        }

        Set-DesktopName -Desktop $desktop -Name $Name
        Write-Host "Renamed Desktop $($Index + 1) to `"$Name`"" -ForegroundColor Green

        return @{ Success = $true; Index = $Index; Name = $Name }
    }
    catch {
        Write-Error "Failed to rename desktop: $_"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Move-TeamerWindow {
    <#
    .SYNOPSIS
        Moves a window to a different desktop
    .DESCRIPTION
        Moves a window by handle to a target desktop.
        Target desktop must be index 2+ (Desktop 3+).
    .PARAMETER Hwnd
        The window handle
    .PARAMETER ToDesktop
        The target desktop index (must be 2+)
    .EXAMPLE
        Move-TeamerWindow -Hwnd 12345 -ToDesktop 2
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int64]$Hwnd,

        [Parameter(Mandatory)]
        [int]$ToDesktop
    )

    Assert-SafeDesktopOperation -Index $ToDesktop -Operation "move window to"

    try {
        $targetDesktop = Get-Desktop -Index $ToDesktop
        if (-not $targetDesktop) {
            throw "Target desktop index $ToDesktop does not exist"
        }

        Move-Window -Desktop $targetDesktop -Hwnd $Hwnd

        $title = [TeamerWin32]::GetWindowTitle([IntPtr]$Hwnd)
        $shortTitle = if ($title.Length -gt 30) { $title.Substring(0, 27) + "..." } else { $title }

        Write-Host "Moved `"$shortTitle`" to Desktop $($ToDesktop + 1)" -ForegroundColor Green

        return @{ Success = $true; Hwnd = $Hwnd; ToDesktop = $ToDesktop }
    }
    catch {
        Write-Error "Failed to move window: $_"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Switch-TeamerDesktop {
    <#
    .SYNOPSIS
        Switches to a virtual desktop
    .DESCRIPTION
        Switches the active desktop. Works on any desktop (including protected).
    .PARAMETER Index
        The 0-based desktop index
    .EXAMPLE
        Switch-TeamerDesktop -Index 2
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$Index
    )

    try {
        $desktop = Get-Desktop -Index $Index
        if (-not $desktop) {
            throw "Desktop index $Index does not exist"
        }

        Switch-Desktop -Desktop $desktop

        $desktopName = ""
        try { $desktopName = Get-DesktopName -Desktop $desktop } catch { }
        if ([string]::IsNullOrWhiteSpace($desktopName)) { $desktopName = "Desktop $($Index + 1)" }

        Write-Host "Switched to Desktop $($Index + 1) `"$desktopName`"" -ForegroundColor Green

        return @{ Success = $true; Index = $Index; Name = $desktopName }
    }
    catch {
        Write-Error "Failed to switch desktop: $_"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Pin-TeamerWindow {
    <#
    .SYNOPSIS
        Pins a window to all desktops
    .DESCRIPTION
        Pins a window so it appears on all virtual desktops.
    .PARAMETER Hwnd
        The window handle
    .PARAMETER Unpin
        If specified, unpins the window instead
    .EXAMPLE
        Pin-TeamerWindow -Hwnd 12345
    .EXAMPLE
        Pin-TeamerWindow -Hwnd 12345 -Unpin
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int64]$Hwnd,

        [Parameter()]
        [switch]$Unpin
    )

    try {
        $title = [TeamerWin32]::GetWindowTitle([IntPtr]$Hwnd)
        $shortTitle = if ($title.Length -gt 30) { $title.Substring(0, 27) + "..." } else { $title }

        if ($Unpin) {
            Unpin-Window -Hwnd $Hwnd
            Write-Host "Unpinned `"$shortTitle`"" -ForegroundColor Green
        }
        else {
            Pin-Window -Hwnd $Hwnd
            Write-Host "Pinned `"$shortTitle`" to all desktops" -ForegroundColor Green
        }

        return @{ Success = $true; Hwnd = $Hwnd; Pinned = (-not $Unpin) }
    }
    catch {
        Write-Error "Failed to pin/unpin window: $_"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

#endregion

#region DELETE Functions

function Remove-TeamerDesktop {
    <#
    .SYNOPSIS
        Removes a virtual desktop
    .DESCRIPTION
        Removes a desktop. Only works on Desktop 3+ (index 2+).
        Windows on the removed desktop will be moved to an adjacent desktop.
    .PARAMETER Index
        The 0-based desktop index (must be 2+)
    .PARAMETER Force
        Skip confirmation prompt
    .EXAMPLE
        Remove-TeamerDesktop -Index 2
    .EXAMPLE
        Remove-TeamerDesktop -Index 2 -Force
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$Index,

        [Parameter()]
        [switch]$Force
    )

    Assert-SafeDesktopOperation -Index $Index -Operation "remove"

    try {
        $desktop = Get-Desktop -Index $Index
        if (-not $desktop) {
            throw "Desktop index $Index does not exist"
        }

        $desktopName = ""
        try { $desktopName = Get-DesktopName -Desktop $desktop } catch { }
        if ([string]::IsNullOrWhiteSpace($desktopName)) { $desktopName = "Desktop $($Index + 1)" }

        # Confirmation
        if (-not $Force) {
            Write-Host "Are you sure you want to remove Desktop $($Index + 1) `"$desktopName`"?" -ForegroundColor Yellow
            Write-Host "Windows will be moved to an adjacent desktop." -ForegroundColor Yellow
            $confirm = Read-Host "Type 'yes' to confirm"

            if ($confirm -ne 'yes') {
                Write-Host "Operation cancelled." -ForegroundColor Cyan
                return @{ Success = $false; Cancelled = $true }
            }
        }

        Remove-Desktop -Desktop $desktop
        Write-Host "Removed Desktop $($Index + 1) `"$desktopName`"" -ForegroundColor Green

        return @{ Success = $true; Index = $Index; Name = $desktopName }
    }
    catch {
        Write-Error "Failed to remove desktop: $_"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Close-TeamerWindow {
    <#
    .SYNOPSIS
        Closes a window
    .DESCRIPTION
        Closes a window by handle. Only works on windows on Desktop 3+ (index 2+).
    .PARAMETER Hwnd
        The window handle
    .PARAMETER Force
        Close windows on any desktop (USE WITH CAUTION)
    .EXAMPLE
        Close-TeamerWindow -Hwnd 12345
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int64]$Hwnd,

        [Parameter()]
        [switch]$Force
    )

    try {
        # Check which desktop the window is on
        $windowDesktopIndex = Get-WindowDesktopIndex -Hwnd $Hwnd

        if (-not $Force -and $windowDesktopIndex -ge 0) {
            Assert-SafeDesktopOperation -Index $windowDesktopIndex -Operation "close window on"
        }

        $title = [TeamerWin32]::GetWindowTitle([IntPtr]$Hwnd)
        $shortTitle = if ($title.Length -gt 30) { $title.Substring(0, 27) + "..." } else { $title }

        $result = [TeamerWin32]::CloseWindow([IntPtr]$Hwnd)

        if ($result) {
            Write-Host "Closed `"$shortTitle`"" -ForegroundColor Green
            return @{ Success = $true; Hwnd = $Hwnd }
        }
        else {
            throw "CloseWindow returned false"
        }
    }
    catch {
        Write-Error "Failed to close window: $_"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

#endregion

#region Export Functions

# Note: Export-ModuleMember only works when loaded as a module (.psm1)
# This script is designed to be dot-sourced: . .\Manage-Teamer.ps1
# All functions are automatically available after dot-sourcing

#endregion

# Display welcome message when dot-sourced
$protectedNamesList = $script:PROTECTED_DESKTOP_NAMES -join ", "

Write-Host ""
Write-Host "Teamer CRUD Module Loaded" -ForegroundColor Green
Write-Host "Protected Desktops (by name): $protectedNamesList" -ForegroundColor Yellow
Write-Host "Manageable: Any desktop NOT in protected list" -ForegroundColor Cyan
Write-Host ""
Write-Host "Commands:" -ForegroundColor Cyan
Write-Host "  Show-TeamerTree               - Display system tree" -ForegroundColor White
Write-Host "  New-TeamerDesktop             - Create new desktop" -ForegroundColor White
Write-Host "  New-TeamerTerminal            - Launch terminal" -ForegroundColor White
Write-Host "  Rename-TeamerDesktop          - Rename desktop (non-protected)" -ForegroundColor White
Write-Host "  Move-TeamerWindow             - Move window" -ForegroundColor White
Write-Host "  Switch-TeamerDesktop          - Switch desktop" -ForegroundColor White
Write-Host "  Pin-TeamerWindow              - Pin/unpin window" -ForegroundColor White
Write-Host "  Remove-TeamerDesktop          - Remove desktop (non-protected)" -ForegroundColor White
Write-Host "  Close-TeamerWindow            - Close window (non-protected desktop)" -ForegroundColor White
Write-Host ""
