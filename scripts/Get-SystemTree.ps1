<#
.SYNOPSIS
    Captures system configuration as a hierarchical tree: Screen -> Virtual Desktop -> Window -> Tabs

.DESCRIPTION
    Builds a tree structure showing:
    - Screens (monitors) with their properties
    - Virtual desktops on each screen (with windows present)
    - Windows classified as terminals or applications
    - For Windows Terminal: attempts to enumerate tabs

.PARAMETER OutputPath
    Directory to save the config file. Defaults to E:\Projects\teamer\config\

.EXAMPLE
    .\Get-SystemTree.ps1

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

    [DllImport("user32.dll")]
    public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    public const int GWL_STYLE = -16;
    public const int GWL_EXSTYLE = -20;
    public const uint WS_VISIBLE = 0x10000000;
    public const uint WS_CAPTION = 0x00C00000;
    public const uint WS_EX_TOOLWINDOW = 0x00000080;
    public const uint WS_EX_APPWINDOW = 0x00040000;
    public const uint MONITOR_DEFAULTTONEAREST = 2;

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
}
"@
#endregion

#region Helper Functions

function Ensure-Dependencies {
    Write-Host "Checking dependencies..." -ForegroundColor Cyan

    $loadedModule = Get-Module -Name VirtualDesktop -ErrorAction SilentlyContinue
    if ($loadedModule) {
        Write-Host "PSVirtualDesktop module already loaded." -ForegroundColor Green
        return
    }

    $availableModule = Get-Module -ListAvailable -Name VirtualDesktop -ErrorAction SilentlyContinue
    if (-not $availableModule) {
        Write-Host "PSVirtualDesktop module not found. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name VirtualDesktop -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
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

    try {
        Import-Module VirtualDesktop -Force -ErrorAction Stop
        Write-Host "PSVirtualDesktop module loaded." -ForegroundColor Green
    }
    catch {
        $modulePaths = @(
            "$env:USERPROFILE\Documents\PowerShell\Modules\VirtualDesktop",
            "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\VirtualDesktop",
            "$env:USERPROFILE\OneDrive\Documents\PowerShell\Modules\VirtualDesktop",
            "$env:USERPROFILE\OneDrive\Documents\WindowsPowerShell\Modules\VirtualDesktop"
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
                catch { }
            }
        }

        if (-not $imported) {
            Write-Error "Could not load VirtualDesktop module."
            throw "Module import failed"
        }
    }

    $null = Get-Command Get-DesktopCount -ErrorAction Stop
}

function Get-Screens {
    Write-Host "Querying screens..." -ForegroundColor Cyan

    Add-Type -AssemblyName System.Windows.Forms
    $screens = [System.Windows.Forms.Screen]::AllScreens
    $videoControllers = Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue

    $result = @()
    $index = 1

    foreach ($screen in $screens) {
        $screenInfo = [ordered]@{
            number       = $index
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
            bounds       = [ordered]@{
                left   = $screen.Bounds.Left
                top    = $screen.Bounds.Top
                right  = $screen.Bounds.Right
                bottom = $screen.Bounds.Bottom
            }
            refreshRate  = if ($videoControllers) { $videoControllers[0].CurrentRefreshRate } else { 60 }
            desktops     = @()  # Will be populated later
        }
        $result += $screenInfo
        $index++
    }

    Write-Host "Found $($result.Count) screen(s)." -ForegroundColor Green
    return $result
}

function Get-VirtualDesktops {
    Write-Host "Querying virtual desktops..." -ForegroundColor Cyan

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
            index     = $i
            name      = $desktopName
            isCurrent = ($i -eq $currentIndex)
            windows   = @()  # Will be populated later
        }
    }

    Write-Host "Found $desktopCount virtual desktop(s). Current: $currentIndex" -ForegroundColor Green
    return $desktops
}

function Get-WindowsTerminalProfiles {
    # Load Windows Terminal settings to get profile info
    $settingsPath = [Environment]::GetFolderPath('LocalApplicationData') +
                    '\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json'

    if (Test-Path $settingsPath) {
        try {
            $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json
            $profiles = @{}

            foreach ($profile in $settings.profiles.list) {
                $profiles[$profile.guid] = [ordered]@{
                    name       = $profile.name
                    hidden     = $profile.hidden
                    source     = $profile.source
                    tabColor   = $profile.tabColor
                    icon       = $profile.icon
                }
            }
            return $profiles
        }
        catch {
            return @{}
        }
    }
    return @{}
}

function Get-WindowsTerminalTabs {
    param([int64]$Hwnd)

    # Windows Terminal doesn't expose tabs via simple API
    # We can try UI Automation to enumerate tab items
    $tabs = @()

    try {
        Add-Type -AssemblyName UIAutomationClient
        Add-Type -AssemblyName UIAutomationTypes

        $automation = [System.Windows.Automation.AutomationElement]
        $window = $automation::FromHandle([IntPtr]$Hwnd)

        if ($window) {
            # Find TabItem elements
            $tabCondition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                [System.Windows.Automation.ControlType]::TabItem
            )

            $tabItems = $window.FindAll(
                [System.Windows.Automation.TreeScope]::Descendants,
                $tabCondition
            )

            foreach ($tab in $tabItems) {
                $name = $tab.Current.Name
                if (-not [string]::IsNullOrWhiteSpace($name)) {
                    $tabs += [ordered]@{
                        name     = $name
                        isActive = $false  # Would need more work to determine
                    }
                }
            }
        }
    }
    catch {
        # UI Automation failed, fall back to window title
        $title = [Win32Window]::GetWindowTitle([IntPtr]$Hwnd)
        if ($title) {
            $tabs += [ordered]@{
                name     = $title
                isActive = $true
                note     = "Active tab from window title"
            }
        }
    }

    return $tabs
}

function Detect-ShellType {
    param([string]$Title)

    # WSL/Linux detection patterns (check first as they're more specific)
    # Linux prompts: "user: ~", "user@host:", "username: /path"
    if ($Title -match '(?i)(ubuntu|debian|kali|fedora|opensuse|arch|manjaro|wsl)') {
        $distro = $Matches[1]
        return [ordered]@{
            shell  = 'wsl'
            distro = $distro
            label  = "WSL ($distro)"
        }
    }

    # Linux-style prompt detection: "username: ~" or "username: /path" or "user@host"
    if ($Title -match '^[a-z][a-z0-9_-]*:\s*[~/]' -or $Title -match '^[a-z][a-z0-9_-]*@[a-z]') {
        return [ordered]@{
            shell  = 'wsl'
            distro = 'Linux'
            label  = 'WSL (Linux)'
        }
    }

    # PowerShell detection
    if ($Title -match '(?i)(powershell|pwsh|PS\s+[A-Z]:\\)') {
        if ($Title -match '(?i)pwsh|PowerShell\s*7|Preview') {
            return [ordered]@{
                shell = 'powershell-core'
                label = 'PowerShell Core'
            }
        }
        return [ordered]@{
            shell = 'powershell'
            label = 'PowerShell'
        }
    }

    # Git Bash detection
    if ($Title -match '(?i)(mingw|git\s*bash|bash.*git)') {
        return [ordered]@{
            shell = 'git-bash'
            label = 'Git Bash'
        }
    }

    # CMD detection
    if ($Title -match '(?i)(command\s*prompt|cmd\.exe|^[A-Z]:\\.*>)') {
        return [ordered]@{
            shell = 'cmd'
            label = 'Command Prompt'
        }
    }

    # Generic bash/zsh
    if ($Title -match '(?i)(bash|zsh|fish|sh\s*-)') {
        return [ordered]@{
            shell = 'shell'
            label = 'Shell'
        }
    }

    # Unknown shell in terminal
    return [ordered]@{
        shell = 'unknown'
        label = 'Terminal'
    }
}

function Classify-Window {
    param(
        [string]$ProcessName,
        [string]$Title,
        [string]$ClassName
    )

    # Terminal host applications
    $terminalHosts = @{
        'WindowsTerminal' = 'Windows Terminal'
        'ConEmu64'        = 'ConEmu'
        'ConEmuC64'       = 'ConEmu'
        'Hyper'           = 'Hyper'
        'Alacritty'       = 'Alacritty'
        'wezterm-gui'     = 'WezTerm'
        'Tabby'           = 'Tabby'
        'Terminus'        = 'Terminus'
    }

    # Direct shell processes (not in a terminal host)
    $directShells = @{
        'powershell' = @{ shell = 'powershell'; label = 'PowerShell' }
        'pwsh'       = @{ shell = 'powershell-core'; label = 'PowerShell Core' }
        'cmd'        = @{ shell = 'cmd'; label = 'Command Prompt' }
        'wsl'        = @{ shell = 'wsl'; label = 'WSL' }
        'wslhost'    = @{ shell = 'wsl'; label = 'WSL' }
        'ubuntu'     = @{ shell = 'wsl'; distro = 'Ubuntu'; label = 'WSL (Ubuntu)' }
        'debian'     = @{ shell = 'wsl'; distro = 'Debian'; label = 'WSL (Debian)' }
        'kali'       = @{ shell = 'wsl'; distro = 'Kali'; label = 'WSL (Kali)' }
        'mintty'     = @{ shell = 'git-bash'; label = 'Git Bash' }
        'bash'       = @{ shell = 'bash'; label = 'Bash' }
    }

    # Check if it's a terminal host (like Windows Terminal)
    if ($terminalHosts.ContainsKey($ProcessName)) {
        $shellInfo = Detect-ShellType -Title $Title

        return [ordered]@{
            type        = 'terminal'
            terminalApp = $terminalHosts[$ProcessName]
            shell       = $shellInfo.shell
            shellLabel  = $shellInfo.label
            distro      = $shellInfo.distro
        }
    }

    # Check if it's a direct shell process
    if ($directShells.ContainsKey($ProcessName)) {
        $info = $directShells[$ProcessName]
        return [ordered]@{
            type        = 'terminal'
            terminalApp = $null  # Running directly, not in a host
            shell       = $info.shell
            shellLabel  = $info.label
            distro      = $info.distro
        }
    }

    # Check for terminal-like titles (fallback)
    if ($Title -match '(?i)(PowerShell|Command Prompt|cmd\.exe|bash|zsh|ubuntu|wsl|Terminal)') {
        $shellInfo = Detect-ShellType -Title $Title
        return [ordered]@{
            type        = 'terminal'
            terminalApp = 'Unknown Terminal'
            shell       = $shellInfo.shell
            shellLabel  = $shellInfo.label
            distro      = $shellInfo.distro
        }
    }

    return [ordered]@{
        type = 'application'
    }
}

function Get-AllWindows {
    Write-Host "Enumerating windows..." -ForegroundColor Cyan

    $windows = @()
    $allHandles = [Win32Window]::GetAllWindows()
    $processCache = @{}
    $terminalProfiles = Get-WindowsTerminalProfiles

    $skipClasses = @('Progman', 'WorkerW', 'Shell_TrayWnd', 'Shell_SecondaryTrayWnd', 'NotifyIconOverflowWindow')

    foreach ($hwnd in $allHandles) {
        if (-not [Win32Window]::IsAltTabWindow($hwnd)) {
            continue
        }

        $className = [Win32Window]::GetWindowClassName($hwnd)
        if ($className -in $skipClasses) {
            continue
        }

        $title = [Win32Window]::GetWindowTitle($hwnd)

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

        # Get monitor this window is on
        $monitorDevice = [Win32Window]::GetMonitorDeviceName($hwnd)

        # Get desktop assignment
        $desktopIndex = -1
        try {
            $desktop = Get-DesktopFromWindow -Hwnd $hwnd.ToInt64()
            if ($desktop) {
                $desktopIndex = Get-DesktopIndex -Desktop $desktop
            }
        }
        catch { }

        # Check if pinned
        $isPinned = $false
        try {
            $isPinned = Test-WindowPinned -Hwnd $hwnd.ToInt64()
        }
        catch { }

        # Classify window
        $classification = Classify-Window -ProcessName $processName -Title $title -ClassName $className

        $window = [ordered]@{
            hwnd          = $hwnd.ToInt64()
            processName   = $processName
            processId     = $processId
            title         = $title
            monitorDevice = $monitorDevice
            desktopIndex  = $desktopIndex
            position      = [ordered]@{
                x = $rect.Left
                y = $rect.Top
            }
            size          = [ordered]@{
                width  = $rect.Right - $rect.Left
                height = $rect.Bottom - $rect.Top
            }
            state         = $state
            isPinned      = $isPinned
            type          = $classification.type
        }

        # Add terminal-specific info
        if ($classification.type -eq 'terminal') {
            $window.terminalApp = $classification.terminalApp
            $window.shell = $classification.shell
            $window.shellLabel = $classification.shellLabel

            if ($classification.distro) {
                $window.distro = $classification.distro
            }

            # Get tabs for Windows Terminal
            if ($processName -eq 'WindowsTerminal') {
                $tabs = Get-WindowsTerminalTabs -Hwnd $hwnd.ToInt64()
                if ($tabs.Count -gt 0) {
                    # Detect shell type for each tab
                    $enhancedTabs = @()
                    foreach ($tab in $tabs) {
                        $tabShellInfo = Detect-ShellType -Title $tab.name
                        $enhancedTab = [ordered]@{
                            name       = $tab.name
                            shell      = $tabShellInfo.shell
                            shellLabel = $tabShellInfo.label
                            isActive   = $tab.isActive
                        }
                        if ($tabShellInfo.distro) {
                            $enhancedTab.distro = $tabShellInfo.distro
                        }
                        $enhancedTabs += $enhancedTab
                    }
                    $window.tabs = $enhancedTabs
                }
            }
        }

        $windows += $window
    }

    Write-Host "Found $($windows.Count) window(s)." -ForegroundColor Green
    return $windows
}

function Build-Tree {
    param(
        [array]$Screens,
        [array]$Desktops,
        [array]$Windows
    )

    Write-Host "Building tree structure..." -ForegroundColor Cyan

    # Create a deep copy of screens and desktops for the tree
    $tree = @()

    foreach ($screen in $Screens) {
        $screenNode = [ordered]@{
            number      = $screen.number
            deviceName  = $screen.deviceName
            isPrimary   = $screen.isPrimary
            resolution  = $screen.resolution
            refreshRate = $screen.refreshRate
            desktops    = @()
        }

        # For each desktop, find windows on this screen
        foreach ($desktop in $Desktops) {
            $desktopWindows = @($Windows | Where-Object {
                $_.desktopIndex -eq $desktop.index -and
                $_.monitorDevice -eq $screen.deviceName
            })

            # Also include pinned windows (they appear on all desktops)
            $pinnedWindows = @($Windows | Where-Object {
                $_.isPinned -eq $true -and
                $_.monitorDevice -eq $screen.deviceName
            })

            $allWindows = @()
            $allWindows += $desktopWindows
            foreach ($pw in $pinnedWindows) {
                if ($allWindows.hwnd -notcontains $pw.hwnd) {
                    $allWindows += $pw
                }
            }

            if ($allWindows.Count -gt 0) {
                $desktopNode = [ordered]@{
                    index     = $desktop.index
                    name      = $desktop.name
                    isCurrent = $desktop.isCurrent
                    windows   = @()
                }

                foreach ($win in $allWindows) {
                    $windowNode = [ordered]@{
                        title       = $win.title
                        processName = $win.processName
                        type        = $win.type
                        state       = $win.state
                        isPinned    = $win.isPinned
                        hwnd        = $win.hwnd
                    }

                    if ($win.type -eq 'terminal') {
                        $windowNode.terminalApp = $win.terminalApp
                        $windowNode.shell = $win.shell
                        $windowNode.shellLabel = $win.shellLabel

                        if ($win.distro) {
                            $windowNode.distro = $win.distro
                        }

                        if ($win.tabs) {
                            $windowNode.tabs = $win.tabs
                        }
                    }

                    $desktopNode.windows += $windowNode
                }

                $screenNode.desktops += $desktopNode
            }
        }

        $tree += $screenNode
    }

    return $tree
}

function Print-Tree {
    param([array]$Tree)

    Write-Host "`n" -NoNewline
    Write-Host "=" * 70 -ForegroundColor Magenta
    Write-Host "  SYSTEM TREE" -ForegroundColor Magenta
    Write-Host "=" * 70 -ForegroundColor Magenta

    foreach ($screen in $Tree) {
        $primaryMark = if ($screen.isPrimary) { " (Primary)" } else { "" }
        Write-Host "`n[Screen $($screen.number)]$primaryMark $($screen.resolution.width)x$($screen.resolution.height) @ $($screen.refreshRate)Hz" -ForegroundColor Yellow

        if ($screen.desktops.Count -eq 0) {
            Write-Host "  (no windows)" -ForegroundColor DarkGray
            continue
        }

        foreach ($desktop in $screen.desktops) {
            $currentMark = if ($desktop.isCurrent) { " *ACTIVE*" } else { "" }
            Write-Host "  [Desktop: $($desktop.name)]$currentMark" -ForegroundColor Cyan

            foreach ($window in $desktop.windows) {
                $pinnedMark = if ($window.isPinned) { " (pinned)" } else { "" }
                $stateIcon = switch ($window.state) {
                    "Minimized" { " [-]" }
                    "Maximized" { " [+]" }
                    default { "" }
                }

                $displayTitle = if ($window.title.Length -gt 50) {
                    $window.title.Substring(0, 47) + "..."
                } else {
                    $window.title
                }

                if ($window.type -eq 'terminal') {
                    # Terminal window: show host > shell format
                    $shellIcon = switch ($window.shell) {
                        'powershell'      { "[PS]" }
                        'powershell-core' { "[PS7]" }
                        'wsl'             { "[WSL]" }
                        'cmd'             { "[CMD]" }
                        'git-bash'        { "[GIT]" }
                        'shell'           { "[SH]" }
                        default           { "[T]" }
                    }

                    # Use terminalApp name or process name
                    $hostName = if ($window.terminalApp) { $window.terminalApp } else { $window.processName }
                    $shellName = if ($window.shellLabel) { $window.shellLabel } else { "Terminal" }

                    Write-Host "    $shellIcon " -ForegroundColor Green -NoNewline
                    Write-Host "$hostName" -ForegroundColor White -NoNewline
                    Write-Host " > " -ForegroundColor DarkGray -NoNewline
                    Write-Host "$shellName" -ForegroundColor Cyan -NoNewline
                    Write-Host "$stateIcon$pinnedMark" -ForegroundColor White
                    Write-Host "        $displayTitle" -ForegroundColor Gray
                }
                else {
                    # Application window
                    Write-Host "    [A] " -ForegroundColor White -NoNewline
                    Write-Host "$($window.processName)$stateIcon$pinnedMark" -ForegroundColor White
                    Write-Host "        $displayTitle" -ForegroundColor Gray
                }

                # Show tabs for terminals with shell info
                if ($window.tabs) {
                    foreach ($tab in $window.tabs) {
                        $activeIcon = if ($tab.isActive) { ">" } else { " " }
                        $tabShellName = if ($tab.shellLabel) { $tab.shellLabel } else { "Tab" }

                        $tabIcon = switch ($tab.shell) {
                            'powershell'      { "[PS]" }
                            'powershell-core' { "[PS7]" }
                            'wsl'             { "[WSL]" }
                            'cmd'             { "[CMD]" }
                            'git-bash'        { "[GIT]" }
                            default           { "[TAB]" }
                        }

                        Write-Host "        $activeIcon $tabIcon " -ForegroundColor DarkCyan -NoNewline
                        Write-Host "$tabShellName" -ForegroundColor Cyan -NoNewline
                        Write-Host ": $($tab.name)" -ForegroundColor DarkGray
                    }
                }
            }
        }
    }

    Write-Host "`n" -NoNewline
    Write-Host "=" * 70 -ForegroundColor Magenta
    Write-Host "Legend: [PS]=PowerShell [PS7]=PS Core [WSL]=WSL [CMD]=Cmd [GIT]=Git Bash [A]=App" -ForegroundColor DarkGray
    Write-Host "        [-]=Minimized [+]=Maximized" -ForegroundColor DarkGray
    Write-Host "=" * 70 -ForegroundColor Magenta
}

function Export-Tree {
    param(
        [array]$Tree,
        [string]$OutputPath
    )

    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $timestamp = Get-Date -Format "yyyy-MM-ddTHH-mm-ss"
    $filename = "system-tree-$timestamp.json"
    $fullPath = Join-Path $OutputPath $filename

    $config = [ordered]@{
        timestamp = (Get-Date -Format "o")
        tree      = $Tree
    }

    $json = $config | ConvertTo-Json -Depth 10
    $json | Out-File -FilePath $fullPath -Encoding UTF8

    Write-Host "`nTree saved to: $fullPath" -ForegroundColor Green
    return $fullPath
}

#endregion

#region Main Execution

Write-Host "`n" -NoNewline
Write-Host "=" * 60 -ForegroundColor Magenta
Write-Host "  System Tree Discovery Tool" -ForegroundColor Magenta
Write-Host "=" * 60 -ForegroundColor Magenta
Write-Host ""

try {
    Ensure-Dependencies
    Write-Host ""

    # Gather data
    $screens = @(Get-Screens)
    $desktops = @(Get-VirtualDesktops)
    $windows = @(Get-AllWindows)

    # Build hierarchical tree
    $tree = @(Build-Tree -Screens $screens -Desktops $desktops -Windows $windows)

    # Print tree to console
    Print-Tree -Tree $tree

    # Export to JSON
    $outputFile = Export-Tree -Tree $tree -OutputPath $OutputPath

    # Summary
    $totalWindows = ($tree | ForEach-Object { $_.desktops } | ForEach-Object { $_.windows } | Measure-Object).Count
    $terminalCount = ($tree | ForEach-Object { $_.desktops } | ForEach-Object { $_.windows } | Where-Object { $_.type -eq 'terminal' } | Measure-Object).Count

    Write-Host "`nSummary:" -ForegroundColor Cyan
    Write-Host "  Screens:    $($screens.Count)" -ForegroundColor White
    Write-Host "  Desktops:   $($desktops.Count)" -ForegroundColor White
    Write-Host "  Windows:    $totalWindows (Terminals: $terminalCount)" -ForegroundColor White

    return $tree
}
catch {
    Write-Error "Failed to capture system tree: $_"
    throw
}

#endregion
