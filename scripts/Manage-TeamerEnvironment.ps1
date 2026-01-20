<#
.SYNOPSIS
    Teamer Environment Manager - CRUD operations for development environments.

.DESCRIPTION
    Manages project-specific development environments including:
    - Templates: Reusable base configurations
    - Profiles: Shell/terminal profiles
    - Layouts: Multi-monitor window arrangements
    - Environments: Project-specific configurations

    Start behavior is ADDITIVE - creates new desktops alongside existing ones.

.NOTES
    Requires: Manage-Teamer.ps1 (for desktop/window operations)
    Requires: Windows 10 2004+ or Windows 11
#>

#region Constants

$script:PROJECT_ROOT = "E:\Projects\teamer"
$script:ENVIRONMENTS_ROOT = Join-Path $script:PROJECT_ROOT "environments"
$script:TEMPLATES_PATH = Join-Path $script:ENVIRONMENTS_ROOT "templates"
$script:PROFILES_PATH = Join-Path $script:ENVIRONMENTS_ROOT "profiles"
$script:LAYOUTS_PATH = Join-Path $script:ENVIRONMENTS_ROOT "layouts"
$script:PROJECTS_PATH = Join-Path $script:ENVIRONMENTS_ROOT "projects"
$script:SERVICES_PATH = Join-Path $script:ENVIRONMENTS_ROOT "services"
$script:STATE_FILE = Join-Path $script:ENVIRONMENTS_ROOT "state.json"

# Track active environments and their desktop indices (loaded from state file)
$script:ActiveEnvironments = @{}

#endregion

#region State Management

function Save-TeamerState {
    <#
    .SYNOPSIS
        Saves active environments to state file
    #>
    $state = @{
        timestamp = (Get-Date -Format "o")
        activeEnvironments = @{}
    }

    foreach ($key in $script:ActiveEnvironments.Keys) {
        $env = $script:ActiveEnvironments[$key]
        $state.activeEnvironments[$key] = @{
            desktopIndices = $env.DesktopIndices
            startTime = $env.StartTime.ToString("o")
        }
    }

    $json = $state | ConvertTo-Json -Depth 5
    $json | Out-File -FilePath $script:STATE_FILE -Encoding UTF8 -Force
}

function Load-TeamerState {
    <#
    .SYNOPSIS
        Loads active environments from state file
    #>
    if (-not (Test-Path $script:STATE_FILE)) {
        return
    }

    try {
        $content = Get-Content -Path $script:STATE_FILE -Raw -Encoding UTF8
        $state = $content | ConvertFrom-Json

        if ($state.activeEnvironments) {
            foreach ($prop in $state.activeEnvironments.PSObject.Properties) {
                $name = $prop.Name
                $data = $prop.Value

                # Verify the desktops still exist
                $validIndices = @()
                foreach ($index in $data.desktopIndices) {
                    try {
                        $desktop = Get-Desktop -Index $index
                        if ($desktop) {
                            $validIndices += $index
                        }
                    }
                    catch {
                        # Desktop doesn't exist anymore
                    }
                }

                if ($validIndices.Count -gt 0) {
                    $config = Get-TeamerEnvironment -Name $name
                    $script:ActiveEnvironments[$name] = @{
                        Config = $config
                        DesktopIndices = $validIndices
                        StartTime = [DateTime]::Parse($data.startTime)
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Could not load state file: $_"
    }

    # Clean up state file if environments were removed
    Save-TeamerState
}

function Clear-TeamerState {
    <#
    .SYNOPSIS
        Clears the state file
    #>
    $script:ActiveEnvironments = @{}
    if (Test-Path $script:STATE_FILE) {
        Remove-Item -Path $script:STATE_FILE -Force
    }
}

#endregion

#region Module Initialization

# Load Manage-Teamer.ps1 at script scope (not inside a function)
# This ensures all its functions are available in the caller's scope
$script:ManageTeamerPath = Join-Path $script:PROJECT_ROOT "scripts\Manage-Teamer.ps1"
if (-not (Get-Command "New-TeamerDesktop" -ErrorAction SilentlyContinue)) {
    if (Test-Path $script:ManageTeamerPath) {
        . $script:ManageTeamerPath
    }
    else {
        throw "Manage-Teamer.ps1 not found at $script:ManageTeamerPath"
    }
}

# Ensure directories exist
$script:RequiredPaths = @($script:TEMPLATES_PATH, $script:PROFILES_PATH, $script:LAYOUTS_PATH, $script:PROJECTS_PATH, $script:SERVICES_PATH)
foreach ($path in $script:RequiredPaths) {
    if (-not (Test-Path $path)) {
        New-Item -ItemType Directory -Path $path -Force | Out-Null
    }
}

# State loading is deferred until after all functions are defined (see end of script)

#endregion

#region Configuration Helpers

function Get-TeamerConfigPath {
    <#
    .SYNOPSIS
        Resolves a config file path based on type and name
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('template', 'profile', 'layout', 'project', 'service')]
        [string]$Type,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $basePath = switch ($Type) {
        'template' { $script:TEMPLATES_PATH }
        'profile'  { $script:PROFILES_PATH }
        'layout'   { $script:LAYOUTS_PATH }
        'project'  { $script:PROJECTS_PATH }
        'service'  { $script:SERVICES_PATH }
    }

    # Handle nested project directories
    if ($Type -eq 'project') {
        $projectDir = Join-Path $basePath $Name
        return Join-Path $projectDir "environment.json"
    }

    return Join-Path $basePath "$Name.json"
}

function Read-TeamerConfig {
    <#
    .SYNOPSIS
        Reads and parses a JSON config file
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        return $null
    }

    try {
        $content = Get-Content -Path $Path -Raw -Encoding UTF8
        return $content | ConvertFrom-Json
    }
    catch {
        Write-Error "Failed to parse config at $Path : $_"
        return $null
    }
}

function Write-TeamerConfig {
    <#
    .SYNOPSIS
        Writes a config object to JSON file
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [object]$Config
    )

    $parentDir = Split-Path -Path $Path -Parent
    if (-not (Test-Path $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $json = $Config | ConvertTo-Json -Depth 10
    $json | Out-File -FilePath $Path -Encoding UTF8 -Force
}

#endregion

#region Layout Functions

function Get-TeamerLayout {
    <#
    .SYNOPSIS
        Gets a layout configuration
    .PARAMETER Name
        Layout name (without .json extension)
    .EXAMPLE
        Get-TeamerLayout -Name "single-focus"
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        # List all layouts
        $layouts = Get-ChildItem -Path $script:LAYOUTS_PATH -Filter "*.json" -ErrorAction SilentlyContinue
        $result = @()
        foreach ($file in $layouts) {
            $config = Read-TeamerConfig -Path $file.FullName
            if ($config) {
                $result += [PSCustomObject]@{
                    Name        = $file.BaseName
                    DisplayName = $config.name
                    Description = $config.description
                    Zones       = ($config.zones | ForEach-Object { $_.name }) -join ", "
                }
            }
        }
        return $result
    }

    $path = Get-TeamerConfigPath -Type layout -Name $Name
    return Read-TeamerConfig -Path $path
}

function Get-TeamerZone {
    <#
    .SYNOPSIS
        Resolves a zone from a layout configuration
    .PARAMETER Layout
        Layout name
    .PARAMETER Zone
        Zone name
    .EXAMPLE
        Get-TeamerZone -Layout "dual-code-services" -Zone "main"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Layout,

        [Parameter(Mandatory)]
        [string]$Zone
    )

    $layoutConfig = Get-TeamerLayout -Name $Layout
    if (-not $layoutConfig) {
        Write-Warning "Layout '$Layout' not found"
        return $null
    }

    $zoneConfig = $layoutConfig.zones | Where-Object { $_.name -eq $Zone }
    if (-not $zoneConfig) {
        Write-Warning "Zone '$Zone' not found in layout '$Layout'"
        return $null
    }

    return $zoneConfig
}

#endregion

#region Profile Functions

function Get-TeamerProfile {
    <#
    .SYNOPSIS
        Gets a shell profile configuration
    .PARAMETER Name
        Profile name (without .json extension)
    .EXAMPLE
        Get-TeamerProfile -Name "wsl-ubuntu"
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        # List all profiles
        $profiles = Get-ChildItem -Path $script:PROFILES_PATH -Filter "*.json" -ErrorAction SilentlyContinue
        $result = @()
        foreach ($file in $profiles) {
            $config = Read-TeamerConfig -Path $file.FullName
            if ($config) {
                $result += [PSCustomObject]@{
                    Name        = $file.BaseName
                    DisplayName = $config.name
                    Shell       = $config.shell
                    Distribution = $config.distribution
                }
            }
        }
        return $result
    }

    $path = Get-TeamerConfigPath -Type profile -Name $Name
    return Read-TeamerConfig -Path $path
}

function New-TeamerProfile {
    <#
    .SYNOPSIS
        Creates a new shell profile
    .PARAMETER Name
        Profile name
    .PARAMETER Shell
        Shell type: powershell, pwsh, wsl, cmd, git-bash
    .PARAMETER Distribution
        WSL distribution name (for wsl shell type)
    .PARAMETER TerminalProfile
        Windows Terminal profile name
    .EXAMPLE
        New-TeamerProfile -Name "my-wsl" -Shell wsl -Distribution "Ubuntu-22.04"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateSet('powershell', 'pwsh', 'wsl', 'cmd', 'git-bash')]
        [string]$Shell,

        [Parameter()]
        [string]$Distribution,

        [Parameter()]
        [string]$TerminalProfile
    )

    $path = Get-TeamerConfigPath -Type profile -Name $Name
    if (Test-Path $path) {
        throw "Profile '$Name' already exists"
    }

    $config = @{
        '$schema' = "../schemas/profile.schema.json"
        name = $Name
        shell = $Shell
        startupCommands = @()
        environment = @{}
    }

    if ($Distribution) {
        $config.distribution = $Distribution
    }

    if ($TerminalProfile) {
        $config.terminalProfile = $TerminalProfile
    }
    else {
        # Default terminal profiles
        $config.terminalProfile = switch ($Shell) {
            'powershell' { "Windows PowerShell" }
            'pwsh'       { "PowerShell" }
            'wsl'        { if ($Distribution) { $Distribution } else { "Ubuntu" } }
            'cmd'        { "Command Prompt" }
            'git-bash'   { "Git Bash" }
        }
    }

    Write-TeamerConfig -Path $path -Config $config
    Write-Host "Created profile: $Name" -ForegroundColor Green

    return $config
}

function Set-TeamerProfile {
    <#
    .SYNOPSIS
        Updates a shell profile property
    .PARAMETER Name
        Profile name
    .PARAMETER Property
        Property to update
    .PARAMETER Value
        New value
    .EXAMPLE
        Set-TeamerProfile -Name "wsl-ubuntu" -Property "distribution" -Value "Ubuntu-22.04"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string]$Property,

        [Parameter(Mandatory)]
        $Value
    )

    $path = Get-TeamerConfigPath -Type profile -Name $Name
    $config = Read-TeamerConfig -Path $path

    if (-not $config) {
        throw "Profile '$Name' not found"
    }

    # Convert to hashtable for modification
    $configHash = @{}
    $config.PSObject.Properties | ForEach-Object { $configHash[$_.Name] = $_.Value }
    $configHash[$Property] = $Value

    Write-TeamerConfig -Path $path -Config $configHash
    Write-Host "Updated profile '$Name': $Property = $Value" -ForegroundColor Green
}

function Remove-TeamerProfile {
    <#
    .SYNOPSIS
        Removes a shell profile
    .PARAMETER Name
        Profile name
    .PARAMETER Force
        Skip confirmation
    .EXAMPLE
        Remove-TeamerProfile -Name "old-profile" -Force
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [switch]$Force
    )

    $path = Get-TeamerConfigPath -Type profile -Name $Name

    if (-not (Test-Path $path)) {
        throw "Profile '$Name' not found"
    }

    if (-not $Force) {
        $confirm = Read-Host "Remove profile '$Name'? (yes/no)"
        if ($confirm -ne 'yes') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    Remove-Item -Path $path -Force
    Write-Host "Removed profile: $Name" -ForegroundColor Green
}

#endregion

#region Template Functions

function Get-TeamerTemplate {
    <#
    .SYNOPSIS
        Gets a template configuration
    .PARAMETER Name
        Template name (without .json extension)
    .EXAMPLE
        Get-TeamerTemplate -Name "fullstack"
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        # List all templates
        $templates = Get-ChildItem -Path $script:TEMPLATES_PATH -Filter "*.json" -ErrorAction SilentlyContinue
        $result = @()
        foreach ($file in $templates) {
            $config = Read-TeamerConfig -Path $file.FullName
            if ($config) {
                $result += [PSCustomObject]@{
                    Name        = $file.BaseName
                    DisplayName = $config.name
                    Description = $config.description
                    Layout      = $config.layout
                    Desktops    = $config.desktops.Count
                }
            }
        }
        return $result
    }

    $path = Get-TeamerConfigPath -Type template -Name $Name
    return Read-TeamerConfig -Path $path
}

function New-TeamerTemplate {
    <#
    .SYNOPSIS
        Creates a new template from an existing environment or from scratch
    .PARAMETER Name
        Template name
    .PARAMETER From
        Existing environment name to copy from (optional)
    .EXAMPLE
        New-TeamerTemplate -Name "my-template"
    .EXAMPLE
        New-TeamerTemplate -Name "custom-fullstack" -From "my-project"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [string]$From
    )

    $path = Get-TeamerConfigPath -Type template -Name $Name
    if (Test-Path $path) {
        throw "Template '$Name' already exists"
    }

    if ($From) {
        # Copy from existing environment
        $envConfig = Get-TeamerEnvironment -Name $From
        if (-not $envConfig) {
            throw "Environment '$From' not found"
        }

        # Convert to template (remove project-specific fields)
        $templateConfig = @{
            '$schema' = "../schemas/environment.schema.json"
            name = $Name
            description = "Template created from $From"
            layout = $envConfig.layout
            desktops = $envConfig.desktops
            onStart = $envConfig.onStart
            onStop = $envConfig.onStop
        }
    }
    else {
        # Create minimal template
        $templateConfig = @{
            '$schema' = "../schemas/environment.schema.json"
            name = $Name
            description = "Custom template"
            layout = "single-focus"
            desktops = @(
                @{
                    name = "Main"
                    windows = @(
                        @{
                            type = "terminal"
                            profile = "powershell"
                            zone = "main"
                            cwd = "."
                        }
                    )
                }
            )
            onStart = @()
            onStop = @()
        }
    }

    Write-TeamerConfig -Path $path -Config $templateConfig
    Write-Host "Created template: $Name" -ForegroundColor Green

    return $templateConfig
}

#endregion

#region Environment CRUD Functions

function Get-TeamerEnvironment {
    <#
    .SYNOPSIS
        Gets environment configuration(s)
    .PARAMETER Name
        Environment name (optional - lists all if not specified)
    .EXAMPLE
        Get-TeamerEnvironment
    .EXAMPLE
        Get-TeamerEnvironment -Name "my-saas"
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        # List all environments
        $projects = Get-ChildItem -Path $script:PROJECTS_PATH -Directory -ErrorAction SilentlyContinue
        $result = @()
        foreach ($dir in $projects) {
            $configPath = Join-Path $dir.FullName "environment.json"
            if (Test-Path $configPath) {
                $config = Read-TeamerConfig -Path $configPath
                if ($config) {
                    $isActive = $script:ActiveEnvironments.ContainsKey($dir.Name)
                    $result += [PSCustomObject]@{
                        Name        = $dir.Name
                        DisplayName = $config.name
                        Description = $config.description
                        Layout      = $config.layout
                        Desktops    = $config.desktops.Count
                        IsActive    = $isActive
                        WorkingDir  = $config.workingDirectory
                    }
                }
            }
        }
        return $result
    }

    $path = Get-TeamerConfigPath -Type project -Name $Name
    return Read-TeamerConfig -Path $path
}

function New-TeamerEnvironment {
    <#
    .SYNOPSIS
        Creates a new environment configuration
    .PARAMETER Name
        Environment name (used for directory and identification)
    .PARAMETER Template
        Template to base the environment on (optional)
    .PARAMETER WorkingDirectory
        Project working directory (optional)
    .PARAMETER Description
        Environment description (optional)
    .EXAMPLE
        New-TeamerEnvironment -Name "my-saas"
    .EXAMPLE
        New-TeamerEnvironment -Name "my-saas" -Template "fullstack" -WorkingDirectory "E:\Projects\my-saas"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [string]$Template,

        [Parameter()]
        [string]$WorkingDirectory,

        [Parameter()]
        [string]$Description
    )

    $path = Get-TeamerConfigPath -Type project -Name $Name
    if (Test-Path $path) {
        throw "Environment '$Name' already exists"
    }

    # Start with template or base config
    if ($Template) {
        $templateConfig = Get-TeamerTemplate -Name $Template
        if (-not $templateConfig) {
            throw "Template '$Template' not found"
        }
        $config = @{
            '$schema' = "../schemas/environment.schema.json"
            name = if ($Description) { $Description } else { $Name }
            description = "Created from template: $Template"
            layout = $templateConfig.layout
            desktops = $templateConfig.desktops
            onStart = $templateConfig.onStart
            onStop = $templateConfig.onStop
        }
    }
    else {
        $config = @{
            '$schema' = "../schemas/environment.schema.json"
            name = if ($Description) { $Description } else { $Name }
            description = "Custom environment"
            layout = "single-focus"
            desktops = @(
                @{
                    name = "Main"
                    windows = @(
                        @{
                            type = "terminal"
                            profile = "powershell"
                            zone = "main"
                            cwd = "."
                        }
                    )
                }
            )
            onStart = @()
            onStop = @()
        }
    }

    if ($WorkingDirectory) {
        $config.workingDirectory = $WorkingDirectory
    }

    if ($Description) {
        $config.description = $Description
    }

    Write-TeamerConfig -Path $path -Config $config
    Write-Host "Created environment: $Name" -ForegroundColor Green

    return $config
}

function Start-TeamerEnvironment {
    <#
    .SYNOPSIS
        Deploys an environment (creates desktops, launches windows)
    .DESCRIPTION
        Creates new virtual desktops for the environment and launches
        configured windows. This is ADDITIVE - new desktops are created
        alongside existing ones.
    .PARAMETER Name
        Environment name to start
    .EXAMPLE
        Start-TeamerEnvironment -Name "my-saas"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    # Check if already active
    if ($script:ActiveEnvironments.ContainsKey($Name)) {
        Write-Warning "Environment '$Name' is already active"
        return
    }

    $config = Get-TeamerEnvironment -Name $Name
    if (-not $config) {
        throw "Environment '$Name' not found"
    }

    Write-Host "Starting environment: $($config.name)" -ForegroundColor Cyan

    # Run onStart commands
    if ($config.onStart -and $config.onStart.Count -gt 0) {
        $workDir = if ($config.workingDirectory) { $config.workingDirectory } else { $PWD }
        Write-Host "Running startup commands in $workDir..." -ForegroundColor Yellow

        Push-Location $workDir
        try {
            foreach ($cmd in $config.onStart) {
                Write-Host "  > $cmd" -ForegroundColor DarkGray
                Invoke-Expression $cmd
            }
        }
        finally {
            Pop-Location
        }
    }

    # Track created desktops
    $createdDesktops = @()

    # Create desktops and launch windows
    $desktopCount = $config.desktops.Count
    $desktopIndex = 0
    foreach ($desktop in $config.desktops) {
        # Build desktop name: use environment name, append desktop name if multiple desktops
        $desktopName = if ($desktopCount -eq 1) {
            $Name  # Just use environment name for single-desktop environments
        }
        else {
            "$Name - $($desktop.name)"  # "my-project - Code", "my-project - Services"
        }

        Write-Host "Creating desktop: $desktopName" -ForegroundColor Yellow

        $result = New-TeamerDesktop -Name $desktopName
        $desktopIndex++
        if (-not $result.Success) {
            Write-Error "Failed to create desktop: $($desktop.name)"
            continue
        }

        $desktopIndex = $result.Index
        $createdDesktops += $desktopIndex

        # Switch to the new desktop
        Switch-TeamerDesktop -Index $desktopIndex | Out-Null
        Start-Sleep -Milliseconds 500

        # Launch windows on this desktop
        foreach ($window in $desktop.windows) {
            $workDir = if ($config.workingDirectory) { $config.workingDirectory } else { $PWD }

            # Resolve relative cwd
            if ($window.cwd -and $window.cwd -ne ".") {
                $workDir = if ([System.IO.Path]::IsPathRooted($window.cwd)) {
                    $window.cwd
                }
                else {
                    Join-Path $workDir $window.cwd
                }
            }

            switch ($window.type) {
                'terminal' {
                    Write-Host "  Launching terminal: $($window.profile)" -ForegroundColor DarkGray
                    Start-TeamerTerminalFromProfile -DesktopIndex $desktopIndex -ProfileName $window.profile -WorkingDirectory $workDir -Command $window.command -Title $window.title -TabColor $window.tabColor
                }
                'app' {
                    Write-Host "  Launching app: $($window.path)" -ForegroundColor DarkGray
                    $args = if ($window.args) { $window.args } else { @() }
                    Start-TeamerApp -DesktopIndex $desktopIndex -Path $window.path -Arguments $args -WorkingDirectory $workDir
                }
                'browser' {
                    Write-Host "  Opening browser: $($window.url)" -ForegroundColor DarkGray
                    Start-TeamerBrowser -Url $window.url
                }
            }

            Start-Sleep -Milliseconds 300
        }
    }

    # Track this environment as active and save state
    $script:ActiveEnvironments[$Name] = @{
        Config = $config
        DesktopIndices = $createdDesktops
        StartTime = Get-Date
    }
    Save-TeamerState

    Write-Host "Environment '$Name' started successfully" -ForegroundColor Green
    Write-Host "  Desktops created: $($createdDesktops -join ', ')" -ForegroundColor Cyan
}

function Stop-TeamerEnvironment {
    <#
    .SYNOPSIS
        Tears down an environment (removes desktops)
    .DESCRIPTION
        Removes all virtual desktops created by the environment.
        Runs any configured onStop commands first.
    .PARAMETER Name
        Environment name to stop
    .PARAMETER Force
        Skip confirmation prompt
    .EXAMPLE
        Stop-TeamerEnvironment -Name "my-saas"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [switch]$Force
    )

    if (-not $script:ActiveEnvironments.ContainsKey($Name)) {
        Write-Warning "Environment '$Name' is not active"
        return
    }

    $activeEnv = $script:ActiveEnvironments[$Name]
    $config = $activeEnv.Config

    if (-not $Force) {
        $desktopCount = $activeEnv.DesktopIndices.Count
        $confirm = Read-Host "Stop environment '$Name' and remove $desktopCount desktop(s)? (yes/no)"
        if ($confirm -ne 'yes') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    Write-Host "Stopping environment: $($config.name)" -ForegroundColor Cyan

    # Run onStop commands
    if ($config.onStop -and $config.onStop.Count -gt 0) {
        $workDir = if ($config.workingDirectory) { $config.workingDirectory } else { $PWD }
        Write-Host "Running shutdown commands..." -ForegroundColor Yellow

        Push-Location $workDir
        try {
            foreach ($cmd in $config.onStop) {
                Write-Host "  > $cmd" -ForegroundColor DarkGray
                Invoke-Expression $cmd
            }
        }
        finally {
            Pop-Location
        }
    }

    # Remove desktops in reverse order (highest index first)
    # Use underlying Remove-Desktop to bypass session protection
    # (these desktops were created by Teamer, so we trust they can be removed)
    $sortedDesktops = $activeEnv.DesktopIndices | Sort-Object -Descending
    foreach ($index in $sortedDesktops) {
        Write-Host "Removing desktop $($index + 1)..." -ForegroundColor Yellow
        try {
            $desktop = Get-Desktop -Index $index
            if ($desktop) {
                Remove-Desktop -Desktop $desktop
            }
        }
        catch {
            Write-Warning "Could not remove desktop $($index + 1): $_"
        }
    }

    # Remove from active tracking and save state
    $script:ActiveEnvironments.Remove($Name)
    Save-TeamerState

    Write-Host "Environment '$Name' stopped" -ForegroundColor Green
}

function Save-TeamerEnvironment {
    <#
    .SYNOPSIS
        Saves the current system state to an environment configuration
    .DESCRIPTION
        Captures the current desktops and windows created by the environment
        and saves them back to the configuration file.
    .PARAMETER Name
        Environment name to save
    .EXAMPLE
        Save-TeamerEnvironment -Name "my-saas"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    if (-not $script:ActiveEnvironments.ContainsKey($Name)) {
        Write-Warning "Environment '$Name' is not active. Cannot save."
        return
    }

    $activeEnv = $script:ActiveEnvironments[$Name]
    $path = Get-TeamerConfigPath -Type project -Name $Name
    $config = Get-TeamerEnvironment -Name $Name

    # Get current tree state
    $tree = Get-TeamerTree

    # Update desktops in config based on current state
    $updatedDesktops = @()

    foreach ($desktopIndex in $activeEnv.DesktopIndices) {
        # Find this desktop in the tree
        foreach ($screen in $tree) {
            $desktopNode = $screen.desktops | Where-Object { $_.index -eq $desktopIndex }
            if ($desktopNode) {
                $desktopConfig = @{
                    name = $desktopNode.name
                    windows = @()
                }

                foreach ($win in $desktopNode.windows) {
                    $windowConfig = @{
                        type = $win.type
                        zone = "main"  # Default zone
                    }

                    if ($win.type -eq 'terminal') {
                        # Map shell back to profile
                        $windowConfig.profile = switch ($win.shell) {
                            'powershell'      { 'powershell' }
                            'powershell-core' { 'pwsh' }
                            'wsl'             { 'wsl-ubuntu' }
                            'cmd'             { 'cmd' }
                            default           { 'powershell' }
                        }
                    }
                    else {
                        $windowConfig.path = $win.processName
                    }

                    $desktopConfig.windows += $windowConfig
                }

                $updatedDesktops += $desktopConfig
                break
            }
        }
    }

    # Update config
    $configHash = @{}
    $config.PSObject.Properties | ForEach-Object { $configHash[$_.Name] = $_.Value }
    $configHash.desktops = $updatedDesktops

    Write-TeamerConfig -Path $path -Config $configHash
    Write-Host "Saved environment: $Name" -ForegroundColor Green
}

function Remove-TeamerEnvironment {
    <#
    .SYNOPSIS
        Deletes an environment configuration
    .PARAMETER Name
        Environment name to remove
    .PARAMETER Force
        Skip confirmation prompt
    .EXAMPLE
        Remove-TeamerEnvironment -Name "old-project" -Force
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [switch]$Force
    )

    # Check if active
    if ($script:ActiveEnvironments.ContainsKey($Name)) {
        throw "Environment '$Name' is currently active. Stop it first with Stop-TeamerEnvironment."
    }

    $projectDir = Join-Path $script:PROJECTS_PATH $Name
    if (-not (Test-Path $projectDir)) {
        throw "Environment '$Name' not found"
    }

    if (-not $Force) {
        $confirm = Read-Host "Remove environment '$Name'? This cannot be undone. (yes/no)"
        if ($confirm -ne 'yes') {
            Write-Host "Cancelled." -ForegroundColor Yellow
            return
        }
    }

    Remove-Item -Path $projectDir -Recurse -Force
    Write-Host "Removed environment: $Name" -ForegroundColor Green
}

#endregion

#region Window Launch Helpers

function Start-TeamerTerminalFromProfile {
    <#
    .SYNOPSIS
        Launches a terminal using a profile configuration
    #>
    param(
        [Parameter(Mandatory)]
        [int]$DesktopIndex,

        [Parameter(Mandatory)]
        [string]$ProfileName,

        [Parameter()]
        [string]$WorkingDirectory,

        [Parameter()]
        [string]$Command,

        [Parameter()]
        [string]$Title,

        [Parameter()]
        [string]$TabColor
    )

    $profile = Get-TeamerProfile -Name $ProfileName
    if (-not $profile) {
        Write-Warning "Profile '$ProfileName' not found, using powershell"
        New-TeamerTerminal -DesktopIndex $DesktopIndex -Shell 'ps'
        return
    }

    # Use Windows Terminal with specific profile
    $terminalProfile = $profile.terminalProfile

    # Build Windows Terminal arguments
    $wtArgs = @("-p", "`"$terminalProfile`"")

    # Add working directory
    if ($WorkingDirectory -and (Test-Path $WorkingDirectory)) {
        $wtArgs += "-d"
        $wtArgs += "`"$WorkingDirectory`""
    }

    # Add title (window override > profile)
    $effectiveTitle = if ($Title) { $Title } elseif ($profile.title) { $profile.title } else { $null }
    if ($effectiveTitle) {
        $wtArgs += "--title"
        $wtArgs += "`"$effectiveTitle`""
    }

    # Add tab color (window override > profile)
    $effectiveColor = if ($TabColor) { $TabColor } elseif ($profile.tabColor) { $profile.tabColor } else { $null }
    if ($effectiveColor) {
        $wtArgs += "--tabColor"
        $wtArgs += "`"$effectiveColor`""
    }

    # Build the command to run
    $cmdParts = @()

    # Add startup commands from profile
    if ($profile.startupCommands -and $profile.startupCommands.Count -gt 0) {
        $cmdParts += $profile.startupCommands
    }

    # Add custom command
    if ($Command) {
        $cmdParts += $Command
    }

    # Launch terminal
    Start-Process -FilePath "wt.exe" -ArgumentList ($wtArgs -join " ")
}

function Start-TeamerApp {
    <#
    .SYNOPSIS
        Launches an application
    #>
    param(
        [Parameter(Mandatory)]
        [int]$DesktopIndex,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter()]
        [string[]]$Arguments = @(),

        [Parameter()]
        [string]$WorkingDirectory
    )

    $startParams = @{
        FilePath = $Path
        PassThru = $true
    }

    if ($Arguments.Count -gt 0) {
        $startParams.ArgumentList = $Arguments
    }

    if ($WorkingDirectory -and (Test-Path $WorkingDirectory)) {
        $startParams.WorkingDirectory = $WorkingDirectory
    }

    try {
        Start-Process @startParams | Out-Null
    }
    catch {
        Write-Warning "Failed to launch $Path : $_"
    }
}

function Start-TeamerBrowser {
    <#
    .SYNOPSIS
        Opens a URL in the default browser
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Url
    )

    Start-Process $Url
}

#endregion

#region State Loading (deferred)

# Load saved state now that all functions are defined
Load-TeamerState

#endregion

#region Display Welcome

Write-Host ""
Write-Host "Teamer Environment Manager Loaded" -ForegroundColor Green
Write-Host ""
Write-Host "Environment Commands:" -ForegroundColor Cyan
Write-Host "  Get-TeamerEnvironment       - List or get environment(s)" -ForegroundColor White
Write-Host "  New-TeamerEnvironment       - Create new environment" -ForegroundColor White
Write-Host "  Start-TeamerEnvironment     - Deploy environment (ADDITIVE)" -ForegroundColor White
Write-Host "  Stop-TeamerEnvironment      - Tear down environment" -ForegroundColor White
Write-Host "  Save-TeamerEnvironment      - Snapshot current state" -ForegroundColor White
Write-Host "  Remove-TeamerEnvironment    - Delete environment config" -ForegroundColor White
Write-Host ""
Write-Host "Profile Commands:" -ForegroundColor Cyan
Write-Host "  Get-TeamerProfile           - List or get profile(s)" -ForegroundColor White
Write-Host "  New-TeamerProfile           - Create shell profile" -ForegroundColor White
Write-Host "  Set-TeamerProfile           - Update profile property" -ForegroundColor White
Write-Host "  Remove-TeamerProfile        - Delete profile" -ForegroundColor White
Write-Host ""
Write-Host "Template Commands:" -ForegroundColor Cyan
Write-Host "  Get-TeamerTemplate          - List or get template(s)" -ForegroundColor White
Write-Host "  New-TeamerTemplate          - Create template" -ForegroundColor White
Write-Host ""
Write-Host "Layout Commands:" -ForegroundColor Cyan
Write-Host "  Get-TeamerLayout            - List or get layout(s)" -ForegroundColor White
Write-Host "  Get-TeamerZone              - Resolve zone from layout" -ForegroundColor White
Write-Host ""

#endregion
