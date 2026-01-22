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

# Dynamically resolve project root from script location for portability
$script:PROJECT_ROOT = Split-Path -Parent $PSScriptRoot
$script:ENVIRONMENTS_ROOT = Join-Path $script:PROJECT_ROOT "environments"
$script:TEMPLATES_PATH = Join-Path $script:ENVIRONMENTS_ROOT "templates"
$script:PROFILES_PATH = Join-Path $script:ENVIRONMENTS_ROOT "profiles"
$script:LAYOUTS_PATH = Join-Path $script:ENVIRONMENTS_ROOT "layouts"
$script:PROJECTS_PATH = Join-Path $script:ENVIRONMENTS_ROOT "projects"
$script:SERVICES_PATH = Join-Path $script:ENVIRONMENTS_ROOT "services"
$script:STATE_FILE = Join-Path $script:ENVIRONMENTS_ROOT "state.json"
$script:LOG_FILE = Join-Path $script:ENVIRONMENTS_ROOT "teamer.log"

# Track active environments and their desktop indices (loaded from state file)
$script:ActiveEnvironments = @{}

# Dangerous command patterns to block in lifecycle commands
$script:BLOCKED_COMMAND_PATTERNS = @(
    'Remove-Item.*-Recurse.*-Force',          # Recursive force delete
    'rm\s+-rf\s+/',                           # Unix-style dangerous rm
    'Format-Volume',                          # Disk formatting
    'Clear-Disk',                             # Disk clearing
    'Stop-Computer',                          # Shutdown
    'Restart-Computer',                       # Restart
    'Remove-PSDrive',                         # Remove drives
    '>\s*\$null\s*2>&1.*Remove',             # Hidden destructive commands
    'Invoke-WebRequest.*\|\s*Invoke-Expression', # Download and execute
    'iex\s*\(.*Net\.WebClient',              # Download and execute variant
    'Start-Process.*-Verb\s+RunAs'           # Elevation attempts
)

#endregion

#region Logging

function Write-TeamerLog {
    <#
    .SYNOPSIS
        Writes a message to the Teamer log file
    .PARAMETER Message
        The message to log
    .PARAMETER Level
        Log level: Info, Warning, Error
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    try {
        $logEntry | Out-File -FilePath $script:LOG_FILE -Append -Encoding UTF8
    }
    catch {
        # Silently fail if logging fails - don't interrupt main operations
    }
}

function Get-TeamerLog {
    <#
    .SYNOPSIS
        Gets recent log entries
    .PARAMETER Lines
        Number of lines to retrieve (default: 50)
    #>
    param(
        [Parameter()]
        [int]$Lines = 50
    )

    if (Test-Path $script:LOG_FILE) {
        Get-Content -Path $script:LOG_FILE -Tail $Lines
    }
    else {
        Write-Host "No log file found at $script:LOG_FILE" -ForegroundColor Yellow
    }
}

function Clear-TeamerLog {
    <#
    .SYNOPSIS
        Clears the Teamer log file
    #>
    if (Test-Path $script:LOG_FILE) {
        Remove-Item -Path $script:LOG_FILE -Force
        Write-Host "Log file cleared." -ForegroundColor Green
    }
}

#endregion

#region Command Sanitization

function Test-CommandSafe {
    <#
    .SYNOPSIS
        Validates a command against blocked patterns for security
    .PARAMETER Command
        The command string to validate
    .OUTPUTS
        PSCustomObject with IsValid and Reason properties
    #>
    param(
        [Parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [string]$Command
    )

    # Check for empty commands
    if ([string]::IsNullOrWhiteSpace($Command)) {
        return [PSCustomObject]@{ IsValid = $false; Reason = "Empty command" }
    }

    # Check against blocked patterns
    foreach ($pattern in $script:BLOCKED_COMMAND_PATTERNS) {
        if ($Command -match $pattern) {
            return [PSCustomObject]@{
                IsValid = $false
                Reason = "Command matches blocked pattern: $pattern"
            }
        }
    }

    return [PSCustomObject]@{ IsValid = $true; Reason = $null }
}

function Invoke-TeamerCommand {
    <#
    .SYNOPSIS
        Safely executes a lifecycle command with validation and logging
    .PARAMETER Command
        The command to execute
    .PARAMETER WorkingDirectory
        Working directory for the command
    .PARAMETER Phase
        The lifecycle phase (onStart/onStop) for logging
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Command,

        [Parameter()]
        [string]$WorkingDirectory,

        [Parameter()]
        [ValidateSet('onStart', 'onStop')]
        [string]$Phase = 'onStart'
    )

    # Validate command
    $validation = Test-CommandSafe -Command $Command
    if (-not $validation.IsValid) {
        $errorMsg = "BLOCKED: $($validation.Reason) - Command: $Command"
        Write-TeamerLog -Message $errorMsg -Level Error
        Write-Warning $errorMsg
        return @{ Success = $false; Error = $validation.Reason }
    }

    Write-TeamerLog -Message "[$Phase] Executing: $Command" -Level Info

    try {
        if ($WorkingDirectory -and (Test-Path $WorkingDirectory)) {
            Push-Location $WorkingDirectory
        }

        # Use Start-Process for better isolation instead of Invoke-Expression
        # For simple commands, we still use Invoke-Expression but with validation
        $result = Invoke-Expression $Command 2>&1

        Write-TeamerLog -Message "[$Phase] Completed: $Command" -Level Info

        return @{ Success = $true; Output = $result }
    }
    catch {
        $errorMsg = "[$Phase] Failed: $Command - Error: $_"
        Write-TeamerLog -Message $errorMsg -Level Error
        Write-Warning $errorMsg
        return @{ Success = $false; Error = $_.Exception.Message }
    }
    finally {
        if ($WorkingDirectory -and (Test-Path $WorkingDirectory)) {
            Pop-Location
        }
    }
}

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
    Write-TeamerLog -Message "Starting environment: $Name" -Level Info

    # Run onStart commands with validation
    if ($config.onStart -and $config.onStart.Count -gt 0) {
        $workDir = if ($config.workingDirectory) { $config.workingDirectory } else { $PWD }
        Write-Host "Running startup commands in $workDir..." -ForegroundColor Yellow

        foreach ($cmd in $config.onStart) {
            Write-Host "  > $cmd" -ForegroundColor DarkGray
            $cmdResult = Invoke-TeamerCommand -Command $cmd -WorkingDirectory $workDir -Phase 'onStart'
            if (-not $cmdResult.Success) {
                Write-Warning "Startup command failed: $cmd"
            }
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

        # IMPORTANT: Switch to target desktop FIRST, then launch windows
        # This ensures windows open on the correct desktop
        Write-Host "  Switching to desktop $($desktopIndex + 1)..." -ForegroundColor DarkGray
        Switch-TeamerDesktop -Index $desktopIndex | Out-Null
        Start-Sleep -Milliseconds 500

        # Verify we're on the correct desktop before proceeding
        $currentDesktop = Get-CurrentDesktop
        $currentIndex = Get-DesktopIndex -Desktop $currentDesktop
        $targetDesktop = Get-Desktop -Index $desktopIndex
        if ($currentIndex -ne $desktopIndex) {
            Write-Warning "Failed to switch to desktop $desktopIndex, retrying..."
            Switch-TeamerDesktop -Index $desktopIndex | Out-Null
            Start-Sleep -Milliseconds 500
        }

        Write-Host "  Launching windows on desktop $($desktopIndex + 1)..." -ForegroundColor DarkGray

        # Load layout for grid positioning
        $layout = $null
        $grid = $null
        if ($config.layout) {
            $layout = Get-TeamerLayout -Name $config.layout
            if ($layout -and $layout.grid) {
                $grid = $layout.grid
            }
        }

        # Group terminals by tabGroup
        $terminalGroups = @{}
        $ungroupedWindows = @()

        foreach ($window in $desktop.windows) {
            if ($window.type -eq 'terminal' -and $window.tabGroup) {
                if (-not $terminalGroups.ContainsKey($window.tabGroup)) {
                    $terminalGroups[$window.tabGroup] = @()
                }
                $terminalGroups[$window.tabGroup] += $window
            }
            else {
                $ungroupedWindows += $window
            }
        }

        $baseWorkDir = if ($config.workingDirectory) { $config.workingDirectory } else { $PWD }

        # Collect window handles as we launch them
        $launchedWindows = @()

        # Launch terminal tab groups
        foreach ($groupName in $terminalGroups.Keys) {
            $terminals = $terminalGroups[$groupName]
            Write-Host "    Terminal group '$groupName' ($($terminals.Count) tabs)" -ForegroundColor DarkGray
            Start-TeamerTerminalTabGroup -DesktopIndex $desktopIndex -Terminals $terminals -BaseWorkingDirectory $baseWorkDir
            Start-Sleep -Milliseconds 1000
        }

        # Launch ungrouped windows
        foreach ($window in $ungroupedWindows) {
            $workDir = $baseWorkDir

            # Resolve relative cwd
            if ($window.cwd -and $window.cwd -ne ".") {
                $workDir = if ([System.IO.Path]::IsPathRooted($window.cwd)) {
                    $window.cwd
                }
                else {
                    Join-Path $baseWorkDir $window.cwd
                }
            }

            switch ($window.type) {
                'terminal' {
                    Write-Host "    Terminal: $($window.profile)" -ForegroundColor DarkGray
                    Start-TeamerTerminalFromProfile -DesktopIndex $desktopIndex -ProfileName $window.profile -WorkingDirectory $workDir -Command $window.command -Title $window.title -TabColor $window.tabColor
                }
                'app' {
                    Write-Host "    App: $($window.path)" -ForegroundColor DarkGray
                    $appArgs = if ($window.args) { $window.args } else { @() }
                    Start-TeamerApp -DesktopIndex $desktopIndex -Path $window.path -Arguments $appArgs -WorkingDirectory $workDir
                }
                'browser' {
                    Write-Host "    Browser: $($window.url)" -ForegroundColor DarkGray
                    Start-TeamerBrowser -Url $window.url
                }
            }

            Start-Sleep -Milliseconds 500
        }

        # Wait for all windows to fully initialize
        Start-Sleep -Milliseconds 1500

        # Apply tiling if grid is defined for this desktop
        if ($desktop.grid) {
            Write-Host "  Applying tiling ($($desktop.grid.rows)x$($desktop.grid.cols) grid)..." -ForegroundColor DarkGray
            Apply-TeamerDesktopTiling -Desktop $desktop -BaseWorkingDirectory $baseWorkDir -DesktopIndex $desktopIndex
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
    Write-TeamerLog -Message "Stopping environment: $Name" -Level Info

    # Run onStop commands with validation
    if ($config.onStop -and $config.onStop.Count -gt 0) {
        $workDir = if ($config.workingDirectory) { $config.workingDirectory } else { $PWD }
        Write-Host "Running shutdown commands..." -ForegroundColor Yellow

        foreach ($cmd in $config.onStop) {
            Write-Host "  > $cmd" -ForegroundColor DarkGray
            $cmdResult = Invoke-TeamerCommand -Command $cmd -WorkingDirectory $workDir -Phase 'onStop'
            if (-not $cmdResult.Success) {
                Write-Warning "Shutdown command failed: $cmd"
            }
        }
    }

    # Remove desktops by NAME (more reliable than index which can shift)
    # Build desktop names using the same logic as Start-TeamerEnvironment
    $desktopCount = $config.desktops.Count
    $desktopNames = @()
    foreach ($desktopConfig in $config.desktops) {
        $desktopName = if ($desktopCount -eq 1) {
            $Name  # Environment name for single-desktop environments
        }
        else {
            "$Name - $($desktopConfig.name)"  # "my-project - Code" for multi-desktop
        }
        $desktopNames += $desktopName
    }

    # Find and remove each desktop by name
    # IMPORTANT: Close windows BEFORE removing desktop to prevent them from moving to active desktop
    $allDesktops = Get-DesktopList
    foreach ($desktopName in $desktopNames) {
        $desktop = $allDesktops | Where-Object { $_.Name -eq $desktopName }
        if ($desktop) {
            $desktopIndex = $desktop.Number

            # Switch to target desktop first
            Write-Host "Switching to desktop '$desktopName' to close windows..." -ForegroundColor DarkGray
            try {
                Switch-TeamerDesktop -Index $desktopIndex | Out-Null
                Start-Sleep -Milliseconds 300
            }
            catch {
                Write-Warning "Could not switch to desktop: $_"
            }

            # Get all windows on this desktop and close them
            Write-Host "Closing windows on '$desktopName'..." -ForegroundColor Yellow
            $targetDesktop = Get-Desktop -Index $desktopIndex
            $windowsOnDesktop = Get-Process | Where-Object {
                $_.MainWindowHandle -ne [IntPtr]::Zero
            } | ForEach-Object {
                $proc = $_
                try {
                    $windowDesktop = Get-DesktopFromWindow -Hwnd $proc.MainWindowHandle
                    if ($windowDesktop -and (Get-DesktopIndex -Desktop $windowDesktop) -eq $desktopIndex) {
                        $proc
                    }
                }
                catch {
                    # Window might not be on any desktop or already closed
                }
            }

            foreach ($proc in $windowsOnDesktop) {
                if ($proc) {
                    Write-Host "  Closing: $($proc.ProcessName)" -ForegroundColor DarkGray
                    try {
                        $proc.CloseMainWindow() | Out-Null
                        # Give window time to close gracefully
                        Start-Sleep -Milliseconds 200
                    }
                    catch {
                        Write-Warning "Could not close $($proc.ProcessName): $_"
                    }
                }
            }

            # Small delay to ensure windows are closed
            Start-Sleep -Milliseconds 500

            # Now remove the desktop
            Write-Host "Removing desktop '$desktopName'..." -ForegroundColor Yellow
            try {
                Remove-Desktop -Desktop (Get-Desktop -Index $desktopIndex)
            }
            catch {
                Write-Warning "Could not remove desktop '$desktopName': $_"
            }
        }
        else {
            Write-Host "Desktop '$desktopName' not found (may already be removed)" -ForegroundColor DarkGray
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

function Convert-ToWslPath {
    <#
    .SYNOPSIS
        Converts a Windows path to a WSL mount path
    .EXAMPLE
        Convert-ToWslPath "E:\Projects\teamer"
        Returns: /mnt/e/Projects/teamer
    #>
    param(
        [Parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [string]$WindowsPath
    )

    if ([string]::IsNullOrWhiteSpace($WindowsPath)) {
        return $null
    }

    # Extract drive letter and convert to lowercase
    if ($WindowsPath -match '^([A-Za-z]):(.*)$') {
        $driveLetter = $Matches[1].ToLower()
        $restOfPath = $Matches[2] -replace '\\', '/'
        return "/mnt/$driveLetter$restOfPath"
    }

    # If no drive letter, just convert backslashes
    return $WindowsPath -replace '\\', '/'
}

function Start-TeamerTerminalFromProfile {
    <#
    .SYNOPSIS
        Launches a terminal using a profile configuration
    .DESCRIPTION
        Supports different shell types: powershell, pwsh, wsl, cmd, git-bash
        For WSL, properly invokes wsl.exe with the specified distribution.
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

    # Build Windows Terminal arguments
    $wtArgs = @()

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

    # Handle different shell types
    $shell = $profile.shell
    switch ($shell) {
        'wsl' {
            # For WSL, we need to use wsl.exe with distribution
            $distribution = if ($profile.distribution) { $profile.distribution } else { "Ubuntu" }

            # Convert Windows path to WSL path if needed
            $wslWorkDir = $null
            if ($WorkingDirectory -and (Test-Path $WorkingDirectory)) {
                $wslWorkDir = Convert-ToWslPath $WorkingDirectory
            }

            # Build the WSL command
            $wslCmd = "wsl.exe -d $distribution"

            # Add startup commands and custom command
            $cmdParts = @()
            if ($wslWorkDir) {
                $cmdParts += "cd '$wslWorkDir'"
            }
            if ($profile.startupCommands -and $profile.startupCommands.Count -gt 0) {
                $cmdParts += $profile.startupCommands
            }
            if ($Command) {
                $cmdParts += $Command
            }

            # If there are commands, wrap them in bash -c, otherwise just launch WSL
            if ($cmdParts.Count -gt 0) {
                $bashCmd = $cmdParts -join ' && '
                # Escape for Windows command line
                $bashCmd = $bashCmd -replace '"', '\"'
                $wslCmd += " -- bash -c `"$bashCmd && exec bash`""
            }

            $wtArgs += $wslCmd
        }
        'pwsh' {
            # PowerShell Core
            $wtArgs += "-p"
            $wtArgs += "`"PowerShell`""

            if ($WorkingDirectory -and (Test-Path $WorkingDirectory)) {
                $wtArgs += "-d"
                $wtArgs += "`"$WorkingDirectory`""
            }
        }
        'cmd' {
            # Command Prompt
            $wtArgs += "-p"
            $wtArgs += "`"Command Prompt`""

            if ($WorkingDirectory -and (Test-Path $WorkingDirectory)) {
                $wtArgs += "-d"
                $wtArgs += "`"$WorkingDirectory`""
            }
        }
        'git-bash' {
            # Git Bash
            $wtArgs += "-p"
            $wtArgs += "`"Git Bash`""

            if ($WorkingDirectory -and (Test-Path $WorkingDirectory)) {
                $wtArgs += "-d"
                $wtArgs += "`"$WorkingDirectory`""
            }
        }
        default {
            # PowerShell (Windows PowerShell)
            $wtArgs += "-p"
            $wtArgs += "`"Windows PowerShell`""

            if ($WorkingDirectory -and (Test-Path $WorkingDirectory)) {
                $wtArgs += "-d"
                $wtArgs += "`"$WorkingDirectory`""
            }
        }
    }

    # Launch terminal
    $argString = $wtArgs -join " "
    Write-Host "    wt.exe $argString" -ForegroundColor DarkGray
    Start-Process -FilePath "wt.exe" -ArgumentList $argString
}

function Start-TeamerApp {
    <#
    .SYNOPSIS
        Launches an application and moves it to the target desktop.
        If the app is already running, reuses the existing window instead of launching a new instance.
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

    $appName = [System.IO.Path]::GetFileNameWithoutExtension($Path)

    # Check if app is already running with a window
    $existingProc = Get-Process -ErrorAction SilentlyContinue | Where-Object {
        $_.ProcessName -match "^$appName$" -and $_.MainWindowHandle -ne [IntPtr]::Zero
    } | Select-Object -First 1

    if ($existingProc) {
        # App already running - just move it to target desktop
        Write-Host "    $appName already running, moving to desktop $($DesktopIndex + 1)" -ForegroundColor DarkGray
        try {
            $targetDesktop = Get-Desktop -Index $DesktopIndex
            if ($targetDesktop) {
                Move-Window -Desktop $targetDesktop -Hwnd $existingProc.MainWindowHandle
            }
        }
        catch {
            Write-Warning "Could not move existing $appName to desktop $DesktopIndex : $_"
        }
        return
    }

    # App not running - launch it
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
        $process = Start-Process @startParams

        # Wait for the main window to appear
        $maxWait = 15  # seconds
        $waited = 0
        $hwnd = $null

        while ($waited -lt $maxWait) {
            Start-Sleep -Milliseconds 500
            $waited += 0.5

            # Method 1: Direct process MainWindowHandle
            $process.Refresh()
            if ($process.MainWindowHandle -ne [IntPtr]::Zero) {
                $hwnd = $process.MainWindowHandle
                break
            }

            # Method 2: Find by process name (apps like Excel spawn new processes)
            $procs = Get-Process -ErrorAction SilentlyContinue | Where-Object {
                $_.ProcessName -match "^$appName$" -and $_.MainWindowHandle -ne [IntPtr]::Zero
            }
            if ($procs) {
                $newest = $procs | Sort-Object StartTime -Descending | Select-Object -First 1
                if ($newest) {
                    $hwnd = $newest.MainWindowHandle
                    break
                }
            }
        }

        if ($hwnd -and $hwnd -ne [IntPtr]::Zero) {
            # Move window to target desktop
            try {
                $targetDesktop = Get-Desktop -Index $DesktopIndex
                if ($targetDesktop) {
                    Move-Window -Desktop $targetDesktop -Hwnd $hwnd
                    Write-Host "    Moved $appName to desktop $($DesktopIndex + 1)" -ForegroundColor DarkGray
                }
            }
            catch {
                Write-Warning "Could not move $Path to desktop $DesktopIndex : $_"
            }
        }
        else {
            Write-Warning "Could not find main window for $Path within $maxWait seconds"
        }
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

function Start-TeamerTerminalTabGroup {
    <#
    .SYNOPSIS
        Launches multiple terminals as tabs in a single Windows Terminal window
    .DESCRIPTION
        Takes an array of terminal configurations and launches them all in one window.
        Each terminal becomes a tab. Supports different profiles, titles, and colors per tab.
    #>
    param(
        [Parameter(Mandatory)]
        [int]$DesktopIndex,

        [Parameter(Mandatory)]
        [array]$Terminals,

        [Parameter()]
        [string]$BaseWorkingDirectory
    )

    if ($Terminals.Count -eq 0) {
        return
    }

    # Build Windows Terminal command with multiple tabs
    # Format: wt [options] ; new-tab [options] ; new-tab [options]
    $wtParts = @()

    $isFirst = $true
    foreach ($term in $Terminals) {
        $tabArgs = @()

        # Get the profile config
        $profileConfig = Get-TeamerProfile -Name $term.profile
        $shell = if ($profileConfig) { $profileConfig.shell } else { "powershell" }

        # Resolve working directory
        $workDir = $BaseWorkingDirectory
        if ($term.cwd -and $term.cwd -ne ".") {
            $workDir = if ([System.IO.Path]::IsPathRooted($term.cwd)) {
                $term.cwd
            }
            else {
                Join-Path $BaseWorkingDirectory $term.cwd
            }
        }

        # Add title
        $effectiveTitle = if ($term.title) { $term.title } elseif ($profileConfig -and $profileConfig.title) { $profileConfig.title } else { $null }
        if ($effectiveTitle) {
            $tabArgs += "--title"
            $tabArgs += "`"$effectiveTitle`""
        }

        # Add tab color
        $effectiveColor = if ($term.tabColor) { $term.tabColor } elseif ($profileConfig -and $profileConfig.tabColor) { $profileConfig.tabColor } else { $null }
        if ($effectiveColor) {
            $tabArgs += "--tabColor"
            $tabArgs += "`"$effectiveColor`""
        }

        # Handle shell type
        switch ($shell) {
            'wsl' {
                $distribution = if ($profileConfig -and $profileConfig.distribution) { $profileConfig.distribution } else { "Ubuntu" }

                # Convert Windows path to WSL path
                $wslWorkDir = $null
                if ($workDir -and (Test-Path $workDir)) {
                    $wslWorkDir = Convert-ToWslPath $workDir
                }

                # Use Windows Terminal's WSL profile directly for cleaner tab handling
                $termProfile = if ($profileConfig -and $profileConfig.terminalProfile) { $profileConfig.terminalProfile } else { "Ubuntu" }
                $tabArgs += "-p"
                $tabArgs += "`"$termProfile`""

                # Use --startingDirectory for WSL path (Windows Terminal handles conversion)
                if ($workDir -and (Test-Path $workDir)) {
                    $tabArgs += "-d"
                    $tabArgs += "`"$workDir`""
                }
            }
            'pwsh' {
                $tabArgs += "-p"
                $tabArgs += "`"PowerShell`""
                if ($workDir -and (Test-Path $workDir)) {
                    $tabArgs += "-d"
                    $tabArgs += "`"$workDir`""
                }
            }
            'cmd' {
                $tabArgs += "-p"
                $tabArgs += "`"Command Prompt`""
                if ($workDir -and (Test-Path $workDir)) {
                    $tabArgs += "-d"
                    $tabArgs += "`"$workDir`""
                }
            }
            default {
                $tabArgs += "-p"
                $tabArgs += "`"Windows PowerShell`""
                if ($workDir -and (Test-Path $workDir)) {
                    $tabArgs += "-d"
                    $tabArgs += "`"$workDir`""
                }
            }
        }

        if ($isFirst) {
            $wtParts += ($tabArgs -join " ")
            $isFirst = $false
        }
        else {
            $wtParts += "`; new-tab " + ($tabArgs -join " ")
        }
    }

    $argString = $wtParts -join " "
    Write-Host "    wt.exe $argString" -ForegroundColor DarkGray
    Start-Process -FilePath "wt.exe" -ArgumentList $argString
}

#endregion

#region Grid and Window Positioning

# Win32 APIs are loaded from shared TeamerWin32.ps1 module via Manage-Teamer.ps1
# The Get-TeamerWindowFrameOffset function uses DWM API with fallback to hardcoded values

function Get-TeamerScreenBounds {
    <#
    .SYNOPSIS
        Gets the bounds of the primary screen
    #>
    Add-Type -AssemblyName System.Windows.Forms
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen
    return @{
        X = $screen.WorkingArea.X
        Y = $screen.WorkingArea.Y
        Width = $screen.WorkingArea.Width
        Height = $screen.WorkingArea.Height
    }
}

function Get-TeamerGridCellBounds {
    <#
    .SYNOPSIS
        Calculates the pixel bounds for a grid cell position.
        Ensures consistent gaps by calculating exact pixel positions.
    .PARAMETER Grid
        Grid configuration with rows, cols, gap, margin
    .PARAMETER Row
        Starting row (0-indexed)
    .PARAMETER Col
        Starting column (0-indexed)
    .PARAMETER RowSpan
        Number of rows to span
    .PARAMETER ColSpan
        Number of columns to span
    #>
    param(
        [Parameter(Mandatory)]
        [object]$Grid,

        [Parameter(Mandatory)]
        [int]$Row,

        [Parameter(Mandatory)]
        [int]$Col,

        [Parameter()]
        [int]$RowSpan = 1,

        [Parameter()]
        [int]$ColSpan = 1
    )

    $screen = Get-TeamerScreenBounds

    $rows = $Grid.rows
    $cols = $Grid.cols
    $gap = if ($null -ne $Grid.gap) { $Grid.gap } else { 2 }
    $margin = if ($null -ne $Grid.margin) { $Grid.margin } else { 2 }

    # Calculate total space taken by gaps and margins
    $totalGapsX = ($cols - 1) * $gap
    $totalGapsY = ($rows - 1) * $gap
    $totalMarginsX = 2 * $margin
    $totalMarginsY = 2 * $margin

    # Available space for cells
    $availableWidth = $screen.Width - $totalMarginsX - $totalGapsX
    $availableHeight = $screen.Height - $totalMarginsY - $totalGapsY

    # Base cell size (may have remainder)
    $baseCellWidth = [math]::Floor($availableWidth / $cols)
    $baseCellHeight = [math]::Floor($availableHeight / $rows)

    # Distribute remainder pixels to last cells
    $remainderX = $availableWidth - ($baseCellWidth * $cols)
    $remainderY = $availableHeight - ($baseCellHeight * $rows)

    # Calculate X position: margin + (cells before * cellWidth) + (gaps before * gap) + extra pixels for earlier cells
    $x = $screen.X + $margin
    for ($c = 0; $c -lt $Col; $c++) {
        $x += $baseCellWidth + $gap
        if ($c -ge ($cols - $remainderX)) { $x += 1 }  # Extra pixel for last cells
    }

    # Calculate Y position
    $y = $screen.Y + $margin
    for ($r = 0; $r -lt $Row; $r++) {
        $y += $baseCellHeight + $gap
        if ($r -ge ($rows - $remainderY)) { $y += 1 }  # Extra pixel for last cells
    }

    # Calculate width spanning multiple columns
    $width = 0
    for ($c = $Col; $c -lt ($Col + $ColSpan); $c++) {
        $width += $baseCellWidth
        if ($c -ge ($cols - $remainderX)) { $width += 1 }  # Extra pixel for last cells
        if ($c -lt ($Col + $ColSpan - 1)) { $width += $gap }  # Add gap between spanned cells
    }

    # Calculate height spanning multiple rows
    $height = 0
    for ($r = $Row; $r -lt ($Row + $RowSpan); $r++) {
        $height += $baseCellHeight
        if ($r -ge ($rows - $remainderY)) { $height += 1 }  # Extra pixel for last cells
        if ($r -lt ($Row + $RowSpan - 1)) { $height += $gap }  # Add gap between spanned rows
    }

    return @{
        X = $x
        Y = $y
        Width = $width
        Height = $height
    }
}

function Move-TeamerWindow {
    <#
    .SYNOPSIS
        Moves and resizes a window by its process handle
    #>
    param(
        [Parameter(Mandatory)]
        [System.Diagnostics.Process]$Process,

        [Parameter(Mandatory)]
        [int]$X,

        [Parameter(Mandatory)]
        [int]$Y,

        [Parameter(Mandatory)]
        [int]$Width,

        [Parameter(Mandatory)]
        [int]$Height
    )

    # Wait for window handle
    $attempts = 0
    while ($Process.MainWindowHandle -eq [IntPtr]::Zero -and $attempts -lt 20) {
        Start-Sleep -Milliseconds 100
        $Process.Refresh()
        $attempts++
    }

    if ($Process.MainWindowHandle -ne [IntPtr]::Zero) {
        # Get invisible border offset to compensate
        $frameOffset = Get-TeamerWindowFrameOffset -Handle $Process.MainWindowHandle
        $adjustedX = $X - $frameOffset.Left
        $adjustedY = $Y - $frameOffset.Top
        $adjustedWidth = $Width + $frameOffset.Left + $frameOffset.Right
        $adjustedHeight = $Height + $frameOffset.Top + $frameOffset.Bottom

        [TeamerWin32]::MoveWindow($Process.MainWindowHandle, $adjustedX, $adjustedY, $adjustedWidth, $adjustedHeight, $true) | Out-Null
        return $true
    }

    return $false
}

#endregion

#region Tiling Functions

$script:TILING_FILE = Join-Path $script:ENVIRONMENTS_ROOT "tiling.json"

function Get-TeamerTiling {
    <#
    .SYNOPSIS
        Gets tiling configuration for a desktop
    .PARAMETER DesktopName
        Name of the desktop to get tiling for
    #>
    param(
        [Parameter()]
        [string]$DesktopName
    )

    if (-not (Test-Path $script:TILING_FILE)) {
        return $null
    }

    $content = Get-Content -Path $script:TILING_FILE -Raw -Encoding UTF8
    $tiling = $content | ConvertFrom-Json

    if ([string]::IsNullOrWhiteSpace($DesktopName)) {
        return $tiling.desktops
    }

    return $tiling.desktops.$DesktopName
}

function Set-TeamerTiling {
    <#
    .SYNOPSIS
        Sets tiling configuration for a desktop
    .PARAMETER DesktopName
        Name of the desktop
    .PARAMETER Grid
        Grid configuration hashtable with rows, cols, gap, margin
    .PARAMETER Windows
        Array of window position configurations
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DesktopName,

        [Parameter(Mandatory)]
        [hashtable]$Grid,

        [Parameter(Mandatory)]
        [array]$Windows
    )

    # Load existing or create new
    $tiling = @{ desktops = @{} }
    if (Test-Path $script:TILING_FILE) {
        $content = Get-Content -Path $script:TILING_FILE -Raw -Encoding UTF8
        $existing = $content | ConvertFrom-Json
        # Convert to hashtable
        foreach ($prop in $existing.desktops.PSObject.Properties) {
            $tiling.desktops[$prop.Name] = $prop.Value
        }
    }

    # Set the desktop tiling
    $tiling.desktops[$DesktopName] = @{
        grid = $Grid
        windows = $Windows
    }

    # Save
    $tiling.'$schema' = "schemas/tiling.schema.json"
    $json = $tiling | ConvertTo-Json -Depth 10
    $json | Out-File -FilePath $script:TILING_FILE -Encoding UTF8 -Force

    Write-Host "Saved tiling for desktop: $DesktopName" -ForegroundColor Green
}

function Get-DesktopWindows {
    <#
    .SYNOPSIS
        Gets all visible windows on the current desktop only
    #>
    $currentDesktop = Get-CurrentDesktop
    $windows = @()

    Get-Process | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | ForEach-Object {
        # Check if this window is on the current desktop
        try {
            $windowDesktop = Get-DesktopFromWindow -Hwnd $_.MainWindowHandle
            if ($windowDesktop -and (Test-CurrentDesktop -Desktop $windowDesktop)) {
                $windows += @{
                    Process = $_
                    Handle = $_.MainWindowHandle
                    Title = $_.MainWindowTitle
                    ProcessName = $_.ProcessName
                }
            }
        }
        catch {
            # Window might not be on any desktop (pinned or special)
        }
    }
    return $windows
}

function Apply-TeamerTiling {
    <#
    .SYNOPSIS
        Applies tiling configuration to windows on current desktop
    .PARAMETER DesktopName
        Name of the desktop configuration to apply
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DesktopName
    )

    $config = Get-TeamerTiling -DesktopName $DesktopName
    if (-not $config) {
        Write-Warning "No tiling configuration found for desktop: $DesktopName"
        return
    }

    $grid = $config.grid
    $windowConfigs = $config.windows

    Write-Host "Applying tiling for: $DesktopName ($($grid.rows)x$($grid.cols) grid)" -ForegroundColor Cyan

    # Get all windows on current desktop
    $desktopWindows = Get-DesktopWindows

    foreach ($winConfig in $windowConfigs) {
        $match = $winConfig.match
        $matchedWindow = $null

        foreach ($win in $desktopWindows) {
            $matched = $false

            if ($match.process) {
                if ($win.ProcessName -match $match.process) {
                    $matched = $true
                }
            }
            if ($match.title) {
                if ($win.Title -match $match.title) {
                    $matched = $true
                }
            }

            if ($matched) {
                $matchedWindow = $win
                break
            }
        }

        if ($matchedWindow) {
            $row = $winConfig.row
            $col = $winConfig.col
            $rowSpan = if ($winConfig.rowSpan) { $winConfig.rowSpan } else { 1 }
            $colSpan = if ($winConfig.colSpan) { $winConfig.colSpan } else { 1 }

            $bounds = Get-TeamerGridCellBounds -Grid $grid -Row $row -Col $col -RowSpan $rowSpan -ColSpan $colSpan

            Write-Host "  Moving $($matchedWindow.ProcessName) to ($row,$col) span($rowSpan,$colSpan)" -ForegroundColor DarkGray

            [TeamerWin32]::MoveWindow(
                $matchedWindow.Handle,
                $bounds.X,
                $bounds.Y,
                $bounds.Width,
                $bounds.Height,
                $true
            ) | Out-Null
        }
        else {
            $matchStr = if ($match.process) { $match.process } elseif ($match.title) { $match.title } else { "unknown" }
            Write-Host "  No window found matching: $matchStr" -ForegroundColor Yellow
        }
    }

    Write-Host "Tiling applied." -ForegroundColor Green
}

function Apply-TeamerDesktopTiling {
    <#
    .SYNOPSIS
        Applies tiling to windows based on desktop config from environment
    .DESCRIPTION
        Uses the grid config from desktop and row/col from each window definition
        to position windows. Matches windows by process name or title.
        First moves all matched windows to the target desktop, then applies tiling.
    .PARAMETER Desktop
        Desktop configuration object from environment
    .PARAMETER BaseWorkingDirectory
        Base working directory for the environment
    .PARAMETER DesktopIndex
        Target desktop index to move windows to
    #>
    param(
        [Parameter(Mandatory)]
        [object]$Desktop,

        [Parameter()]
        [string]$BaseWorkingDirectory,

        [Parameter()]
        [int]$DesktopIndex = -1
    )

    $grid = $Desktop.grid
    if (-not $grid) {
        Write-Warning "No grid defined for desktop: $($Desktop.name)"
        return
    }

    # Get target desktop object if index provided
    $targetDesktop = $null
    if ($DesktopIndex -ge 0) {
        $targetDesktop = Get-Desktop -Index $DesktopIndex
    }

    # Get ALL windows (not just current desktop) so we can find and move them
    $allWindows = @()
    Get-Process | Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } | ForEach-Object {
        $allWindows += @{
            Process = $_
            Handle = $_.MainWindowHandle
            Title = $_.MainWindowTitle
            ProcessName = $_.ProcessName
        }
    }

    # Track which windows we've already matched (to avoid matching same window twice)
    $usedHandles = @{}

    foreach ($window in $Desktop.windows) {
        # Skip windows without position info
        if ($null -eq $window.row -and $null -eq $window.col) {
            continue
        }

        $row = if ($null -ne $window.row) { $window.row } else { 0 }
        $col = if ($null -ne $window.col) { $window.col } else { 0 }
        $rowSpan = if ($window.rowSpan) { $window.rowSpan } else { 1 }
        $colSpan = if ($window.colSpan) { $window.colSpan } else { 1 }

        # Determine what to match based on window type
        $matchPattern = $null

        switch ($window.type) {
            'terminal' {
                $matchPattern = "WindowsTerminal"
            }
            'app' {
                $appName = [System.IO.Path]::GetFileNameWithoutExtension($window.path)
                $matchPattern = "^$appName$"
            }
            'browser' {
                $matchPattern = "^(chrome|firefox|msedge|brave)$"
            }
        }

        if (-not $matchPattern) {
            continue
        }

        # Find matching window that we haven't used yet
        $matchedWindow = $null
        foreach ($win in $allWindows) {
            if ($win.ProcessName -match $matchPattern -and -not $usedHandles.ContainsKey($win.Handle)) {
                $matchedWindow = $win
                $usedHandles[$win.Handle] = $true
                break
            }
        }

        if ($matchedWindow) {
            # First, move window to target desktop if specified
            if ($targetDesktop) {
                try {
                    Move-Window -Desktop $targetDesktop -Hwnd $matchedWindow.Handle 2>$null | Out-Null
                }
                catch {
                    # Window might already be on target desktop
                }
            }

            # Calculate bounds and position window
            $bounds = Get-TeamerGridCellBounds -Grid $grid -Row $row -Col $col -RowSpan $rowSpan -ColSpan $colSpan

            # Get invisible border offset for this window
            # Windows 10/11 have invisible borders (~7px) that affect positioning
            $frameOffset = Get-TeamerWindowFrameOffset -Handle $matchedWindow.Handle

            # Adjust position to compensate for invisible borders
            # We move the window frame left/up by the border amount so visible part is at target position
            $adjustedX = $bounds.X - $frameOffset.Left
            $adjustedY = $bounds.Y - $frameOffset.Top
            # Width/height need to include the invisible borders
            $adjustedWidth = $bounds.Width + $frameOffset.Left + $frameOffset.Right
            $adjustedHeight = $bounds.Height + $frameOffset.Top + $frameOffset.Bottom

            Write-Host "    $($matchedWindow.ProcessName) -> X=$($bounds.X) Y=$($bounds.Y) W=$($bounds.Width) H=$($bounds.Height) (adjusted: X=$adjustedX Y=$adjustedY W=$adjustedWidth H=$adjustedHeight)" -ForegroundColor DarkGray

            # Use SetWindowPos for more reliable positioning
            [TeamerWin32]::SetWindowPos(
                $matchedWindow.Handle,
                [IntPtr]::Zero,
                $adjustedX,
                $adjustedY,
                $adjustedWidth,
                $adjustedHeight,
                [TeamerWin32]::SWP_NOZORDER -bor [TeamerWin32]::SWP_NOACTIVATE
            ) | Out-Null

            # Also use MoveWindow as backup
            [TeamerWin32]::MoveWindow(
                $matchedWindow.Handle,
                $adjustedX,
                $adjustedY,
                $adjustedWidth,
                $adjustedHeight,
                $true
            ) | Out-Null
        }
        else {
            Write-Host "    No window found for pattern: $matchPattern" -ForegroundColor Yellow
        }
    }
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
Write-Host "Tiling Commands:" -ForegroundColor Cyan
Write-Host "  Get-TeamerTiling            - Get tiling config for desktop" -ForegroundColor White
Write-Host "  Set-TeamerTiling            - Set tiling config for desktop" -ForegroundColor White
Write-Host "  Apply-TeamerTiling          - Apply tiling to current windows" -ForegroundColor White
Write-Host ""
Write-Host "Logging Commands:" -ForegroundColor Cyan
Write-Host "  Get-TeamerLog               - View recent log entries" -ForegroundColor White
Write-Host "  Clear-TeamerLog             - Clear the log file" -ForegroundColor White
Write-Host ""
Write-Host "Security:" -ForegroundColor Cyan
Write-Host "  Test-CommandSafe            - Validate command safety" -ForegroundColor White
Write-Host ""

#endregion
