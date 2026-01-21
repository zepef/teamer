@{
    # Script module or binary module file associated with this manifest
    RootModule = 'scripts\Manage-TeamerEnvironment.ps1'

    # Version number of this module
    ModuleVersion = '1.0.0'

    # ID used to uniquely identify this module
    GUID = 'f8e3c7a1-9d4b-4e6f-8c2a-1b5d9e3f7a2c'

    # Author of this module
    Author = 'Teamer Contributors'

    # Company or vendor of this module
    CompanyName = 'Unknown'

    # Copyright statement for this module
    Copyright = '(c) 2024-2026. All rights reserved.'

    # Description of the functionality provided by this module
    Description = 'Windows Virtual Desktop Orchestration and Development Environment Management System. Provides automated workspace setup, multi-desktop management, and window tiling capabilities for developers.'

    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '5.1'

    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules = @()

    # Assemblies that must be loaded prior to importing this module
    RequiredAssemblies = @()

    # Script files (.ps1) that are run in the caller's environment prior to importing this module
    ScriptsToProcess = @(
        'scripts\TeamerWin32.ps1',
        'scripts\Manage-Teamer.ps1'
    )

    # Type files (.ps1xml) to be loaded when importing this module
    TypesToProcess = @()

    # Format files (.ps1xml) to be loaded when importing this module
    FormatsToProcess = @()

    # Modules to import as nested modules of the module specified in RootModule
    NestedModules = @()

    # Functions to export from this module
    FunctionsToExport = @(
        # Environment Management
        'Get-TeamerEnvironment',
        'New-TeamerEnvironment',
        'Start-TeamerEnvironment',
        'Stop-TeamerEnvironment',
        'Save-TeamerEnvironment',
        'Remove-TeamerEnvironment',

        # Profile Management
        'Get-TeamerProfile',
        'New-TeamerProfile',
        'Set-TeamerProfile',
        'Remove-TeamerProfile',

        # Template Management
        'Get-TeamerTemplate',
        'New-TeamerTemplate',

        # Layout Management
        'Get-TeamerLayout',
        'Get-TeamerZone',

        # Desktop Management (from Manage-Teamer.ps1)
        'New-TeamerDesktop',
        'New-TeamerTerminal',
        'Get-TeamerTree',
        'Show-TeamerTree',
        'Rename-TeamerDesktop',
        'Move-TeamerWindow',
        'Switch-TeamerDesktop',
        'Pin-TeamerWindow',
        'Remove-TeamerDesktop',
        'Close-TeamerWindow',

        # Tiling
        'Get-TeamerTiling',
        'Set-TeamerTiling',
        'Apply-TeamerTiling',
        'Apply-TeamerDesktopTiling',
        'Get-TeamerGridCellBounds',
        'Get-TeamerScreenBounds',

        # Logging
        'Write-TeamerLog',
        'Get-TeamerLog',
        'Clear-TeamerLog',

        # Command Safety
        'Test-CommandSafe',
        'Invoke-TeamerCommand',

        # Win32 Helpers
        'Get-TeamerWindowFrameOffset',
        'Move-TeamerWindowPosition'
    )

    # Cmdlets to export from this module
    CmdletsToExport = @()

    # Variables to export from this module
    VariablesToExport = @()

    # Aliases to export from this module
    AliasesToExport = @()

    # DSC resources to export from this module
    DscResourcesToExport = @()

    # List of all modules packaged with this module
    ModuleList = @()

    # List of all files packaged with this module
    FileList = @(
        'Teamer.psd1',
        'scripts\TeamerWin32.ps1',
        'scripts\Manage-Teamer.ps1',
        'scripts\Manage-TeamerEnvironment.ps1',
        'scripts\Get-SystemTree.ps1',
        'scripts\Get-SystemConfig.ps1'
    )

    # Private data to pass to the module specified in RootModule
    PrivateData = @{
        PSData = @{
            # Tags applied to this module for discoverability in PowerShell Gallery
            Tags = @('VirtualDesktop', 'Windows', 'Desktop', 'Development', 'Environment', 'Workspace', 'Tiling', 'WindowManager')

            # A URL to the license for this module
            LicenseUri = ''

            # A URL to the main website for this project
            ProjectUri = ''

            # A URL to an icon representing this module
            IconUri = ''

            # ReleaseNotes of this module
            ReleaseNotes = @'
## Version 1.0.0
- Initial release
- Virtual desktop CRUD operations
- Environment lifecycle management (start/stop)
- Multi-shell support (PowerShell, WSL, CMD, Git Bash)
- Grid-based window tiling
- Tab grouping for terminals
- State persistence
- Template and profile system
- Command sanitization and logging
- DWM API-based window positioning
'@

            # Prerelease tag (e.g., 'beta', 'preview')
            Prerelease = ''

            # Flag to indicate whether the module requires explicit user acceptance for install/update/save
            RequireLicenseAcceptance = $false

            # External dependent modules of this module
            ExternalModuleDependencies = @('VirtualDesktop')
        }
    }

    # HelpInfo URI of this module
    HelpInfoURI = ''

    # Default prefix for commands exported from this module
    DefaultCommandPrefix = ''
}
