# Teamer

Windows Virtual Desktop Orchestration and Development Environment Management System.

Teamer provides automated workspace setup, multi-desktop management, and window tiling capabilities for developers on Windows 10/11.

## Features

- **Virtual Desktop Management** - Create, rename, switch, and remove virtual desktops
- **Environment Orchestration** - Define reproducible development environments with JSON configs
- **Multi-Shell Support** - PowerShell, PowerShell Core, WSL, CMD, Git Bash
- **Grid-Based Tiling** - Position windows in customizable grid layouts
- **Tab Grouping** - Group multiple terminals as tabs in one Windows Terminal window
- **Profile System** - Reusable shell profiles with startup commands and environment variables
- **Template System** - Reusable environment templates for different project types
- **State Persistence** - Track active environments across sessions
- **Desktop Protection** - Safeguard system desktops from accidental modification

## Requirements

- Windows 10 2004+ or Windows 11
- PowerShell 5.1 or PowerShell 7+
- Windows Terminal (recommended)
- PSVirtualDesktop module (auto-installed if missing)

## Installation

### Option 1: Clone and Dot-Source

```powershell
# Clone the repository
git clone https://github.com/your-repo/teamer.git E:\Projects\teamer

# Load the module (dot-source)
. E:\Projects\teamer\scripts\Manage-TeamerEnvironment.ps1
```

### Option 2: Import as Module

```powershell
# Import the module
Import-Module E:\Projects\teamer\Teamer.psd1
```

### Option 3: Add to Profile

Add to your PowerShell profile for automatic loading:

```powershell
# Add to $PROFILE
. "E:\Projects\teamer\scripts\Manage-TeamerEnvironment.ps1"
```

## Quick Start

### 1. List Available Templates

```powershell
Get-TeamerTemplate
```

### 2. Create an Environment

```powershell
New-TeamerEnvironment -Name "my-project" -Template "fullstack" -WorkingDirectory "C:\Projects\my-project"
```

### 3. Start the Environment

```powershell
Start-TeamerEnvironment -Name "my-project"
```

### 4. Stop the Environment

```powershell
Stop-TeamerEnvironment -Name "my-project"
```

## Core Commands

### Environment Management

| Command | Description |
|---------|-------------|
| `Get-TeamerEnvironment` | List all environments or get a specific one |
| `New-TeamerEnvironment` | Create a new environment configuration |
| `Start-TeamerEnvironment` | Deploy environment (create desktops, launch windows) |
| `Stop-TeamerEnvironment` | Tear down environment (close windows, remove desktops) |
| `Save-TeamerEnvironment` | Snapshot current state to configuration |
| `Remove-TeamerEnvironment` | Delete environment configuration |

### Desktop Management

| Command | Description |
|---------|-------------|
| `New-TeamerDesktop` | Create a new virtual desktop |
| `Show-TeamerTree` | Display system tree (screens, desktops, windows) |
| `Switch-TeamerDesktop` | Switch to a specific desktop |
| `Rename-TeamerDesktop` | Rename a desktop |
| `Remove-TeamerDesktop` | Remove a desktop |

### Window Management

| Command | Description |
|---------|-------------|
| `New-TeamerTerminal` | Launch a terminal on a specific desktop |
| `Move-TeamerWindow` | Move a window to a different desktop |
| `Pin-TeamerWindow` | Pin/unpin a window to all desktops |
| `Close-TeamerWindow` | Close a window |

### Profile Management

| Command | Description |
|---------|-------------|
| `Get-TeamerProfile` | List or get shell profiles |
| `New-TeamerProfile` | Create a new shell profile |
| `Set-TeamerProfile` | Update a profile property |
| `Remove-TeamerProfile` | Delete a profile |

### Tiling

| Command | Description |
|---------|-------------|
| `Get-TeamerTiling` | Get tiling configuration |
| `Set-TeamerTiling` | Set tiling configuration |
| `Apply-TeamerTiling` | Apply tiling to current windows |

### Logging

| Command | Description |
|---------|-------------|
| `Get-TeamerLog` | View recent log entries |
| `Clear-TeamerLog` | Clear the log file |

## Configuration

### Environment Configuration

Environments are defined in `environments/projects/<name>/environment.json`:

```json
{
  "$schema": "../schemas/environment.schema.json",
  "name": "My Project",
  "description": "Full-stack development environment",
  "workingDirectory": "C:\\Projects\\my-project",
  "layout": "dual-code-services",
  "desktops": [
    {
      "name": "Code",
      "grid": { "rows": 1, "cols": 2, "gap": 2, "margin": 2 },
      "windows": [
        { "type": "terminal", "profile": "wsl-ubuntu", "row": 0, "col": 0 },
        { "type": "app", "path": "code", "row": 0, "col": 1 }
      ]
    },
    {
      "name": "Services",
      "windows": [
        { "type": "terminal", "profile": "powershell", "zone": "main" }
      ]
    }
  ],
  "onStart": ["npm install", "docker-compose up -d"],
  "onStop": ["docker-compose down"]
}
```

### Profile Configuration

Shell profiles are defined in `environments/profiles/<name>.json`:

```json
{
  "$schema": "../schemas/profile.schema.json",
  "name": "WSL Ubuntu",
  "shell": "wsl",
  "distribution": "Ubuntu",
  "terminalProfile": "Ubuntu",
  "startupCommands": ["cd ~", "source ~/.bashrc"],
  "environment": {
    "NODE_ENV": "development"
  },
  "tabColor": "#E95420"
}
```

### Layout Configuration

Layouts define multi-monitor arrangements in `environments/layouts/<name>.json`:

```json
{
  "$schema": "../schemas/layout.schema.json",
  "name": "Dual Code Services",
  "description": "Code on primary, services on secondary",
  "grid": { "rows": 2, "cols": 2 },
  "zones": [
    { "name": "main", "monitor": "primary", "arrangement": "maximized" },
    { "name": "services", "monitor": "secondary", "arrangement": "tiled" }
  ]
}
```

## Directory Structure

```
teamer/
├── Teamer.psd1                    # Module manifest
├── README.md                      # This file
├── scripts/
│   ├── TeamerWin32.ps1           # Shared Win32 API definitions
│   ├── Manage-Teamer.ps1         # Core desktop/window operations
│   ├── Manage-TeamerEnvironment.ps1  # Environment management
│   ├── Get-SystemTree.ps1        # System discovery
│   └── Get-SystemConfig.ps1      # Configuration capture
├── environments/
│   ├── schemas/                  # JSON Schema definitions
│   ├── templates/                # Reusable environment templates
│   ├── profiles/                 # Shell profiles
│   ├── layouts/                  # Multi-monitor layouts
│   ├── projects/                 # Project-specific environments
│   ├── services/                 # External service configs
│   └── state.json               # Runtime state
└── tests/                        # Pester tests
```

## Safety Features

### Desktop Protection

Teamer protects system desktops from accidental modification. Protected desktops (by name):

- Main
- Code
- Desktop 1, Desktop 2, Desktop 3

Only desktops created by Teamer can be modified or removed.

### Command Sanitization

Lifecycle commands (`onStart`/`onStop`) are validated against dangerous patterns:

- Recursive force deletions
- Disk formatting commands
- System shutdown/restart
- Download-and-execute patterns
- Privilege escalation attempts

### Logging

All operations are logged to `environments/teamer.log`:

```powershell
# View recent logs
Get-TeamerLog -Lines 100

# Clear logs
Clear-TeamerLog
```

## Supported Shells

| Shell | Profile Value | Description |
|-------|---------------|-------------|
| PowerShell | `powershell` | Windows PowerShell 5.1 |
| PowerShell Core | `pwsh` | PowerShell 7+ |
| WSL | `wsl` | Windows Subsystem for Linux |
| Command Prompt | `cmd` | Classic CMD |
| Git Bash | `git-bash` | MINGW64 Git Bash |

## Examples

### Create a Data Science Environment

```powershell
New-TeamerEnvironment -Name "data-analysis" -Template "data-science" -WorkingDirectory "C:\Projects\analysis"
Start-TeamerEnvironment -Name "data-analysis"
```

### View System Tree

```powershell
# Show all windows
Show-TeamerTree

# Show only terminals
Show-TeamerTree -Filter terminals

# Show only manageable desktops
Show-TeamerTree -Manageable
```

### Create Custom Profile

```powershell
New-TeamerProfile -Name "python-dev" -Shell wsl -Distribution "Ubuntu"
Set-TeamerProfile -Name "python-dev" -Property "startupCommands" -Value @("source venv/bin/activate", "export PYTHONPATH=.")
```

### Manual Window Tiling

```powershell
# Set up a 2x2 grid tiling for current desktop
Set-TeamerTiling -DesktopName "Code" -Grid @{rows=2; cols=2; gap=4; margin=4} -Windows @(
    @{ match = @{ process = "WindowsTerminal" }; row = 0; col = 0 },
    @{ match = @{ process = "Code" }; row = 0; col = 1; rowSpan = 2 }
)

# Apply the tiling
Apply-TeamerTiling -DesktopName "Code"
```

## Troubleshooting

### PSVirtualDesktop Module Not Found

```powershell
Install-Module -Name VirtualDesktop -Scope CurrentUser -Force
```

### Windows Not Positioning Correctly

Check DPI scaling settings and ensure all monitors use the same scale factor. Use `debug-tiling.ps1` for diagnostics:

```powershell
.\scripts\debug-tiling.ps1
```

### Environment Won't Start

Check the log file for errors:

```powershell
Get-TeamerLog -Lines 50
```

## Testing

Teamer includes a comprehensive Pester test suite covering:

- Command sanitization (security)
- Profile, template, and layout management
- Grid calculations
- WSL path conversion
- Logging functionality
- Win32 module integration
- Desktop protection

### Running Tests

```powershell
# Run all tests
Invoke-Pester -Path .\tests\Teamer.Tests.ps1

# Run with verbose output (Pester 5.x)
Invoke-Pester -Path .\tests\Teamer.Tests.ps1 -Output Detailed
```

### Test Coverage

| Category | Tests | Status |
|----------|-------|--------|
| Command Sanitization | 12 | Passing |
| Profile Management | 3 | Passing |
| Template Management | 3 | Passing |
| Layout Management | 2 | Passing |
| Grid Calculations | 7 | Passing |
| WSL Path Conversion | 5 | Passing |
| Logging | 4 | Passing |
| Win32 Module | 3 | Passing |
| Desktop Protection | 3 | Passing |
| **Total** | **42** | **All Passing** |

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Run tests: `Invoke-Pester .\tests\`
5. Submit a pull request

## Changelog

### v1.0.0 (2026-01-21)

- Initial release
- Virtual desktop CRUD operations
- Environment lifecycle management (start/stop)
- Multi-shell support (PowerShell, WSL, CMD, Git Bash)
- Grid-based window tiling with DWM API integration
- Tab grouping for terminals
- State persistence
- Template and profile system
- Command sanitization and security logging
- Comprehensive Pester test suite (42 tests)
- Module manifest for proper PowerShell module distribution

## License

MIT License - See LICENSE file for details.
