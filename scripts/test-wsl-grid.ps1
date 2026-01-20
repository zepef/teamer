# Test script for WSL terminals, tab groups, and grid layout
. "$PSScriptRoot\Manage-TeamerEnvironment.ps1" | Out-Null

# Stop previous test if running
if ((Get-TeamerEnvironment | Where-Object { $_.Name -eq 'wsl-grid-test' -and $_.IsActive })) {
    Write-Host "Stopping previous wsl-grid-test..." -ForegroundColor Yellow
    Stop-TeamerEnvironment -Name "wsl-grid-test" -Force
}

# Remove old config
$projectDir = "E:\Projects\teamer\environments\projects\wsl-grid-test"
if (Test-Path $projectDir) {
    Remove-Item $projectDir -Recurse -Force
}

Write-Host "=== Creating WSL + Grid + Tab Groups Test ===" -ForegroundColor Magenta

# Create environment config
$config = @{
    '$schema' = "../schemas/environment.schema.json"
    name = "WSL Grid Test"
    description = "Testing WSL terminals, tab groups, and grid layout"
    workingDirectory = "E:\Projects\teamer"
    layout = "grid-3x2"
    desktops = @(
        @{
            name = "Development"
            windows = @(
                # Tab group 1: Shell terminals (will be tabs in one window)
                @{
                    type = "terminal"
                    profile = "wsl-ubuntu"
                    tabGroup = "shells"
                    title = "Ubuntu"
                    tabColor = "#E95420"
                    row = 0
                    col = 0
                    rowSpan = 2
                    colSpan = 1
                },
                @{
                    type = "terminal"
                    profile = "powershell"
                    tabGroup = "shells"
                    title = "PowerShell"
                    tabColor = "#012456"
                },
                # Standalone terminal (no tab group)
                @{
                    type = "terminal"
                    profile = "powershell"
                    title = "Logs"
                    tabColor = "#4EC9B0"
                    row = 2
                    col = 0
                    rowSpan = 1
                    colSpan = 2
                },
                # App in sidebar
                @{
                    type = "app"
                    path = "notepad.exe"
                    row = 0
                    col = 1
                    rowSpan = 2
                    colSpan = 1
                }
            )
        }
    )
    onStart = @()
    onStop = @()
}

# Save the config
New-Item -ItemType Directory -Path $projectDir -Force | Out-Null
$config | ConvertTo-Json -Depth 10 | Out-File "$projectDir\environment.json" -Encoding UTF8

Write-Host "Created wsl-grid-test environment" -ForegroundColor Green

# Start the environment
Write-Host "Starting environment..." -ForegroundColor Yellow
Start-TeamerEnvironment -Name "wsl-grid-test"

Write-Host ""
Write-Host "You should see:" -ForegroundColor Cyan
Write-Host "  - Desktop named 'wsl-grid-test'" -ForegroundColor White
Write-Host "  - Terminal window with 2 tabs:" -ForegroundColor White
Write-Host "    - Ubuntu (orange tab) - real WSL shell" -ForegroundColor White
Write-Host "    - PowerShell (blue tab)" -ForegroundColor White
Write-Host "  - Standalone 'Logs' terminal (green tab)" -ForegroundColor White
Write-Host "  - Notepad window" -ForegroundColor White
Write-Host ""
Write-Host "Grid layout (3x2):" -ForegroundColor Cyan
Write-Host "  [Shells tabs] [Notepad  ]" -ForegroundColor White
Write-Host "  [           ] [         ]" -ForegroundColor White
Write-Host "  [   Logs (spans both cols)   ]" -ForegroundColor White
