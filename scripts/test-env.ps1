# Test script for environment with notepad and WSL
. "$PSScriptRoot\Manage-TeamerEnvironment.ps1" | Out-Null

# Stop color-test if running
if ((Get-TeamerEnvironment | Where-Object { $_.Name -eq 'color-test' })) {
    Write-Host "Stopping color-test..." -ForegroundColor Yellow
    try {
        $desktop = Get-Desktop -Index 2
        if ($desktop) {
            Remove-Desktop -Desktop $desktop
        }
    } catch {}
    Remove-Item "E:\Projects\teamer\environments\projects\color-test" -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "=== Creating test with Notepad and WSL ===" -ForegroundColor Magenta

# Create environment config
$config = @{
    '$schema' = "../schemas/environment.schema.json"
    name = "Dev Environment"
    description = "Testing with notepad and WSL"
    workingDirectory = "E:\Projects\teamer"
    layout = "single-focus"
    desktops = @(
        @{
            name = "Workspace"
            windows = @(
                @{
                    type = "terminal"
                    profile = "wsl-ubuntu"
                    zone = "main"
                    title = "Ubuntu Shell"
                    tabColor = "#E95420"
                },
                @{
                    type = "terminal"
                    profile = "powershell"
                    zone = "main"
                    title = "PowerShell"
                    tabColor = "#012456"
                },
                @{
                    type = "app"
                    path = "notepad.exe"
                    zone = "main"
                }
            )
        }
    )
    onStart = @()
    onStop = @()
}

# Save the config
$projectDir = "E:\Projects\teamer\environments\projects\dev-test"
if (-not (Test-Path $projectDir)) {
    New-Item -ItemType Directory -Path $projectDir -Force | Out-Null
}
$config | ConvertTo-Json -Depth 10 | Out-File "$projectDir\environment.json" -Encoding UTF8

Write-Host "Created dev-test environment" -ForegroundColor Green

# Start the environment
Write-Host "Starting environment..." -ForegroundColor Yellow
Start-TeamerEnvironment -Name "dev-test"

Write-Host ""
Write-Host "You should see:" -ForegroundColor Cyan
Write-Host "  - Desktop named 'dev-test'" -ForegroundColor White
Write-Host "  - Ubuntu terminal with orange tab (#E95420)" -ForegroundColor White
Write-Host "  - PowerShell terminal with blue tab (#012456)" -ForegroundColor White
Write-Host "  - Notepad window" -ForegroundColor White
