# Test script for WSL and Excel
. "$PSScriptRoot\Manage-TeamerEnvironment.ps1" | Out-Null

# Create test config
$projectDir = "E:\Projects\teamer\environments\projects\wsl-excel-test"
if (Test-Path $projectDir) { Remove-Item $projectDir -Recurse -Force }
New-Item -ItemType Directory -Path $projectDir -Force | Out-Null

$config = @{
    '$schema' = "../schemas/environment.schema.json"
    name = "WSL + Excel Test"
    description = "Testing WSL terminal and Excel app"
    workingDirectory = "E:\Projects\teamer"
    layout = "single-focus"
    desktops = @(
        @{
            name = "Work"
            windows = @(
                @{
                    type = "terminal"
                    profile = "wsl-ubuntu"
                    title = "Ubuntu"
                    tabColor = "#E95420"
                },
                @{
                    type = "app"
                    path = "excel.exe"
                }
            )
        }
    )
    onStart = @()
    onStop = @()
}

$config | ConvertTo-Json -Depth 10 | Out-File "$projectDir\environment.json" -Encoding UTF8

Write-Host "Starting WSL + Excel environment..." -ForegroundColor Cyan
Start-TeamerEnvironment -Name "wsl-excel-test"

Write-Host ""
Write-Host "You should see:" -ForegroundColor Green
Write-Host "  - Desktop named 'wsl-excel-test'" -ForegroundColor White
Write-Host "  - Ubuntu terminal (orange tab)" -ForegroundColor White
Write-Host "  - Microsoft Excel" -ForegroundColor White
