# Test script for environment with integrated tiling
. "$PSScriptRoot\Manage-TeamerEnvironment.ps1" | Out-Null

# Clean up previous test
$projectDir = "E:\Projects\teamer\environments\projects\tiling-test"
if (Test-Path $projectDir) { Remove-Item $projectDir -Recurse -Force }

Write-Host "=== Environment with Tiling Test ===" -ForegroundColor Magenta

# Create environment config with grid tiling
$config = @{
    '$schema' = "../schemas/environment.schema.json"
    name = "Tiling Test Environment"
    description = "Testing environment-integrated tiling"
    workingDirectory = "E:\Projects\teamer"
    desktops = @(
        @{
            name = "Dev"
            grid = @{
                rows = 2
                cols = 2
                gap = 2
                margin = 2
            }
            windows = @(
                @{
                    type = "terminal"
                    profile = "wsl-ubuntu"
                    title = "Ubuntu"
                    tabColor = "#E95420"
                    row = 0
                    col = 0
                    rowSpan = 2
                    colSpan = 1
                },
                @{
                    type = "app"
                    path = "excel.exe"
                    row = 0
                    col = 1
                    rowSpan = 1
                    colSpan = 1
                },
                @{
                    type = "app"
                    path = "notepad.exe"
                    row = 1
                    col = 1
                    rowSpan = 1
                    colSpan = 1
                }
            )
        }
    )
    onStart = @()
    onStop = @()
}

# Save and start
New-Item -ItemType Directory -Path $projectDir -Force | Out-Null
$config | ConvertTo-Json -Depth 10 | Out-File "$projectDir\environment.json" -Encoding UTF8

Write-Host "Starting environment with tiling..." -ForegroundColor Yellow
Start-TeamerEnvironment -Name "tiling-test"

Write-Host ""
Write-Host "Expected layout (2x2 grid):" -ForegroundColor Cyan
Write-Host "  [Ubuntu    ] [Excel  ]" -ForegroundColor White
Write-Host "  [          ] [Notepad]" -ForegroundColor White
