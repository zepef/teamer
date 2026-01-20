# Stop all active test environments
. "$PSScriptRoot\Manage-TeamerEnvironment.ps1" | Out-Null

$active = Get-TeamerEnvironment | Where-Object { $_.IsActive }
if ($active) {
    foreach ($env in $active) {
        Write-Host "Stopping $($env.Name)..." -ForegroundColor Yellow
        Stop-TeamerEnvironment -Name $env.Name -Force
    }
    Write-Host "All environments stopped." -ForegroundColor Green
} else {
    Write-Host "No active environments." -ForegroundColor Cyan
}
