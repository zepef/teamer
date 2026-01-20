# Debug grid calculations
. "$PSScriptRoot\Manage-TeamerEnvironment.ps1" | Out-Null

$screen = Get-TeamerScreenBounds
Write-Host "Screen: $($screen.Width) x $($screen.Height)" -ForegroundColor Cyan

$grid = @{ rows = 2; cols = 2; gap = 2; margin = 2 }
Write-Host "Grid: 2x2, gap=2, margin=2" -ForegroundColor Cyan

$positions = @(
    @{ name = "Terminal"; row = 0; col = 0; rowSpan = 2; colSpan = 1 },
    @{ name = "Excel"; row = 0; col = 1; rowSpan = 1; colSpan = 1 },
    @{ name = "Notepad"; row = 1; col = 1; rowSpan = 1; colSpan = 1 }
)

Write-Host ""
foreach ($pos in $positions) {
    $b = Get-TeamerGridCellBounds -Grid $grid -Row $pos.row -Col $pos.col -RowSpan $pos.rowSpan -ColSpan $pos.colSpan
    $endX = $b.X + $b.Width
    $endY = $b.Y + $b.Height
    Write-Host "$($pos.name): X=$($b.X) Y=$($b.Y) W=$($b.Width) H=$($b.Height) (ends at X=$endX Y=$endY)"
}

Write-Host ""
Write-Host "Checking gaps:" -ForegroundColor Yellow

# Terminal ends, Excel starts
$term = Get-TeamerGridCellBounds -Grid $grid -Row 0 -Col 0 -RowSpan 2 -ColSpan 1
$excel = Get-TeamerGridCellBounds -Grid $grid -Row 0 -Col 1 -RowSpan 1 -ColSpan 1
$notepad = Get-TeamerGridCellBounds -Grid $grid -Row 1 -Col 1 -RowSpan 1 -ColSpan 1

Write-Host "Left margin: $($term.X) (should be 2)"
Write-Host "Gap Terminal-Excel: $($excel.X - ($term.X + $term.Width)) (should be 2)"
Write-Host "Right margin: $($screen.Width - ($excel.X + $excel.Width)) (should be 2)"
Write-Host "Top margin: $($term.Y) (should be 2)"
Write-Host "Gap Excel-Notepad: $($notepad.Y - ($excel.Y + $excel.Height)) (should be 2)"
Write-Host "Bottom margin: $($screen.Height - ($notepad.Y + $notepad.Height)) (should be 2)"
