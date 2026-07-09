# Exports the Date/Player/Guesses log from the Wordle Excel file to games.csv.
# Works whether the workbook is currently open in the user's Excel or not:
#   - If it's already open (e.g. they're mid-entry), reads live values via COM
#     from that session without touching/closing/saving it.
#   - If it's not open, launches a hidden, read-only Excel instance, reads it,
#     and closes that instance only (never touches the user's own session).

$ErrorActionPreference = "Stop"
$excelPath = "C:\Users\ethan\OneDrive\Desktop\National Wordle League\Excel Wordle Data.xlsx"
$csvPath = Join-Path $PSScriptRoot "games.csv"

$attached = $false
$excel = $null
$wb = $null

try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    foreach ($w in $excel.Workbooks) {
        if ($w.Name -eq "Excel Wordle Data.xlsx") { $wb = $w }
    }
    if ($wb) { $attached = $true }
} catch {
    $excel = $null
}

if (-not $wb) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($excelPath, [System.Reflection.Missing]::Value, $true)  # ReadOnly
}

try {
    $ws = $wb.Sheets.Item("Sheet1")
    $rows = @("Date,Player,Guesses")
    $r = 4
    while ($true) {
        $dCell = $ws.Cells.Item($r, 16)   # P
        $pCell = $ws.Cells.Item($r, 17)   # Q
        $gCell = $ws.Cells.Item($r, 18)   # R
        if ([string]::IsNullOrEmpty($dCell.Value2) -and [string]::IsNullOrEmpty($pCell.Value2)) { break }
        if (-not [string]::IsNullOrEmpty($dCell.Value2) -and -not [string]::IsNullOrEmpty($pCell.Value2) -and -not [string]::IsNullOrEmpty($gCell.Value2)) {
            $dateVal = [datetime]::FromOADate($dCell.Value2).ToString("yyyy-MM-dd")
            $player = ($pCell.Value2).ToString().Trim()
            $guesses = [int]$gCell.Value2
            $rows += "$dateVal,$player,$guesses"
        }
        $r++
    }
    $rows | Out-File -FilePath $csvPath -Encoding utf8
    Write-Output "Wrote $($rows.Count - 1) games to $csvPath"
}
finally {
    if (-not $attached) {
        $wb.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
