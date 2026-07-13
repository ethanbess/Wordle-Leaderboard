# Daily leaderboard refresh: export latest games -> recompute scores -> publish to GitHub.
# Scheduled to run every 15 minutes via Windows Task Scheduler.

Set-Location $PSScriptRoot
$logPath = Join-Path $PSScriptRoot "update_log.txt"

function Write-Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    Write-Output $line
    Add-Content -Path $logPath -Value $line
}

$git = "C:\Program Files\Git\cmd\git.exe"
$python = "C:\Users\ethan\AppData\Local\Microsoft\WindowsApps\python.exe"

try {
    & powershell -NoProfile -ExecutionPolicy Bypass -File "$PSScriptRoot\export_games.ps1"
    & $python "$PSScriptRoot\generate_leaderboard.py"

    & $git add games.csv leaderboard.json
    $hasChanges = & $git status --porcelain -- games.csv leaderboard.json
    if (-not $hasChanges) {
        Write-Log "No changes to publish."
        exit 0
    }

    & $git commit -m "Daily leaderboard update: $(Get-Date -Format 'yyyy-MM-dd')" | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Log "ERROR: git commit failed (exit $LASTEXITCODE). Changes remain staged for next run."
        exit 1
    }

    & $git push 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Log "ERROR: git push failed (exit $LASTEXITCODE) - likely a network/connectivity issue. Commit was made locally and will be pushed automatically on the next successful run."
        exit 1
    }

    Write-Log "Leaderboard updated and pushed successfully."
    exit 0
}
catch {
    Write-Log "ERROR: unexpected failure - $($_.Exception.Message)"
    exit 1
}
