# Daily leaderboard refresh: export latest games -> recompute scores -> publish to GitHub.
# Scheduled to run once a day via Windows Task Scheduler.

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

& powershell -NoProfile -ExecutionPolicy Bypass -File "$PSScriptRoot\export_games.ps1"
python "$PSScriptRoot\generate_leaderboard.py"

git add games.csv leaderboard.json
$hasChanges = git status --porcelain
if ($hasChanges) {
    git commit -m "Daily leaderboard update: $(Get-Date -Format 'yyyy-MM-dd')"
    git push
    Write-Output "Leaderboard updated and pushed."
} else {
    Write-Output "No changes to publish today."
}
