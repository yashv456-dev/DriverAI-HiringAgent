# PowerShell launcher for DriverAI Hiring Agent P2 - Live Scoring
# Usage: .\RunP2.ps1
param(
    [int]$BatchSize = 500,
    [switch]$DryRun,
    [switch]$Scorecards
)

$Host.UI.RawUI.WindowTitle = "DriverAI Hiring Agent P2 - Live Scoring"
Set-Location -Path $PSScriptRoot

$PythonExe = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $PythonExe)) {
    $PythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source
}

if (-not $PythonExe) {
    Write-Error "[error] Python executable not found. Please run Launch.bat once to set up the environment."
    exit 1
}

$ArgsList = @("bot.py", "--score-sharepoint", "--batch-size", $BatchSize)
if ($DryRun) { $ArgsList += "--dry-run" }
if ($Scorecards) { $ArgsList += "--scorecards" }

Write-Host "[run] Starting P2 scoring pass against SharePoint queue..." -ForegroundColor Cyan
& $PythonExe @ArgsList
$ExitCode = $LASTEXITCODE

if ($ExitCode -eq 0) {
    Write-Host "[success] P2 scoring pass completed successfully." -ForegroundColor Green
} else {
    Write-Host "[error] P2 exited with code $ExitCode. Check P2_Logs or console output above." -ForegroundColor Red
}
exit $ExitCode
