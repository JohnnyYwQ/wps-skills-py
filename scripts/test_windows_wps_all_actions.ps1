param(
    [string]$PythonExe = "python",
    [string]$OutputDir = "",
    [switch]$SkipInteractive
)

$ErrorActionPreference = "Stop"

if ($env:OS -ne "Windows_NT") {
    throw "This script must run on a real Windows machine with WPS Office"
}

$parseScript = Join-Path $PSScriptRoot "test_windows_powershell_parse.ps1"
$runner = Join-Path $PSScriptRoot "test_windows_wps_all_actions.py"

Write-Host "Step 1/2: parse every generated Windows PowerShell bridge"
& $parseScript -PythonExe $PythonExe
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Write-Host "Step 2/2: execute all 235 Action cases against real WPS"
$runnerArguments = @("-X", "utf8", $runner)
if (-not [string]::IsNullOrWhiteSpace($OutputDir)) {
    $runnerArguments += @("--output-dir", $OutputDir)
}
if ($SkipInteractive) {
    $runnerArguments += "--skip-interactive"
}

& $PythonExe @runnerArguments
$runnerExitCode = $LASTEXITCODE
if ($runnerExitCode -ne 0) {
    Write-Host "Real-WPS full Action suite failed. Send back report.json, failures.log, actions.jsonl, and traces/."
    exit $runnerExitCode
}

Write-Host "Real-WPS full Action suite passed."
