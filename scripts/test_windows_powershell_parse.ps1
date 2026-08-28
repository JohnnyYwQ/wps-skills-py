param(
    [string]$PythonExe = "python"
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$tempRoot = [System.IO.Path]::Combine(
    [System.IO.Path]::GetTempPath(),
    "wps-bridge-parse-$([Guid]::NewGuid().ToString('N'))"
)

function Invoke-Python {
    param([string[]]$Arguments)

    & $PythonExe @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Python command failed with exit code $LASTEXITCODE"
    }
}

$previousRepoRoot = $env:WPS_BRIDGE_REPO_ROOT
$previousOutputRoot = $env:WPS_BRIDGE_PARSE_ROOT

try {
    [System.IO.Directory]::CreateDirectory($tempRoot) | Out-Null
    $env:WPS_BRIDGE_REPO_ROOT = $repoRoot
    $env:WPS_BRIDGE_PARSE_ROOT = $tempRoot

    Invoke-Python -Arguments @("--version")

    $exportCode = @'
import os
import sys
from pathlib import Path

repo_root = Path(os.environ["WPS_BRIDGE_REPO_ROOT"])
output_root = Path(os.environ["WPS_BRIDGE_PARSE_ROOT"])
sys.path.insert(0, str(repo_root / "bridge"))

import wps_excel
import wps_ppt
import wps_word

for name, script in (
    ("excel", wps_excel.PS_BRIDGE_SCRIPT),
    ("ppt", wps_ppt.PS_BRIDGE_SCRIPT),
    ("word", wps_word.PS_BRIDGE_SCRIPT),
):
    (output_root / f"wps_{name}.ps1").write_text(script, encoding="utf-8-sig")
'@
    Invoke-Python -Arguments @("-c", $exportCode)

    $failed = $false
    foreach ($app in @("excel", "ppt", "word")) {
        $scriptPath = Join-Path $tempRoot "wps_$app.ps1"
        $tokens = $null
        $parseErrors = $null
        [System.Management.Automation.Language.Parser]::ParseFile(
            $scriptPath,
            [ref]$tokens,
            [ref]$parseErrors
        ) | Out-Null

        if ($parseErrors.Count -eq 0) {
            Write-Host "[PASS] $app bridge: Windows PowerShell parse succeeded"
            continue
        }

        $failed = $true
        Write-Host "[FAIL] $app bridge: $($parseErrors.Count) parse error(s)"
        foreach ($parseError in $parseErrors) {
            $line = $parseError.Extent.StartLineNumber
            $column = $parseError.Extent.StartColumnNumber
            Write-Host "  $scriptPath`:$line`:$column $($parseError.Message)"
        }
    }

    if ($failed) {
        throw "PowerShell bridge parse regression test failed"
    }
} finally {
    $env:WPS_BRIDGE_REPO_ROOT = $previousRepoRoot
    $env:WPS_BRIDGE_PARSE_ROOT = $previousOutputRoot
    if (
        [System.IO.Directory]::Exists($tempRoot) -and
        [System.IO.Path]::GetFileName($tempRoot).StartsWith("wps-bridge-parse-")
    ) {
        [System.IO.Directory]::Delete($tempRoot, $true)
    }
}

Write-Host "Windows PowerShell bridge parse regression test passed."
