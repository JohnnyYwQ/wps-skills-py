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
        throw "Python 命令失败，exit code: $LASTEXITCODE"
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
            Write-Host "[PASS] $app bridge: Windows PowerShell 解析成功"
            continue
        }

        $failed = $true
        Write-Host "[FAIL] $app bridge: $($parseErrors.Count) 个解析错误"
        foreach ($parseError in $parseErrors) {
            $line = $parseError.Extent.StartLineNumber
            $column = $parseError.Extent.StartColumnNumber
            Write-Host "  $scriptPath`:$line`:$column $($parseError.Message)"
        }
    }

    if ($failed) {
        throw "PowerShell bridge 解析回归测试失败"
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

Write-Host "Windows PowerShell bridge 解析回归测试通过。"
