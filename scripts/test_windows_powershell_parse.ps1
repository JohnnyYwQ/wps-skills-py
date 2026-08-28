param(
    [string]$PythonExe = "python"
)

$ErrorActionPreference = "Stop"
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

try {
    [System.IO.Directory]::CreateDirectory($tempRoot) | Out-Null
    Invoke-Python -Arguments @("--version")
    $exportScript = Join-Path $PSScriptRoot "export_windows_powershell_bridges.py"
    Invoke-Python -Arguments @($exportScript, "--output-dir", $tempRoot)

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
    if (
        [System.IO.Directory]::Exists($tempRoot) -and
        [System.IO.Path]::GetFileName($tempRoot).StartsWith("wps-bridge-parse-")
    ) {
        [System.IO.Directory]::Delete($tempRoot, $true)
    }
}

Write-Host "Windows PowerShell bridge parse regression test passed."
