param(
    [string]$PythonExe = "python",
    [switch]$KeepArtifacts
)

$ErrorActionPreference = "Stop"
$callScript = Join-Path $PSScriptRoot "call.py"
$artifactRoot = [System.IO.Path]::Combine(
    [System.IO.Path]::GetTempPath(),
    "wps-bridge-regression-$([Guid]::NewGuid().ToString('N'))"
)
$workbookCreated = $false
$testPassed = $false

function Invoke-WpsAction {
    param(
        [string]$Label,
        [string]$Action,
        [string]$App,
        [hashtable]$Params = @{}
    )

    $requestJson = $Params | ConvertTo-Json -Depth 20 -Compress
    $rawResponse = $requestJson | & $PythonExe -X utf8 $callScript $Action --app $App --stdin
    $exitCode = $LASTEXITCODE
    $responseText = ($rawResponse | Out-String).Trim()

    try {
        $response = $responseText | ConvertFrom-Json
    } catch {
        throw "[$Label] Action CLI returned invalid JSON (exit code $exitCode): $responseText"
    }

    if (-not $response.success) {
        $code = if ($response.code) { $response.code } else { "UNKNOWN" }
        $errorMessage = if ($response.error) { $response.error } else { $responseText }
        throw "[$Label] failed ($code): $errorMessage"
    }
    if ($response.data -is [System.Array]) {
        throw "[$Label] data must be an object, got an array: $responseText"
    }

    Write-Host "[PASS] $Label"
    return $response
}

function Assert-ExportedImage {
    param(
        [string]$Label,
        [string]$Path
    )

    if (-not [System.IO.File]::Exists($Path)) {
        throw "[$Label] output file was not created: $Path"
    }
    $length = (Get-Item -LiteralPath $Path).Length
    if ($length -le 0) {
        throw "[$Label] output file is empty: $Path"
    }
    Write-Host "[PASS] $Label output is valid ($length bytes)"
}

if ($env:OS -ne "Windows_NT") {
    throw "This script must run on a real Windows machine"
}

try {
    [System.IO.Directory]::CreateDirectory($artifactRoot) | Out-Null
    Write-Host "Test artifact directory: $artifactRoot"
    Write-Host "Do not switch the active WPS document while this test is running."

    & (Join-Path $PSScriptRoot "test_windows_powershell_parse.ps1") -PythonExe $PythonExe

    foreach ($app in @("excel", "ppt", "word")) {
        Invoke-WpsAction -Label "$app bridge initialization" -Action "getAppInfo" -App $app | Out-Null
    }

    $created = Invoke-WpsAction -Label "Create test workbook" -Action "createWorkbook" -App "excel"
    $workbookCreated = $true
    Write-Host "Test workbook: $($created.data.name)"

    Invoke-WpsAction -Label "Write chart data" -Action "setRangeData" -App "excel" -Params @{
        range = "A1:B4"
        data = @(
            @("Category", "Value"),
            @("A", 10),
            @("B", 20),
            @("C", 15)
        )
    } | Out-Null

    $chart = Invoke-WpsAction -Label "Create chart with data labels" -Action "createChart" -App "excel" -Params @{
        dataRange = "A1:B4"
        chartType = "column"
        title = "bridge regression"
        showLegend = $true
        showDataLabels = $true
        position = @{
            left = 100
            top = 100
            width = 480
            height = 300
        }
    }
    $chartName = [string]$chart.data.chartName
    if ([string]::IsNullOrWhiteSpace($chartName)) {
        throw "createChart did not return chartName"
    }

    $chartImage = Join-Path $artifactRoot "chart.png"
    $chartExport = Invoke-WpsAction -Label "Export chart image" -Action "exportChartAsImage" -App "excel" -Params @{
        chartName = $chartName
        outputPath = $chartImage
        format = "PNG"
    }
    if ($chartExport.data -isnot [PSCustomObject]) {
        throw "exportChartAsImage data is not an object"
    }
    Assert-ExportedImage -Label "Export chart image" -Path $chartImage

    $rangeImage = Join-Path $artifactRoot "range.png"
    $rangeExport = Invoke-WpsAction -Label "Export range image" -Action "exportRangeAsImage" -App "excel" -Params @{
        range = "A1:B4"
        outputPath = $rangeImage
        format = "PNG"
    }
    if ($rangeExport.data -isnot [PSCustomObject]) {
        throw "exportRangeAsImage data is not an object"
    }
    Assert-ExportedImage -Label "Export range image" -Path $rangeImage

    $testPassed = $true
} finally {
    if ($workbookCreated) {
        try {
            Invoke-WpsAction -Label "Close test workbook" -Action "closeWorkbook" -App "excel" -Params @{
                save = $false
            } | Out-Null
        } catch {
            Write-Warning "Could not close the test workbook: $($_.Exception.Message)"
        }
    }

    if ($testPassed -and -not $KeepArtifacts) {
        if (
            [System.IO.Directory]::Exists($artifactRoot) -and
            [System.IO.Path]::GetFileName($artifactRoot).StartsWith("wps-bridge-regression-")
        ) {
            [System.IO.Directory]::Delete($artifactRoot, $true)
        }
    } else {
        Write-Host "Test artifacts retained at: $artifactRoot"
    }
}

Write-Host "Windows WPS bridge real-machine regression test passed."
