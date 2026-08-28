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
    $rawResponse = $requestJson | & $PythonExe $callScript $Action --app $App --stdin
    $exitCode = $LASTEXITCODE
    $responseText = ($rawResponse | Out-String).Trim()

    try {
        $response = $responseText | ConvertFrom-Json
    } catch {
        throw "[$Label] Action CLI 未返回合法 JSON（exit code $exitCode）：$responseText"
    }

    if (-not $response.success) {
        $code = if ($response.code) { $response.code } else { "UNKNOWN" }
        $errorMessage = if ($response.error) { $response.error } else { $responseText }
        throw "[$Label] 失败（$code）：$errorMessage"
    }
    if ($response.data -is [System.Array]) {
        throw "[$Label] data 必须是 object，实际为 array：$responseText"
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
        throw "[$Label] 未生成文件：$Path"
    }
    $length = (Get-Item -LiteralPath $Path).Length
    if ($length -le 0) {
        throw "[$Label] 生成了空文件：$Path"
    }
    Write-Host "[PASS] $Label 文件有效（$length bytes）"
}

if ($env:OS -ne "Windows_NT") {
    throw "此脚本只能在 Windows 实机运行"
}

try {
    [System.IO.Directory]::CreateDirectory($artifactRoot) | Out-Null
    Write-Host "测试产物目录：$artifactRoot"
    Write-Host "测试期间请勿手动切换 WPS 的活动文档。"

    & (Join-Path $PSScriptRoot "test_windows_powershell_parse.ps1") -PythonExe $PythonExe

    foreach ($app in @("excel", "ppt", "word")) {
        Invoke-WpsAction -Label "$app bridge 初始化" -Action "getAppInfo" -App $app | Out-Null
    }

    $created = Invoke-WpsAction -Label "创建测试工作簿" -Action "createWorkbook" -App "excel"
    $workbookCreated = $true
    Write-Host "测试工作簿：$($created.data.name)"

    Invoke-WpsAction -Label "写入图表数据" -Action "setRangeData" -App "excel" -Params @{
        range = "A1:B4"
        data = @(
            @("类别", "数值"),
            @("A", 10),
            @("B", 20),
            @("C", 15)
        )
    } | Out-Null

    $chart = Invoke-WpsAction -Label "创建带数据标签的图表" -Action "createChart" -App "excel" -Params @{
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
        throw "createChart 未返回 chartName"
    }

    $chartImage = Join-Path $artifactRoot "chart.png"
    $chartExport = Invoke-WpsAction -Label "导出图表图片" -Action "exportChartAsImage" -App "excel" -Params @{
        chartName = $chartName
        outputPath = $chartImage
        format = "PNG"
    }
    if ($chartExport.data -isnot [PSCustomObject]) {
        throw "exportChartAsImage 的 data 不是 object"
    }
    Assert-ExportedImage -Label "导出图表图片" -Path $chartImage

    $rangeImage = Join-Path $artifactRoot "range.png"
    $rangeExport = Invoke-WpsAction -Label "导出区域图片" -Action "exportRangeAsImage" -App "excel" -Params @{
        range = "A1:B4"
        outputPath = $rangeImage
        format = "PNG"
    }
    if ($rangeExport.data -isnot [PSCustomObject]) {
        throw "exportRangeAsImage 的 data 不是 object"
    }
    Assert-ExportedImage -Label "导出区域图片" -Path $rangeImage

    $testPassed = $true
} finally {
    if ($workbookCreated) {
        try {
            Invoke-WpsAction -Label "关闭测试工作簿" -Action "closeWorkbook" -App "excel" -Params @{
                save = $false
            } | Out-Null
        } catch {
            Write-Warning "测试工作簿自动关闭失败：$($_.Exception.Message)"
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
        Write-Host "测试产物保留在：$artifactRoot"
    }
}

Write-Host "Windows WPS bridge 实机回归测试全部通过。"
