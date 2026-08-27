#!/usr/bin/env python3
"""
WPS Excel 控制器 - 纯 Python 实现
通过 COM 自动化（Windows）或命令行（Linux）控制 WPS Excel。

不依赖 MCP，不依赖外网，不依赖 Node.js/JS 环境。
Windows: 通过 subprocess 调用 PowerShell COM（PowerShell 是 Windows 内置组件）
Linux:   通过 subprocess 调用 WPS 命令行工具

所有 80 个 action 均通过 execute(action, params) 统一入口调用。
"""

import json
import subprocess
import sys
import os
import platform
import time
import tempfile
import queue
import threading
from typing import Any, Dict, Optional

from powershell_contracts import render_chart_type_converter
from line_process import stop_line_process
from action_timing import bounded_timeout
from windows_com import describe_powershell_startup_failure, resolve_com_runtime

# ==================== 平台检测 ====================
IS_WINDOWS = platform.system() == "Windows"
IS_LINUX = platform.system() == "Linux"
IS_MACOS = platform.system() == "Darwin"

# WPS Excel COM ProgID
EXCEL_PROGID = "Ket.Application"

# 单个 action 执行超时（秒）：超过则强杀 PowerShell 桥接进程，避免 COM 弹框导致永久卡死
EXEC_TIMEOUT = 120

# PowerShell 桥接脚本（持久进程模式）
PS_BRIDGE_SCRIPT = r'''
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$error.Clear()
$global:WpsActivationError = $null
$global:ExcelActionsWithoutActiveWorkbook = @(
    "ping", "getOpenWorkbooks", "getAppInfo", "reconnect", "switchWorkbook",
    "createWorkbook", "openWorkbook"
)

# 创建/获取 COM 对象
function Get-ExcelApp {
    try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application') }
    catch {
        try { return New-Object -ComObject 'Ket.Application' }
        catch {
            $hr = '0x{0:X8}' -f $_.Exception.HResult
            $global:WpsActivationError = "New-Object failed ($hr): $($_.Exception.Message)"
            try {
                $clsid = (Get-ItemProperty -Path 'HKCU:\Software\Classes\Ket.Application\CLSID' -ErrorAction Stop).'(default)'
                if ($clsid) {
                    $type = [Type]::GetTypeFromCLSID([Guid]$clsid)
                    if ($type) { return [Activator]::CreateInstance($type) }
                }
            } catch {}
        }
    }
    return $null
}

$global:excel = Get-ExcelApp
if ($global:excel) {
    try { $global:excel.Visible = $true } catch {}
    # 抑制“是否覆盖”等模态对话框，避免 COM 无法应答导致保存卡死
    try { $global:excel.DisplayAlerts = $false } catch {}
    Write-Host '{"ready":true}'
} else {
    $message = if ($global:WpsActivationError) {
        $global:WpsActivationError
    } else {
        '无法连接WPS Excel，请确认WPS已安装'
    }
    Write-Host (@{ready=$false; error=$message} | ConvertTo-Json -Compress)
    exit 1
}

# ==================== Action 实现函数 ====================

function Exec-ping($p) {
    return @{success=$true; data=@{message="pong"; timestamp=(Get-Date -Format "o")}}
}

# 将 1-based 列号转换为 A1 列标（修复 [char](64+$i) 在 >26 列时出错的问题）
function ConvertToA1($col) {
    # 必须用整数运算：$col 来自 JSON 是 Double，若直接 % / Floor 会保持 Double，
    # 导致 [char](65 + 0.0) 触发“无法将值 65 转换为 System.Char”错误。
    [int]$n = $col
    $s = ""
    while ($n -gt 0) {
        $n = $n - 1
        [int]$r = $n % 26
        $s = [char](65 + $r) + $s
        $n = [int][Math]::Floor($n / 26)
    }
    return $s
}

''' + render_chart_type_converter() + r'''
function Exec-getOpenWorkbooks($p) {
    $names = @()
    foreach ($wb in $global:excel.Workbooks) { $names += $wb.Name }
    return @{success=$true; data=@{workbooks=$names; count=$names.Count}}
}

function Exec-getActiveWorkbook($p) {
    $wb = $global:excel.ActiveWorkbook
    if (-not $wb) { return @{success=$false; error="没有打开的工作簿"} }
    $sheets = @()
    for ($i=1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
    return @{success=$true; data=@{name=$wb.Name; path=$wb.FullName; sheetCount=$wb.Sheets.Count; sheets=$sheets}}
}

function Exec-openWorkbook($p) {
    $wb = $global:excel.Workbooks.Open($p.filePath)
    return @{success=$true; data=@{name=$wb.Name; path=$wb.FullName}}
}

function Exec-switchWorkbook($p) {
    $wb = $global:excel.Workbooks.Item($p.name)
    $wb.Activate()
    return @{success=$true}
}

function Exec-closeWorkbook($p) {
    $name = $p.name
    if (-not $name) { $wb = $global:excel.ActiveWorkbook } 
    else { $wb = $global:excel.Workbooks.Item($name) }
    $save = if ($p.save -ne $false) { $true } else { $false }
    $wb.Close($save)
    return @{success=$true}
}

function Exec-createWorkbook($p) {
    $wb = $global:excel.Workbooks.Add()
    return @{success=$true; data=@{name=$wb.Name}}
}

function Exec-getCellValue($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $cell = $sheet.Cells.Item($p.row, $p.col)
    return @{success=$true; data=@{value=$cell.Value2; text=$cell.Text; formula=$cell.Formula}}
}

function Exec-setCellValue($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Cells.Item($p.row, $p.col).Value2 = $p.value
    return @{success=$true}
}

function Exec-getFormula($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $cell = $sheet.Range($p.cell)
    return @{success=$true; data=@{cell=$p.cell; formula=$cell.Formula; value=$cell.Value2}}
}

function Exec-getCellInfo($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $cell = $sheet.Range($p.cell)
    return @{success=$true; data=@{
        value=$cell.Value2; text=$cell.Text; formula=$cell.Formula
        row=$cell.Row; column=$cell.Column
        fontName=$cell.Font.Name; fontSize=$cell.Font.Size; bold=$cell.Font.Bold
        numberFormat=$cell.NumberFormat
    }}
}

function Exec-clearRange($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    $type = $p.type
    if ($type -eq "contents") { $range.ClearContents() }
    elseif ($type -eq "formats") { $range.ClearFormats() }
    else { $range.Clear() }
    return @{success=$true}
}

function Exec-setFormula($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Range($p.cell).Formula = $p.formula
    return @{success=$true; data=@{cell=$p.cell; formula=$p.formula}}
}

function Exec-getContext($p) {
    $app = $global:excel
    $wb = $app.ActiveWorkbook
    if (-not $wb) { return @{success=$false; error="没有打开的工作簿"} }
    $sheet = $app.ActiveSheet
    $usedRange = $sheet.UsedRange
    $headers = @()
    if ($usedRange.Rows.Count -gt 0) {
        $colCount = [Math]::Min($usedRange.Columns.Count, 26)
        for ($i=1; $i -le $colCount; $i++) {
            $colLetter = ConvertToA1 $i
            # 用绝对坐标读取表头，避免 WPS 下 $range.Cells.Item 的 0 基索引偏移
            $headers += @{column=$colLetter; value=$sheet.Cells.Item($usedRange.Row, $usedRange.Column + $i - 1).Value2}
        }
    }
    $sheets = @()
    for ($i=1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
    return @{success=$true; data=@{
        workbookName=$wb.Name; currentSheet=$sheet.Name; allSheets=$sheets
        selectedCell=$app.Selection.Address()
        headers=$headers; usedRangeAddress=$usedRange.Address()
    }}
}

function Exec-diagnoseFormula($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $cell = $sheet.Range($p.cell)
    $val = $cell.Value2
    $formula = $cell.Formula
    $errorType = $null
    $diagnosis = ""
    $suggestion = ""
    $errors = @{"-2146826281"="#DIV/0!"; "-2146826246"="#VALUE!"; "-2146826245"="#REF!"; "-2146826252"="#NAME?"; "-2146826288"="#N/A"; "-2146826259"="#NUM!"; "-2146826280"="#NULL!"}
    $errKey = "$val"
    if ($errors.ContainsKey($errKey)) {
        $errorType = $errors[$errKey]
        switch ($errorType) {
            "#DIV/0!" { $diagnosis="除数为零"; $suggestion="检查除数是否为0或空单元格" }
            "#VALUE!" { $diagnosis="参数类型错误"; $suggestion="检查参数是否为正确类型" }
            "#REF!" { $diagnosis="引用了不存在的单元格"; $suggestion="检查引用范围是否被删除" }
            "#NAME?" { $diagnosis="函数名或名称错误"; $suggestion="检查函数名拼写是否正确" }
            "#N/A" { $diagnosis="查找未找到匹配值"; $suggestion="检查查找值是否存在" }
            Default { $diagnosis="公式错误"; $suggestion="检查公式" }
        }
    }
    $precedents = @()
    try { $prec = $cell.Precedents; foreach ($c in $prec) { $precedents += $c.Address() } } catch {}
    return @{success=$true; data=@{
        cell=$p.cell; formula=$formula; currentValue=$val
        errorType=$errorType; diagnosis=$diagnosis; suggestion=$suggestion; precedents=$precedents
    }}
}

function Exec-evaluateFormula($p) {
    $result = $global:excel.Evaluate($p.formula)
    return @{success=$true; data=@{result=$result}}
}

function Exec-setPrintArea($p) {
    $wb = $global:excel.ActiveWorkbook
    $wb.ActiveSheet.PageSetup.PrintArea = $p.range
    return @{success=$true}
}

function Exec-setZoom($p) {
    $wb = $global:excel.ActiveWorkbook
    $wb.ActiveWindow.Zoom = $p.percent
    return @{success=$true}
}

function Exec-getRangeData($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    # 注意：WPS 的 $range.Value2 批量读取存在偏移/丢数据缺陷（数据整体右下一格、末行末列丢失），
    # 且 $range.Cells.Item(r,c) 在 WPS 中是 0 基索引（Excel 为 1 基），直接用会整体错位。
    # 因此改用 $sheet.Cells.Item 的绝对 1 基坐标逐格读取，确保与写入一致（已验证可用）。
    $startRow = $range.Row
    $startCol = $range.Column
    $rows = $range.Rows.Count
    $cols = $range.Columns.Count
    $data = @()
    for ($r=0; $r -lt $rows; $r++) {
        $row = @()
        for ($c=0; $c -lt $cols; $c++) {
            $row += $sheet.Cells.Item($startRow + $r, $startCol + $c).Value2
        }
        $data += ,$row
    }
    return @{success=$true; data=@{data=$data; rows=$rows; cols=$cols}}
}

function Exec-setRangeData($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $startCell = $sheet.Range($p.range)
    $data = $p.data
    $rows = $data.Count
    if ($rows -eq 0) { return @{success=$true; data=@{rows=0; cols=0}} }
    $cols = $data[0].Count
    # 构造真正的二维数组，单次 COM 写入
    $arr = New-Object 'object[,]' $rows, $cols
    for ($r=0; $r -lt $rows; $r++) {
        for ($c=0; $c -lt $cols; $c++) { $arr[$r, $c] = $data[$r][$c] }
    }
    $dest = $startCell.Resize($rows, $cols)
    $dest.Value2 = $arr
    return @{success=$true; data=@{rows=$rows; cols=$cols}}
}

function Exec-cleanData($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    $results = @()
    foreach ($op in $p.operations) {
        switch ($op) {
            "trim" {
                for ($r=1; $r -le $range.Rows.Count; $r++) {
                    for ($c=1; $c -le $range.Columns.Count; $c++) {
                        $cell = $sheet.Cells.Item($range.Row + $r - 1, $range.Column + $c - 1)
                        if ($cell.Value2 -is [string]) { $cell.Value2 = $cell.Value2.Trim() }
                    }
                }
                $results += @{operation="trim"; success=$true; message="已去除前后空格"}
            }
            "remove_duplicates" {
                $cols = @()
                for ($i=1; $i -le $range.Columns.Count; $i++) { $cols += $i }
                $range.RemoveDuplicates($cols, 1)
                $results += @{operation="remove_duplicates"; success=$true; message="已删除重复行"}
            }
            "remove_empty_rows" {
                $deleted = 0
                for ($r=$range.Rows.Count; $r -ge 1; $r--) {
                    $isEmpty = $true
                    for ($c=1; $c -le $range.Columns.Count; $c++) {
                        $v = $sheet.Cells.Item($range.Row + $r - 1, $range.Column + $c - 1).Value2
                        if ($null -ne $v -and "$v" -ne "") { $isEmpty=$false; break }
                    }
                    if ($isEmpty) { $range.Rows.Item($r).Delete(); $deleted++ }
                }
                $results += @{operation="remove_empty_rows"; success=$true; message="删除了$deleted行空行"}
            }
            "unify_date" {
                $ud_ok = 0; $ud_fail = 0
                for ($r=1; $r -le $range.Rows.Count; $r++) {
                    for ($c=1; $c -le $range.Columns.Count; $c++) {
                        $cell = $sheet.Cells.Item($range.Row + $r - 1, $range.Column + $c - 1)
                        try { $cell.Value2 = [DateTime]($cell.Value2).ToString("yyyy-MM-dd"); $ud_ok++ }
                        catch { $ud_fail++ }
                    }
                }
                $results += @{operation="unify_date"; success=($ud_fail -eq 0); message="已统一 $ud_ok 个，失败 $ud_fail 个"; unified=$ud_ok; failed=$ud_fail}
            }
        }
    }
    return @{success=$true; data=@{range=$p.range; operations=$results}}
}

function Exec-removeDuplicates($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    $origCount = $range.Rows.Count
    $cols = @()
    if ($p.columns -and $p.columns.Count -gt 0) {
        foreach ($c in $p.columns) { $cols += [int]$c }
    } else {
        for ($i=1; $i -le $range.Columns.Count; $i++) { $cols += $i }
    }
    $hasHeader = if ($p.hasHeader -ne $false) { 1 } else { 0 }
    $range.RemoveDuplicates($cols, $hasHeader)
    $newCount = $range.Rows.Count
    return @{success=$true; data=@{originalCount=$origCount; removedCount=($origCount-$newCount); remainingCount=$newCount}}
}

function Exec-sortRange($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    $col = $p.column
    $keyRange = $range.Columns.Item($col)
    $order = if ($p.ascending -ne $false) { 1 } else { 2 }
    $range.Sort($keyRange, $order)
    return @{success=$true}
}

function Exec-findReplace($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $usedRange = $sheet.UsedRange
    $matchCase = if ($p.matchCase) { true } else { $false }
    $what = $p.find
    $replacement = $p.replace
    $found = $usedRange.Find($what, [Type]::Missing, -4162, 1, 1, 1, $matchCase)
    $count = 0
    while ($null -ne $found) {
        $found.Value2 = $replacement
        $count++
        $found = $usedRange.FindNext($found)
        if ($count -gt 10000) { break }
    }
    return @{success=$true; data=@{count=$count}}
}

function Exec-insertRows($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $count = if ($p.count) { $p.count } else { 1 }
    $sheet.Range("A" + $p.row + ":A" + ($p.row + $count - 1)).Insert(-4121)
    return @{success=$true}
}

function Exec-addCellComment($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $cell = $sheet.Range($p.cell)
    $cell.ClearComments()
    $cell.AddComment($p.comment)
    return @{success=$true}
}

function Exec-protectSheet($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $doProtect = if ($p.protect -ne $false) { $true } else { $false }
    if ($doProtect) {
        $pw = if ($p.password) { $p.password } else { "" }
        $sheet.Protect($pw)
    } else {
        $pw = if ($p.password) { $p.password } else { "" }
        $sheet.Unprotect($pw)
    }
    return @{success=$true}
}

function Exec-addConditionalFormat($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $range = $sheet.Range($p.range)
    $range.FormatConditions.Delete()
    $range.FormatConditions.Add(1, $p.condition, [Type]::Missing)
    return @{success=$true}
}

function Exec-protectWorkbook($p) {
    $wb = $global:excel.ActiveWorkbook
    $doProtect = $p.protect
    if ($doProtect) { $wb.Protect($p.password, $true, $false) }
    else { $wb.Unprotect($p.password) }
    return @{success=$true}
}

function Exec-autoFilter($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $range = $sheet.Range($p.range)
    if ($p.criteria) { $range.AutoFilter($p.field, $p.criteria) }
    else { $range.AutoFilter() }
    return @{success=$true}
}

function Exec-copyRange($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Range($p.sourceRange).Copy($sheet.Range($p.targetRange))
    return @{success=$true}
}

function Exec-pasteRange($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Range($p.targetRange).PasteSpecial()
    return @{success=$true}
}

function Exec-fillSeries($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $src = $sheet.Range($p.sourceRange)
    $tgt = $sheet.Range($p.targetRange)
    $src.AutoFill($tgt, 0)
    return @{success=$true}
}

function Exec-transpose($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    $data = $global:excel.Transpose($range.Value2)
    $startCell = $sheet.Cells.Item(1, $range.Column + $range.Columns.Count + 1)
    $destRange = $sheet.Range($startCell, $startCell.Offset($range.Columns.Count - 1, $range.Rows.Count - 1))
    $destRange.Value2 = $data
    return @{success=$true; data=@{destination=$destRange.Address()}}
}

function Exec-textToColumns($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    $delim = if ($p.delimiter) { [char]$p.delimiter[0] } else { [char]"," }
    $range.TextToColumns($range, 1, 1, $false, $false, $false, $false, $false, $true, $delim)
    return @{success=$true}
}

function Exec-subtotal($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $range = $sheet.Range($p.range)
    $groupBy = $p.groupBy
    $func = switch ($p.function) { "sum" {1} "count" {2} "average" {1} default {1} }
    $summaryRange = $sheet.Range($p.summaryRange)
    $range.Subtotal($groupBy, $func, @($summaryRange.Column))
    return @{success=$true}
}

function Exec-createChart($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $dataRange = $sheet.Range($p.dataRange)
    $chartType = Convert-ChartType $p.chartType
    $left = if ($p.position.left) { $p.position.left } else { 100 }
    $top = if ($p.position.top) { $p.position.top } else { 100 }
    $width = if ($p.position.width) { $p.position.width } else { 480 }
    $height = if ($p.position.height) { $p.position.height } else { 300 }
    $chartObj = $sheet.ChartObjects().Add($left, $top, $width, $height)
    $chartObj.Chart.SetSourceData($dataRange)
    $chartObj.Chart.ChartType = $chartType
    if ($p.title) { $chartObj.Chart.HasTitle = $true; $chartObj.Chart.ChartTitle.Text = $p.title }
    if ($p.showLegend -ne $false) { $chartObj.Chart.HasLegend = $true }
    if ($p.showDataLabels) { $chartObj.Chart.HasAxis = $true }
    $chartTypeName = if ($p.chartTypeName) { $p.chartTypeName } else { [string]$p.chartType }
    return @{success=$true; data=@{chartName=$chartObj.Name; chartIndex=1; dataRange=$p.dataRange; chartType=$chartTypeName; position=@{left=$left; top=$top; width=$width; height=$height}}}
}

function Exec-updateChart($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    if ($p.chartName) { $chartObj = $sheet.ChartObjects($p.chartName) }
    else { $chartObj = $sheet.ChartObjects().Item($p.chartIndex) }
    $updated = @()
    if ($p.title) { $chartObj.Chart.HasTitle=$true; $chartObj.Chart.ChartTitle.Text=$p.title; $updated += "title" }
    if ($null -ne $p.chartType) { $chartObj.Chart.ChartType=(Convert-ChartType $p.chartType); $updated += "chartType" }
    if ($null -ne $p.showLegend) { $chartObj.Chart.HasLegend=$p.showLegend; $updated += "showLegend" }
    if ($p.dataRange) { $chartObj.Chart.SetSourceData($sheet.Range($p.dataRange)); $updated += "dataRange" }
    return @{success=$true; data=@{chartName=$chartObj.Name; updatedProperties=$updated}}
}

function Exec-exportChartAsImage($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $chartObj = $sheet.ChartObjects($p.chartName)
    $outputPath = $p.outputPath
    $format = if ($p.format) { $p.format } else { "PNG" }
    $chartObj.Chart.Export($outputPath, $format)
    return @{success=$true; data=@{chartName=$p.chartName; outputPath=$outputPath; format=$format}}
}

function Exec-exportRangeAsImage($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    $outputPath = $p.outputPath
    $format = if ($p.format) { $p.format } else { "PNG" }
    $range.CopyPicture(1, 2)
    $tempChart = $sheet.ChartObjects().Add(0, 0, $range.Width, $range.Height)
    $tempChart.Activate()
    $tempChart.Chart.Paste()
    $tempChart.Chart.Export($outputPath, $format)
    $tempChart.Delete()
    return @{success=$true; data=@{range=$p.range; outputPath=$outputPath; format=$format}}
}

function Exec-createPivotTable($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $srcRange = $sheet.Range($p.sourceRange)
    $destSheet = if ($p.destinationSheet) { $wb.Sheets.Item($p.destinationSheet) } else { $sheet }
    $destCell = $destSheet.Range($p.destinationCell)
    $cache = $wb.PivotCaches().Add(1, $srcRange)
    $pt = $cache.CreatePivotTable($destCell)
    if ($p.rowFields) { foreach ($f in $p.rowFields) { $pt.PivotFields($f).Orientation = 1 } }
    if ($p.columnFields) { foreach ($f in $p.columnFields) { $pt.PivotFields($f).Orientation = 2 } }
    if ($p.valueFields) {
        foreach ($vf in $p.valueFields) {
            $agg = if ($vf.aggregation) { $vf.aggregation.ToString().ToUpper() } else { "SUM" }
            $fn = switch ($agg) { "SUM"{1} "AVERAGE"{2} "COUNT"{3} "MAX"{4} "MIN"{5} "PRODUCT"{6} default{1} }
            $pt.AddDataField($pt.PivotFields($vf.field), "$agg of $($vf.field)", $fn)
        }
    }
    return @{success=$true; data=@{tableName=$pt.Name}}
}

function Exec-updatePivotTable($p) {
    $wb = $global:excel.ActiveWorkbook
    foreach ($sheet in $wb.Sheets) {
        foreach ($pt in $sheet.PivotTables) {
            if ($pt.Name -eq $p.tableName) {
                $pt.RefreshTable()
                return @{success=$true}
            }
        }
    }
    return @{success=$false; error="未找到透视表: $($p.tableName)"}
}

function Exec-createSheet($p) {
    $wb = $global:excel.ActiveWorkbook
    $newSheet = $wb.Sheets.Add()
    if ($p.name) { $newSheet.Name = $p.name }
    $index = $newSheet.Index
    return @{success=$true; data=@{name=$newSheet.Name; index=($index-1)}}
}

function Exec-deleteSheet($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.Sheets.Item($p.name)
    $global:excel.DisplayAlerts = $false
    $sheet.Delete()
    $global:excel.DisplayAlerts = $true
    return @{success=$true; data=@{deleted=$p.name}}
}

function Exec-renameSheet($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.Sheets.Item($p.oldName)
    $sheet.Name = $p.newName
    return @{success=$true; data=@{oldName=$p.oldName; newName=$p.newName}}
}

function Exec-copySheet($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.Sheets.Item($p.name)
    $newSheet = $sheet.Copy()
    $newSheet = $wb.Sheets.Item($wb.Sheets.Count)
    if ($p.newName) { $newSheet.Name = $p.newName }
    return @{success=$true; data=@{sourceName=$p.name; newName=$newSheet.Name; index=($newSheet.Index-1)}}
}

function Exec-getSheetList($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheets = @()
    for ($i=1; $i -le $wb.Sheets.Count; $i++) {
        $s = $wb.Sheets.Item($i)
        $sheets += @{name=$s.Name; index=($i-1); active=($s.Name -eq $wb.ActiveSheet.Name)}
    }
    return @{success=$true; data=@{sheets=$sheets; count=$wb.Sheets.Count}}
}

function Exec-switchSheet($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.Sheets.Item($p.name)
    $sheet.Activate()
    return @{success=$true; data=@{activatedSheet=$p.name}}
}

function Exec-moveSheet($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.Sheets.Item($p.name)
    $sheet.Move($wb.Sheets.Item($p.position + 1))
    return @{success=$true; data=@{name=$p.name; newPosition=$p.position}}
}

function Exec-getSelection($p) {
    $app = $global:excel
    $sel = $app.Selection
    return @{success=$true; data=@{
        address=$sel.Address(); rowCount=$sel.Rows.Count; columnCount=$sel.Columns.Count
        sheet=$app.ActiveSheet.Name
    }}
}

function Exec-deleteRows($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $count = if ($p.count) { $p.count } else { 1 }
    $sheet.Range("A" + $p.row + ":A" + ($p.row + $count - 1)).Delete()
    return @{success=$true}
}

function Exec-insertColumns($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $count = if ($p.count) { $p.count } else { 1 }
    $col = ConvertToA1 $p.column
    $endCol = ConvertToA1 ($p.column + $count - 1)
    $sheet.Range($col + "1:" + $endCol + "1").Insert(-4159)
    return @{success=$true}
}

function Exec-deleteColumns($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $count = if ($p.count) { $p.count } else { 1 }
    $col = ConvertToA1 $p.column
    $endCol = ConvertToA1 ($p.column + $count - 1)
    $sheet.Range($col + "1:" + $endCol + "1").Delete()
    return @{success=$true}
}

function Exec-freezePanes($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $freeze = if ($p.freeze -ne $false) { $true } else { $false }
    if ($freeze) {
        $row = if ($p.row) { $p.row } else { 1 }
        $col = if ($p.column) { $p.column } else { 1 }
        $sheet.Cells.Item($row + 1, $col + 1).Select()
        $wb.ActiveWindow.FreezePanes = $true
    } else {
        $wb.ActiveWindow.FreezePanes = $false
    }
    return @{success=$true}
}

function Exec-createNamedRange($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $sheet.Names.Add($p.name, $sheet.Range($p.range))
    return @{success=$true; data=@{name=$p.name; range=$p.range}}
}

function Exec-hideColumns($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $col = ConvertToA1 $p.column
    $count = if ($p.count) { $p.count } else { 1 }
    $endCol = ConvertToA1 ($p.column + $count - 1)
    $range = $sheet.Range($col + "1:" + $endCol + "1")
    if ($p.hide) { $range.EntireColumn.Hidden = $true }
    else { $range.EntireColumn.Hidden = $false }
    return @{success=$true}
}

function Exec-autoSum($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = $wb.ActiveSheet
    $range = $sheet.Range($p.range)
    $sheet.Range($p.targetCell).Formula = "=SUM(" + $p.range + ")"
    return @{success=$true; data=@{range=$p.range; targetCell=$p.targetCell}}
}

function Exec-setCellFormat($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    $f = $p.format
    if ($null -ne $f.bold) { $range.Font.Bold = $f.bold }
    if ($null -ne $f.italic) { $range.Font.Italic = $f.italic }
    if ($null -ne $f.fontSize) { $range.Font.Size = $f.fontSize }
    if ($null -ne $f.fontName) { $range.Font.Name = $f.fontName }
    if ($null -ne $f.fontColor) { $range.Font.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb([Convert]::ToInt32($f.fontColor.Substring(1,2),16), [Convert]::ToInt32($f.fontColor.Substring(3,2),16), [Convert]::ToInt32($f.fontColor.Substring(5,2),16))) }
    if ($null -ne $f.bgColor) { $range.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb([Convert]::ToInt32($f.bgColor.Substring(1,2),16), [Convert]::ToInt32($f.bgColor.Substring(3,2),16), [Convert]::ToInt32($f.bgColor.Substring(5,2),16))) }
    if ($null -ne $f.underline) { $range.Font.Underline = if ($f.underline) { 2 } else { -4142 } }
    if ($null -ne $f.horizontalAlignment) {
        $align = switch ($f.horizontalAlignment) { "left" {1} "center" {-4108} "right" {-4152} default {-4131} }
        $range.HorizontalAlignment = $align
    }
    if ($null -ne $f.verticalAlignment) {
        $valign = switch ($f.verticalAlignment) { "top" {-4160} "center" {-4108} "bottom" {-4107} default {-4108} }
        $range.VerticalAlignment = $valign
    }
    if ($null -ne $f.wrapText) { $range.WrapText = $f.wrapText }
    return @{success=$true}
}

function Exec-setCellStyle($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Range($p.range).Style = $p.style
    return @{success=$true}
}

function Exec-setBorder($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    $style = switch ($p.borderStyle) { "thin" {1} "medium" {-4138} "thick" {4} "double" {-4119} "none" {-4142} default {1} }
    $color = if ($p.color) { 
        $c = $p.color
        [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb([Convert]::ToInt32($c.Substring(1,2),16), [Convert]::ToInt32($c.Substring(3,2),16), [Convert]::ToInt32($c.Substring(5,2),16)))
    } else { 0 }
    $borders = $range.Borders
    $pos = switch ($p.position) { "top" {5} "bottom" {9} "left" {7} "right" {10} "outline" {-4120} default {1} }
    if ($pos -eq 1) { $borders.Item(1).LineStyle = $style; $borders.Item(1).Color = $color }
    elseif ($pos -eq -4120) { foreach ($i in @(5,7,8,9,10)) { $borders.Item($i).LineStyle = $style; $borders.Item($i).Color = $color } }
    else { $borders.Item($pos).LineStyle = $style; $borders.Item($pos).Color = $color }
    return @{success=$true}
}

function Exec-setNumberFormat($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Range($p.range).NumberFormat = $p.format
    return @{success=$true}
}

function Exec-mergeCells($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Range($p.range).Merge()
    return @{success=$true}
}

function Exec-unmergeCells($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Range($p.range).UnMerge()
    return @{success=$true}
}

function Exec-setColumnWidth($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Columns.Item($p.column).ColumnWidth = $p.width
    return @{success=$true}
}

function Exec-setRowHeight($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Rows.Item($p.row).RowHeight = $p.height
    return @{success=$true}
}

function Exec-hideRows($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $count = if ($p.count) { $p.count } else { 1 }
    $hide = if ($p.hide -ne $false) { $true } else { $false }
    $sheet.Rows.Item($p.row.ToString() + ":" + ($p.row + $count - 1).ToString()).Hidden = $hide
    return @{success=$true}
}

function Exec-showRows($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $count = if ($p.count) { $p.count } else { 1 }
    $sheet.Rows.Item($p.row.ToString() + ":" + ($p.row + $count - 1).ToString()).Hidden = $false
    return @{success=$true}
}

function Exec-showColumns($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $count = if ($p.count) { $p.count } else { 1 }
    $col = ConvertToA1 $p.column
    $endCol = ConvertToA1 ($p.column + $count - 1)
    $sheet.Range($col + "1:" + $endCol + "1").EntireColumn.Hidden = $false
    return @{success=$true}
}

function Exec-groupRows($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Range($p.range).Rows.Group()
    return @{success=$true}
}

function Exec-addDataValidation($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    $valType = switch ($p.type) { "list" {3} "whole" {1} "decimal" {2} "date" {4} "textLength" {6} "custom" {7} default {3} }
    $range.Validation.Delete()
    $range.Validation.Add($valType, 1, [Type]::Missing, $p.formula)
    return @{success=$true}
}

function Exec-deleteCellComment($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Range($p.cell).ClearComments()
    return @{success=$true}
}

function Exec-getCellComments($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    $comments = @()
    for ($r=1; $r -le $range.Rows.Count; $r++) {
        for ($c=1; $c -le $range.Columns.Count; $c++) {
            $cell = $sheet.Cells.Item($range.Row + $r - 1, $range.Column + $c - 1)
            if ($cell.Comment) { $comments += @{cell=$cell.Address(); text=$cell.Comment.Text()} }
        }
    }
    return @{success=$true; data=@{comments=$comments}}
}

function Exec-unprotectSheet($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $pw = if ($p.password) { $p.password } else { "" }
    $sheet.Unprotect($pw)
    return @{success=$true}
}

function Exec-lockCells($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $range = $sheet.Range($p.range)
    $locked = if ($p.lock -ne $false) { $true } else { $false }
    $range.Locked = $locked
    return @{success=$true}
}

function Exec-setArrayFormula($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $sheet.Range($p.range).FormulaArray = $p.formula
    return @{success=$true}
}

function Exec-insertExcelImage($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $imgPath = if ($p.imagePath) { $p.imagePath } else { $p.path }
    $destCell = if ($p.range) { $sheet.Range($p.range) } else { $sheet.Cells.Item(1,1) }
    $shape = $sheet.Shapes.AddPicture($imgPath, 0, 1, $destCell.Left, $destCell.Top, -1, -1)
    return @{success=$true; data=@{name=$shape.Name}}
}

function Exec-setHyperlink($p) {
    $wb = $global:excel.ActiveWorkbook
    $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $wb.ActiveSheet }
    $cell = $sheet.Range($p.cell)
    $cell.Hyperlinks.Delete()
    $sheet.Hyperlinks.Add($cell, $p.address, "", "", $p.textToDisplay)
    return @{success=$true}
}

# ==================== 通用 action（各应用 COM 不同，逐应用实现） ====================
function Exec-save($p) {
    $wb = $global:excel.ActiveWorkbook
    if (-not $wb) { return @{success=$false; error="无活动工作簿"} }
    try { $wb.Save(); return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Invoke-ExcelSafeTargetWrite($targetPath, $overwrite, [scriptblock]$write) {
    if (-not (Test-Path -LiteralPath $targetPath)) {
        try {
            & $write
            return @{success=$true}
        } catch {
            return @{success=$false; code="TARGET_WRITE_FAILED"; error=$_.Exception.Message}
        }
    }
    if (-not $overwrite) {
        return @{success=$false; code="TARGET_EXISTS"; error="目标文件已存在: $targetPath"}
    }

    try {
        $fullPath = [System.IO.Path]::GetFullPath($targetPath)
        $directory = [System.IO.Path]::GetDirectoryName($fullPath)
        $backupPath = [System.IO.Path]::Combine(
            $directory,
            "." + [System.IO.Path]::GetFileName($fullPath) + "." + [Guid]::NewGuid() + ".wps-backup"
        )
        [System.IO.File]::Copy($fullPath, $backupPath, $false)
    } catch {
        return @{success=$false; code="OVERWRITE_NOT_SAFE"; error="无法安全备份现有目标: $($_.Exception.Message)"}
    }

    $keepBackup = $true
    try {
        & $write
        $keepBackup = $false
        return @{success=$true}
    } catch {
        $writeError = $_.Exception.Message
        try {
            [System.IO.File]::Copy($backupPath, $fullPath, $true)
            $keepBackup = $false
            return @{success=$false; code="TARGET_WRITE_FAILED"; error="目标写入失败，已恢复原文件: $writeError"}
        } catch {
            return @{success=$false; code="OVERWRITE_RESTORE_FAILED"; error="目标写入失败且无法恢复；备份保留在 $backupPath: $writeError"}
        }
    } finally {
        if (-not $keepBackup) {
            try { [System.IO.File]::Delete($backupPath) } catch {}
        }
    }
}

function Exec-saveAs($p) {
    $wb = $global:excel.ActiveWorkbook
    if (-not $wb) { return @{success=$false; error="无活动工作簿"} }
    $writeResult = Invoke-ExcelSafeTargetWrite $p.filePath $p.overwrite { $wb.SaveAs($p.filePath) }
    if (-not $writeResult.success) { return $writeResult }
    $sz = (Get-Item -LiteralPath $p.filePath).Length
    return @{success=$true; data=@{path=$p.filePath; size=$sz}}
}
function Exec-convertToPDF($p) {
    $wb = $global:excel.ActiveWorkbook
    if (-not $wb) { return @{success=$false; error="无活动工作簿"} }
    $out = if ($p.outputPath) { $p.outputPath } else { $wb.Path + '\' + $wb.Name + '.pdf' }
    $writeResult = Invoke-ExcelSafeTargetWrite $out $p.overwrite { $wb.ExportAsFixedFormat(0, $out) }
    if (-not $writeResult.success) { return $writeResult }
    return @{success=$true; data=@{path=$out}}
}
function Exec-convertFormat($p) {
    $wb = $global:excel.ActiveWorkbook
    if (-not $wb) { return @{success=$false; error="无活动工作簿"} }
    $out = if ($p.outputPath) { $p.outputPath } else { $wb.Path + '\' + $wb.Name + '.' + $p.targetFormat }
    $writeResult = Invoke-ExcelSafeTargetWrite $out $p.overwrite { $wb.SaveAs($out) }
    if (-not $writeResult.success) { return $writeResult }
    return @{success=$true; data=@{path=$out}}
}
function Exec-reconnect($p) {
    try { $global:excel = $null; $global:excel = Get-ExcelApp } catch {}
    if ($global:excel) {
        try { $global:excel.Visible = $true } catch {}
        try { $global:excel.DisplayAlerts = $false } catch {}
        return @{success=$true; data=@{ready=$true; app="WPS表格"}}
    }
    return @{success=$false; error="重连失败：WPS 表格可能已退出，请先打开 WPS 表格"}
}
function Exec-getSelectedText($p) { try { return @{success=$true; data=@{text=$global:excel.Selection.Text}} } catch { return @{success=$false; error=$_.Exception.Message} } }
function Exec-setSelectedText($p) { try { $global:excel.Selection.Text = $p.text; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} } }
function Exec-getAppInfo($p) { try { return @{success=$true; data=@{app="WPS表格"; version=$global:excel.Version}} } catch { return @{success=$false; error=$_.Exception.Message} } }

function Test-ExcelActionRequiresActiveWorkbook($action) {
    return $global:ExcelActionsWithoutActiveWorkbook -notcontains $action
}

function Get-ExcelActiveWorkbook {
    try { return $global:excel.ActiveWorkbook } catch { return $null }
}

# ==================== 主循环（必须放在所有函数定义之后）====================
# 从 stdin 读取 JSON 命令，执行后输出带 reqId 的 JSON 结果。
# 放在末尾是因为 PowerShell 只在执行到函数定义语句时才创建该函数，
# 若主循环早于函数定义运行，调用 Exec-* 会报“无法识别”的错误。
while ($true) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { break }
    $line = $line.Trim()
    if ($line -eq "EXIT") { break }
    if ($line -eq "") { continue }

    $rid = $null
    $attempt = $null
    $traceId = $null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $cmd = $line | ConvertFrom-Json
        $action = $cmd.action
        $params = $cmd.params
        $rid = $cmd.reqId
        $attempt = $cmd.attempt
        $traceId = $cmd.traceId

        if ((Test-ExcelActionRequiresActiveWorkbook $action) -and -not (Get-ExcelActiveWorkbook)) {
            $result = @{
                success=$false
                code="NO_ACTIVE_DOCUMENT"
                error="没有活动工作簿；请先创建或打开工作簿"
            }
        } else {
            $result = & "Exec-$action" $params
        }
        $sw.Stop()
        # 规范化结果，确保始终是带 reqId 的 hashtable，避免 ConvertTo-Json 产出空串/截断
        if ($null -eq $result) { $result = @{success=$true; data=$null} }
        if ($result -isnot [hashtable]) { $result = @{success=$true; data=$result} }
        $result['reqId'] = $rid
        $result['_trace'] = @{traceId=$traceId; backendElapsedMs=[Math]::Round($sw.Elapsed.TotalMilliseconds, 2); attempt=$attempt}
        Write-Host ($result | ConvertTo-Json -Depth 20 -Compress)
    } catch {
        $sw.Stop()
        Write-Host (@{success=$false; error=($_.Exception.Message -replace '"','\"'); reqId=$rid; _trace=@{traceId=$traceId; backendElapsedMs=[Math]::Round($sw.Elapsed.TotalMilliseconds, 2); attempt=$attempt}} | ConvertTo-Json -Depth 20 -Compress)
    }
    [Console]::Out.Flush()
}
'''


class WpsExcelController:
    """WPS Excel 控制器 - 通过 COM 自动化操作 WPS Excel"""

    def __init__(self, trace=None, deadline=None):
        self.platform = platform.system()
        self._ps_process = None
        self._ready = False
        # 子进程 stdout 读取队列与后台读线程（用于超时强杀与 reqId 关联）
        self._queue = None
        self._reader = None
        self._stderr_queue = None
        self._stderr_reader = None
        self._stop = None
        self._id_counter = 0
        self._id_lock = threading.Lock()
        self._linux_bridge = None
        self._linux_error = ""

        if IS_WINDOWS:
            self._init_windows(trace=trace, deadline=deadline)
        elif IS_LINUX:
            self._init_linux(trace=trace)
        else:
            raise RuntimeError(f"不支持的平台: {self.platform}")

    # ==================== Windows: PowerShell COM 桥接 ====================

    def _init_windows(self, trace=None, deadline=None):
        """初始化 Windows PowerShell COM 桥接进程"""
        started = time.perf_counter()
        try:
            com_runtime = resolve_com_runtime(EXCEL_PROGID)
            if not com_runtime.available:
                raise RuntimeError(
                    f"{EXCEL_PROGID} COM 注册不完整: {com_runtime.diagnostic}"
                )
            selected_registration = com_runtime.selected_registration
            if trace:
                trace.event(
                    "powershell.process.starting",
                    app="excel",
                    progId=EXCEL_PROGID,
                    clsid=(selected_registration.clsid if selected_registration else None),
                    registryViewBits=com_runtime.selected_view_bitness,
                    powershellExecutable=com_runtime.powershell_executable,
                )
            # 写入临时 PS1 脚本
            # 注意：必须带 BOM(utf-8-sig)。PowerShell 5.1 在中文 Windows 上默认按系统 ANSI
            # (GBK) 解析无 BOM 的 .ps1，脚本中的中文会被读成乱码从而破坏字符串字面量，
            # 导致整脚本语法解析失败、桥接无法就绪。带 BOM 后 PS 才按 UTF-8 正确解析。
            self._ps_script = tempfile.NamedTemporaryFile(
                mode='w', suffix='.ps1', delete=False, encoding='utf-8-sig'
            )
            self._ps_script.write(PS_BRIDGE_SCRIPT)
            self._ps_script.close()

            # 启动持久 PowerShell 进程
            self._ps_process = subprocess.Popen(
                [com_runtime.powershell_executable, '-NoProfile', '-NoLogo',
                 '-ExecutionPolicy', 'Bypass', '-File', self._ps_script.name],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8',
                errors='replace',
                bufsize=1
            )

            self._stop = threading.Event()
            self._stderr_queue = queue.Queue()
            self._stderr_reader = threading.Thread(target=self._stderr_loop, daemon=True)
            self._stderr_reader.start()

            # 读取就绪信号（带超时保护，避免 PS 卡死导致永久阻塞）
            ready_line = self._read_line_with_timeout(
                self._ps_process.stdout,
                bounded_timeout(30, deadline),
            )
            ready = json.loads(ready_line) if ready_line else {}
            self._ready = ready.get('ready', False)
            if not self._ready:
                err = describe_powershell_startup_failure(
                    self._ps_process,
                    self._stderr_reader,
                    self._stderr_queue,
                    ready.get('error'),
                )
                self._kill_ps()
                raise RuntimeError(f"WPS Excel 连接失败: {err}")

            # 启动后台读线程，持续把 stdout 行放入队列（供 _exec_windows 按 reqId 取回执）
            self._queue = queue.Queue()
            self._reader = threading.Thread(target=self._reader_loop, daemon=True)
            self._reader.start()
            if trace:
                trace.event(
                    "powershell.process.ready",
                    app="excel",
                    progId=EXCEL_PROGID,
                    processPid=self._ps_process.pid,
                    elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                )

        except Exception as e:
            self._ready = False
            if trace:
                trace.event(
                    "powershell.process.failed",
                    status="error",
                    app="excel",
                    progId=EXCEL_PROGID,
                    error=f"{type(e).__name__}: {e}",
                    elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                )
            self._kill_ps()
            raise RuntimeError(f"PowerShell 桥接初始化失败: {e}")

    # ==================== 子进程 stdout 读取与超时 ====================

    def _read_line_with_timeout(self, stream, timeout):
        """在独立线程中读取一行，超时返回空串"""
        q = queue.Queue()

        def _reader():
            try:
                q.put(stream.readline())
            except Exception:
                q.put('')

        t = threading.Thread(target=_reader, daemon=True)
        t.start()
        t.join(timeout)
        if t.is_alive():
            return ''
        try:
            return q.get_nowait()
        except queue.Empty:
            return ''

    def _reader_loop(self):
        """后台持续读取 PS stdout 行到队列"""
        try:
            while not self._stop.is_set() and self._ps_process is not None:
                line = self._ps_process.stdout.readline()
                if not line:
                    break
                self._queue.put(line)
        except Exception:
            pass

    def _stderr_loop(self):
        """持续排空 PowerShell stderr，避免管道写满后阻塞。"""
        try:
            while not self._stop.is_set() and self._ps_process is not None:
                line = self._ps_process.stderr.readline()
                if not line:
                    break
                self._stderr_queue.put(line)
        except Exception:
            pass

    def _drain_stderr(self, trace=None, attempt=None):
        if getattr(self, "_stderr_queue", None) is None:
            return
        while True:
            try:
                raw = self._stderr_queue.get_nowait().strip()
            except queue.Empty:
                break
            if raw and trace:
                trace.event(
                    "powershell.stderr",
                    status="error",
                    app="excel",
                    attempt=attempt,
                    error=raw,
                )

    def _next_id(self):
        with self._id_lock:
            self._id_counter += 1
            return self._id_counter

    def _kill_ps(self):
        """Stop PowerShell with the bounded line-process cleanup protocol."""
        process = self._ps_process
        self._ps_process = None
        if self._stop is not None:
            self._stop.set()
        stop_line_process(process)

    def _read_result(self, req_id, trace=None, attempt=1, started=None, deadline=None):
        started = started or time.perf_counter()
        timeout = bounded_timeout(EXEC_TIMEOUT, deadline)
        response_deadline = time.monotonic() + timeout
        while True:
            remaining = response_deadline - time.monotonic()
            if remaining <= 0:
                self._ready = False
                self._drain_stderr(trace, attempt)
                if trace:
                    trace.event(
                        "powershell.response.timeout",
                        status="timeout",
                        app="excel",
                        reqId=req_id,
                        attempt=attempt,
                        elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                    )
                self._kill_ps()
                return {
                    "success": False,
                    "code": "ACTION_EXECUTION_TIMEOUT",
                    "error": f"action 执行超时（>{EXEC_TIMEOUT}s），已终止 WPS 桥接进程",
                }
            try:
                raw = self._queue.get(timeout=remaining)
            except queue.Empty:
                self._ready = False
                self._drain_stderr(trace, attempt)
                if trace:
                    trace.event(
                        "powershell.response.timeout",
                        status="timeout",
                        app="excel",
                        reqId=req_id,
                        attempt=attempt,
                        elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                    )
                self._kill_ps()
                return {
                    "success": False,
                    "code": "ACTION_EXECUTION_TIMEOUT",
                    "error": f"action 执行超时（>{EXEC_TIMEOUT}s）",
                }
            raw = raw.strip()
            if not raw:
                continue
            try:
                obj = json.loads(raw)
            except Exception:
                if trace:
                    trace.event(
                        "powershell.stdout.noise",
                        status="warning",
                        app="excel",
                        reqId=req_id,
                        attempt=attempt,
                        lineLength=len(raw),
                        **trace.debug_fields(stdout=raw),
                    )
                continue
            if obj.get("reqId") == req_id:
                backend_trace = obj.pop("_trace", {}) or {}
                self._drain_stderr(trace, attempt)
                if trace:
                    trace.event(
                        "powershell.response.received",
                        status="success" if obj.get("success") else "error",
                        app="excel",
                        reqId=req_id,
                        attempt=attempt,
                        error=obj.get("error"),
                        backendTraceId=backend_trace.get("traceId"),
                        backendAttempt=backend_trace.get("attempt"),
                        backendElapsedMs=backend_trace.get("backendElapsedMs"),
                        elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                        **trace.debug_fields(response=obj),
                    )
                return obj
            if trace:
                trace.event(
                    "powershell.response.mismatched",
                    status="warning",
                    app="excel",
                    reqId=req_id,
                    receivedReqId=obj.get("reqId"),
                    attempt=attempt,
                )
            continue

    def _run_windows_attempt(
        self, action, params, attempt, trace=None, deadline=None, correlation_id=None,
    ):
        req_id = correlation_id if correlation_id is not None else self._next_id()
        command = {
            "reqId": req_id,
            "traceId": trace.trace_id if trace else None,
            "attempt": attempt,
            "action": action,
            "params": params,
        }
        cmd = json.dumps(command, ensure_ascii=True)
        self._drain_stderr(trace, attempt)
        started = time.perf_counter()
        if trace:
            trace.event(
                "powershell.request.sent",
                app="excel",
                action=action,
                reqId=req_id,
                attempt=attempt,
                processPid=getattr(self._ps_process, "pid", None),
                **trace.debug_fields(params=params),
            )
        try:
            self._ps_process.stdin.write(cmd + "\n")
            self._ps_process.stdin.flush()
        except Exception as e:
            self._ready = False
            if trace:
                trace.event(
                    "powershell.request.failed",
                    status="error",
                    app="excel",
                    action=action,
                    reqId=req_id,
                    attempt=attempt,
                    error=f"{type(e).__name__}: {e}",
                )
            self._kill_ps()
            return {"success": False, "error": f"发送命令失败: {e}"}
        return self._read_result(
            req_id,
            trace=trace,
            attempt=attempt,
            started=started,
            deadline=deadline,
        )

    def _exec_windows(
        self, action: str, params: dict, trace=None, deadline=None, correlation_id=None,
    ) -> dict:
        """通过 PowerShell COM 执行一次 Action。"""
        if not self._ps_process or self._ps_process.poll() is not None:
            self._ready = False
            if trace:
                trace.event(
                    "powershell.process.unavailable",
                    status="error",
                    app="excel",
                    action=action,
                )
            return {"success": False, "error": "PowerShell 进程已退出"}

        return self._run_windows_attempt(
            action, params, attempt=1, trace=trace, deadline=deadline,
            correlation_id=correlation_id,
        )

    # ==================== Linux: 文件级后端（vendored openpyxl） ====================

    def _init_linux(self, trace=None):
        """初始化 Linux 桥接：懒加载 LinuxExcelBridge（openpyxl 文件级后端）"""
        # 保证 bridge 目录在 sys.path 中（以任意 cwd 启动均可）
        bridge_dir = os.path.dirname(os.path.abspath(__file__))
        if bridge_dir not in sys.path:
            sys.path.insert(0, bridge_dir)
        try:
            from linux_excel import LinuxExcelBridge
        except Exception as e:
            self._ready = False
            self._linux_bridge = None
            self._linux_error = f"Linux 后端加载失败（检查 vendor/openpyxl 是否完整）: {e}"
            if trace:
                trace.event("linux.backend.init.failed", status="error", app="excel", error=self._linux_error)
            return
        self._linux_bridge = LinuxExcelBridge()
        self._ready = True
        if trace:
            trace.event("linux.backend.ready", app="excel", backend="openpyxl")

    def _exec_linux(self, action: str, params: dict, trace=None) -> dict:
        """Linux 上委托给 LinuxExcelBridge（与 Windows 同一套 action 契约）"""
        if getattr(self, "_linux_bridge", None) is None:
            self._init_linux(trace=trace)
        if getattr(self, "_linux_bridge", None) is None:
            return {"success": False, "error": self._linux_error}
        try:
            result = self._linux_bridge.execute(action, params)
            if trace:
                trace.event(
                    "linux.backend.completed",
                    status="success" if result.get("success") else "error",
                    app="excel",
                    action=action,
                    error=result.get("error"),
                )
            return result
        except Exception as e:
            return {"success": False, "error": f"Linux 后端执行异常: {e}"}

    # ==================== 统一执行入口 ====================

    def execute(
        self, action: str, params: dict = None, trace=None, deadline=None, correlation_id=None,
    ) -> dict:
        """统一执行入口 - 所有 action 通过此处调用"""
        if params is None:
            params = {}

        # Runtime owns Windows controller recreation and retry policy.
        if not self._ready and action != "ping":
            try:
                if IS_LINUX:
                    self._init_linux(trace=trace)
            except Exception as e:
                return {"success": False, "error": f"WPS Excel 未连接，且重连失败: {e}"}

        if not self._ready and action != "ping":
            return {"success": False, "error": "WPS Excel 未连接"}

        if IS_WINDOWS:
            return self._exec_windows(
                action, params, trace=trace, deadline=deadline,
                correlation_id=correlation_id,
            )
        elif IS_LINUX:
            return self._exec_linux(action, params, trace=trace)
        else:
            return {"success": False, "error": f"不支持的平台: {self.platform}"}

    def ping(self, trace=None, deadline=None) -> bool:
        """检测 WPS 连接状态"""
        try:
            result = self.execute("ping", trace=trace, deadline=deadline)
            return result.get("success", False)
        except Exception:
            return False

    def close(self):
        """关闭连接并清理资源"""
        self._ready = False
        stop_event = getattr(self, "_stop", None)
        if stop_event is not None:
            stop_event.set()
        process = getattr(self, "_ps_process", None)
        self._ps_process = None
        stop_line_process(process)

        # 清理临时脚本
        if hasattr(self, '_ps_script') and self._ps_script:
            try:
                os.unlink(self._ps_script.name)
            except Exception:
                pass

    def __del__(self):
        self.close()


# ==================== 单例 ====================
_controller = None

def get_controller(trace=None, deadline=None) -> WpsExcelController:
    """获取单例控制器"""
    global _controller
    if _controller is None or not _controller._ready:
        _controller = WpsExcelController(trace=trace, deadline=deadline)
    return _controller


# ==================== 测试入口 ====================
if __name__ == "__main__":
    print("WPS Excel 控制器测试")
    print(f"平台: {platform.system()}")
    print()

    try:
        ctrl = WpsExcelController()
        print(f"连接就绪: {ctrl._ready}")
        print()

        # 测试 ping
        print("=== ping ===")
        print(ctrl.execute("ping"))
        print()

        # 测试获取工作簿
        print("=== getActiveWorkbook ===")
        print(ctrl.execute("getActiveWorkbook"))
        print()

        # 测试获取工作表列表
        print("=== getSheetList ===")
        print(ctrl.execute("getSheetList"))
        print()

        # 测试获取上下文
        print("=== getContext ===")
        print(ctrl.execute("getContext"))
        print()

    except Exception as e:
        print(f"错误: {e}")
    finally:
        if '_controller' in globals() and _controller:
            _controller.close()
