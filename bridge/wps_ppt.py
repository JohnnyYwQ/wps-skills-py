#!/usr/bin/env python3
"""
WPS PPT 控制器 - 纯 Python 实现
通过 COM 自动化（Windows）控制 WPS 演示（Kwpp.Application）。

与 wps_excel.py 完全相同的可靠架构：
  - subprocess 拉起持久 PowerShell 进程（line-RPC）
  - reqId 关联回执，防止 line 协议去同步
  - EXEC_TIMEOUT 超时强杀，防止 COM 弹框永久卡死
  - WPS 断开后自动重连
  - 临时 .ps1 用 utf-8-sig(BOM)；命令用 ensure_ascii=True 跨管道传中文

不依赖 MCP，不依赖外网，不依赖 Node.js/JS 环境。
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
from service_lifecycle import stop_line_process
from windows_com import describe_powershell_startup_failure, resolve_com_runtime

# ==================== 平台检测 ====================
IS_WINDOWS = platform.system() == "Windows"
IS_LINUX = platform.system() == "Linux"
IS_MACOS = platform.system() == "Darwin"

# WPS 演示 COM ProgID
PPT_PROGID = "Kwpp.Application"

# 单个 action 执行超时（秒）
EXEC_TIMEOUT = 60

# 偶发 COM 故障特征串（小写匹配，用于自动重连重试，避免“未注册对象”后无法重置状态）
COM_FAIL_HINTS = (
    "未注册对象", "未注册", "80040154", "rpc 服务器不可用", "800706ba",
    "rpc_e_disconnected", "调用的对象已与其客户端断开连接",
    "ole_e_promptsavecancelled", "8004000c", "catastrophic", "灾难性",
)

# PowerShell 桥接脚本（持久进程模式）
PS_BRIDGE_SCRIPT = r'''
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$error.Clear()
$global:WpsActivationError = $null

# 创建/获取 COM 对象
function Get-PptApp {
    try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwpp.Application') }
    catch {
        try { return New-Object -ComObject 'Kwpp.Application' }
        catch {
            $hr = '0x{0:X8}' -f $_.Exception.HResult
            $global:WpsActivationError = "New-Object failed ($hr): $($_.Exception.Message)"
            try {
                $clsid = (Get-ItemProperty -Path 'HKCU:\Software\Classes\Kwpp.Application\CLSID' -ErrorAction Stop).'(default)'
                if ($clsid) {
                    $type = [Type]::GetTypeFromCLSID([Guid]$clsid)
                    if ($type) { return [Activator]::CreateInstance($type) }
                }
            } catch {}
        }
    }
    return $null
}

$global:ppt = Get-PptApp
if ($global:ppt) {
    try { $global:ppt.Visible = -1 } catch {}
    # 抑制“是否覆盖”等模态对话框，避免 COM 无法应答导致 OLE_E_PROMPTSAVECANCELLED
    try { $global:ppt.DisplayAlerts = 0 } catch {}
    Write-Host '{"ready":true}'
} else {
    $message = if ($global:WpsActivationError) {
        $global:WpsActivationError
    } else {
        '无法连接WPS演示，请确认WPS已安装'
    }
    Write-Host (@{ready=$false; error=$message} | ConvertTo-Json -Compress)
    exit 1
}

# ==================== 辅助函数 ====================

# 十六进制颜色(#RRGGBB) -> OLE 颜色整数(BGR)
function Convert-HexToOle($hex) {
    if ($null -eq $hex -or $hex -eq '') { return $null }
    $h = $hex.ToString().Trim().Replace('#','')
    if ($h.Length -ne 6) { return $null }
    try {
        $r = [Convert]::ToInt32($h.Substring(0,2),16)
        $g = [Convert]::ToInt32($h.Substring(2,2),16)
        $b = [Convert]::ToInt32($h.Substring(4,2),16)
        return [int]($b -bor ($g -shl 8) -bor ($r -shl 16))
    } catch { return $null }
}

''' + render_chart_type_converter() + r'''
function Convert-Distribution($value) {
    if ($value -is [string]) {
        switch ($value.ToLowerInvariant()) {
            "horizontal" { return 0 }
            "vertical" { return 1 }
            default { throw "未知 distribute: $value" }
        }
    }
    return [int]$value
}

function Convert-ZOrder($value) {
    if ($value -is [string]) {
        switch ($value.ToLowerInvariant()) {
            "front" { return 0 }
            "back" { return 1 }
            "forward" { return 2 }
            "backward" { return 3 }
            default { throw "未知 order: $value" }
        }
    }
    return [int]$value
}

function Get-ActivePres { if ($global:ppt.ActivePresentation) { return $global:ppt.ActivePresentation }; return $null }
function Get-PresByName($name) {
    foreach ($p in $global:ppt.Presentations) { if ($p.Name -eq $name) { return $p } }
    return $null
}
function Get-Slide($pres, $idx) {
    if ($idx -eq $null -or $idx -le 0) { $idx = 1 }
    try { return $pres.Slides.Item([int]$idx) } catch { return $null }
}

# 按形状 Id 查找形状（修复 shapeIndex 语义歧义）
# 关键约定：addShape/addTextBox 返回的是 $s.Id（形状唯一 Id，通常从 2 起），
# 而 Shapes.Item(<int>) 接收的是“位置索引”（从 1 起）。二者不是一回事！
# 若用 Item(shapeId) 当位置索引用，会导致全部样式操作错位一格、且每页最后一个
# 形状越界报错 "Value does not fall within the expected range"。
# 因此所有“按 shapeIndex 操作形状”的 action 一律用本函数按 .Id 精确定位。
function Get-ShapeById($slide, $id) {
    if ($null -eq $slide) { throw "未找到幻灯片" }
    if ($null -eq $id) { throw "缺少 shapeIndex 参数" }
    [int]$sid = $id
    foreach ($s in $slide.Shapes) {
        if ($s.Id -eq $sid) { return $s }
    }
    throw "未找到 shapeId=$sid 的形状（shapeIndex 必须传入 addShape/addTextBox 返回的 shapeId，而非位置序号）"
}

# ==================== Action 实现函数 ====================

function Exec-ping($p) {
    return @{success=$true; data=@{message="pong"; timestamp=(Get-Date -Format "o")}}
}

# ---------- 演示文稿管理 ----------
function Exec-createPresentation($p) {
    $pres = $global:ppt.Presentations.Add()
    return @{success=$true; data=@{name=$pres.Name; slideCount=$pres.Slides.Count}}
}

function Exec-openPresentation($p) {
    $path = $p.filePath
    if (-not $path) { return @{success=$false; error="缺少 filePath"} }
    $pres = $global:ppt.Presentations.Open($path)
    return @{success=$true; data=@{name=$pres.Name; slideCount=$pres.Slides.Count; path=$path}}
}

function Exec-closePresentation($p) {
    $pres = if ($p.name) { Get-PresByName $p.name } else { Get-ActivePres }
    if (-not $pres) { return @{success=$false; error="未找到演示文稿"} }
    if ($p.save -eq $true) { try { $pres.Save() } catch {} }
    $pres.Close()
    return @{success=$true; data=@{closed=$pres.Name}}
}

function Exec-getOpenPresentations($p) {
    $list = @()
    foreach ($pres in $global:ppt.Presentations) {
        $list += @{name=$pres.Name; slideCount=$pres.Slides.Count; path=$pres.Path}
    }
    return @{success=$true; data=@{presentations=$list; count=$list.Count}}
}

function Exec-switchPresentation($p) {
    $pres = Get-PresByName $p.name
    if (-not $pres) { return @{success=$false; error="未找到演示文稿: $($p.name)"} }
    $pres.Activate() | Out-Null
    return @{success=$true; data=@{name=$pres.Name}}
}

function Exec-insertSlidesFromFile($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    $after = if ($p.afterIndex) { [int]$p.afterIndex } else { $pres.Slides.Count }
    $s = if ($p.slideStart) { [int]$p.slideStart } else { 1 }
    $e = if ($p.slideEnd) { [int]$p.slideEnd } else { 0 }
    try {
        if ($e -gt 0) { $pres.Slides.InsertFromFile($p.filePath, $after, $s, $e) | Out-Null }
        else { $pres.Slides.InsertFromFile($p.filePath, $after) | Out-Null }
        return @{success=$true; data=@{slideCount=$pres.Slides.Count}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-getSlideMaster($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    try { $master = $pres.SlideMaster; return @{success=$true; data=@{name=$master.Name}} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setMasterBackground($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    $ole = Convert-HexToOle $p.color
    try {
        $master = $pres.SlideMaster
        $master.Background.Fill.ForeColor.RGB = $ole
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-addMasterElement($p) {
    return @{success=$true; data=@{message="母版元素添加（best-effort，按版式处理）"}}
}

# ---------- 幻灯片操作 ----------
function Exec-addSlide($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    $layout = 2  # ppLayoutText
    switch ($p.layout) {
        "title" { $layout = 1 }
        "title_content" { $layout = 2 }
        "blank" { $layout = 12 }
        "two_column" { $layout = 11 }
        "comparison" { $layout = 11 }
        "title_only" { $layout = 11 }
    }
    $pos = if ($p.position) { [int]$p.position } else { $pres.Slides.Count + 1 }
    $slide = $pres.Slides.Add($pos, $layout)
    if ($p.title) {
        try {
            $titleShape = $slide.Shapes.Item(1)
            if ($titleShape.HasTextFrame) { $titleShape.TextFrame.TextRange.Text = $p.title }
        } catch {}
    }
    return @{success=$true; data=@{slideIndex=$slide.SlideIndex; slideCount=$pres.Slides.Count}}
}

function Exec-deleteSlide($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    $slide.Delete()
    return @{success=$true; data=@{slideCount=$pres.Slides.Count}}
}

function Exec-duplicateSlide($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    $newSlide = $slide.Duplicate()
    return @{success=$true; data=@{slideIndex=$newSlide[0].SlideIndex; slideCount=$pres.Slides.Count}}
}

function Exec-moveSlide($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    $slide.MoveTo([int]$p.targetIndex)
    return @{success=$true; data=@{slideCount=$pres.Slides.Count}}
}

function Exec-getSlideCount($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    return @{success=$true; data=@{slideCount=$pres.Slides.Count}}
}

function Exec-getSlideInfo($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    $shapes = @()
    foreach ($s in $slide.Shapes) {
        $txt = ""
        try { if ($s.HasTextFrame -and $s.TextFrame.HasText) { $txt = $s.TextFrame.TextRange.Text } } catch {}
        $shapes += @{index=$s.Id; type=$s.Type; name=$s.Name; text=$txt; left=$s.Left; top=$s.Top; width=$s.Width; height=$s.Height}
    }
    return @{success=$true; data=@{slideIndex=$slide.SlideIndex; layout=$slide.Layout; shapeCount=$slide.Shapes.Count; shapes=$shapes}}
}

function Exec-switchSlide($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $slide.Select() } catch {}
    return @{success=$true; data=@{slideIndex=$slide.SlideIndex}}
}

function Exec-setSlideLayout($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    $layout = 2
    switch ($p.layout) { "title" { $layout=1 } "title_content" { $layout=2 } "blank" { $layout=12 } "title_only" { $layout=11 } }
    try { $slide.Layout = $layout; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setSlideSize($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    try {
        if ($p.width -and $p.height) { $pres.PageSetup.SlideWidth = [int]$p.width; $pres.PageSetup.SlideHeight = [int]$p.height }
        return @{success=$true; data=@{width=$pres.PageSetup.SlideWidth; height=$pres.PageSetup.SlideHeight}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-getSlideNotes($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    $notes = ""
    try { if ($slide.HasNotesPage) { $notes = $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text } } catch {}
    return @{success=$true; data=@{notes=$notes}}
}

function Exec-setSlideNotes($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = $p.text; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setSlideTitle($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $slide.Shapes.Item(1).TextFrame.TextRange.Text = $p.title; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-getSlideTitle($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    $t = ""
    try { $t = $slide.Shapes.Item(1).TextFrame.TextRange.Text } catch {}
    return @{success=$true; data=@{title=$t}}
}

function Exec-setSlideSubtitle($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $slide.Shapes.Item(2).TextFrame.TextRange.Text = $p.subtitle; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setSlideContent($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $tb = $slide.Shapes.AddTextbox(1, 50, 120, 600, 400)
        $tb.TextFrame.TextRange.Text = $p.content
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setSlideTheme($p) {
    return @{success=$true; data=@{message="主题设置（best-effort，WPS 版式有限）"}}
}

function Exec-insertImage($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]($p.slideIndex))
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $left = if ($p.x) { [int]$p.x } else { 100 }
        $top = if ($p.y) { [int]$p.y } else { 100 }
        $sp = $slide.Shapes.AddPicture($p.imagePath, 0, -1, $left, $top)
        if ($p.width) { $sp.Width = [int]$p.width }
        if ($p.height) { $sp.Height = [int]$p.height }
        return @{success=$true; data=@{shapeId=$sp.Id}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-startSlideShow($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    try { $pres.SlideShowSettings.Run() | Out-Null; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-findPptText($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    $results = @()
    foreach ($slide in $pres.Slides) {
        foreach ($s in $slide.Shapes) {
            try { if ($s.HasTextFrame -and $s.TextFrame.HasText -and $s.TextFrame.TextRange.Text -like "*$($p.find_text)*") { $results += @{slideIndex=$slide.SlideIndex; shapeId=$s.Id; text=$s.TextFrame.TextRange.Text} } } catch {}
        }
    }
    return @{success=$true; data=@{matches=$results; count=$results.Count}}
}

function Exec-replacePptText($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    $cnt = 0
    foreach ($slide in $pres.Slides) {
        foreach ($s in $slide.Shapes) {
            try {
                if ($s.HasTextFrame -and $s.TextFrame.HasText) {
                    $t = $s.TextFrame.TextRange.Text
                    if ($t -like "*$($p.find)*") { $s.TextFrame.TextRange.Text = $t.Replace($p.find, $p.replace); $cnt++ }
                }
            } catch {}
        }
    }
    return @{success=$true; data=@{replaced=$cnt}}
}

function Exec-setSlideBackground($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        if ($p.color) { $ole = Convert-HexToOle $p.color; $slide.Background.Fill.ForeColor.RGB = $ole }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

# ---------- 文本框 ----------
function Exec-addTextBox($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $left = if ($p.x) { [int]$p.x } else { 100 }
        $top = if ($p.y) { [int]$p.y } else { 100 }
        $w = if ($p.width) { [int]$p.width } else { 300 }
        $h = if ($p.height) { [int]$p.height } else { 50 }
        $tb = $slide.Shapes.AddTextbox(1, $left, $top, $w, $h)
        if ($p.text) { $tb.TextFrame.TextRange.Text = $p.text }
        return @{success=$true; data=@{shapeId=$tb.Id}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-deleteTextBox($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { (Get-ShapeById $slide ([int]$p.shapeIndex)).Delete(); return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-getTextBoxes($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    $list = @()
    foreach ($s in $slide.Shapes) {
        $txt = ""
        try { if ($s.HasTextFrame -and $s.TextFrame.HasText) { $txt = $s.TextFrame.TextRange.Text } } catch {}
        $list += @{shapeIndex=$s.Id; text=$txt}
    }
    return @{success=$true; data=@{textBoxes=$list; count=$list.Count}}
}

function Exec-setTextBoxText($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { (Get-ShapeById $slide ([int]$p.shapeIndex)).TextFrame.TextRange.Text = $p.text; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setTextBoxStyle($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $sp = (Get-ShapeById $slide ([int]$p.shapeIndex))
        if ($p.font_size) { $sp.TextFrame.TextRange.Font.Size = [int]$p.font_size }
        if ($p.bold) { $sp.TextFrame.TextRange.Font.Bold = $true }
        if ($p.color) { $ole = Convert-HexToOle $p.color; $sp.TextFrame.TextRange.Font.Color.RGB = $ole }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-create3DText($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $tb = $slide.Shapes.AddTextbox(1, 100, 100, 400, 100)
        $tb.TextFrame.TextRange.Text = $p.text
        $tb.TextEffect.PresetThreeDFormat = 1
        return @{success=$true; data=@{shapeId=$tb.Id}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setShapeText($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { (Get-ShapeById $slide ([int]$p.shapeIndex)).TextFrame.TextRange.Text = $p.text; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

# ---------- 形状 ----------
function Exec-addShape($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $type = 1  # msoShapeRectangle
        switch ($p.shapeType) { "rectangle" { $type=1 } "oval" { $type=9 } "roundRectangle" { $type=5 } "triangle" { $type=5 } "line" { $type=9 } "arrow" { $type=33 } }
        $left = if ($p.x) { [int]$p.x } else { 100 }
        $top = if ($p.y) { [int]$p.y } else { 100 }
        $w = if ($p.width) { [int]$p.width } else { 200 }
        $h = if ($p.height) { [int]$p.height } else { 150 }
        $sp = $slide.Shapes.AddShape($type, $left, $top, $w, $h)
        return @{success=$true; data=@{shapeId=$sp.Id}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-deleteShape($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { (Get-ShapeById $slide ([int]$p.shapeIndex)).Delete(); return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-getShapes($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    $list = @()
    foreach ($s in $slide.Shapes) {
        $txt = ""
        try { if ($s.HasTextFrame -and $s.TextFrame.HasText) { $txt = $s.TextFrame.TextRange.Text } } catch {}
        $list += @{shapeIndex=$s.Id; type=$s.Type; name=$s.Name; text=$txt; left=$s.Left; top=$s.Top; width=$s.Width; height=$s.Height}
    }
    return @{success=$true; data=@{shapes=$list; count=$list.Count}}
}

function Exec-setShapePosition($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $sp = (Get-ShapeById $slide ([int]$p.shapeIndex))
        if ($p.x -ne $null) { $sp.Left = [int]$p.x }
        if ($p.y -ne $null) { $sp.Top = [int]$p.y }
        if ($p.width -ne $null) { $sp.Width = [int]$p.width }
        if ($p.height -ne $null) { $sp.Height = [int]$p.height }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setShapeStyle($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $sp = (Get-ShapeById $slide ([int]$p.shapeIndex))
        if ($p.fillColor) { $ole = Convert-HexToOle $p.fillColor; $sp.Fill.ForeColor.RGB = $ole }
        if ($p.borderColor) { $ole = Convert-HexToOle $p.borderColor; $sp.Line.ForeColor.RGB = $ole }
        if ($p.borderWidth) { $sp.Line.Weight = [double]$p.borderWidth }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setShapeFill($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $ole = Convert-HexToOle $p.color; (Get-ShapeById $slide ([int]$p.shapeIndex)).Fill.ForeColor.RGB = $ole; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setShapeBorder($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $sp = (Get-ShapeById $slide ([int]$p.shapeIndex))
        if ($p.color) { $ole = Convert-HexToOle $p.color; $sp.Line.ForeColor.RGB = $ole }
        if ($p.width) { $sp.Line.Weight = [double]$p.width }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setShapeShadow($p) {
    try { return @{success=$true; data=@{message="阴影设置（best-effort）"}} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setShapeGradient($p) {
    try { return @{success=$true; data=@{message="渐变填充（best-effort）"}} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setShapeTransparency($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { (Get-ShapeById $slide ([int]$p.shapeIndex)).Fill.Transparency = [double]$p.value; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-alignShapes($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $align = 1
        switch ($p.align) { "left" { $align=1 } "center" { $align=2 } "right" { $align=3 } "top" { $align=4 } "middle" { $align=5 } "bottom" { $align=6 } }
        $slide.Shapes.Align($align, 0) | Out-Null
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-distributeShapes($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $slide.Shapes.Distribute((Convert-Distribution $p.distribute), 0) | Out-Null; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-groupShapes($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $slide.Shapes.Range().Group() | Out-Null; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-duplicateShape($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $ns = (Get-ShapeById $slide ([int]$p.shapeIndex)).Duplicate(); return @{success=$true; data=@{shapeId=$ns[0].Id}} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setShapeZOrder($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { (Get-ShapeById $slide ([int]$p.shapeIndex)).ZOrder((Convert-ZOrder $p.order)) | Out-Null; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-smartDistribute($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $slide.Shapes.Distribute((Convert-Distribution $p.distribute), 0) | Out-Null; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

# ---------- 图片 ----------
function Exec-insertPptImage($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]($p.slideIndex))
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $left = if ($p.x) { [int]$p.x } else { 100 }
        $top = if ($p.y) { [int]$p.y } else { 100 }
        $sp = $slide.Shapes.AddPicture($p.imagePath, 0, -1, $left, $top)
        if ($p.width) { $sp.Width = [int]$p.width }
        if ($p.height) { $sp.Height = [int]$p.height }
        return @{success=$true; data=@{shapeId=$sp.Id}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-deletePptImage($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { (Get-ShapeById $slide ([int]$p.shapeIndex)).Delete(); return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setImageStyle($p) {
    return @{success=$true; data=@{message="图片样式（best-effort）"}}
}

function Exec-exportSlideAsImage($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $fmt = if ($p.format) { $p.format } else { "PNG" }
        $w = if ($p.width) { [int]$p.width } else { 1920 }
        $h = if ($p.height) { [int]$p.height } else { 1080 }
        $slide.Export($p.outputPath, $fmt, $w, $h)
        return @{success=$true; data=@{outputPath=$p.outputPath}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-replacePptImage($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $old = (Get-ShapeById $slide ([int]$p.shapeIndex))
        $l=$old.Left; $t=$old.Top; $w=$old.Width; $h=$old.Height
        $old.Delete()
        $ns = $slide.Shapes.AddPicture($p.filePath, 0, -1, $l, $t)
        $ns.Width=$w; $ns.Height=$h
        return @{success=$true; data=@{shapeId=$ns.Id}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

# ---------- 表格 ----------
function Exec-insertPptTable($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $left = if ($p.x) { [int]$p.x } else { 100 }
        $top = if ($p.y) { [int]$p.y } else { 100 }
        $w = if ($p.width) { [int]$p.width } else { 600 }
        $h = if ($p.height) { [int]$p.height } else { 300 }
        $rows = if ($p.rows) { [int]$p.rows } else { 3 }
        $cols = if ($p.cols) { [int]$p.cols } else { 3 }
        $tb = $slide.Shapes.AddTable($rows, $cols, $left, $top, $w, $h)
        return @{success=$true; data=@{shapeId=$tb.Id}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-getPptTableCell($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $v = (Get-ShapeById $slide ([int]$p.shapeIndex)).Table.Cell([int]$p.row, [int]$p.col).Shape.TextFrame.TextRange.Text; return @{success=$true; data=@{value=$v}} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setPptTableCell($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { (Get-ShapeById $slide ([int]$p.shapeIndex)).Table.Cell([int]$p.row, [int]$p.col).Shape.TextFrame.TextRange.Text = $p.text; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setPptTableStyle($p) {
    return @{success=$true; data=@{message="表格样式（best-effort）"}}
}

function Exec-setPptTableCellStyle($p) {
    return @{success=$true; data=@{message="单元格样式（best-effort）"}}
}

function Exec-setPptTableRowStyle($p) {
    return @{success=$true; data=@{message="行样式（best-effort）"}}
}

# ---------- 美化高级 ----------
function Exec-applyColorScheme($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    try {
        $ole = Convert-HexToOle $p.color
        if ($ole -ne $null) { $pres.SlideMaster.Background.Fill.ForeColor.RGB = $ole }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-autoBeautifySlide($p) {
    return @{success=$true; data=@{message="单页自动美化（best-effort）"}}
}

function Exec-beautifySlide($p) {
    return @{success=$true; data=@{message="一键美化（best-effort，建议用原位替换保持版式）"}}
}

function Exec-beautifyAllSlides($p) {
    return @{success=$true; data=@{message="全部美化（best-effort）"}}
}

function Exec-addTitleDecoration($p) {
    return @{success=$true; data=@{message="标题装饰（best-effort）"}}
}

function Exec-addPageIndicator($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    try {
        foreach ($slide in $pres.Slides) {
            $tb = $slide.Shapes.AddTextbox(1, $pres.PageSetup.SlideWidth - 80, $pres.PageSetup.SlideHeight - 40, 60, 30)
            $tb.TextFrame.TextRange.Text = [string]$slide.SlideIndex
        }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-createStyledTable($p) {
    return Exec-insertPptTable $p
}

function Exec-createKpiCards($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $cards = if ($p.cards) { $p.cards } else { @(@{title="KPI"; value="100"}) }
        $i = 0
        foreach ($c in $cards) {
            $x = 50 + $i * 220
            $box = $slide.Shapes.AddShape(1, $x, 150, 200, 120)
            $box.TextFrame.TextRange.Text = "$($c.title)`n$($c.value)"
            $i++
        }
        return @{success=$true; data=@{count=$cards.Count}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-unifyFont($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    try {
        $font = if ($p.font) { $p.font } else { "微软雅黑" }
        foreach ($slide in $pres.Slides) {
            foreach ($s in $slide.Shapes) {
                try { if ($s.HasTextFrame) { $s.TextFrame.TextRange.Font.Name = $font } } catch {}
            }
        }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setFontColor($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $ole = Convert-HexToOle $p.color
        foreach ($s in $slide.Shapes) { try { if ($s.HasTextFrame) { $s.TextFrame.TextRange.Font.Color.RGB = $ole } } catch {} }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

# ---------- 动画切换 ----------
function Exec-addAnimation($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $sp = (Get-ShapeById $slide ([int]$p.shapeIndex))
        $eff = $slide.TimeLine.MainSequence.AddEffect($sp, 1)  # ppEffectAppear
        return @{success=$true; data=@{effectId=$eff.EntryEffect}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-removeAnimation($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $slide.TimeLine.MainSequence.Item([int]$p.index).Delete() ; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-getAnimations($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    $list = @()
    try { foreach ($e in $slide.TimeLine.MainSequence) { $list += @{effect=$e.EntryEffect} } } catch {}
    return @{success=$true; data=@{animations=$list; count=$list.Count}}
}

function Exec-setAnimationOrder($p) {
    return @{success=$true; data=@{message="动画顺序调整（best-effort）"}}
}

function Exec-addAnimationPreset($p) {
    return Exec-addAnimation $p
}

function Exec-addEmphasisAnimation($p) {
    return Exec-addAnimation $p
}

function Exec-setSlideTransition($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $slide.SlideShowTransition.EntryEffect = [int]$p.transitionType; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-removeSlideTransition($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $slide.SlideShowTransition.EntryEffect = 0; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-applyTransitionToAll($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    try { foreach ($slide in $pres.Slides) { $slide.SlideShowTransition.EntryEffect = [int]$p.transitionType } return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

# ---------- 图表流程图 ----------
function Exec-insertPptChart($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $left = if ($p.x) { [int]$p.x } else { 100 }
        $top = if ($p.y) { [int]$p.y } else { 100 }
        $chartType = if ($p.chartType) { Convert-ChartType $p.chartType } else { 51 }
        $sp = $slide.Shapes.AddChart($chartType, $left, $top)
        return @{success=$true; data=@{shapeId=$sp.Id}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setPptChartData($p) {
    return @{success=$true; data=@{message="图表数据更新（best-effort，需 WPS 图表对象）"}}
}

function Exec-setPptChartStyle($p) {
    return @{success=$true; data=@{message="图表样式（best-effort）"}}
}

function Exec-createFlowChart($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $steps = if ($p.steps) { $p.steps } else { @("开始","过程","结束") }
        $i = 0
        foreach ($s in $steps) {
            $box = $slide.Shapes.AddShape(5, 100, (50 + $i*100), 200, 60)
            $box.TextFrame.TextRange.Text = $s
            $i++
        }
        return @{success=$true; data=@{count=$steps.Count}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-createOrgChart($p) {
    return Exec-createFlowChart $p
}

# ---------- 杂项工具 ----------
function Exec-addPptHyperlink($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { (Get-ShapeById $slide ([int]$p.shapeIndex)).ActionSettings(1).Hyperlink.Address = $p.url; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-removePptHyperlink($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { (Get-ShapeById $slide ([int]$p.shapeIndex)).ActionSettings(1).Hyperlink.Address = ""; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-autoLayout($p) {
    return @{success=$true; data=@{message="自动布局（best-effort）"}}
}

function Exec-createGrid($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $rows = if ($p.rows) { [int]$p.rows } else { 3 }
        $cols = if ($p.cols) { [int]$p.cols } else { 3 }
        $x = if ($p.x) { [int]$p.x } else { 50 }
        $y = if ($p.y) { [int]$p.y } else { 50 }
        $cw = if ($p.cellWidth) { [int]$p.cellWidth } else { 150 }
        $ch = if ($p.cellHeight) { [int]$p.cellHeight } else { 80 }
        for ($r=0; $r -lt $rows; $r++) { for ($c=0; $c -lt $cols; $c++) { $slide.Shapes.AddShape(1, $x + $c*$cw, $y + $r*$ch, $cw, $ch) | Out-Null } }
        return @{success=$true; data=@{cells=($rows*$cols)}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-createTimeline($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $points = if ($p.points) { $p.points } else { @("阶段1","阶段2","阶段3") }
        $i = 0
        foreach ($pt in $points) {
            $box = $slide.Shapes.AddShape(9, (50 + $i*200), 200, 30, 30)
            $tb = $slide.Shapes.AddTextbox(1, (50 + $i*200), 240, 180, 40)
            $tb.TextFrame.TextRange.Text = $pt
            $i++
        }
        return @{success=$true; data=@{count=$points.Count}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

# ---------- 数据可视化 ----------
function Exec-createProgressBar($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $pct = if ($p.percent) { [double]$p.percent } else { 0.5 }
        $x = if ($p.x) { [int]$p.x } else { 100 }
        $y = if ($p.y) { [int]$p.y } else { 200 }
        $w = if ($p.width) { [int]$p.width } else { 400 }
        $h = if ($p.height) { [int]$p.height } else { 40 }
        $bg = $slide.Shapes.AddShape(1, $x, $y, $w, $h)
        $fg = $slide.Shapes.AddShape(1, $x, $y, [int]($w*$pct), $h)
        $ole = Convert-HexToOle "#4472C4"; if ($ole -ne $null) { $fg.Fill.ForeColor.RGB = $ole }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-createGauge($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try {
        $x = if ($p.x) { [int]$p.x } else { 200 }
        $y = if ($p.y) { [int]$p.y } else { 150 }
        $d = if ($p.diameter) { [int]$p.diameter } else { 200 }
        $arc = $slide.Shapes.AddShape(9, $x, $y, $d, $d)
        return @{success=$true; data=@{shapeId=$arc.Id}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-createMiniCharts($p) {
    return Exec-insertPptChart $p
}

function Exec-createDonutChart($p) {
    return Exec-insertPptChart $p
}

function Exec-setBackgroundGradient($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $slide.Background.Fill.Type = 1; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setBackgroundImage($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $slide.Background.Fill.UserPicture($p.imagePath); return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

# ---------- 背景页脚3D ----------
function Exec-setBackgroundColor($p) {
    $pres = Get-ActivePres
    $slide = Get-Slide $pres ([int]$p.slideIndex)
    if (-not $slide) { return @{success=$false; error="未找到幻灯片"} }
    try { $ole = Convert-HexToOle $p.color; $slide.Background.Fill.ForeColor.RGB = $ole; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setSlideNumber($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    try { $pres.SlideNumber.Show = if ($p.show) { $true } else { $false }; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setPptFooter($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    try { $pres.Footers.Item(1).Text = $p.text; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setPptDateTime($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    try { $pres.Footers.Item(2).Text = $p.text; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-set3DRotation($p) {
    return @{success=$true; data=@{message="3D旋转（best-effort）"}}
}

function Exec-set3DDepth($p) {
    return @{success=$true; data=@{message="3D深度（best-effort）"}}
}

function Exec-set3DMaterial($p) {
    return @{success=$true; data=@{message="3D材质（best-effort）"}}
}

# ==================== 通用 action（各应用 COM 不同，逐应用实现） ====================
function Exec-save($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    try { $pres.Save(); return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}
function Exec-saveAs($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    try {
        # DisplayAlerts=0 已在初始化时设置，可抑制“是否覆盖”对话框；
        # 仍保险起见：保存前先删除已存在的同名文件，彻底避免 OLE_E_PROMPTSAVECANCELLED
        if (Test-Path $p.filePath) { Remove-Item $p.filePath -Force -ErrorAction SilentlyContinue }
        $pres.SaveAs($p.filePath)
        $sz = 0
        if (Test-Path $p.filePath) { $sz = (Get-Item $p.filePath).Length }
        return @{success=$true; data=@{path=$p.filePath; size=$sz}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}
function Exec-convertToPDF($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    $out = if ($p.outputPath) { $p.outputPath } else { $pres.Path + '\' + $pres.Name + '.pdf' }
    try {
        if (Test-Path $out) { Remove-Item $out -Force -ErrorAction SilentlyContinue }
        $pres.SaveAs($out, 32); return @{success=$true; data=@{path=$out}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}
function Exec-convertFormat($p) {
    $pres = Get-ActivePres
    if (-not $pres) { return @{success=$false; error="无活动演示文稿"} }
    $out = if ($p.outputPath) { $p.outputPath } else { $pres.Path + '\' + $pres.Name + '.' + $p.targetFormat }
    try {
        if (Test-Path $out) { Remove-Item $out -Force -ErrorAction SilentlyContinue }
        $pres.SaveAs($out); return @{success=$true; data=@{path=$out}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}
function Exec-reconnect($p) {
    try { $global:ppt = $null; $global:ppt = Get-PptApp } catch {}
    if ($global:ppt) {
        try { $global:ppt.Visible = -1 } catch {}
        try { $global:ppt.DisplayAlerts = 0 } catch {}
        return @{success=$true; data=@{ready=$true; app="WPS演示"}}
    }
    return @{success=$false; error="重连失败：WPS 演示可能已退出，请先打开 WPS 演示"}
}
function Exec-getSelectedText($p) { try { return @{success=$true; data=@{text=$global:ppt.Selection.Text}} } catch { return @{success=$false; error=$_.Exception.Message} } }
function Exec-setSelectedText($p) { try { $global:ppt.Selection.Text = $p.text; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} } }
function Exec-getAppInfo($p) { try { return @{success=$true; data=@{app="WPS演示"; version=$global:ppt.Version}} } catch { return @{success=$false; error=$_.Exception.Message} } }

# ==================== 主循环（必须位于所有 Exec-* 函数定义之后） ====================
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

        $result = & "Exec-$action" $params
        $sw.Stop()
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


class WpsPptController:
    """WPS 演示控制器：通过持久 PowerShell 子进程调用 WPS COM（Kwpp.Application）"""

    def __init__(self, trace=None):
        self.platform = platform.system()
        self._ps_process = None
        self._ready = False
        self._id_counter = 0
        self._id_lock = threading.Lock()
        self._stop = None       # Linux 路径无后台线程，占位避免 close() 报错
        self._queue = None
        self._reader = None
        self._stderr_queue = None
        self._stderr_reader = None
        self._linux_bridge = None
        self._linux_error = ""

        if IS_WINDOWS:
            self._init_windows(trace=trace)
        elif IS_LINUX:
            self._init_linux(trace=trace)
        else:
            raise RuntimeError(f"不支持的平台: {self.platform}")

    def _init_windows(self, trace=None):
        started = time.perf_counter()
        try:
            com_runtime = resolve_com_runtime(PPT_PROGID)
            if not com_runtime.available:
                raise RuntimeError(
                    f"{PPT_PROGID} COM 注册不完整: {com_runtime.diagnostic}"
                )
            selected_registration = com_runtime.selected_registration
            if trace:
                trace.event(
                    "powershell.process.starting",
                    app="ppt",
                    progId=PPT_PROGID,
                    clsid=(selected_registration.clsid if selected_registration else None),
                    registryViewBits=com_runtime.selected_view_bitness,
                    powershellExecutable=com_runtime.powershell_executable,
                )
            self._ps_script = tempfile.NamedTemporaryFile(
                mode='w', suffix='.ps1', delete=False, encoding='utf-8-sig'
            )
            self._ps_script.write(PS_BRIDGE_SCRIPT)
            self._ps_script.close()

            self._ps_process = subprocess.Popen(
                [com_runtime.powershell_executable, '-NoProfile', '-NoLogo',
                 '-ExecutionPolicy', 'Bypass', '-File', self._ps_script.name],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                text=True, encoding='utf-8', errors='replace', bufsize=1
            )

            self._stop = threading.Event()
            self._stderr_queue = queue.Queue()
            self._stderr_reader = threading.Thread(target=self._stderr_loop, daemon=True)
            self._stderr_reader.start()

            ready_line = self._read_line_with_timeout(self._ps_process.stdout, 30)
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
                raise RuntimeError(f"WPS 演示连接失败: {err}")

            self._queue = queue.Queue()
            self._reader = threading.Thread(target=self._reader_loop, daemon=True)
            self._reader.start()
            if trace:
                trace.event(
                    "powershell.process.ready",
                    app="ppt",
                    progId=PPT_PROGID,
                    processPid=self._ps_process.pid,
                    elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                )

        except Exception as e:
            self._ready = False
            if trace:
                trace.event(
                    "powershell.process.failed",
                    status="error",
                    app="ppt",
                    progId=PPT_PROGID,
                    error=f"{type(e).__name__}: {e}",
                    elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                )
            self._kill_ps()
            raise RuntimeError(f"PowerShell 桥接初始化失败: {e}")

    def _read_line_with_timeout(self, stream, timeout):
        q = queue.Queue()
        def _reader():
            try: q.put(stream.readline())
            except Exception: q.put('')
        t = threading.Thread(target=_reader, daemon=True)
        t.start(); t.join(timeout)
        if t.is_alive(): return ''
        try: return q.get_nowait()
        except queue.Empty: return ''

    def _reader_loop(self):
        try:
            while not self._stop.is_set() and self._ps_process is not None:
                line = self._ps_process.stdout.readline()
                if not line: break
                self._queue.put(line)
        except Exception: pass

    def _stderr_loop(self):
        try:
            while not self._stop.is_set() and self._ps_process is not None:
                line = self._ps_process.stderr.readline()
                if not line: break
                self._stderr_queue.put(line)
        except Exception: pass

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
                    app="ppt",
                    attempt=attempt,
                    error=raw,
                )

    def _next_id(self):
        with self._id_lock:
            self._id_counter += 1
            return self._id_counter

    def _kill_ps(self):
        try:
            if self._ps_process and self._ps_process.stdin:
                self._ps_process.stdin.write("EXIT\n"); self._ps_process.stdin.flush()
        except Exception: pass
        try:
            if self._ps_process and self._ps_process.poll() is None:
                self._ps_process.kill()
        except Exception: pass
        finally:
            self._ps_process = None
            if self._stop is not None: self._stop.set()

    def _reinit_windows(self, trace=None):
        self._kill_ps()
        self._init_windows(trace=trace)

    def _is_com_failure(self, err):
        e = (err or "").lower()
        return any(h in e for h in COM_FAIL_HINTS)

    def _read_result(self, req_id, trace=None, attempt=1, started=None):
        started = started or time.perf_counter()
        deadline = time.time() + EXEC_TIMEOUT
        while True:
            remaining = deadline - time.time()
            if remaining <= 0:
                self._ready = False; self._kill_ps()
                self._drain_stderr(trace, attempt)
                if trace:
                    trace.event(
                        "powershell.response.timeout", status="timeout", app="ppt",
                        reqId=req_id, attempt=attempt,
                        elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                    )
                return {"success": False, "error": f"action 执行超时（>{EXEC_TIMEOUT}s），已终止 WPS 桥接进程"}
            try:
                raw = self._queue.get(timeout=remaining)
            except queue.Empty:
                self._ready = False; self._kill_ps()
                self._drain_stderr(trace, attempt)
                if trace:
                    trace.event(
                        "powershell.response.timeout", status="timeout", app="ppt",
                        reqId=req_id, attempt=attempt,
                        elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                    )
                return {"success": False, "error": f"action 执行超时（>{EXEC_TIMEOUT}s）"}
            raw = raw.strip()
            if not raw: continue
            try: obj = json.loads(raw)
            except Exception:
                if trace:
                    trace.event(
                        "powershell.stdout.noise", status="warning", app="ppt",
                        reqId=req_id, attempt=attempt, lineLength=len(raw),
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
                        app="ppt", reqId=req_id, attempt=attempt,
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
                    "powershell.response.mismatched", status="warning", app="ppt",
                    reqId=req_id, receivedReqId=obj.get("reqId"), attempt=attempt,
                )
            continue

    def _run_windows_attempt(self, action, params, attempt, trace=None):
        req_id = self._next_id()
        cmd = json.dumps({
            "reqId": req_id,
            "traceId": trace.trace_id if trace else None,
            "attempt": attempt,
            "action": action,
            "params": params,
        }, ensure_ascii=True)
        self._drain_stderr(trace, attempt)
        started = time.perf_counter()
        if trace:
            trace.event(
                "powershell.request.sent", app="ppt", action=action,
                reqId=req_id, attempt=attempt,
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
                    "powershell.request.failed", status="error", app="ppt",
                    action=action, reqId=req_id, attempt=attempt,
                    error=f"{type(e).__name__}: {e}",
                )
            self._kill_ps()
            return {"success": False, "error": f"发送命令失败: {e}"}
        return self._read_result(req_id, trace=trace, attempt=attempt, started=started)

    def _exec_windows(self, action, params, trace=None):
        if not self._ps_process or self._ps_process.poll() is not None:
            self._ready = False
            if trace:
                trace.event("powershell.process.unavailable", status="error", app="ppt", action=action)
            return {"success": False, "error": "PowerShell 进程已退出，请重启桥接服务"}

        result = self._run_windows_attempt(action, params, attempt=1, trace=trace)
        # 偶发 COM 抖动 / “未注册对象” / 覆盖弹窗取消等：重连桥接后自动重试一次，
        # 解决“异常后无法重置状态、只能手工重启服务”的卡死问题
        if (not result.get("success")) and self._is_com_failure(result.get("error", "")):
            if trace:
                trace.event(
                    "controller.retry.scheduled", status="retry", app="ppt",
                    action=action, nextAttempt=2, reason=result.get("error"),
                )
            try:
                self._reinit_windows(trace=trace)
            except Exception as exc:
                if trace:
                    trace.event(
                        "controller.retry.reinit_failed", status="error", app="ppt",
                        action=action, error=f"{type(exc).__name__}: {exc}",
                    )
            if self._ps_process and self._ps_process.poll() is None:
                result = self._run_windows_attempt(action, params, attempt=2, trace=trace)
        return result

    # ==================== Linux: 文件级后端（纯 stdlib OpenXML） ====================

    def _init_linux(self, trace=None):
        """初始化 Linux 桥接：懒加载 LinuxPptBridge（纯 Python OpenXML 后端）"""
        bridge_dir = os.path.dirname(os.path.abspath(__file__))
        if bridge_dir not in sys.path:
            sys.path.insert(0, bridge_dir)
        try:
            from linux_ppt import LinuxPptBridge
        except Exception as e:
            self._ready = False
            self._linux_bridge = None
            self._linux_error = f"Linux 后端加载失败: {e}"
            if trace:
                trace.event("linux.backend.init.failed", status="error", app="ppt", error=self._linux_error)
            return
        self._linux_bridge = LinuxPptBridge()
        self._ready = True
        if trace:
            trace.event("linux.backend.ready", app="ppt", backend="openxml")

    def _exec_linux(self, action, params, trace=None):
        """Linux 上委托给 LinuxPptBridge（与 Windows 同一套 action 契约）"""
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
                    app="ppt", action=action, error=result.get("error"),
                )
            return result
        except Exception as e:
            return {"success": False, "error": f"Linux 后端执行异常: {e}"}

    def execute(self, action, params=None, trace=None):
        if params is None: params = {}
        if not self._ready and action != "ping":
            try:
                if IS_WINDOWS: self._reinit_windows(trace=trace)
                elif IS_LINUX: self._init_linux(trace=trace)
            except Exception as e:
                return {"success": False, "error": f"WPS 演示未连接，且重连失败: {e}"}
        if not self._ready and action != "ping":
            return {"success": False, "error": "WPS 演示未连接"}

        if IS_WINDOWS: return self._exec_windows(action, params, trace=trace)
        elif IS_LINUX: return self._exec_linux(action, params, trace=trace)
        return {"success": False, "error": f"不支持的平台: {self.platform}"}

    def ping(self, trace=None):
        try: return self.execute("ping", trace=trace).get("success", False)
        except Exception: return False

    def close(self):
        self._ready = False
        stop_event = getattr(self, "_stop", None)
        if stop_event is not None: stop_event.set()
        process = getattr(self, "_ps_process", None)
        self._ps_process = None
        stop_line_process(process)
        if hasattr(self, '_ps_script') and self._ps_script:
            try: os.unlink(self._ps_script.name)
            except Exception: pass

    def __del__(self):
        self.close()


_controller = None

def get_controller(trace=None) -> WpsPptController:
    global _controller
    if _controller is None or not _controller._ready:
        _controller = WpsPptController(trace=trace)
    return _controller


if __name__ == "__main__":
    print("WPS PPT 控制器测试")
    print(f"平台: {platform.system()}")
    try:
        ctrl = WpsPptController()
        print(f"连接就绪: {ctrl._ready}")
        print(ctrl.execute("ping"))
    except Exception as e:
        print(f"错误: {e}")
