#!/usr/bin/env python3
"""
WPS Word 控制器 - 纯 Python 实现
通过 COM 自动化（Windows）控制 WPS 文字（Kwps.Application）。

与 wps_excel.py / wps_ppt.py 完全相同的可靠架构（reqId / 超时强杀 / 重连 / BOM+ensure_ascii）。
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

IS_WINDOWS = platform.system() == "Windows"
IS_LINUX = platform.system() == "Linux"
IS_MACOS = platform.system() == "Darwin"

WORD_PROGID = "Kwps.Application"
EXEC_TIMEOUT = 60

# 偶发 COM 故障特征串（小写匹配，用于自动重连重试，避免“未注册对象”后无法重置状态）
COM_FAIL_HINTS = (
    "未注册对象", "未注册", "80040154", "rpc 服务器不可用", "800706ba",
    "rpc_e_disconnected", "调用的对象已与其客户端断开连接",
    "ole_e_promptsavecancelled", "8004000c", "catastrophic", "灾难性",
)

PS_BRIDGE_SCRIPT = r'''
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$error.Clear()

function Get-WordApp {
    try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application') }
    catch {
        try { return New-Object -ComObject 'Kwps.Application' }
        catch {
            try {
                $clsid = (Get-ItemProperty -Path 'HKCU:\Software\Classes\Kwps.Application\CLSID' -ErrorAction Stop).'(default)'
                if ($clsid) {
                    $type = [Type]::GetTypeFromCLSID([Guid]$clsid)
                    if ($type) { return [Activator]::CreateInstance($type) }
                }
            } catch {}
        }
    }
    return $null
}

$global:word = Get-WordApp
if ($global:word) {
    try { $global:word.Visible = $true } catch {}
    # 抑制“是否覆盖”等模态对话框，避免 COM 无法应答导致保存卡死
    try { $global:word.DisplayAlerts = $false } catch {}
    Write-Host '{"ready":true}'
} else {
    Write-Host '{"ready":false,"error":"无法连接WPS文字，请确认WPS已安装"}'
    exit 1
}

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

function Get-ActiveDoc { if ($global:word.ActiveDocument) { return $global:word.ActiveDocument }; return $null }
function Get-DocByName($name) {
    foreach ($d in $global:word.Documents) { if ($d.Name -eq $name) { return $d } }
    return $null
}
# 解析 range 参数 -> 返回 Range 对象
function Resolve-Range($doc, $range) {
    if ($null -eq $range -or $range -eq '' -or $range -eq 'all') { return $doc.Content }
    if ($range -is [int] -or ($range -match '^\d+$')) { return $doc.Paragraphs.Item([int]$range).Range }
    # 字符串关键词：查找并返回命中范围
    $r = $doc.Content
    [void]$r.Find.Execute($range, $false, $false, $false, $false, $false, $true, 1, $false, "", 0)
    if ($r.Find.Found) { return $r }
    return $doc.Content
}

# ==================== Action 实现函数 ====================

function Exec-ping($p) {
    return @{success=$true; data=@{message="pong"; timestamp=(Get-Date -Format "o")}}
}

# ---------- 格式化 ----------
function Exec-setFont($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $rng = Resolve-Range $doc $p.range
        if ($p.font_name) { $rng.Font.Name = $p.font_name }
        if ($p.font_size) { $rng.Font.Size = [double]$p.font_size }
        if ($p.bold -eq $true) { $rng.Font.Bold = $true }
        if ($p.bold -eq $false) { $rng.Font.Bold = $false }
        if ($p.italic -eq $true) { $rng.Font.Italic = $true }
        if ($p.underline -eq $true) { $rng.Font.Underline = 1 }
        if ($p.color) { $ole = Convert-HexToOle $p.color; if ($ole -ne $null) { $rng.Font.Color = $ole } }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-applyStyle($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $rng = Resolve-Range $doc $p.range
        $rng.Style = $p.style_name
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setTextColor($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $rng = Resolve-Range $doc $p.range
        $ole = Convert-HexToOle $p.color
        if ($ole -ne $null) { $rng.Font.Color = $ole; return @{success=$true} }
        return @{success=$false; error="颜色格式错误"}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setLineSpacing($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $para = if ($p.paragraphIndex) { $doc.Paragraphs.Item([int]$p.paragraphIndex) } else { $doc.Paragraphs.Item(1) }
        $para.Range.ParagraphFormat.LineSpacing = [double]$p.lineSpacing
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

# ---------- 内容 ----------
function Exec-insertText($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $rng = $doc.Content
        if ($p.position -eq 'start') { $rng.Collapse(1) }
        else { $rng.Collapse(0) }  # 末尾
        if ($p.new_paragraph -eq $true) { $rng.InsertParagraphAfter() }
        $rng.InsertAfter($p.text)
        if ($p.style) { $rng.Style = $p.style }
        return @{success=$true; data=@{inserted=$true}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-findReplace($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $rng = $doc.Content
        $replaceMode = if ($p.replace_all -eq $true) { 2 } else { 1 }
        [void]$rng.Find.Execute($p.find_text, [bool]$p.match_case, [bool]$p.match_whole_word, $false, $false, $false, $true, 1, $false, $p.replace_text, $replaceMode)
        return @{success=$true; data=@{done=$true}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-insertTable($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $rows = if ($p.rows) { [int]$p.rows } else { 3 }
        $cols = if ($p.cols) { [int]$p.cols } else { 3 }
        $rng = $doc.Content; $rng.Collapse(0)
        $tbl = $doc.Tables.Add($rng, $rows, $cols)
        return @{success=$true; data=@{rows=$rows; cols=$cols}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-insertImage($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $rng = $doc.Content; $rng.Collapse(0)
        $img = $doc.InlineShapes.AddPicture($p.imagePath, $false, $true, $rng)
        if ($p.width) { $img.Width = [int]$p.width }
        if ($p.height) { $img.Height = [int]$p.height }
        return @{success=$true; data=@{width=$img.Width; height=$img.Height}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-addComment($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $rng = $doc.Content; $rng.Collapse(0)
        $doc.Comments.Add($rng, $p.text) | Out-Null
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-insertPageBreak($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try { $doc.Content.InsertBreak(7); return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-insertBookmark($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $rng = $doc.Content; $rng.Collapse(0)
        $doc.Bookmarks.Add($p.name, $rng) | Out-Null
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-insertSectionBreak($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try { $doc.Content.InsertBreak(2); return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setParagraph($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $para = if ($p.paragraphIndex) { $doc.Paragraphs.Item([int]$p.paragraphIndex) } else { $doc.Paragraphs.Item(1) }
        if ($p.alignment) {
            $a = 0
            switch ($p.alignment) { "left" { $a=0 } "center" { $a=1 } "right" { $a=2 } "justify" { $a=3 } }
            $para.Range.ParagraphFormat.Alignment = $a
        }
        if ($p.lineSpacing) { $para.Range.ParagraphFormat.LineSpacing = [double]$p.lineSpacing }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-setPageSetup($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $ps = $doc.PageSetup
        if ($p.orientation) { $ps.Orientation = if ($p.orientation -eq 'landscape') { 1 } else { 0 } }
        if ($p.marginTop) { $ps.TopMargin = [double]$p.marginTop }
        if ($p.marginBottom) { $ps.BottomMargin = [double]$p.marginBottom }
        if ($p.marginLeft) { $ps.LeftMargin = [double]$p.marginLeft }
        if ($p.marginRight) { $ps.RightMargin = [double]$p.marginRight }
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

# ---------- 文档管理 ----------
function Exec-getActiveDocument($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    return @{success=$true; data=@{name=$doc.Name; path=$doc.Path; paragraphs=$doc.Paragraphs.Count}}
}

function Exec-getOpenDocuments($p) {
    $list = @()
    foreach ($d in $global:word.Documents) { $list += @{name=$d.Name; path=$d.Path} }
    return @{success=$true; data=@{documents=$list; count=$list.Count}}
}

function Exec-switchDocument($p) {
    $doc = Get-DocByName $p.name
    if (-not $doc) { return @{success=$false; error="未找到文档: $($p.name)"} }
    $doc.Activate() | Out-Null
    return @{success=$true; data=@{name=$doc.Name}}
}

function Exec-createDocument($p) {
    if ($p.template) {
        $doc = $global:word.Documents.Add($p.template)
    } else {
        $doc = $global:word.Documents.Add()
    }
    return @{success=$true; data=@{name=$doc.Name; paragraphs=$doc.Paragraphs.Count}}
}

function Exec-openDocument($p) {
    if (-not $p.filePath) { return @{success=$false; error="缺少 filePath"} }
    $doc = $global:word.Documents.Open($p.filePath)
    return @{success=$true; data=@{name=$doc.Name; paragraphs=$doc.Paragraphs.Count}}
}

function Exec-getDocumentText($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    $txt = $doc.Content.Text
    if ($p.start -ne $null -or $p.end -ne $null) {
        $s = if ($p.start) { [int]$p.start } else { 0 }
        $e = if ($p.end) { [int]$p.end } else { $txt.Length }
        if ($e -gt $txt.Length) { $e = $txt.Length }
        $txt = $txt.Substring($s, [Math]::Max(0, $e-$s))
    }
    return @{success=$true; data=@{text=$txt; length=$txt.Length}}
}

function Exec-insertHeader($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try { $doc.Sections.Item(1).Headers.Item(1).Range.Text = $p.text; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-insertFooter($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try { $doc.Sections.Item(1).Footers.Item(1).Range.Text = $p.text; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-generateTOC($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $rng = $doc.Range(0,0)
        $levels = if ($p.levels) { [int]$p.levels } else { 3 }
        $doc.TablesOfContents.Add($rng, $true, $levels) | Out-Null
        return @{success=$true; data=@{levels=$levels}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

# ---------- 模板填写 ----------
function Exec-getDocumentParagraphs($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    $list = @()
    $s = if ($p.start_paragraph) { [int]$p.start_paragraph } else { 1 }
    $e = if ($p.end_paragraph) { [int]$p.end_paragraph } else { $doc.Paragraphs.Count }
    for ($i=$s; $i -le $e; $i++) {
        $para = $doc.Paragraphs.Item($i)
        $list += @{index=$i; text=$para.Range.Text; style=$para.Style.NameLocal}
    }
    return @{success=$true; data=@{paragraphs=$list; count=$list.Count}}
}

function Exec-findInDocument($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    $results = @()
    $rng = $doc.Content
    # 顺序查找所有命中
    do {
        [void]$rng.Find.Execute($p.find_text, [bool]$p.match_case, [bool]$p.match_whole_word, $false, $false, $false, $true, 1, $false, "", 0)
        if ($rng.Find.Found) {
            $results += @{start=$rng.Start; end=$rng.End; text=$rng.Text}
            $rng.Collapse(0)
        } else { break }
        if ($p.max_results -and $results.Count -ge [int]$p.max_results) { break }
    } while ($true)
    return @{success=$true; data=@{matches=$results; count=$results.Count}}
}

function Exec-smartFillField($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        $rng = $doc.Content
        [void]$rng.Find.Execute($p.keyword, $false, $false, $false, $false, $false, $true, 1, $false, "", 0)
        if (-not $rng.Find.Found) { return @{success=$false; error="未找到关键字: $($p.keyword)"} }
        # 移动到关键字之后，替换同段落剩余占位内容
        $rng.Collapse(0)
        $paraEnd = $rng.Paragraphs.Item(1).Range.End
        $fillRng = $doc.Range($rng.End, $paraEnd)
        $fillRng.Text = $p.value
        return @{success=$true; data=@{filled=$p.keyword}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

function Exec-replaceBookmarkContent($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        if (-not $doc.Bookmarks.Exists($p.name)) { return @{success=$false; error="书签不存在: $($p.name)"} }
        $doc.Bookmarks.Item($p.name).Range.Text = $p.text
        return @{success=$true}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}

# ==================== 通用 action（各应用 COM 不同，逐应用实现） ====================
function Exec-save($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try { $doc.Save(); return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} }
}
function Exec-saveAs($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    try {
        if (Test-Path $p.filePath) { Remove-Item $p.filePath -Force -ErrorAction SilentlyContinue }
        $doc.SaveAs($p.filePath)
        $sz = 0
        if (Test-Path $p.filePath) { $sz = (Get-Item $p.filePath).Length }
        return @{success=$true; data=@{path=$p.filePath; size=$sz}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}
function Exec-convertToPDF($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    $out = if ($p.outputPath) { $p.outputPath } else { $doc.Path + '\' + $doc.Name + '.pdf' }
    try {
        if (Test-Path $out) { Remove-Item $out -Force -ErrorAction SilentlyContinue }
        $doc.ExportAsFixedFormat($out, 17); return @{success=$true; data=@{path=$out}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}
function Exec-convertFormat($p) {
    $doc = Get-ActiveDoc
    if (-not $doc) { return @{success=$false; error="无活动文档"} }
    $out = if ($p.outputPath) { $p.outputPath } else { $doc.Path + '\' + $doc.Name + '.' + $p.targetFormat }
    try {
        if (Test-Path $out) { Remove-Item $out -Force -ErrorAction SilentlyContinue }
        $doc.SaveAs($out); return @{success=$true; data=@{path=$out}}
    } catch { return @{success=$false; error=$_.Exception.Message} }
}
function Exec-reconnect($p) {
    try { $global:word = $null; $global:word = Get-WordApp } catch {}
    if ($global:word) {
        try { $global:word.Visible = $true } catch {}
        try { $global:word.DisplayAlerts = $false } catch {}
        return @{success=$true; data=@{ready=$true; app="WPS文字"}}
    }
    return @{success=$false; error="重连失败：WPS 文字可能已退出，请先打开 WPS 文字"}
}
function Exec-getSelectedText($p) { try { return @{success=$true; data=@{text=$global:word.Selection.Text}} } catch { return @{success=$false; error=$_.Exception.Message} } }
function Exec-setSelectedText($p) { try { $global:word.Selection.Text = $p.text; return @{success=$true} } catch { return @{success=$false; error=$_.Exception.Message} } }
function Exec-getAppInfo($p) { try { return @{success=$true; data=@{app="WPS文字"; version=$global:word.Version}} } catch { return @{success=$false; error=$_.Exception.Message} } }

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


class WpsWordController:
    """WPS 文字控制器：通过持久 PowerShell 子进程调用 WPS COM（Kwps.Application）"""

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
        if trace:
            trace.event("powershell.process.starting", app="word", progId=WORD_PROGID)
        try:
            self._ps_script = tempfile.NamedTemporaryFile(
                mode='w', suffix='.ps1', delete=False, encoding='utf-8-sig'
            )
            self._ps_script.write(PS_BRIDGE_SCRIPT)
            self._ps_script.close()

            self._ps_process = subprocess.Popen(
                ['powershell.exe', '-NoProfile', '-NoLogo',
                 '-ExecutionPolicy', 'Bypass', '-File', self._ps_script.name],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                text=True, encoding='utf-8', bufsize=1
            )

            self._stop = threading.Event()
            self._stderr_queue = queue.Queue()
            self._stderr_reader = threading.Thread(target=self._stderr_loop, daemon=True)
            self._stderr_reader.start()

            ready_line = self._read_line_with_timeout(self._ps_process.stdout, 30)
            ready = json.loads(ready_line) if ready_line else {}
            self._ready = ready.get('ready', False)
            if not self._ready:
                err = ready.get('error', '未知错误')
                self._kill_ps()
                raise RuntimeError(f"WPS 文字连接失败: {err}")

            self._queue = queue.Queue()
            self._reader = threading.Thread(target=self._reader_loop, daemon=True)
            self._reader.start()
            if trace:
                trace.event(
                    "powershell.process.ready",
                    app="word",
                    progId=WORD_PROGID,
                    processPid=self._ps_process.pid,
                    elapsedMs=round((time.perf_counter() - started) * 1000, 2),
                )

        except Exception as e:
            self._ready = False
            if trace:
                trace.event(
                    "powershell.process.failed",
                    status="error",
                    app="word",
                    progId=WORD_PROGID,
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
                    app="word",
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
                        "powershell.response.timeout", status="timeout", app="word",
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
                        "powershell.response.timeout", status="timeout", app="word",
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
                        "powershell.stdout.noise", status="warning", app="word",
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
                        app="word", reqId=req_id, attempt=attempt,
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
                    "powershell.response.mismatched", status="warning", app="word",
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
                "powershell.request.sent", app="word", action=action,
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
                    "powershell.request.failed", status="error", app="word",
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
                trace.event("powershell.process.unavailable", status="error", app="word", action=action)
            return {"success": False, "error": "PowerShell 进程已退出，请重启桥接服务"}

        result = self._run_windows_attempt(action, params, attempt=1, trace=trace)
        # 偶发 COM 抖动 / “未注册对象” / 覆盖弹窗取消等：重连桥接后自动重试一次
        if (not result.get("success")) and self._is_com_failure(result.get("error", "")):
            if trace:
                trace.event(
                    "controller.retry.scheduled", status="retry", app="word",
                    action=action, nextAttempt=2, reason=result.get("error"),
                )
            try:
                self._reinit_windows(trace=trace)
            except Exception as exc:
                if trace:
                    trace.event(
                        "controller.retry.reinit_failed", status="error", app="word",
                        action=action, error=f"{type(exc).__name__}: {exc}",
                    )
            if self._ps_process and self._ps_process.poll() is None:
                result = self._run_windows_attempt(action, params, attempt=2, trace=trace)
        return result

    # ==================== Linux: 文件级后端（纯 stdlib OpenXML） ====================

    def _init_linux(self, trace=None):
        """初始化 Linux 桥接：懒加载 LinuxWordBridge（纯 Python OpenXML 后端）"""
        bridge_dir = os.path.dirname(os.path.abspath(__file__))
        if bridge_dir not in sys.path:
            sys.path.insert(0, bridge_dir)
        try:
            from linux_word import LinuxWordBridge
        except Exception as e:
            self._ready = False
            self._linux_bridge = None
            self._linux_error = f"Linux 后端加载失败: {e}"
            if trace:
                trace.event("linux.backend.init.failed", status="error", app="word", error=self._linux_error)
            return
        self._linux_bridge = LinuxWordBridge()
        self._ready = True
        if trace:
            trace.event("linux.backend.ready", app="word", backend="openxml")

    def _exec_linux(self, action, params, trace=None):
        """Linux 上委托给 LinuxWordBridge（与 Windows 同一套 action 契约）"""
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
                    app="word", action=action, error=result.get("error"),
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
                return {"success": False, "error": f"WPS 文字未连接，且重连失败: {e}"}
        if not self._ready and action != "ping":
            return {"success": False, "error": "WPS 文字未连接"}

        if IS_WINDOWS: return self._exec_windows(action, params, trace=trace)
        elif IS_LINUX: return self._exec_linux(action, params, trace=trace)
        return {"success": False, "error": f"不支持的平台: {self.platform}"}

    def ping(self, trace=None):
        try: return self.execute("ping", trace=trace).get("success", False)
        except Exception: return False

    def close(self):
        if self._stop is not None: self._stop.set()
        try:
            if self._ps_process and self._ps_process.stdin:
                try: self._ps_process.stdin.write("EXIT\n"); self._ps_process.stdin.flush()
                except Exception: pass
                if self._ps_process.poll() is None: self._ps_process.kill()
        except Exception: pass
        finally:
            self._ps_process = None
        if hasattr(self, '_ps_script') and self._ps_script:
            try: os.unlink(self._ps_script.name)
            except Exception: pass

    def __del__(self):
        self.close()


_controller = None

def get_controller(trace=None) -> WpsWordController:
    global _controller
    if _controller is None or not _controller._ready:
        _controller = WpsWordController(trace=trace)
    return _controller


if __name__ == "__main__":
    print("WPS Word 控制器测试")
    print(f"平台: {platform.system()}")
    try:
        ctrl = WpsWordController()
        print(f"连接就绪: {ctrl._ready}")
        print(ctrl.execute("ping"))
    except Exception as e:
        print(f"错误: {e}")
