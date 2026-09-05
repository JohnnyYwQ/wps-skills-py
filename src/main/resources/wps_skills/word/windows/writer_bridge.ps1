[CmdletBinding()]
param(
    [switch]$DebugCloseCreatedDocument
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$script:Application = $null
$script:Documents = $null
$script:Document = $null
$script:DocumentId = $null
$script:DocumentCreatedBySession = $false
$script:ApplicationCreatedBySession = $false
$script:DebugCloseCreatedDocument = [bool]$DebugCloseCreatedDocument
$script:AuthorizedPath = $null
$script:Revision = $null
$script:Fingerprint = $null
$script:PreparationId = $null
$script:PreparedIdentity = $null
$script:PreparedPath = $null
$script:PreparedCanonicalPath = $null
$script:CoordinationMutex = $null
$script:CoordinationMutexName = $null
$script:CoordinationStatePath = $null
$script:CoordinationGuardId = $null
$script:CoordinationLeaseId = $null

class StaleDocumentRevisionException : System.Exception {
    StaleDocumentRevisionException([string]$message) : base($message) {}
}

class PersistenceLocatorRequiredException : System.Exception {
    PersistenceLocatorRequiredException([string]$message) : base($message) {}
}

class ContentAnchorBoundaryException : System.Exception {
    ContentAnchorBoundaryException([string]$message) : base($message) {}
}

class ContentVerificationException : System.Exception {
    ContentVerificationException([string]$message) : base($message) {}
}

class DocumentLeaseConflictException : System.Exception {
    DocumentLeaseConflictException([string]$message) : base($message) {}
}

class DocumentQuarantinedException : System.Exception {
    DocumentQuarantinedException([string]$message) : base($message) {}
}

if ($null -eq ('WpsSkills.NativeFileIdentity' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace WpsSkills {
    [StructLayout(LayoutKind.Sequential)]
    public struct ByHandleFileInformation {
        public uint FileAttributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME CreationTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastAccessTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastWriteTime;
        public uint VolumeSerialNumber;
        public uint FileSizeHigh;
        public uint FileSizeLow;
        public uint NumberOfLinks;
        public uint FileIndexHigh;
        public uint FileIndexLow;
    }

    public static class NativeFileIdentity {
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool GetFileInformationByHandle(
            SafeFileHandle handle,
            out ByHandleFileInformation information
        );
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WindowRect {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MonitorInfo {
        public int Size;
        public WindowRect Monitor;
        public WindowRect Work;
        public uint Flags;
    }

    public static class NativeWindowPlacement {
        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, out WindowRect rect);

        [DllImport("user32.dll")]
        public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);

        [DllImport("user32.dll")]
        public static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo information);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetWindowPos(
            IntPtr hwnd,
            IntPtr insertAfter,
            int x,
            int y,
            int width,
            int height,
            uint flags
        );
    }
}
'@
}

function Write-BridgeRecord {
    param([Parameter(Mandatory = $true)]$Record)

    $json = $Record | ConvertTo-Json -Compress -Depth 24
    $ascii = New-Object Text.StringBuilder
    foreach ($character in $json.ToCharArray()) {
        $codePoint = [int][char]$character
        if ($codePoint -le 0x7f) {
            [void]$ascii.Append($character)
        }
        else {
            [void]$ascii.AppendFormat('\u{0:x4}', $codePoint)
        }
    }
    [Console]::Out.WriteLine($ascii.ToString())
    [Console]::Out.Flush()
}

function New-SuccessRecord {
    param(
        [Parameter(Mandatory = $true)][string]$RequestId,
        [Parameter(Mandatory = $true)]$Data
    )

    return [ordered]@{
        requestId = $RequestId
        outcome = 'succeeded'
        data = $Data
    }
}

function New-FailureRecord {
    param(
        [Parameter(Mandatory = $true)][string]$RequestId,
        [Parameter(Mandatory = $true)][ValidateSet('failed', 'unknown')]
        [string]$Outcome,
        [Parameter(Mandatory = $true)][string]$Code,
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $true)]
        [ValidateSet('unchanged', 'lost', 'unprovable')]
        [string]$BindingDisposition
    )

    return [ordered]@{
        requestId = $RequestId
        outcome = $Outcome
        error = [ordered]@{
            code = $Code
            message = $Message
        }
        bindingDisposition = $BindingDisposition
    }
}

function Get-ObjectPropertyNames {
    param([Parameter(Mandatory = $true)]$Value)

    return @($Value.PSObject.Properties | ForEach-Object { $_.Name })
}

function Test-ExactFields {
    param(
        [Parameter(Mandatory = $true)]$Value,
        [Parameter(Mandatory = $true)][string[]]$Expected
    )

    if ($null -eq $Value -or $Value -isnot [pscustomobject]) {
        return $false
    }
    $actual = @(Get-ObjectPropertyNames -Value $Value | Sort-Object)
    $wanted = @($Expected | Sort-Object)
    return $null -eq (Compare-Object -ReferenceObject $wanted -DifferenceObject $actual)
}

function Get-StableFileIdentity {
    param([Parameter(Mandatory = $true)][string]$Path)

    $stream = $null
    try {
        $stream = [IO.File]::Open(
            $Path,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
        )
        $information = New-Object WpsSkills.ByHandleFileInformation
        if (-not [WpsSkills.NativeFileIdentity]::GetFileInformationByHandle(
            $stream.SafeFileHandle,
            [ref]$information
        )) {
            throw [ComponentModel.Win32Exception]::new(
                [Runtime.InteropServices.Marshal]::GetLastWin32Error()
            )
        }
        $fileIndex = ([uint64]$information.FileIndexHigh -shl 32) -bor [uint64]$information.FileIndexLow
        return 'file-{0:x8}-{1:x16}' -f [uint32]$information.VolumeSerialNumber, $fileIndex
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-CoordinationHash {
    param([Parameter(Mandatory = $true)][string]$Identity)

    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Identity)
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Write-CoordinationState {
    param(
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][bool]$InFlight
    )

    if ($null -eq $script:CoordinationMutex) { return }
    $directory = [IO.Path]::GetDirectoryName($script:CoordinationStatePath)
    [IO.Directory]::CreateDirectory($directory) | Out-Null
    $record = [ordered]@{
        identity = $script:PreparedIdentity
        ownerPid = $PID
        mode = $Mode
        inFlight = $InFlight
        updatedUtc = [DateTime]::UtcNow.ToString('o')
    }
    $temporary = $script:CoordinationStatePath + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    try {
        [IO.File]::WriteAllText(
            $temporary,
            ($record | ConvertTo-Json -Compress),
            (New-Object Text.UTF8Encoding($false))
        )
        Move-Item -LiteralPath $temporary -Destination $script:CoordinationStatePath -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporary) {
            Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        }
    }
}

function Set-CoordinationInFlight {
    param([Parameter(Mandatory = $true)][bool]$InFlight)

    if ($null -eq $script:CoordinationMutex) { return }
    $mode = if ([string]::IsNullOrEmpty($script:CoordinationLeaseId)) { 'guard' } else { 'lease' }
    Write-CoordinationState -Mode $mode -InFlight $InFlight
}

function Invoke-CoordinatedWpsCall {
    param([Parameter(Mandatory = $true)][scriptblock]$Action)

    if ($null -eq $script:CoordinationMutex) {
        throw 'A WPS call requires an acquired document guard.'
    }
    Set-CoordinationInFlight -InFlight $true
    try {
        return & $Action
    }
    finally {
        Set-CoordinationInFlight -InFlight $false
    }
}

function Release-CoordinationResources {
    param([Parameter(Mandatory = $true)][bool]$Clean)

    if ($null -eq $script:CoordinationMutex) { return }
    try {
        if ($Clean -and -not [string]::IsNullOrEmpty($script:CoordinationStatePath)) {
            Remove-Item -LiteralPath $script:CoordinationStatePath -Force -ErrorAction SilentlyContinue
        }
        $script:CoordinationMutex.ReleaseMutex()
    }
    catch {
        if ($Clean) { throw }
    }
    finally {
        $script:CoordinationMutex.Dispose()
        $script:CoordinationMutex = $null
        $script:CoordinationMutexName = $null
        $script:CoordinationStatePath = $null
        $script:CoordinationGuardId = $null
        $script:CoordinationLeaseId = $null
    }
}

function Release-ComReference {
    param($Value)

    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
        }
        catch {
        }
    }
}

function Connect-WpsApplication {
    $created = $false
    try {
        $script:Application = [Runtime.InteropServices.Marshal]::GetActiveObject('KWPS.Application')
    }
    catch {
        $script:Application = New-Object -ComObject 'KWPS.Application'
        $created = $true
    }
    $script:ApplicationCreatedBySession = $created
    $script:Application.Visible = $true
    $script:Documents = $script:Application.Documents
    return $created
}

function Repair-BoundDocumentWindowBounds {
    param([Parameter(Mandatory = $true)]$Window)

    # WPS's COM Width/Height can diverge from the visible top-level frame.
    # Resolve the bound document's HWND to its root and repair only an
    # implausibly small restored rectangle. 2 is GA_ROOT and
    # MONITOR_DEFAULTTONEAREST; 0x14 is SWP_NOZORDER | SWP_NOACTIVATE.
    $childHwnd = [IntPtr]([int64]$Window.Hwnd)
    if ($childHwnd -eq [IntPtr]::Zero) {
        return
    }
    $rootHwnd = [WpsSkills.NativeWindowPlacement]::GetAncestor($childHwnd, 2)
    if ($rootHwnd -eq [IntPtr]::Zero) {
        return
    }

    $rootRect = New-Object WpsSkills.WindowRect
    $monitor = [WpsSkills.NativeWindowPlacement]::MonitorFromWindow($rootHwnd, 2)
    $monitorInfo = New-Object WpsSkills.MonitorInfo
    $monitorInfo.Size = [Runtime.InteropServices.Marshal]::SizeOf($monitorInfo)
    if (
        $monitor -eq [IntPtr]::Zero -or
        -not [WpsSkills.NativeWindowPlacement]::GetWindowRect(
            $rootHwnd,
            [ref]$rootRect
        ) -or
        -not [WpsSkills.NativeWindowPlacement]::GetMonitorInfo(
            $monitor,
            [ref]$monitorInfo
        )
    ) {
        return
    }

    $workWidth = $monitorInfo.Work.Right - $monitorInfo.Work.Left
    $workHeight = $monitorInfo.Work.Bottom - $monitorInfo.Work.Top
    $currentWidth = $rootRect.Right - $rootRect.Left
    $currentHeight = $rootRect.Bottom - $rootRect.Top
    if (
        $workWidth -le 0 -or
        $workHeight -le 0 -or
        (
            $currentWidth -ge ($workWidth * 0.68) -and
            $currentHeight -ge ($workHeight * 0.68)
        )
    ) {
        return
    }

    $targetWidth = [int][math]::Round($workWidth * 0.8)
    $targetHeight = [int][math]::Round($workHeight * 0.8)
    $targetLeft = $monitorInfo.Work.Left + [int][math]::Round(
        ($workWidth - $targetWidth) / 2
    )
    $targetTop = $monitorInfo.Work.Top + [int][math]::Round(
        ($workHeight - $targetHeight) / 2
    )
    [void][WpsSkills.NativeWindowPlacement]::SetWindowPos(
        $rootHwnd,
        [IntPtr]::Zero,
        $targetLeft,
        $targetTop,
        $targetWidth,
        $targetHeight,
        0x14
    )
}

function Show-BoundDocument {
    param([Parameter(Mandatory = $true)][bool]$UseNormalWindowState)

    $script:Document.Activate()
    $window = $null
    try {
        $window = $script:Document.ActiveWindow
        if ($UseNormalWindowState) {
            # 0 is wdWindowStateNormal on the outer WPS application. The
            # document child remains free to fill that normal application.
            $script:Application.WindowState = 0
        }
        Repair-BoundDocumentWindowBounds -Window $window
    }
    catch {
        # Window presentation must not invalidate an exact document binding.
    }
    finally {
        Release-ComReference -Value $window
    }
}

function Invoke-DebugCreatedDocumentCleanup {
    if (
        -not $script:DebugCloseCreatedDocument -or
        -not $script:DocumentCreatedBySession -or
        $null -eq $script:Document
    ) {
        return
    }

    Set-CoordinationInFlight -InFlight $true
    try {
        if (Test-BoundDocumentLive) {
            # 0 is wdDoNotSaveChanges. This switch exists only for an explicit
            # debug Session and can target only that Session's created document.
            $script:Document.Close(0)
        }
        $script:DocumentCreatedBySession = $false
        if ($script:ApplicationCreatedBySession) {
            try {
                if ([int]$script:Documents.Count -eq 0) {
                    $script:Application.Quit(0)
                }
            }
            catch {
                # Closing the owned test document is the required cleanup;
                # quitting an otherwise empty application is best effort.
            }
        }
    }
    finally {
        Set-CoordinationInFlight -InFlight $false
    }
}

function New-ContentRevision {
    return 'word-' + [guid]::NewGuid().ToString('N')
}

function Normalize-WordText {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) {
        return ''
    }

    # WPS exposes table structure through private control markers in Range.Text:
    # every cell ends with CR+BEL and every row adds another CR+BEL.  Translate
    # the doubled row marker first so the portable representation is TSV-like.
    $normalized = $Text.Replace("`r`a`r`a", "`n")
    $normalized = $normalized.Replace("`r`a", "`t")
    $normalized = $normalized.Replace("`r", "`n")

    # Keep the caller-facing text independent of host-specific Word markers.
    $normalized = $normalized.Replace(
        ([string][char]11),
        ([string][char]0x2028)
    )
    $normalized = $normalized.Replace(
        ([string][char]1),
        ([string][char]0xFFFC)
    )
    $normalized = $normalized.Replace(
        ([string][char]7),
        "`t"
    )
    $normalized = $normalized.Replace(
        ([string][char]0x2029),
        "`n"
    )

    # Field delimiters and any other remaining C0/C1 controls are structural
    # implementation details, not document text.  Preserve only the controls
    # admitted by wordNormalizedText; surrogate pairs remain untouched.
    foreach ($code in 0..31) {
        if ($code -notin @(9, 10, 12)) {
            $normalized = $normalized.Replace(
                ([string][char]$code),
                ''
            )
        }
    }
    foreach ($code in 127..159) {
        $normalized = $normalized.Replace(
            ([string][char]$code),
            ''
        )
    }
    return $normalized
}

function Normalize-StoryText {
    param([AllowNull()][string]$Text)

    $normalized = Normalize-WordText -Text $Text
    $normalized = $normalized.Replace(([string][char]12), "`n")
    $normalized = $normalized.Replace(([string][char]0x2028), "`n")
    $normalized = $normalized.Replace(([string][char]0xFFFC), '')
    while ($normalized.EndsWith("`n")) {
        $normalized = $normalized.Substring(0, $normalized.Length - 1)
    }
    return $normalized
}

function Get-DocumentPersistenceState {
    if ([string]::IsNullOrEmpty([string]$script:Document.Path)) {
        return 'unsaved'
    }
    if ([bool]$script:Document.Saved) {
        return 'saved'
    }
    return 'modified'
}

function Get-DocumentEnd {
    $range = $null
    try {
        $range = $script:Document.Content
        return [Math]::Max(0, ([int]$range.End - 1))
    }
    finally {
        Release-ComReference -Value $range
    }
}

function Get-DocumentFingerprint {
    $content = $null
    try {
        $content = $script:Document.Content
        $sectionFacts = @(
            Get-SectionSnapshots
        ) | ConvertTo-Json -Compress -Depth 12
        $facts = @(
            [string]$content.Text,
            [string]$script:Document.Paragraphs.Count,
            [string]$script:Document.Sections.Count,
            [string]$script:Document.Tables.Count,
            [string]$script:Document.InlineShapes.Count,
            [string]$script:Document.Shapes.Count,
            $sectionFacts
        ) -join ([char]0x1f)
        $bytes = [Text.Encoding]::UTF8.GetBytes($facts)
        $sha = [Security.Cryptography.SHA256]::Create()
        try {
            return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
        }
        finally {
            $sha.Dispose()
        }
    }
    finally {
        Release-ComReference -Value $content
    }
}

function Sync-ContentRevision {
    $observed = Get-DocumentFingerprint
    if ($null -eq $script:Fingerprint) {
        $script:Fingerprint = $observed
    }
    elseif ($observed -ne $script:Fingerprint) {
        $script:Fingerprint = $observed
        $script:Revision = New-ContentRevision
    }
    return $observed
}

function Assert-BoundDocument {
    param([Parameter(Mandatory = $true)][string]$DocumentId)

    if (
        $null -eq $script:Document -or
        [string]::IsNullOrEmpty($script:DocumentId) -or
        $DocumentId -ne $script:DocumentId
    ) {
        throw 'The bridge document reference is not bound.'
    }
}

function Test-BoundDocumentLive {
    if ($null -eq $script:Document) {
        return $false
    }
    $application = $null
    try {
        $null = $script:Document.Name
        $application = $script:Document.Application
        $null = $application.Name
        return $true
    }
    catch {
        return $false
    }
    finally {
        Release-ComReference -Value $application
    }
}

function New-RangeValue {
    param(
        [Parameter(Mandatory = $true)][int]$Start,
        [Parameter(Mandatory = $true)][int]$End
    )

    return [ordered]@{
        start = $Start
        end = $End
        revision = $script:Revision
    }
}

function Resolve-BodyAnchor {
    param([Parameter(Mandatory = $true)]$Anchor)

    $documentEnd = Get-DocumentEnd
    switch ([string]$Anchor.kind) {
        'documentStart' { return 0 }
        'documentEnd' { return $documentEnd }
        'before' {
            if ([string]$Anchor.range.revision -ne $script:Revision) {
                throw [StaleDocumentRevisionException]::new(
                    'The requested Content Range has a stale revision.'
                )
            }
            $start = [int]$Anchor.range.start
            $end = [int]$Anchor.range.end
            if ($start -lt 0 -or $end -lt $start -or $end -gt $documentEnd) {
                throw 'The requested Content Range is outside the document body.'
            }
            return $start
        }
        'after' {
            if ([string]$Anchor.range.revision -ne $script:Revision) {
                throw [StaleDocumentRevisionException]::new(
                    'The requested Content Range has a stale revision.'
                )
            }
            $start = [int]$Anchor.range.start
            $end = [int]$Anchor.range.end
            if ($start -lt 0 -or $end -lt $start -or $end -gt $documentEnd) {
                throw 'The requested Content Range is outside the document body.'
            }
            return $end
        }
        default { throw 'Unsupported body anchor.' }
    }
}

function Test-ParagraphBoundary {
    param([Parameter(Mandatory = $true)][int]$Position)

    if ($Position -eq 0) {
        return $true
    }
    $probe = $null
    try {
        $probe = $script:Document.Range($Position - 1, $Position)
        return ([string]$probe.Text) -eq "`r"
    }
    finally {
        Release-ComReference -Value $probe
    }
}

function Convert-HexColorToOle {
    param([Parameter(Mandatory = $true)][string]$Color)

    $red = [Convert]::ToInt32($Color.Substring(1, 2), 16)
    $green = [Convert]::ToInt32($Color.Substring(3, 2), 16)
    $blue = [Convert]::ToInt32($Color.Substring(5, 2), 16)
    return $red -bor ($green -shl 8) -bor ($blue -shl 16)
}

function Set-TextFormatPatch {
    param(
        [Parameter(Mandatory = $true)]$Range,
        [Parameter(Mandatory = $true)]$Patch
    )

    $font = $null
    try {
        $font = $Range.Font
        $names = @(Get-ObjectPropertyNames -Value $Patch)
        if ($names -contains 'fontFamily') {
            $font.Name = [string]$Patch.fontFamily
        }
        if ($names -contains 'westernFontFamily') {
            $font.NameAscii = [string]$Patch.westernFontFamily
            $font.NameOther = [string]$Patch.westernFontFamily
        }
        if ($names -contains 'eastAsiaFontFamily') {
            $font.NameFarEast = [string]$Patch.eastAsiaFontFamily
        }
        if ($names -contains 'fontSizePt') {
            $font.Size = [double]$Patch.fontSizePt
        }
        if ($names -contains 'bold') { $font.Bold = [bool]$Patch.bold }
        if ($names -contains 'italic') { $font.Italic = [bool]$Patch.italic }
        if ($names -contains 'underline') {
            $font.Underline = if (
                [string]$Patch.underline -eq 'single'
            ) { 1 } else { 0 }
        }
        if ($names -contains 'color') {
            $font.Color = Convert-HexColorToOle `
                -Color ([string]$Patch.color)
        }
    }
    finally {
        Release-ComReference -Value $font
    }
}

function Set-ParagraphFormatPatch {
    param(
        [Parameter(Mandatory = $true)]$Range,
        [Parameter(Mandatory = $true)]$Patch
    )

    $format = $null
    try {
        $format = $Range.ParagraphFormat
        $names = @(Get-ObjectPropertyNames -Value $Patch)
        if ($names -contains 'alignment') {
            $format.Alignment = switch ([string]$Patch.alignment) {
                'left' { 0 }
                'center' { 1 }
                'right' { 2 }
                'justify' { 3 }
            }
        }
        if ($names -contains 'lineSpacing') {
            $lineSpacing = $Patch.lineSpacing
            switch ([string]$lineSpacing.kind) {
                'single' { $format.LineSpacingRule = 0 }
                'oneAndHalf' { $format.LineSpacingRule = 1 }
                'double' { $format.LineSpacingRule = 2 }
                'atLeast' {
                    $format.LineSpacingRule = 3
                    $format.LineSpacing = [double]$lineSpacing.points
                }
                'exact' {
                    $format.LineSpacingRule = 4
                    $format.LineSpacing = [double]$lineSpacing.points
                }
                'multiple' {
                    $format.LineSpacingRule = 5
                    $format.LineSpacing = 12.0 * [double]$lineSpacing.value
                }
            }
        }
        if ($names -contains 'spaceBeforePt') { $format.SpaceBefore = [double]$Patch.spaceBeforePt }
        if ($names -contains 'spaceAfterPt') { $format.SpaceAfter = [double]$Patch.spaceAfterPt }
        if ($names -contains 'leftIndentPt') { $format.LeftIndent = [double]$Patch.leftIndentPt }
        if ($names -contains 'rightIndentPt') { $format.RightIndent = [double]$Patch.rightIndentPt }
        if ($names -contains 'firstLineIndentPt') { $format.FirstLineIndent = [double]$Patch.firstLineIndentPt }
    }
    finally {
        Release-ComReference -Value $format
    }
}

function Get-BooleanFormatValue {
    param($Value)

    if ($Value -eq -1 -or $Value -eq $true) {
        return $true
    }
    if ($Value -eq 0 -or $Value -eq $false) {
        return $false
    }
    return $null
}

function Get-TextFormatSnapshot {
    param([Parameter(Mandatory = $true)]$Range)

    $font = $null
    try {
        $font = $Range.Font
        $underline = switch ([int]$font.Underline) {
            0 { 'none' }
            1 { 'single' }
            default { 'other' }
        }
        $colorValue = [long]$font.Color
        $color = if ($colorValue -eq -16777216 -or $colorValue -eq 9999999) {
            if ($colorValue -eq -16777216) { 'automatic' } else { $null }
        }
        else {
            $red = $colorValue -band 0xff
            $green = ($colorValue -shr 8) -band 0xff
            $blue = ($colorValue -shr 16) -band 0xff
            '#{0:X2}{1:X2}{2:X2}' -f $red, $green, $blue
        }
        $fontName = [string]$font.Name
        $westernFontName = [string]$font.NameAscii
        $eastAsiaFontName = [string]$font.NameFarEast
        $fontSize = [double]$font.Size
        return [ordered]@{
            fontFamily = if ($fontName) { $fontName } else { $null }
            westernFontFamily = if ($westernFontName) {
                $westernFontName
            }
            else {
                $null
            }
            eastAsiaFontFamily = if ($eastAsiaFontName) {
                $eastAsiaFontName
            }
            else {
                $null
            }
            fontSizePt = if ($fontSize -ge 0 -and $fontSize -ne 9999999) { $fontSize } else { $null }
            bold = Get-BooleanFormatValue -Value $font.Bold
            italic = Get-BooleanFormatValue -Value $font.Italic
            underline = $underline
            color = $color
        }
    }
    finally {
        Release-ComReference -Value $font
    }
}

function Get-LineSpacingSnapshot {
    param([Parameter(Mandatory = $true)]$ParagraphFormat)

    $rule = [int]$ParagraphFormat.LineSpacingRule
    $spacing = [double]$ParagraphFormat.LineSpacing
    switch ($rule) {
        0 { return [ordered]@{ kind = 'single' } }
        1 { return [ordered]@{ kind = 'oneAndHalf' } }
        2 { return [ordered]@{ kind = 'double' } }
        3 { return [ordered]@{ kind = 'atLeast'; points = $spacing } }
        4 { return [ordered]@{ kind = 'exact'; points = $spacing } }
        5 {
            return [ordered]@{
                kind = 'multiple'
                value = $spacing / 12.0
            }
        }
        default { return [ordered]@{ kind = 'other' } }
    }
}

function Get-ParagraphFormatSnapshot {
    param([Parameter(Mandatory = $true)]$Range)

    $format = $null
    try {
        $format = $Range.ParagraphFormat
        $alignment = switch ([int]$format.Alignment) {
            0 { 'left' }
            1 { 'center' }
            2 { 'right' }
            3 { 'justify' }
            default { 'other' }
        }
        return [ordered]@{
            alignment = $alignment
            lineSpacing = Get-LineSpacingSnapshot -ParagraphFormat $format
            spaceBeforePt = [double]$format.SpaceBefore
            spaceAfterPt = [double]$format.SpaceAfter
            leftIndentPt = [double]$format.LeftIndent
            rightIndentPt = [double]$format.RightIndent
            firstLineIndentPt = [double]$format.FirstLineIndent
        }
    }
    finally {
        Release-ComReference -Value $format
    }
}

function Get-StorySnapshot {
    param(
        [Parameter(Mandatory = $true)]$Collection,
        [Parameter(Mandatory = $true)][string]$Area,
        [Parameter(Mandatory = $true)][string]$Variant,
        [Parameter(Mandatory = $true)][int]$Index,
        [Parameter(Mandatory = $true)][int]$SectionIndex
    )

    $story = $null
    $range = $null
    try {
        $story = $Collection.Item($Index)
        $exists = [bool]$story.Exists
        $link = if ($SectionIndex -eq 0) {
            $false
        }
        else {
            [bool]$story.LinkToPrevious
        }
        $text = ''
        if ($exists) {
            $range = $story.Range
            $text = Normalize-StoryText -Text ([string]$range.Text)
        }
        if ($text.Length -gt 32768) {
            throw 'Header or footer text exceeds the supported limit.'
        }
        return [ordered]@{
            area = $Area
            variant = $Variant
            exists = $exists
            linkToPrevious = $link
            text = $text
        }
    }
    finally {
        Release-ComReference -Value $range
        Release-ComReference -Value $story
    }
}

function Get-SectionSnapshots {
    $count = [int]$script:Document.Sections.Count
    if ($count -lt 1 -or $count -gt 64) {
        throw 'Section count exceeds the supported limit.'
    }
    $snapshots = @()
    for ($sectionIndex = 0; $sectionIndex -lt $count; $sectionIndex++) {
        $section = $null
        $pageSetup = $null
        $headers = $null
        $footers = $null
        try {
            $section = $script:Document.Sections.Item($sectionIndex + 1)
            $pageSetup = $section.PageSetup
            $headers = $section.Headers
            $footers = $section.Footers
            $stories = @(
                Get-StorySnapshot -Collection $headers -Area 'header' -Variant 'primary' -Index 1 -SectionIndex $sectionIndex
                Get-StorySnapshot -Collection $headers -Area 'header' -Variant 'firstPage' -Index 2 -SectionIndex $sectionIndex
                Get-StorySnapshot -Collection $headers -Area 'header' -Variant 'evenPages' -Index 3 -SectionIndex $sectionIndex
                Get-StorySnapshot -Collection $footers -Area 'footer' -Variant 'primary' -Index 1 -SectionIndex $sectionIndex
                Get-StorySnapshot -Collection $footers -Area 'footer' -Variant 'firstPage' -Index 2 -SectionIndex $sectionIndex
                Get-StorySnapshot -Collection $footers -Area 'footer' -Variant 'evenPages' -Index 3 -SectionIndex $sectionIndex
            )
            $snapshots += [ordered]@{
                index = $sectionIndex
                layout = [ordered]@{
                    orientation = if ([int]$pageSetup.Orientation -eq 1) { 'landscape' } else { 'portrait' }
                    margins = [ordered]@{
                        top = [ordered]@{ value = [double]$pageSetup.TopMargin; unit = 'pt' }
                        right = [ordered]@{ value = [double]$pageSetup.RightMargin; unit = 'pt' }
                        bottom = [ordered]@{ value = [double]$pageSetup.BottomMargin; unit = 'pt' }
                        left = [ordered]@{ value = [double]$pageSetup.LeftMargin; unit = 'pt' }
                    }
                }
                headerFooter = [ordered]@{
                    firstPageEnabled = [bool]$pageSetup.DifferentFirstPageHeaderFooter
                    evenPagesEnabled = [bool]$pageSetup.OddAndEvenPagesHeaderFooter
                    stories = $stories
                }
            }
        }
        finally {
            Release-ComReference -Value $footers
            Release-ComReference -Value $headers
            Release-ComReference -Value $pageSetup
            Release-ComReference -Value $section
        }
    }
    return $snapshots
}

function Resolve-BodyScope {
    param([Parameter(Mandatory = $true)]$Scope)

    $documentEnd = Get-DocumentEnd
    if ($Scope.kind -eq 'document') {
        return @(0, $documentEnd)
    }
    if ($Scope.kind -ne 'range') {
        throw 'Unsupported body scope.'
    }
    if ([string]$Scope.range.revision -ne $script:Revision) {
        throw [StaleDocumentRevisionException]::new(
            'The requested Content Range has a stale revision.'
        )
    }
    $start = [int]$Scope.range.start
    $end = [int]$Scope.range.end
    if ($start -lt 0 -or $end -lt $start -or $end -gt $documentEnd) {
        throw 'The requested Content Range is outside the document body.'
    }
    return @($start, $end)
}

function Get-ParagraphSnapshots {
    param(
        [Parameter(Mandatory = $true)][int]$ScopeStart,
        [Parameter(Mandatory = $true)][int]$ScopeEnd,
        [Parameter(Mandatory = $true)][int]$ReturnedEnd,
        [Parameter(Mandatory = $true)][int]$MaxParagraphs,
        [Parameter(Mandatory = $true)][int]$MaxRuns
    )

    $snapshots = @()
    $runCount = 0
    $coveredEnd = $ScopeStart
    $limited = $false
    $paragraphs = $script:Document.Paragraphs
    try {
        for ($index = 1; $index -le [int]$paragraphs.Count; $index++) {
            $paragraph = $null
            $paragraphRange = $null
            $clipped = $null
            try {
                $paragraph = $paragraphs.Item($index)
                $paragraphRange = $paragraph.Range
                $paragraphStart = [Math]::Max($ScopeStart, [int]$paragraphRange.Start)
                $paragraphEnd = [Math]::Min(
                    $ReturnedEnd,
                    [Math]::Min($ScopeEnd, [int]$paragraphRange.End)
                )
                if ($paragraphEnd -le $paragraphStart) {
                    continue
                }
                if ($paragraphStart -ge $ReturnedEnd) {
                    break
                }
                if ($snapshots.Count -ge $MaxParagraphs) {
                    $limited = $true
                    break
                }
                $clipped = $script:Document.Range($paragraphStart, $paragraphEnd)
                $text = Normalize-WordText -Text ([string]$clipped.Text)
                $runs = @()
                if ($text.Length -gt 0 -and $runCount -lt $MaxRuns) {
                    $runs += [ordered]@{
                        range = New-RangeValue -Start $paragraphStart -End $paragraphEnd
                        text = $text
                        format = Get-TextFormatSnapshot -Range $clipped
                    }
                    $runCount++
                }
                elseif ($text.Length -gt 0) {
                    $limited = $true
                    break
                }
                $outlineLevel = [int]$paragraph.OutlineLevel
                $snapshot = [ordered]@{
                    kind = if ($outlineLevel -ge 1 -and $outlineLevel -le 9) { 'heading' } else { 'paragraph' }
                    range = New-RangeValue -Start $paragraphStart -End $paragraphEnd
                    complete = (
                        $paragraphStart -eq [int]$paragraphRange.Start -and
                        $paragraphEnd -eq [Math]::Min(
                            (Get-DocumentEnd),
                            [int]$paragraphRange.End
                        )
                    )
                    text = $text
                    runs = $runs
                    format = Get-ParagraphFormatSnapshot -Range $clipped
                }
                if ($snapshot.kind -eq 'heading') {
                    $snapshot['level'] = $outlineLevel
                }
                $snapshots += $snapshot
                $coveredEnd = $paragraphEnd
            }
            finally {
                Release-ComReference -Value $clipped
                Release-ComReference -Value $paragraphRange
                Release-ComReference -Value $paragraph
            }
        }
    }
    finally {
        Release-ComReference -Value $paragraphs
    }
    if (-not $limited) {
        $coveredEnd = $ReturnedEnd
    }
    return [ordered]@{
        paragraphs = @($snapshots)
        returnedEnd = $coveredEnd
    }
}

function Get-DocumentHeadingCount {
    $headingCount = 0
    $paragraphs = $script:Document.Paragraphs
    try {
        for ($index = 1; $index -le [int]$paragraphs.Count; $index++) {
            $paragraph = $null
            try {
                $paragraph = $paragraphs.Item($index)
                $outlineLevel = [int]$paragraph.OutlineLevel
                if ($outlineLevel -ge 1 -and $outlineLevel -le 9) {
                    $headingCount++
                }
            }
            finally {
                Release-ComReference -Value $paragraph
            }
        }
    }
    finally {
        Release-ComReference -Value $paragraphs
    }
    return $headingCount
}

function Assert-TextFormatPatch {
    param(
        [Parameter(Mandatory = $true)]$Range,
        [Parameter(Mandatory = $true)]$Patch
    )

    $observed = Get-TextFormatSnapshot -Range $Range
    $names = @(Get-ObjectPropertyNames -Value $Patch)
    if ($names -contains 'fontFamily' -and [string]$observed.fontFamily -ne [string]$Patch.fontFamily) {
        throw [ContentVerificationException]::new('The inserted font family did not read back exactly.')
    }
    if (
        $names -contains 'westernFontFamily' -and
        [string]$observed.westernFontFamily -ne
            [string]$Patch.westernFontFamily
    ) {
        throw [ContentVerificationException]::new(
            'The inserted Western font family did not read back exactly.'
        )
    }
    if (
        $names -contains 'eastAsiaFontFamily' -and
        [string]$observed.eastAsiaFontFamily -ne
            [string]$Patch.eastAsiaFontFamily
    ) {
        throw [ContentVerificationException]::new(
            'The inserted East Asian font family did not read back exactly.'
        )
    }
    if ($names -contains 'westernFontFamily') {
        $font = $null
        try {
            $font = $Range.Font
            if (
                [string]$font.NameOther -ne
                    [string]$Patch.westernFontFamily
            ) {
                throw [ContentVerificationException]::new(
                    'The inserted Western fallback font did not read back exactly.'
                )
            }
        }
        finally {
            Release-ComReference -Value $font
        }
    }
    if ($names -contains 'fontSizePt' -and [Math]::Abs([double]$observed.fontSizePt - [double]$Patch.fontSizePt) -gt 0.05) {
        throw [ContentVerificationException]::new('The inserted font size did not read back exactly.')
    }
    if ($names -contains 'bold' -and $observed.bold -ne [bool]$Patch.bold) {
        throw [ContentVerificationException]::new('The inserted bold state did not read back exactly.')
    }
    if ($names -contains 'italic' -and $observed.italic -ne [bool]$Patch.italic) {
        throw [ContentVerificationException]::new('The inserted italic state did not read back exactly.')
    }
    if ($names -contains 'underline' -and [string]$observed.underline -ne [string]$Patch.underline) {
        throw [ContentVerificationException]::new('The inserted underline state did not read back exactly.')
    }
    if ($names -contains 'color' -and [string]$observed.color -ne ([string]$Patch.color).ToUpperInvariant()) {
        throw [ContentVerificationException]::new('The inserted color did not read back exactly.')
    }
}

function Assert-ParagraphFormatPatch {
    param(
        [Parameter(Mandatory = $true)]$Range,
        [Parameter(Mandatory = $true)]$Patch
    )

    $observed = Get-ParagraphFormatSnapshot -Range $Range
    $names = @(Get-ObjectPropertyNames -Value $Patch)
    if ($names -contains 'alignment' -and [string]$observed.alignment -ne [string]$Patch.alignment) {
        throw [ContentVerificationException]::new('The inserted paragraph alignment did not read back exactly.')
    }
    foreach ($field in @('spaceBeforePt', 'spaceAfterPt', 'leftIndentPt', 'rightIndentPt', 'firstLineIndentPt')) {
        if ($names -contains $field -and [Math]::Abs([double]$observed.$field - [double]$Patch.$field) -gt 0.05) {
            throw [ContentVerificationException]::new("The inserted paragraph field $field did not read back exactly.")
        }
    }
    if ($names -contains 'lineSpacing') {
        $expected = $Patch.lineSpacing
        if ([string]$observed.lineSpacing.kind -ne [string]$expected.kind) {
            throw [ContentVerificationException]::new('The inserted line-spacing kind did not read back exactly.')
        }
        if (
            @('exact', 'atLeast') -contains [string]$expected.kind -and
            [Math]::Abs([double]$observed.lineSpacing.points - [double]$expected.points) -gt 0.05
        ) {
            throw [ContentVerificationException]::new('The inserted line spacing did not read back exactly.')
        }
        if (
            [string]$expected.kind -eq 'multiple' -and
            [Math]::Abs([double]$observed.lineSpacing.value - [double]$expected.value) -gt 0.01
        ) {
            throw [ContentVerificationException]::new('The inserted line-spacing multiple did not read back exactly.')
        }
    }
}

function Invoke-WriteContent {
    param([Parameter(Mandatory = $true)]$Arguments)

    Assert-BoundDocument -DocumentId ([string]$Arguments.documentId)
    if ([bool]$script:Document.ReadOnly) {
        throw [UnauthorizedAccessException]::new('The bound document is read-only.')
    }
    $operationArguments = $Arguments.operationArguments
    $beforeFingerprint = Sync-ContentRevision
    $revisionBefore = $script:Revision
    $position = Resolve-BodyAnchor -Anchor $operationArguments.anchor
    $blocks = @($operationArguments.blocks)
    $paragraphBlocks = [string]$blocks[0].kind -ne 'text'

    $builder = [Text.StringBuilder]::new()
    if ($paragraphBlocks -and -not (Test-ParagraphBoundary -Position $position)) {
        if ([string]$operationArguments.anchor.kind -eq 'documentEnd') {
            # Content.End-1 precedes Word's final paragraph mark.  Separate a
            # requested paragraph sequence from a non-empty last paragraph.
            [void]$builder.Append("`r")
        }
        else {
            throw [ContentAnchorBoundaryException]::new(
                'Structured paragraphs require a paragraph-boundary Body Anchor.'
            )
        }
    }
    $runPlans = @()
    $blockPlans = @()
    foreach ($block in $blocks) {
        $blockStart = $builder.Length
        foreach ($run in @($block.runs)) {
            $runStart = $builder.Length
            [void]$builder.Append([string]$run.text)
            $runPlans += [ordered]@{
                start = $runStart
                end = $builder.Length
                format = $run.format
            }
        }
        if ($paragraphBlocks) {
            [void]$builder.Append("`r")
        }
        $blockPlans += [ordered]@{
            start = $blockStart
            end = $builder.Length
            kind = [string]$block.kind
            level = if ([string]$block.kind -eq 'heading') { [int]$block.level } else { 0 }
            format = $block.format
        }
    }
    $insertedText = $builder.ToString()
    if ($insertedText.Length -lt 1) {
        throw 'The structured content must produce a non-empty insertion.'
    }

    $insertion = $null
    try {
        $insertion = $script:Document.Range($position, $position)
        $insertion.Text = $insertedText
    }
    finally {
        Release-ComReference -Value $insertion
    }

    foreach ($plan in $runPlans) {
        if ($plan.end -le $plan.start -or $null -eq $plan.format) { continue }
        $runRange = $null
        try {
            $runRange = $script:Document.Range(
                $position + [int]$plan.start,
                $position + [int]$plan.end
            )
            Set-TextFormatPatch -Range $runRange -Patch $plan.format
        }
        finally {
            Release-ComReference -Value $runRange
        }
    }
    if ($paragraphBlocks) {
        foreach ($plan in $blockPlans) {
            $blockRange = $null
            $paragraph = $null
            try {
                $blockRange = $script:Document.Range(
                    $position + [int]$plan.start,
                    $position + [int]$plan.end
                )
                if ($null -ne $plan.format) {
                    Set-ParagraphFormatPatch -Range $blockRange -Patch $plan.format
                }
                $paragraph = $blockRange.Paragraphs.Item(1)
                $paragraph.OutlineLevel = if ($plan.kind -eq 'heading') { [int]$plan.level } else { 10 }
            }
            finally {
                Release-ComReference -Value $paragraph
                Release-ComReference -Value $blockRange
            }
        }
    }

    $insertedEnd = $position + $insertedText.Length
    $readBack = $null
    try {
        $readBack = $script:Document.Range($position, $insertedEnd)
        if ([string]$readBack.Text -ne $insertedText) {
            throw [ContentVerificationException]::new('The inserted text did not read back exactly.')
        }
    }
    finally {
        Release-ComReference -Value $readBack
    }
    foreach ($plan in $runPlans) {
        if ($plan.end -le $plan.start -or $null -eq $plan.format) { continue }
        $runRange = $null
        try {
            $runRange = $script:Document.Range(
                $position + [int]$plan.start,
                $position + [int]$plan.end
            )
            Assert-TextFormatPatch -Range $runRange -Patch $plan.format
        }
        finally {
            Release-ComReference -Value $runRange
        }
    }
    if ($paragraphBlocks) {
        foreach ($plan in $blockPlans) {
            $blockRange = $null
            $paragraph = $null
            try {
                $blockRange = $script:Document.Range(
                    $position + [int]$plan.start,
                    $position + [int]$plan.end
                )
                $paragraph = $blockRange.Paragraphs.Item(1)
                $observedLevel = [int]$paragraph.OutlineLevel
                $expectedLevel = if ($plan.kind -eq 'heading') { [int]$plan.level } else { 10 }
                if ($observedLevel -ne $expectedLevel) {
                    throw [ContentVerificationException]::new('The inserted semantic paragraph kind did not read back exactly.')
                }
                if ($null -ne $plan.format) {
                    Assert-ParagraphFormatPatch -Range $blockRange -Patch $plan.format
                }
            }
            finally {
                Release-ComReference -Value $paragraph
                Release-ComReference -Value $blockRange
            }
        }
    }

    $afterFingerprint = Get-DocumentFingerprint
    if ($afterFingerprint -eq $beforeFingerprint) {
        throw [ContentVerificationException]::new('The write did not change the document revision.')
    }
    $script:Fingerprint = $afterFingerprint
    $script:Revision = New-ContentRevision
    return [ordered]@{
        revisionBefore = $revisionBefore
        revisionAfter = $script:Revision
        range = New-RangeValue -Start $position -End $insertedEnd
    }
}

function Invoke-PrepareExistingDocument {
    param([Parameter(Mandatory = $true)]$Arguments)

    if ($null -ne $script:Document -or $null -ne $script:CoordinationMutex) {
        throw 'The bridge is already bound or coordinated.'
    }
    $path = [string]$Arguments.path
    if (
        [string]::IsNullOrEmpty($path) -or
        -not [IO.Path]::IsPathRooted($path) -or
        -not $path.EndsWith('.docx', [StringComparison]::OrdinalIgnoreCase)
    ) {
        throw 'The document path must be an absolute .docx path.'
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw [IO.FileNotFoundException]::new('The document does not exist.', $path)
    }
    $canonicalPath = [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $path).Path)
    $file = Get-Item -LiteralPath $canonicalPath
    if ($file.Length -lt 1) {
        throw 'The document artifact is empty.'
    }
    $script:PreparationId = 'preparation-' + [guid]::NewGuid().ToString('N')
    $script:PreparedIdentity = Get-StableFileIdentity -Path $canonicalPath
    $script:PreparedPath = $path
    $script:PreparedCanonicalPath = $canonicalPath
    return [ordered]@{
        preparationId = $script:PreparationId
        coordinationIdentity = $script:PreparedIdentity
    }
}

function Invoke-PrepareNewDocument {
    if ($null -ne $script:Document -or $null -ne $script:CoordinationMutex) {
        throw 'The bridge is already bound or coordinated.'
    }
    $script:PreparationId = 'preparation-' + [guid]::NewGuid().ToString('N')
    $script:PreparedIdentity = 'new-' + [guid]::NewGuid().ToString('N')
    $script:PreparedPath = $null
    $script:PreparedCanonicalPath = $null
    return [ordered]@{
        preparationId = $script:PreparationId
        coordinationIdentity = $script:PreparedIdentity
    }
}

function Invoke-AcquireCoordinationGuard {
    param([Parameter(Mandatory = $true)]$Arguments)

    if ($null -ne $script:CoordinationMutex) {
        throw 'The bridge already owns a coordination guard.'
    }
    $identity = [string]$Arguments.coordinationIdentity
    if (
        [string]::IsNullOrEmpty($script:PreparationId) -or
        $identity -ne $script:PreparedIdentity
    ) {
        throw 'The coordination identity does not match the prepared document.'
    }
    $hash = Get-CoordinationHash -Identity $identity
    $mutexName = 'Global\WpsSkills.Document.' + $hash
    $stateDirectory = Join-Path $env:LOCALAPPDATA 'WpsSkills\document-coordination'
    $statePath = Join-Path $stateDirectory ($hash + '.json')
    $mutex = New-Object Threading.Mutex($false, $mutexName)
    $acquired = $false
    $abandoned = $false
    try {
        try {
            $acquired = $mutex.WaitOne(0)
        }
        catch [Threading.AbandonedMutexException] {
            $acquired = $true
            $abandoned = $true
        }
        if (-not $acquired) {
            throw [DocumentLeaseConflictException]::new(
                'Another Action Session owns the document Lease.'
            )
        }
        if (Test-Path -LiteralPath $statePath) {
            $unsafe = $true
            try {
                $previous = [IO.File]::ReadAllText($statePath) | ConvertFrom-Json -ErrorAction Stop
                $unsafe = (
                    [string]$previous.identity -ne $identity -or
                    [bool]$previous.inFlight -or
                    [string]$previous.mode -eq 'quarantine'
                )
            }
            catch {
                $unsafe = $true
            }
            if ($unsafe) {
                try { $mutex.ReleaseMutex() } catch {}
                throw [DocumentQuarantinedException]::new(
                    'The prior document owner ended without proving WPS quiescence.'
                )
            }
            Remove-Item -LiteralPath $statePath -Force
        }
        elseif ($abandoned) {
            # No persisted in-flight marker means the old owner had not entered WPS.
        }
        $script:CoordinationMutex = $mutex
        $script:CoordinationMutexName = $mutexName
        $script:CoordinationStatePath = $statePath
        $script:CoordinationGuardId = 'guard-' + [guid]::NewGuid().ToString('N')
        Write-CoordinationState -Mode 'guard' -InFlight $false
        $mutex = $null
        return [ordered]@{ guardId = $script:CoordinationGuardId }
    }
    finally {
        if ($null -ne $mutex) { $mutex.Dispose() }
    }
}

function Invoke-CommitDocumentLease {
    param([Parameter(Mandatory = $true)]$Arguments)

    if (
        $null -eq $script:CoordinationMutex -or
        [string]$Arguments.guardId -ne $script:CoordinationGuardId -or
        [string]$Arguments.documentId -ne $script:DocumentId
    ) {
        throw 'The acquisition guard and exact document cannot be committed.'
    }
    $script:CoordinationLeaseId = 'lease-' + [guid]::NewGuid().ToString('N')
    Write-CoordinationState -Mode 'lease' -InFlight $false
    return [ordered]@{ leaseId = $script:CoordinationLeaseId }
}

function Invoke-ReleaseDocumentResources {
    param([Parameter(Mandatory = $true)]$Arguments)

    if ($null -eq $script:CoordinationMutex) {
        return [ordered]@{ state = 'not_acquired' }
    }
    if (
        -not [string]::IsNullOrEmpty($script:CoordinationLeaseId) -and
        [string]$Arguments.leaseId -ne $script:CoordinationLeaseId
    ) {
        throw 'The document Lease reference does not match.'
    }
    if (
        [string]::IsNullOrEmpty($script:CoordinationLeaseId) -and
        [string]$Arguments.guardId -ne $script:CoordinationGuardId
    ) {
        throw 'The acquisition guard reference does not match.'
    }
    Invoke-DebugCreatedDocumentCleanup
    Release-CoordinationResources -Clean $true
    return [ordered]@{ state = 'released' }
}

function Invoke-AcquireExistingDocument {
    param([Parameter(Mandatory = $true)]$Arguments)

    if ($null -ne $script:Document) {
        throw 'The bridge is already bound.'
    }
    if (
        [string]$Arguments.preparationId -ne $script:PreparationId -or
        $null -eq $script:CoordinationMutex
    ) {
        throw 'The document was not prepared and guarded.'
    }
    $path = $script:PreparedPath
    $canonicalPath = $script:PreparedCanonicalPath
    if ((Get-StableFileIdentity -Path $canonicalPath) -ne $script:PreparedIdentity) {
        throw 'The document file identity changed after preparation.'
    }
    $file = Get-Item -LiteralPath $canonicalPath
    if ($file.Length -lt 1) {
        throw 'The document artifact is empty.'
    }

    $applicationCreated = Connect-WpsApplication
    # WPS COM activation can create a hidden application even in an
    # interactive Windows session. Establish is user-facing: reveal the
    # application, then activate only the exact document selected below.
    $matches = @()
    for ($index = 1; $index -le [int]$script:Documents.Count; $index++) {
        $candidate = $script:Documents.Item($index)
        try {
            if (-not [string]::IsNullOrEmpty([string]$candidate.Path)) {
                $candidatePath = [IO.Path]::GetFullPath([string]$candidate.FullName)
                if ((Get-StableFileIdentity -Path $candidatePath) -eq $script:PreparedIdentity) {
                    $matches += $candidate
                    $candidate = $null
                }
            }
        }
        finally {
            Release-ComReference -Value $candidate
        }
    }
    if ($matches.Count -gt 1) {
        foreach ($match in $matches) {
            Release-ComReference -Value $match
        }
        throw 'More than one live document claims the requested file.'
    }
    if ($matches.Count -eq 1) {
        $script:Document = $matches[0]
    }
    else {
        $script:Document = $script:Documents.Open($canonicalPath)
    }

    $actualPath = [IO.Path]::GetFullPath([string]$script:Document.FullName)
    if ((Get-StableFileIdentity -Path $actualPath) -ne $script:PreparedIdentity) {
        throw 'WPS returned a different document than the requested file.'
    }
    $script:DocumentCreatedBySession = $false
    Show-BoundDocument -UseNormalWindowState $applicationCreated
    $script:DocumentId = 'document-' + [guid]::NewGuid().ToString('N')
    $script:AuthorizedPath = $path
    $script:Revision = New-ContentRevision
    $script:Fingerprint = Get-DocumentFingerprint

    return [ordered]@{
        documentId = $script:DocumentId
        revision = $script:Revision
        persistenceState = Get-DocumentPersistenceState
        readOnly = [bool]$script:Document.ReadOnly
        artifactFormat = 'docx'
        artifactSizeBytes = [long]$file.Length
    }
}

function Invoke-AcquireNewDocument {
    param([Parameter(Mandatory = $true)]$Arguments)

    if ($null -ne $script:Document) {
        throw 'The bridge is already bound.'
    }
    if (
        [string]$Arguments.preparationId -ne $script:PreparationId -or
        $null -eq $script:CoordinationMutex
    ) {
        throw 'The new document was not prepared and guarded.'
    }
    $applicationCreated = Connect-WpsApplication
    $script:Document = $script:Documents.Add()
    $script:DocumentCreatedBySession = $true
    Show-BoundDocument -UseNormalWindowState $applicationCreated
    $script:DocumentId = 'document-' + [guid]::NewGuid().ToString('N')
    $script:Revision = New-ContentRevision
    $script:Fingerprint = Get-DocumentFingerprint
    return [ordered]@{
        documentId = $script:DocumentId
        revision = $script:Revision
        persistenceState = 'unsaved'
        readOnly = $false
    }
}

function Invoke-InspectDocument {
    param([Parameter(Mandatory = $true)]$Arguments)

    Assert-BoundDocument -DocumentId ([string]$Arguments.documentId)
    $operationArguments = $Arguments.operationArguments
    $before = Sync-ContentRevision
    $scope = Resolve-BodyScope -Scope $operationArguments.scope
    $scopeStart = [int]$scope[0]
    $scopeEnd = [int]$scope[1]
    $maxText = [int]$operationArguments.limits.maxTextCharacters
    $maxParagraphs = [int]$operationArguments.limits.maxParagraphs
    $maxRuns = [int]$operationArguments.limits.maxRuns
    $returnedEnd = [Math]::Min($scopeEnd, $scopeStart + $maxText)
    if ($returnedEnd -gt $scopeStart) {
        $probe = $script:Document.Range($scopeStart, $returnedEnd)
        try {
            $textProbe = [string]$probe.Text
            if (
                $textProbe.Length -gt 0 -and
                [char]::IsHighSurrogate($textProbe[$textProbe.Length - 1])
            ) {
                $returnedEnd--
            }
        }
        finally {
            Release-ComReference -Value $probe
        }
    }
    $paragraphResult = Get-ParagraphSnapshots `
        -ScopeStart $scopeStart `
        -ScopeEnd $scopeEnd `
        -ReturnedEnd $returnedEnd `
        -MaxParagraphs $maxParagraphs `
        -MaxRuns $maxRuns
    $returnedEnd = [int]$paragraphResult.returnedEnd
    $paragraphs = @($paragraphResult.paragraphs)
    $bodyRange = $script:Document.Range($scopeStart, $returnedEnd)
    try {
        $text = Normalize-WordText -Text ([string]$bodyRange.Text)
    }
    finally {
        Release-ComReference -Value $bodyRange
    }

    $sections = @(Get-SectionSnapshots)
    $headingCount = Get-DocumentHeadingCount
    $completeText = $null
    $completeRange = $script:Document.Content
    try {
        $completeText = [string]$completeRange.Text
    }
    finally {
        Release-ComReference -Value $completeRange
    }
    $after = Get-DocumentFingerprint
    if ($before -ne $after) {
        throw 'The document changed while it was being inspected.'
    }

    $truncated = $returnedEnd -lt $scopeEnd
    return [ordered]@{
        revision = $script:Revision
        scopeRange = New-RangeValue -Start $scopeStart -End $scopeEnd
        returnedRange = New-RangeValue -Start $scopeStart -End $returnedEnd
        text = $text
        paragraphs = $paragraphs
        structure = [ordered]@{
            paragraphCount = [int]$script:Document.Paragraphs.Count
            headingCount = $headingCount
            tableCount = [int]$script:Document.Tables.Count
            inlineImageCount = [int]$script:Document.InlineShapes.Count
            floatingImageCount = [int]$script:Document.Shapes.Count
            sectionCount = $sections.Count
            pageBreakCount = ([regex]::Matches($completeText, "`f")).Count
            sectionBreakCount = [Math]::Max(0, $sections.Count - 1)
            sections = $sections
        }
        documentState = [ordered]@{
            persistenceState = Get-DocumentPersistenceState
            readOnly = [bool]$script:Document.ReadOnly
        }
        truncated = $truncated
        remainingRange = if ($truncated) {
            New-RangeValue -Start $returnedEnd -End $scopeEnd
        }
        else {
            $null
        }
    }
}

function Invoke-SaveDocument {
    param([Parameter(Mandatory = $true)]$Arguments)

    Assert-BoundDocument -DocumentId ([string]$Arguments.documentId)
    if ([string]::IsNullOrEmpty($script:AuthorizedPath)) {
        throw [PersistenceLocatorRequiredException]::new(
            'The bound document has no persistence locator.'
        )
    }
    if ([bool]$script:Document.ReadOnly) {
        throw [UnauthorizedAccessException]::new('The bound document is read-only.')
    }
    $authorizedPath = [string]$Arguments.authorizedPath
    if ($authorizedPath -ne $script:AuthorizedPath) {
        throw 'The save locator does not match the bound document.'
    }
    $expected = [IO.Path]::GetFullPath($script:AuthorizedPath)
    $actual = [IO.Path]::GetFullPath([string]$script:Document.FullName)
    if (-not [string]::Equals($expected, $actual, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'The bound document no longer has its established locator.'
    }

    $before = Sync-ContentRevision
    $revision = $script:Revision
    $script:Document.Save()
    $after = Get-DocumentFingerprint
    if ($before -ne $after) {
        throw 'The document content changed while it was being saved.'
    }
    if (-not [bool]$script:Document.Saved) {
        throw 'WPS did not report the document as saved.'
    }
    $file = Get-Item -LiteralPath $actual
    if ($file.Length -lt 1) {
        throw 'The saved document artifact is empty.'
    }
    return [ordered]@{
        revisionBefore = $revision
        revisionAfter = $revision
        artifact = [ordered]@{
            path = $script:AuthorizedPath
            format = 'docx'
            sizeBytes = [long]$file.Length
        }
        documentState = [ordered]@{
            persistenceState = 'saved'
            readOnly = [bool]$script:Document.ReadOnly
        }
    }
}

$writerActionsPath = Join-Path $PSScriptRoot 'writer_actions.ps1'
if (-not (Test-Path -LiteralPath $writerActionsPath -PathType Leaf)) {
    throw 'The Writer Action implementation resource is missing.'
}
. $writerActionsPath

function Invoke-BridgeOperation {
    param(
        [Parameter(Mandatory = $true)][string]$RequestId,
        [Parameter(Mandatory = $true)][string]$Operation,
        [Parameter(Mandatory = $true)]$Arguments
    )

    try {
        switch ($Operation) {
            'prepare_existing_document' {
                try {
                    $data = Invoke-PrepareExistingDocument -Arguments $Arguments
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [IO.FileNotFoundException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'DOCUMENT_NOT_FOUND' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch [UnauthorizedAccessException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'DOCUMENT_ACCESS_DENIED' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'DOCUMENT_OPEN_FAILED' -Message $_.Exception.Message -BindingDisposition unchanged
                }
            }
            'prepare_new_document' {
                try {
                    $data = Invoke-PrepareNewDocument
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'DOCUMENT_CREATE_FAILED' -Message $_.Exception.Message -BindingDisposition unchanged
                }
            }
            'acquire_coordination_guard' {
                try {
                    $data = Invoke-AcquireCoordinationGuard -Arguments $Arguments
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [DocumentLeaseConflictException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'DOCUMENT_LEASE_CONFLICT' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch [DocumentQuarantinedException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'DOCUMENT_QUARANTINED' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch {
                    return New-FailureRecord -RequestId $RequestId -Outcome unknown -Code 'DOCUMENT_BINDING_UNAVAILABLE' -Message $_.Exception.Message -BindingDisposition unprovable
                }
            }
            'commit_document_lease' {
                try {
                    $data = Invoke-CommitDocumentLease -Arguments $Arguments
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch {
                    return New-FailureRecord -RequestId $RequestId -Outcome unknown -Code 'DOCUMENT_BINDING_UNAVAILABLE' -Message $_.Exception.Message -BindingDisposition unprovable
                }
            }
            'release_document_resources' {
                try {
                    $data = Invoke-ReleaseDocumentResources -Arguments $Arguments
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch {
                    return New-FailureRecord -RequestId $RequestId -Outcome unknown -Code 'DOCUMENT_BINDING_UNAVAILABLE' -Message $_.Exception.Message -BindingDisposition unprovable
                }
            }
            'acquire_existing_document' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-AcquireExistingDocument -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [IO.FileNotFoundException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'DOCUMENT_NOT_FOUND' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch {
                    return New-FailureRecord -RequestId $RequestId -Outcome unknown -Code 'DOCUMENT_OPEN_FAILED' -Message $_.Exception.Message -BindingDisposition unprovable
                }
            }
            'acquire_new_document' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-AcquireNewDocument -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch {
                    return New-FailureRecord -RequestId $RequestId -Outcome unknown -Code 'DOCUMENT_CREATE_FAILED' -Message $_.Exception.Message -BindingDisposition unprovable
                }
            }
            'probe_bound_document' {
                $live = Invoke-CoordinatedWpsCall {
                    return (
                        [string]$Arguments.documentId -eq $script:DocumentId -and
                        (Test-BoundDocumentLive)
                    )
                }
                return New-SuccessRecord -RequestId $RequestId -Data ([ordered]@{ live = $live })
            }
            'insert_structured_body_content' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-WriteContent -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [StaleDocumentRevisionException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'STALE_CONTENT_RANGE' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch [ContentAnchorBoundaryException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'CONTENT_ANCHOR_NOT_PARAGRAPH_BOUNDARY' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch [UnauthorizedAccessException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'DOCUMENT_READ_ONLY' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch [ContentVerificationException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome unknown -Code 'CONTENT_VERIFICATION_FAILED' -Message $_.Exception.Message -BindingDisposition unprovable
                }
                catch {
                    return New-FailureRecord -RequestId $RequestId -Outcome unknown -Code 'CONTENT_WRITE_FAILED' -Message $_.Exception.Message -BindingDisposition unprovable
                }
            }
            'read_revision_coherent_snapshot' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-InspectDocument -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [StaleDocumentRevisionException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'STALE_DOCUMENT_REVISION' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'CONTENT_READ_FAILED' -Message $_.Exception.Message -BindingDisposition unchanged
                }
            }
            'find_literal_body_content' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-FindContent -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [StaleDocumentRevisionException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'STALE_CONTENT_RANGE' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch {
                    return New-WriterOperationFailureRecord `
                        -RequestId $RequestId `
                        -ErrorRecord $_ `
                        -DefaultCode 'CONTENT_READ_FAILED' `
                        -DefaultOutcome failed `
                        -DefaultBindingDisposition unchanged
                }
            }
            'replace_body_content' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-ReplaceContent -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [StaleDocumentRevisionException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'STALE_CONTENT_RANGE' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch {
                    return New-WriterOperationFailureRecord `
                        -RequestId $RequestId `
                        -ErrorRecord $_ `
                        -DefaultCode 'CONTENT_WRITE_FAILED' `
                        -DefaultOutcome unknown `
                        -DefaultBindingDisposition unprovable
                }
            }
            'insert_plain_text_table' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-InsertTable -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [StaleDocumentRevisionException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'STALE_CONTENT_RANGE' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch [ContentAnchorBoundaryException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'CONTENT_ANCHOR_NOT_PARAGRAPH_BOUNDARY' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch {
                    return New-WriterOperationFailureRecord `
                        -RequestId $RequestId `
                        -ErrorRecord $_ `
                        -DefaultCode 'TABLE_APPLY_FAILED' `
                        -DefaultOutcome unknown `
                        -DefaultBindingDisposition unprovable
                }
            }
            'insert_embedded_image' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-InsertImage -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [StaleDocumentRevisionException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'STALE_CONTENT_RANGE' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch [ContentAnchorBoundaryException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'CONTENT_ANCHOR_NOT_PARAGRAPH_BOUNDARY' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch {
                    return New-WriterOperationFailureRecord `
                        -RequestId $RequestId `
                        -ErrorRecord $_ `
                        -DefaultCode 'IMAGE_APPLY_FAILED' `
                        -DefaultOutcome unknown `
                        -DefaultBindingDisposition unprovable
                }
            }
            'update_header_footer_stories' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-SetHeaderFooter -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [StaleDocumentRevisionException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'STALE_DOCUMENT_REVISION' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch {
                    return New-WriterOperationFailureRecord `
                        -RequestId $RequestId `
                        -ErrorRecord $_ `
                        -DefaultCode 'HEADER_FOOTER_APPLY_FAILED' `
                        -DefaultOutcome unknown `
                        -DefaultBindingDisposition unprovable
                }
            }
            'update_page_layout' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-SetPageLayout -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [StaleDocumentRevisionException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'STALE_DOCUMENT_REVISION' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch {
                    return New-WriterOperationFailureRecord `
                        -RequestId $RequestId `
                        -ErrorRecord $_ `
                        -DefaultCode 'PAGE_LAYOUT_APPLY_FAILED' `
                        -DefaultOutcome unknown `
                        -DefaultBindingDisposition unprovable
                }
            }
            'insert_body_break' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-InsertBreak -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [StaleDocumentRevisionException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'STALE_CONTENT_RANGE' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch [ContentAnchorBoundaryException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'CONTENT_ANCHOR_NOT_PARAGRAPH_BOUNDARY' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch {
                    return New-WriterOperationFailureRecord `
                        -RequestId $RequestId `
                        -ErrorRecord $_ `
                        -DefaultCode 'BREAK_APPLY_FAILED' `
                        -DefaultOutcome unknown `
                        -DefaultBindingDisposition unprovable
                }
            }
            'save_existing_artifact' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-SaveDocument -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch [PersistenceLocatorRequiredException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'PERSISTENCE_LOCATOR_REQUIRED' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch [UnauthorizedAccessException] {
                    return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'DOCUMENT_READ_ONLY' -Message $_.Exception.Message -BindingDisposition unchanged
                }
                catch {
                    return New-FailureRecord -RequestId $RequestId -Outcome unknown -Code 'OUTPUT_WRITE_FAILED' -Message $_.Exception.Message -BindingDisposition unprovable
                }
            }
            'export_pdf_artifact' {
                try {
                    $data = Invoke-CoordinatedWpsCall {
                        Invoke-ExportPdf -Arguments $Arguments
                    }
                    return New-SuccessRecord -RequestId $RequestId -Data $data
                }
                catch {
                    return New-WriterOperationFailureRecord `
                        -RequestId $RequestId `
                        -ErrorRecord $_ `
                        -DefaultCode 'OUTPUT_WRITE_FAILED' `
                        -DefaultOutcome unknown `
                        -DefaultBindingDisposition unchanged
                }
            }
            default {
                return New-FailureRecord -RequestId $RequestId -Outcome failed -Code 'WORD_CAPABILITY_UNAVAILABLE' -Message 'The Writer bridge operation is not implemented.' -BindingDisposition unchanged
            }
        }
    }
    catch {
        return New-FailureRecord -RequestId $RequestId -Outcome unknown -Code 'RESPONSE_LOST' -Message $_.Exception.Message -BindingDisposition unprovable
    }
}

try {
    while ($null -ne ($line = [Console]::In.ReadLine())) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        $requestId = 'unparsed'
        try {
            $request = $line | ConvertFrom-Json -ErrorAction Stop
            if (-not (Test-ExactFields -Value $request -Expected @('requestId', 'traceId', 'operation', 'arguments'))) {
                throw 'A bridge request has invalid fields.'
            }
            $requestId = [string]$request.requestId
            if ([string]::IsNullOrEmpty($requestId)) {
                throw 'A bridge requestId must be non-empty.'
            }
            if ([string]::IsNullOrEmpty([string]$request.traceId)) {
                throw 'A bridge traceId must be non-empty.'
            }
            if ([string]::IsNullOrEmpty([string]$request.operation)) {
                throw 'A bridge operation must be non-empty.'
            }
            if ($request.arguments -isnot [pscustomobject]) {
                throw 'Bridge arguments must be an object.'
            }
            $response = Invoke-BridgeOperation `
                -RequestId $requestId `
                -Operation ([string]$request.operation) `
                -Arguments $request.arguments
        }
        catch {
            $response = New-FailureRecord `
                -RequestId $requestId `
                -Outcome failed `
                -Code 'INVALID_PARAMS' `
                -Message $_.Exception.Message `
                -BindingDisposition unchanged
        }
        Write-BridgeRecord -Record $response
    }
}
finally {
    try {
        Release-CoordinationResources -Clean $true
    }
    catch {
    }
    Release-ComReference -Value $script:Document
    Release-ComReference -Value $script:Documents
    Release-ComReference -Value $script:Application
    $script:Document = $null
    $script:Documents = $null
    $script:Application = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
