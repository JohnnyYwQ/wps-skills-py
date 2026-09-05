function Throw-WriterOperationFailure {
    param(
        [Parameter(Mandatory = $true)][string]$Code,
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('failed', 'unknown')][string]$Outcome = 'failed',
        [ValidateSet('unchanged', 'lost', 'unprovable')]
        [string]$BindingDisposition = 'unchanged'
    )

    $exception = [InvalidOperationException]::new($Message)
    $exception.Data['WpsCode'] = $Code
    $exception.Data['WpsOutcome'] = $Outcome
    $exception.Data['WpsBindingDisposition'] = $BindingDisposition
    throw $exception
}

function New-WriterOperationFailureRecord {
    param(
        [Parameter(Mandatory = $true)][string]$RequestId,
        [Parameter(Mandatory = $true)]$ErrorRecord,
        [Parameter(Mandatory = $true)][string]$DefaultCode,
        [ValidateSet('failed', 'unknown')][string]$DefaultOutcome,
        [ValidateSet('unchanged', 'lost', 'unprovable')]
        [string]$DefaultBindingDisposition
    )

    $exception = $ErrorRecord.Exception
    $code = $DefaultCode
    $outcome = $DefaultOutcome
    $disposition = $DefaultBindingDisposition
    if ($null -ne $exception -and $null -ne $exception.Data) {
        if ($exception.Data.Contains('WpsCode')) {
            $code = [string]$exception.Data['WpsCode']
        }
        if ($exception.Data.Contains('WpsOutcome')) {
            $outcome = [string]$exception.Data['WpsOutcome']
        }
        if ($exception.Data.Contains('WpsBindingDisposition')) {
            $disposition = [string]$exception.Data['WpsBindingDisposition']
        }
    }
    return New-FailureRecord `
        -RequestId $RequestId `
        -Outcome $outcome `
        -Code $code `
        -Message $exception.Message `
        -BindingDisposition $disposition
}

function Convert-LengthToPoints {
    param([Parameter(Mandatory = $true)]$Length)

    $value = [double]$Length.value
    switch ([string]$Length.unit) {
        'pt' { return $value }
        'in' { return $value * 72.0 }
        'cm' { return $value * 72.0 / 2.54 }
        'mm' { return $value * 72.0 / 25.4 }
        default {
            Throw-WriterOperationFailure `
                -Code 'INVALID_PARAMS' `
                -Message 'The requested length unit is unsupported.'
        }
    }
}

function Complete-ContentMutation {
    param(
        [Parameter(Mandatory = $true)][string]$BeforeFingerprint,
        [Parameter(Mandatory = $true)][string]$VerificationCode,
        [Parameter(Mandatory = $true)][string]$VerificationMessage,
        [bool]$MustChange = $true
    )

    $afterFingerprint = Get-DocumentFingerprint
    if ($MustChange -and $afterFingerprint -eq $BeforeFingerprint) {
        Throw-WriterOperationFailure `
            -Code $VerificationCode `
            -Message $VerificationMessage `
            -Outcome unknown `
            -BindingDisposition unprovable
    }
    if ($afterFingerprint -ne $BeforeFingerprint) {
        $script:Fingerprint = $afterFingerprint
        $script:Revision = New-ContentRevision
        return $true
    }
    return $false
}

function Test-WordCharacter {
    param([Parameter(Mandatory = $true)][char]$Character)

    return (
        [char]::IsLetterOrDigit($Character) -or
        [char]::IsLetter($Character) -or
        [char]::IsNumber($Character) -or
        $Character -eq '_'
    )
}

function Find-LiteralBodyMatches {
    param(
        [Parameter(Mandatory = $true)]$Query,
        [Parameter(Mandatory = $true)][int]$Maximum
    )

    $scope = Resolve-BodyScope -Scope $Query.scope
    $scopeStart = [int]$scope[0]
    $scopeEnd = [int]$scope[1]
    $range = $null
    try {
        $range = $script:Document.Range($scopeStart, $scopeEnd)
        $bodyText = [string]$range.Text
    }
    finally {
        Release-ComReference -Value $range
    }
    $needle = [string]$Query.text
    $comparison = if ([bool]$Query.caseSensitive) {
        [StringComparison]::Ordinal
    }
    else {
        [StringComparison]::OrdinalIgnoreCase
    }
    $matches = @()
    $cursor = 0
    while ($cursor -le $bodyText.Length - $needle.Length) {
        $found = $bodyText.IndexOf($needle, $cursor, $comparison)
        if ($found -lt 0) { break }
        $end = $found + $needle.Length
        $wholeWord = $true
        if ([bool]$Query.wholeWord) {
            if ($found -gt 0 -and (Test-WordCharacter -Character $bodyText[$found - 1])) {
                $wholeWord = $false
            }
            if (
                $end -lt $bodyText.Length -and
                (Test-WordCharacter -Character $bodyText[$end])
            ) {
                $wholeWord = $false
            }
        }
        if ($wholeWord) {
            $matches += [ordered]@{
                start = $scopeStart + $found
                end = $scopeStart + $end
                text = $bodyText.Substring($found, $needle.Length)
            }
            $cursor = $end
            if ($matches.Count -ge $Maximum) { break }
        }
        else {
            $cursor = $found + 1
        }
    }
    return [ordered]@{
        scopeStart = $scopeStart
        scopeEnd = $scopeEnd
        matches = @($matches)
    }
}

function Invoke-FindContent {
    param([Parameter(Mandatory = $true)]$Arguments)

    Assert-BoundDocument -DocumentId ([string]$Arguments.documentId)
    $operationArguments = $Arguments.operationArguments
    $beforeFingerprint = Sync-ContentRevision
    $limit = [int]$operationArguments.limit
    $found = Find-LiteralBodyMatches `
        -Query $operationArguments.query `
        -Maximum ($limit + 1)
    $afterFingerprint = Get-DocumentFingerprint
    if ($afterFingerprint -ne $beforeFingerprint) {
        Throw-WriterOperationFailure `
            -Code 'CONTENT_CHANGED_DURING_ACTION' `
            -Message 'The document changed while content was being found.'
    }
    $allMatches = @($found.matches)
    $truncated = $allMatches.Count -gt $limit
    $returned = @($allMatches | Select-Object -First $limit)
    $matches = @(
        foreach ($match in $returned) {
            [ordered]@{
                range = New-RangeValue -Start $match.start -End $match.end
                text = [string]$match.text
            }
        }
    )
    $remaining = $null
    if ($truncated) {
        $remainingStart = [int]$returned[-1].end
        if ($remainingStart -lt [int]$found.scopeEnd) {
            $remaining = New-RangeValue `
                -Start $remainingStart `
                -End ([int]$found.scopeEnd)
        }
        else {
            $truncated = $false
        }
    }
    return [ordered]@{
        revision = $script:Revision
        scopeRange = New-RangeValue `
            -Start ([int]$found.scopeStart) `
            -End ([int]$found.scopeEnd)
        matches = $matches
        truncated = $truncated
        remainingRange = $remaining
    }
}

function New-StructuredReplacementPlan {
    param([Parameter(Mandatory = $true)]$Replacement)

    if ([string]$Replacement.kind -eq 'delete') {
        return [ordered]@{
            text = ''
            runPlans = @()
            blockPlans = @()
            paragraphBlocks = $false
        }
    }
    $blocks = if ([string]$Replacement.kind -eq 'text') {
        @([pscustomobject]@{
            kind = 'text'
            runs = @($Replacement.runs)
        })
    }
    else {
        @($Replacement.blocks)
    }
    $paragraphBlocks = [string]$blocks[0].kind -ne 'text'
    $builder = [Text.StringBuilder]::new()
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
            level = if ([string]$block.kind -eq 'heading') {
                [int]$block.level
            }
            else {
                0
            }
            format = $block.format
        }
    }
    return [ordered]@{
        text = $builder.ToString()
        runPlans = @($runPlans)
        blockPlans = @($blockPlans)
        paragraphBlocks = $paragraphBlocks
    }
}

function Set-StructuredReplacement {
    param(
        [Parameter(Mandatory = $true)][int]$Start,
        [Parameter(Mandatory = $true)][int]$End,
        [Parameter(Mandatory = $true)]$Plan
    )

    $target = $null
    try {
        $target = $script:Document.Range($Start, $End)
        $target.Text = [string]$Plan.text
    }
    finally {
        Release-ComReference -Value $target
    }
    foreach ($runPlan in @($Plan.runPlans)) {
        if ($null -eq $runPlan.format -or $runPlan.end -le $runPlan.start) {
            continue
        }
        $runRange = $null
        try {
            $runRange = $script:Document.Range(
                $Start + [int]$runPlan.start,
                $Start + [int]$runPlan.end
            )
            Set-TextFormatPatch -Range $runRange -Patch $runPlan.format
        }
        finally {
            Release-ComReference -Value $runRange
        }
    }
    if ([bool]$Plan.paragraphBlocks) {
        foreach ($blockPlan in @($Plan.blockPlans)) {
            $blockRange = $null
            $paragraph = $null
            try {
                $blockRange = $script:Document.Range(
                    $Start + [int]$blockPlan.start,
                    $Start + [int]$blockPlan.end
                )
                if ($null -ne $blockPlan.format) {
                    Set-ParagraphFormatPatch `
                        -Range $blockRange `
                        -Patch $blockPlan.format
                }
                $paragraph = $blockRange.Paragraphs.Item(1)
                $paragraph.OutlineLevel = if ($blockPlan.kind -eq 'heading') {
                    [int]$blockPlan.level
                }
                else {
                    10
                }
            }
            finally {
                Release-ComReference -Value $paragraph
                Release-ComReference -Value $blockRange
            }
        }
    }
}

function Assert-StructuredReplacement {
    param(
        [Parameter(Mandatory = $true)][int]$Start,
        [Parameter(Mandatory = $true)]$Plan
    )

    $end = $Start + ([string]$Plan.text).Length
    $readBack = $null
    try {
        $readBack = $script:Document.Range($Start, $end)
        if ([string]$readBack.Text -ne [string]$Plan.text) {
            Throw-WriterOperationFailure `
                -Code 'CONTENT_VERIFICATION_FAILED' `
                -Message 'Replacement text did not read back exactly.' `
                -Outcome unknown `
                -BindingDisposition unprovable
        }
    }
    finally {
        Release-ComReference -Value $readBack
    }
    foreach ($runPlan in @($Plan.runPlans)) {
        if ($null -eq $runPlan.format -or $runPlan.end -le $runPlan.start) {
            continue
        }
        $runRange = $null
        try {
            $runRange = $script:Document.Range(
                $Start + [int]$runPlan.start,
                $Start + [int]$runPlan.end
            )
            Assert-TextFormatPatch -Range $runRange -Patch $runPlan.format
        }
        finally {
            Release-ComReference -Value $runRange
        }
    }
    if ([bool]$Plan.paragraphBlocks) {
        foreach ($blockPlan in @($Plan.blockPlans)) {
            $blockRange = $null
            $paragraph = $null
            try {
                $blockRange = $script:Document.Range(
                    $Start + [int]$blockPlan.start,
                    $Start + [int]$blockPlan.end
                )
                $paragraph = $blockRange.Paragraphs.Item(1)
                $expectedLevel = if ($blockPlan.kind -eq 'heading') {
                    [int]$blockPlan.level
                }
                else {
                    10
                }
                if ([int]$paragraph.OutlineLevel -ne $expectedLevel) {
                    Throw-WriterOperationFailure `
                        -Code 'CONTENT_VERIFICATION_FAILED' `
                        -Message 'Replacement paragraph kind did not read back.' `
                        -Outcome unknown `
                        -BindingDisposition unprovable
                }
                if ($null -ne $blockPlan.format) {
                    Assert-ParagraphFormatPatch `
                        -Range $blockRange `
                        -Patch $blockPlan.format
                }
            }
            finally {
                Release-ComReference -Value $paragraph
                Release-ComReference -Value $blockRange
            }
        }
    }
}

function Invoke-ReplaceContent {
    param([Parameter(Mandatory = $true)]$Arguments)

    Assert-BoundDocument -DocumentId ([string]$Arguments.documentId)
    if ([bool]$script:Document.ReadOnly) {
        Throw-WriterOperationFailure `
            -Code 'DOCUMENT_READ_ONLY' `
            -Message 'The bound document is read-only.'
    }
    $operationArguments = $Arguments.operationArguments
    $beforeFingerprint = Sync-ContentRevision
    $revisionBefore = $script:Revision
    $documentEnd = Get-DocumentEnd
    $targets = @()
    if ([string]$operationArguments.target.kind -eq 'range') {
        $requested = $operationArguments.target.range
        if ([string]$requested.revision -ne $script:Revision) {
            throw [StaleDocumentRevisionException]::new(
                'The requested Content Range has a stale revision.'
            )
        }
        $start = [int]$requested.start
        $end = [int]$requested.end
        if ($start -lt 0 -or $end -le $start -or $end -gt $documentEnd) {
            Throw-WriterOperationFailure `
                -Code 'CONTENT_RANGE_OUT_OF_BOUNDS' `
                -Message 'The replacement range is outside the document body.'
        }
        $probe = $null
        try {
            $probe = $script:Document.Range($start, $end)
            $targets += [ordered]@{
                start = $start
                end = $end
                text = [string]$probe.Text
            }
        }
        finally {
            Release-ComReference -Value $probe
        }
    }
    else {
        $expected = [int]$operationArguments.target.expectedMatchCount
        $found = Find-LiteralBodyMatches `
            -Query $operationArguments.target.query `
            -Maximum ($expected + 1)
        $targets = @($found.matches)
        if ($targets.Count -ne $expected) {
            Throw-WriterOperationFailure `
                -Code 'MATCH_COUNT_MISMATCH' `
                -Message 'The observed literal match count differs from expectedMatchCount.'
        }
    }
    $plan = New-StructuredReplacementPlan `
        -Replacement $operationArguments.replacement
    if ([bool]$plan.paragraphBlocks -and $targets.Count -gt 0) {
        foreach ($target in $targets) {
            $startIsBoundary = Test-ParagraphBoundary -Position $target.start
            $endIsBoundary = (
                [int]$target.end -eq $documentEnd -or
                (Test-ParagraphBoundary -Position $target.end)
            )
            if (-not $startIsBoundary -or -not $endIsBoundary) {
                Throw-WriterOperationFailure `
                    -Code 'CONTENT_RANGE_NOT_REPLACEABLE' `
                    -Message 'Structured paragraph replacement requires complete paragraphs.'
            }
        }
    }

    foreach ($target in @(
        $targets | Sort-Object { [int]$_.start } -Descending
    )) {
        Set-StructuredReplacement `
            -Start ([int]$target.start) `
            -End ([int]$target.end) `
            -Plan $plan
    }
    $finalRanges = @()
    $offset = 0
    foreach ($target in @(
        $targets | Sort-Object { [int]$_.start }
    )) {
        $finalStart = [int]$target.start + $offset
        Assert-StructuredReplacement -Start $finalStart -Plan $plan
        $finalEnd = $finalStart + ([string]$plan.text).Length
        $finalRanges += [ordered]@{
            start = $finalStart
            end = $finalEnd
        }
        $offset += ([string]$plan.text).Length - (
            [int]$target.end - [int]$target.start
        )
    }
    $changed = Complete-ContentMutation `
        -BeforeFingerprint $beforeFingerprint `
        -VerificationCode 'CONTENT_VERIFICATION_FAILED' `
        -VerificationMessage 'The replacement effect could not be verified.' `
        -MustChange ($targets.Count -gt 0 -and [string]$operationArguments.replacement.kind -eq 'delete')
    $changedCount = if ($changed) { $targets.Count } else { 0 }
    return [ordered]@{
        revisionBefore = $revisionBefore
        revisionAfter = $script:Revision
        matchedCount = $targets.Count
        changedCount = $changedCount
        ranges = @(
            foreach ($range in $finalRanges) {
                New-RangeValue -Start $range.start -End $range.end
            }
        )
    }
}

function Get-TableCellText {
    param([Parameter(Mandatory = $true)]$Cell)

    $range = $null
    try {
        $range = $Cell.Range
        $text = [string]$range.Text
        if ($text.EndsWith("`r`a")) {
            return $text.Substring(0, $text.Length - 2)
        }
        return $text.TrimEnd([char]13, [char]7)
    }
    finally {
        Release-ComReference -Value $range
    }
}

function Invoke-InsertTable {
    param([Parameter(Mandatory = $true)]$Arguments)

    Assert-BoundDocument -DocumentId ([string]$Arguments.documentId)
    if ([bool]$script:Document.ReadOnly) {
        Throw-WriterOperationFailure `
            -Code 'DOCUMENT_READ_ONLY' `
            -Message 'The bound document is read-only.'
    }
    $operationArguments = $Arguments.operationArguments
    $beforeFingerprint = Sync-ContentRevision
    $revisionBefore = $script:Revision
    $position = Resolve-BodyAnchor -Anchor $operationArguments.anchor
    $data = @($operationArguments.data)
    $rowCount = $data.Count
    $columnCount = @($data[0]).Count
    $anchorRange = $null
    $table = $null
    try {
        $anchorRange = $script:Document.Range($position, $position)
        $table = $script:Document.Tables.Add(
            $anchorRange,
            $rowCount,
            $columnCount
        )
        for ($row = 1; $row -le $rowCount; $row++) {
            for ($column = 1; $column -le $columnCount; $column++) {
                $cell = $null
                try {
                    $cell = $table.Cell($row, $column)
                    $cell.Range.Text = [string]$data[$row - 1][$column - 1]
                }
                finally {
                    Release-ComReference -Value $cell
                }
            }
        }
        $header = $null
        try {
            $header = $table.Rows.Item(1)
            $header.HeadingFormat = [bool]$operationArguments.headerRow
        }
        finally {
            Release-ComReference -Value $header
        }
        $observedData = @()
        for ($row = 1; $row -le $rowCount; $row++) {
            $observedRow = @()
            for ($column = 1; $column -le $columnCount; $column++) {
                $cell = $null
                try {
                    $cell = $table.Cell($row, $column)
                    $observedRow += Get-TableCellText -Cell $cell
                }
                finally {
                    Release-ComReference -Value $cell
                }
            }
            $observedData += ,@($observedRow)
        }
        $observedHeader = $null
        try {
            $observedHeader = $table.Rows.Item(1)
            $headerRow = [int]$observedHeader.HeadingFormat -ne 0
        }
        finally {
            Release-ComReference -Value $observedHeader
        }
        if (
            ($observedData | ConvertTo-Json -Compress) -ne
            ($data | ConvertTo-Json -Compress)
        ) {
            Throw-WriterOperationFailure `
                -Code 'TABLE_VERIFICATION_FAILED' `
                -Message 'The inserted table cells did not read back exactly.' `
                -Outcome unknown `
                -BindingDisposition unprovable
        }
        if ($headerRow -ne [bool]$operationArguments.headerRow) {
            Throw-WriterOperationFailure `
                -Code 'TABLE_VERIFICATION_FAILED' `
                -Message 'The repeating table header state did not read back.' `
                -Outcome unknown `
                -BindingDisposition unprovable
        }
        $tableStart = [int]$table.Range.Start
        $tableEnd = [int]$table.Range.End
    }
    finally {
        Release-ComReference -Value $table
        Release-ComReference -Value $anchorRange
    }
    if ($tableStart -ne $position -or $tableEnd -le $tableStart) {
        Throw-WriterOperationFailure `
            -Code 'TABLE_VERIFICATION_FAILED' `
            -Message 'The inserted table range did not match its Body Anchor.' `
            -Outcome unknown `
            -BindingDisposition unprovable
    }
    [void](Complete-ContentMutation `
        -BeforeFingerprint $beforeFingerprint `
        -VerificationCode 'TABLE_VERIFICATION_FAILED' `
        -VerificationMessage 'The inserted table did not change the document.')
    return [ordered]@{
        revisionBefore = $revisionBefore
        revisionAfter = $script:Revision
        table = [ordered]@{
            range = New-RangeValue -Start $tableStart -End $tableEnd
            rowCount = $rowCount
            columnCount = $columnCount
            headerRow = $headerRow
            data = $observedData
        }
    }
}

function Get-SelectedSectionIndexes {
    param([Parameter(Mandatory = $true)]$Selector)

    if ([string]$Selector.revision -ne $script:Revision) {
        throw [StaleDocumentRevisionException]::new(
            'The requested section selection has a stale revision.'
        )
    }
    $count = [int]$script:Document.Sections.Count
    if ($count -lt 1 -or $count -gt 64) {
        Throw-WriterOperationFailure `
            -Code 'CONTENT_LIMIT_EXCEEDED' `
            -Message 'Section count exceeds the supported limit.'
    }
    if ([string]$Selector.kind -eq 'all') {
        return @(0..($count - 1))
    }
    $indexes = @($Selector.indexes | ForEach-Object { [int]$_ })
    $invalidIndexes = @(
        $indexes | Where-Object { $_ -lt 0 -or $_ -ge $count }
    )
    if ($indexes.Count -lt 1 -or $invalidIndexes.Count -gt 0) {
        Throw-WriterOperationFailure `
            -Code 'SECTION_NOT_FOUND' `
            -Message 'A requested section index does not exist.'
    }
    return $indexes
}

function Get-SectionLayoutSnapshot {
    param([Parameter(Mandatory = $true)][int]$Index)

    $section = $null
    $pageSetup = $null
    try {
        $section = $script:Document.Sections.Item($Index + 1)
        $pageSetup = $section.PageSetup
        return [ordered]@{
            orientation = if ([int]$pageSetup.Orientation -eq 1) {
                'landscape'
            }
            else {
                'portrait'
            }
            margins = [ordered]@{
                top = [ordered]@{ value = [double]$pageSetup.TopMargin; unit = 'pt' }
                right = [ordered]@{ value = [double]$pageSetup.RightMargin; unit = 'pt' }
                bottom = [ordered]@{ value = [double]$pageSetup.BottomMargin; unit = 'pt' }
                left = [ordered]@{ value = [double]$pageSetup.LeftMargin; unit = 'pt' }
            }
        }
    }
    finally {
        Release-ComReference -Value $pageSetup
        Release-ComReference -Value $section
    }
}

function Invoke-SetPageLayout {
    param([Parameter(Mandatory = $true)]$Arguments)

    Assert-BoundDocument -DocumentId ([string]$Arguments.documentId)
    if ([bool]$script:Document.ReadOnly) {
        Throw-WriterOperationFailure `
            -Code 'DOCUMENT_READ_ONLY' `
            -Message 'The bound document is read-only.'
    }
    $operationArguments = $Arguments.operationArguments
    $beforeFingerprint = Sync-ContentRevision
    $revisionBefore = $script:Revision
    $indexes = @(Get-SelectedSectionIndexes -Selector $operationArguments.sections)
    $layoutNames = @(Get-ObjectPropertyNames -Value $operationArguments.layout)
    $plans = @()
    foreach ($index in $indexes) {
        $section = $null
        $pageSetup = $null
        try {
            $section = $script:Document.Sections.Item($index + 1)
            $pageSetup = $section.PageSetup
            $orientation = if ($layoutNames -contains 'orientation') {
                [string]$operationArguments.layout.orientation
            }
            elseif ([int]$pageSetup.Orientation -eq 1) {
                'landscape'
            }
            else {
                'portrait'
            }
            $margins = if ($layoutNames -contains 'margins') {
                [ordered]@{
                    top = Convert-LengthToPoints $operationArguments.layout.margins.top
                    right = Convert-LengthToPoints $operationArguments.layout.margins.right
                    bottom = Convert-LengthToPoints $operationArguments.layout.margins.bottom
                    left = Convert-LengthToPoints $operationArguments.layout.margins.left
                }
            }
            else {
                [ordered]@{
                    top = [double]$pageSetup.TopMargin
                    right = [double]$pageSetup.RightMargin
                    bottom = [double]$pageSetup.BottomMargin
                    left = [double]$pageSetup.LeftMargin
                }
            }
            $pageWidth = [double]$pageSetup.PageWidth
            $pageHeight = [double]$pageSetup.PageHeight
            $currentLandscape = [int]$pageSetup.Orientation -eq 1
            $requestedLandscape = $orientation -eq 'landscape'
            if ($currentLandscape -ne $requestedLandscape) {
                $temporary = $pageWidth
                $pageWidth = $pageHeight
                $pageHeight = $temporary
            }
            if (
                $margins.left + $margins.right -ge $pageWidth -or
                $margins.top + $margins.bottom -ge $pageHeight
            ) {
                Throw-WriterOperationFailure `
                    -Code 'PAGE_LAYOUT_INVALID' `
                    -Message 'Requested margins leave no positive page body.'
            }
            $plans += [ordered]@{
                index = $index
                orientation = $orientation
                margins = $margins
                before = Get-SectionLayoutSnapshot -Index $index
            }
        }
        finally {
            Release-ComReference -Value $pageSetup
            Release-ComReference -Value $section
        }
    }
    foreach ($plan in $plans) {
        $section = $null
        $pageSetup = $null
        try {
            $section = $script:Document.Sections.Item($plan.index + 1)
            $pageSetup = $section.PageSetup
            if ($layoutNames -contains 'orientation') {
                $pageSetup.Orientation = if ($plan.orientation -eq 'landscape') { 1 } else { 0 }
            }
            if ($layoutNames -contains 'margins') {
                $pageSetup.TopMargin = [double]$plan.margins.top
                $pageSetup.RightMargin = [double]$plan.margins.right
                $pageSetup.BottomMargin = [double]$plan.margins.bottom
                $pageSetup.LeftMargin = [double]$plan.margins.left
            }
        }
        finally {
            Release-ComReference -Value $pageSetup
            Release-ComReference -Value $section
        }
    }
    $observed = @()
    $changedCount = 0
    foreach ($plan in $plans) {
        $snapshot = Get-SectionLayoutSnapshot -Index $plan.index
        if (
            ($snapshot | ConvertTo-Json -Compress -Depth 8) -ne
            ($plan.before | ConvertTo-Json -Compress -Depth 8)
        ) {
            $changedCount++
        }
        if (
            $layoutNames -contains 'orientation' -and
            $snapshot.orientation -ne $plan.orientation
        ) {
            Throw-WriterOperationFailure `
                -Code 'PAGE_LAYOUT_VERIFICATION_FAILED' `
                -Message 'Page orientation did not read back exactly.' `
                -Outcome unknown `
                -BindingDisposition unprovable
        }
        if ($layoutNames -contains 'margins') {
            foreach ($side in @('top', 'right', 'bottom', 'left')) {
                if (
                    [Math]::Abs(
                        [double]$snapshot.margins.$side.value -
                        [double]$plan.margins.$side
                    ) -gt 0.5
                ) {
                    Throw-WriterOperationFailure `
                        -Code 'PAGE_LAYOUT_VERIFICATION_FAILED' `
                        -Message "Page margin $side did not read back exactly." `
                        -Outcome unknown `
                        -BindingDisposition unprovable
                }
            }
        }
        $observed += [ordered]@{
            index = [int]$plan.index
            layout = $snapshot
        }
    }
    [void](Complete-ContentMutation `
        -BeforeFingerprint $beforeFingerprint `
        -VerificationCode 'PAGE_LAYOUT_VERIFICATION_FAILED' `
        -VerificationMessage 'The page-layout effect could not be verified.' `
        -MustChange ($changedCount -gt 0))
    return [ordered]@{
        selectedSectionCount = $indexes.Count
        changedCount = $changedCount
        revisionBefore = $revisionBefore
        revisionAfter = $script:Revision
        sections = $observed
    }
}

function Invoke-SetHeaderFooter {
    param([Parameter(Mandatory = $true)]$Arguments)

    Assert-BoundDocument -DocumentId ([string]$Arguments.documentId)
    if ([bool]$script:Document.ReadOnly) {
        Throw-WriterOperationFailure `
            -Code 'DOCUMENT_READ_ONLY' `
            -Message 'The bound document is read-only.'
    }
    $operationArguments = $Arguments.operationArguments
    $beforeFingerprint = Sync-ContentRevision
    $revisionBefore = $script:Revision
    $indexes = @(Get-SelectedSectionIndexes -Selector $operationArguments.sections)
    $areaOrder = @{ header = 0; footer = 1 }
    $variantOrder = @{ primary = 0; firstPage = 1; evenPages = 2 }
    $updates = @(
        $operationArguments.updates | Sort-Object `
            @{ Expression = { $areaOrder[[string]$_.area] } }, `
            @{ Expression = { $variantOrder[[string]$_.variant] } }
    )
    $beforeSections = @{}
    foreach ($index in $indexes) {
        $beforeSections[$index] = @(
            Get-SectionSnapshots | Where-Object { $_.index -eq $index }
        )[0].headerFooter
    }
    foreach ($index in $indexes) {
        $section = $null
        $pageSetup = $null
        try {
            $section = $script:Document.Sections.Item($index + 1)
            $pageSetup = $section.PageSetup
            foreach ($update in $updates) {
                if ($update.variant -eq 'firstPage') {
                    $pageSetup.DifferentFirstPageHeaderFooter = $true
                }
                elseif ($update.variant -eq 'evenPages') {
                    $pageSetup.OddAndEvenPagesHeaderFooter = $true
                }
                $collection = $null
                $story = $null
                $range = $null
                try {
                    if ([string]$update.area -eq 'header') {
                        $collection = $section.Headers
                    }
                    else {
                        $collection = $section.Footers
                    }
                    $storyIndex = switch ([string]$update.variant) {
                        'primary' { 1 }
                        'firstPage' { 2 }
                        'evenPages' { 3 }
                    }
                    $story = $collection.Item($storyIndex)
                    switch ([string]$update.operation.kind) {
                        'replace' {
                            if ($index -gt 0) { $story.LinkToPrevious = $false }
                            $story.Exists = $true
                            $range = $story.Range
                            $range.Text = [string]$update.operation.text
                            $immediateText = Normalize-StoryText `
                                -Text ([string]$range.Text)
                            if (
                                $immediateText -ne
                                [string]$update.operation.text
                            ) {
                                Throw-WriterOperationFailure `
                                    -Code 'HEADER_FOOTER_VERIFICATION_FAILED' `
                                    -Message 'Header/footer text did not read back immediately.' `
                                    -Outcome unknown `
                                    -BindingDisposition unprovable
                            }
                        }
                        'clear' {
                            if ($index -gt 0) { $story.LinkToPrevious = $false }
                            $story.Exists = $true
                            $range = $story.Range
                            $range.Text = ''
                            $immediateText = Normalize-StoryText `
                                -Text ([string]$range.Text)
                            if ($immediateText -ne '') {
                                Throw-WriterOperationFailure `
                                    -Code 'HEADER_FOOTER_VERIFICATION_FAILED' `
                                    -Message 'Cleared header/footer text remained non-empty.' `
                                    -Outcome unknown `
                                    -BindingDisposition unprovable
                            }
                        }
                        'linkToPrevious' {
                            if ($index -eq 0) {
                                Throw-WriterOperationFailure `
                                    -Code 'HEADER_FOOTER_LINK_INVALID' `
                                    -Message 'The first section cannot link to a previous story.'
                            }
                            $story.LinkToPrevious = $true
                        }
                    }
                }
                finally {
                    Release-ComReference -Value $range
                    Release-ComReference -Value $story
                    Release-ComReference -Value $collection
                }
            }
        }
        finally {
            Release-ComReference -Value $pageSetup
            Release-ComReference -Value $section
        }
    }
    $allSnapshots = @(Get-SectionSnapshots)
    $stories = @()
    $changedCount = 0
    foreach ($index in $indexes) {
        $sectionSnapshot = @(
            $allSnapshots | Where-Object { $_.index -eq $index }
        )[0]
        if (
            ($sectionSnapshot.headerFooter | ConvertTo-Json -Compress -Depth 10) -ne
            ($beforeSections[$index] | ConvertTo-Json -Compress -Depth 10)
        ) {
            $changedCount++
        }
        foreach ($update in $updates) {
            $story = @(
                $sectionSnapshot.headerFooter.stories | Where-Object {
                    $_.area -eq $update.area -and
                    $_.variant -eq $update.variant
                }
            )[0]
            $variantEnabled = if ($update.variant -eq 'firstPage') {
                [bool]$sectionSnapshot.headerFooter.firstPageEnabled
            }
            elseif ($update.variant -eq 'evenPages') {
                [bool]$sectionSnapshot.headerFooter.evenPagesEnabled
            }
            else {
                $true
            }
            $operationKind = [string]$update.operation.kind
            if (
                $operationKind -eq 'replace' -and
                (
                    [string]$story.text -ne [string]$update.operation.text -or
                    [bool]$story.linkToPrevious -or
                    -not [bool]$story.exists
                )
            ) {
                Throw-WriterOperationFailure `
                    -Code 'HEADER_FOOTER_VERIFICATION_FAILED' `
                    -Message 'A replaced header/footer story did not read back.' `
                    -Outcome unknown `
                    -BindingDisposition unprovable
            }
            if (
                $operationKind -eq 'clear' -and
                (
                    [string]$story.text -ne '' -or
                    [bool]$story.linkToPrevious
                )
            ) {
                Throw-WriterOperationFailure `
                    -Code 'HEADER_FOOTER_VERIFICATION_FAILED' `
                    -Message 'A cleared header/footer story did not read back.' `
                    -Outcome unknown `
                    -BindingDisposition unprovable
            }
            if (
                $operationKind -eq 'linkToPrevious' -and
                -not [bool]$story.linkToPrevious
            ) {
                Throw-WriterOperationFailure `
                    -Code 'HEADER_FOOTER_VERIFICATION_FAILED' `
                    -Message 'A linked header/footer story did not read back.' `
                    -Outcome unknown `
                    -BindingDisposition unprovable
            }
            $stories += [ordered]@{
                sectionIndex = [int]$index
                area = [string]$story.area
                variant = [string]$story.variant
                variantEnabled = $variantEnabled
                exists = [bool]$story.exists
                linkToPrevious = [bool]$story.linkToPrevious
                text = [string]$story.text
            }
        }
    }
    [void](Complete-ContentMutation `
        -BeforeFingerprint $beforeFingerprint `
        -VerificationCode 'HEADER_FOOTER_VERIFICATION_FAILED' `
        -VerificationMessage 'The header/footer effect could not be verified.' `
        -MustChange ($changedCount -gt 0))
    return [ordered]@{
        selectedSectionCount = $indexes.Count
        changedCount = $changedCount
        revisionBefore = $revisionBefore
        revisionAfter = $script:Revision
        stories = $stories
    }
}

function Get-BodyCharacterPositions {
    param([Parameter(Mandatory = $true)][char]$Character)

    $documentEnd = Get-DocumentEnd
    $range = $null
    try {
        $range = $script:Document.Range(0, $documentEnd)
        $text = [string]$range.Text
    }
    finally {
        Release-ComReference -Value $range
    }
    $positions = @()
    $cursor = 0
    $needle = [string]$Character
    while ($cursor -lt $text.Length) {
        $found = $text.IndexOf(
            $needle,
            $cursor,
            [StringComparison]::Ordinal
        )
        if ($found -lt 0) { break }
        $positions += $found
        $cursor = $found + 1
    }
    return @($positions)
}

function Get-SectionIndexAtPosition {
    param([Parameter(Mandatory = $true)][int]$Position)

    $sectionCount = [int]$script:Document.Sections.Count
    $resolvedIndex = -1
    for ($index = 0; $index -lt $sectionCount; $index++) {
        $section = $null
        $sectionRange = $null
        try {
            $section = $script:Document.Sections.Item($index + 1)
            $sectionRange = $section.Range
            if ([int]$sectionRange.Start -le $Position) {
                $resolvedIndex = $index
            }
        }
        finally {
            Release-ComReference -Value $sectionRange
            Release-ComReference -Value $section
        }
    }
    if ($resolvedIndex -lt 0) {
        Throw-WriterOperationFailure `
            -Code 'BREAK_VERIFICATION_FAILED' `
            -Message 'The section containing the resolved Body Anchor was not observable.'
    }
    return $resolvedIndex
}

function Get-ObservedSectionBreakType {
    param([Parameter(Mandatory = $true)][int]$FollowingSectionIndex)

    $section = $null
    $pageSetup = $null
    try {
        $section = $script:Document.Sections.Item($FollowingSectionIndex + 1)
        $pageSetup = $section.PageSetup
        $sectionStart = [int]$pageSetup.SectionStart
    }
    catch {
        Throw-WriterOperationFailure `
            -Code 'BREAK_VERIFICATION_FAILED' `
            -Message 'The following section start type could not be read back.' `
            -Outcome unknown `
            -BindingDisposition unprovable
    }
    finally {
        Release-ComReference -Value $pageSetup
        Release-ComReference -Value $section
    }
    $observedType = switch ($sectionStart) {
        0 { 'sectionContinuous' }
        2 { 'sectionNextPage' }
        3 { 'sectionEvenPage' }
        4 { 'sectionOddPage' }
        default {
            Throw-WriterOperationFailure `
                -Code 'BREAK_VERIFICATION_FAILED' `
                -Message 'The following section has an unsupported start type.' `
                -Outcome unknown `
                -BindingDisposition unprovable
        }
    }
    return $observedType
}

function Invoke-InsertBreak {
    param([Parameter(Mandatory = $true)]$Arguments)

    Assert-BoundDocument -DocumentId ([string]$Arguments.documentId)
    if ([bool]$script:Document.ReadOnly) {
        Throw-WriterOperationFailure `
            -Code 'DOCUMENT_READ_ONLY' `
            -Message 'The bound document is read-only.'
    }
    $operationArguments = $Arguments.operationArguments
    $beforeFingerprint = Sync-ContentRevision
    $revisionBefore = $script:Revision
    $position = Resolve-BodyAnchor -Anchor $operationArguments.anchor
    $documentEndBefore = Get-DocumentEnd
    $pageBreaksBefore = @(
        Get-BodyCharacterPositions -Character ([char]12)
    )
    $sectionCountBefore = [int]$script:Document.Sections.Count
    $isPageBreak = [string]$operationArguments.type -eq 'page'
    $anchorSectionIndex = if ($isPageBreak) {
        -1
    }
    else {
        Get-SectionIndexAtPosition -Position $position
    }
    $breakConstant = switch ([string]$operationArguments.type) {
        'page' { 7 }
        'sectionNextPage' { 2 }
        'sectionContinuous' { 3 }
        'sectionEvenPage' { 4 }
        'sectionOddPage' { 5 }
    }
    $range = $null
    try {
        $range = $script:Document.Range($position, $position)
        $range.InsertBreak($breakConstant)
    }
    finally {
        Release-ComReference -Value $range
    }
    $sectionCountAfter = [int]$script:Document.Sections.Count
    $breakStart = $position
    $observedBreakType = 'page'
    $followingSectionIndex = -1
    if ($isPageBreak) {
        $documentEndAfter = Get-DocumentEnd
        $pageBreaksAfter = @(
            Get-BodyCharacterPositions -Character ([char]12)
        )
        $delta = $documentEndAfter - $documentEndBefore
        $expectedExisting = @{}
        foreach ($existing in $pageBreaksBefore) {
            $mapped = if ($existing -lt $position) {
                [int]$existing
            }
            else {
                [int]$existing + $delta
            }
            $expectedExisting[[string]$mapped] = $true
        }
        $newMarkers = @(
            $pageBreaksAfter | Where-Object {
                -not $expectedExisting.ContainsKey([string][int]$_)
            }
        )
        if (
            $pageBreaksAfter.Count -ne $pageBreaksBefore.Count + 1 -or
            $newMarkers.Count -ne 1
        ) {
            Throw-WriterOperationFailure `
                -Code 'BREAK_VERIFICATION_FAILED' `
                -Message 'The inserted page-break marker was not uniquely observable.' `
                -Outcome unknown `
                -BindingDisposition unprovable
        }
        $breakStart = [int]$newMarkers[0]
        if ($sectionCountAfter -ne $sectionCountBefore) {
            Throw-WriterOperationFailure `
                -Code 'BREAK_VERIFICATION_FAILED' `
                -Message 'A page break unexpectedly changed the section count.' `
                -Outcome unknown `
                -BindingDisposition unprovable
        }
    }
    else {
        if ($sectionCountAfter -ne $sectionCountBefore + 1) {
            Throw-WriterOperationFailure `
                -Code 'BREAK_VERIFICATION_FAILED' `
                -Message 'The section break did not add exactly one section.' `
                -Outcome unknown `
                -BindingDisposition unprovable
        }
        $followingSectionIndex = $anchorSectionIndex + 1
        if ($followingSectionIndex -ge $sectionCountAfter) {
            Throw-WriterOperationFailure `
                -Code 'BREAK_VERIFICATION_FAILED' `
                -Message 'The resulting following section was not observable.' `
                -Outcome unknown `
                -BindingDisposition unprovable
        }
        $observedBreakType = Get-ObservedSectionBreakType `
            -FollowingSectionIndex $followingSectionIndex
        if ($observedBreakType -ne [string]$operationArguments.type) {
            Throw-WriterOperationFailure `
                -Code 'BREAK_VERIFICATION_FAILED' `
                -Message 'The following section start type did not match the request.' `
                -Outcome unknown `
                -BindingDisposition unprovable
        }
    }
    [void](Complete-ContentMutation `
        -BeforeFingerprint $beforeFingerprint `
        -VerificationCode 'BREAK_VERIFICATION_FAILED' `
        -VerificationMessage 'The break did not change the document.')
    $breakResult = [ordered]@{
        type = $observedBreakType
        range = New-RangeValue -Start $breakStart -End ($breakStart + 1)
        sectionCountBefore = $sectionCountBefore
        sectionCountAfter = $sectionCountAfter
    }
    if (-not $isPageBreak) {
        $breakResult['followingSectionIndex'] = $followingSectionIndex
    }
    return [ordered]@{
        revisionBefore = $revisionBefore
        revisionAfter = $script:Revision
        break = $breakResult
    }
}

function Get-ImageSourceSnapshot {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (
        [string]::IsNullOrEmpty($Path) -or
        -not [IO.Path]::IsPathRooted($Path) -or
        -not (Test-Path -LiteralPath $Path -PathType Leaf)
    ) {
        Throw-WriterOperationFailure `
            -Code 'IMAGE_SOURCE_NOT_FOUND' `
            -Message 'The requested image source does not exist.'
    }
    $source = Get-Item -LiteralPath $Path
    if ($source.Length -lt 1 -or $source.Length -gt 26214400) {
        Throw-WriterOperationFailure `
            -Code 'IMAGE_SOURCE_LIMIT_EXCEEDED' `
            -Message 'The image source size is outside the supported limit.'
    }
    $input = $null
    $output = $null
    $stage = $null
    try {
        $input = [IO.File]::Open(
            $source.FullName,
            [IO.FileMode]::Open,
            [IO.FileAccess]::Read,
            [IO.FileShare]::Read
        )
        $signature = New-Object byte[] 8
        $read = $input.Read($signature, 0, $signature.Length)
        $input.Position = 0
        $png = (
            $read -ge 8 -and
            ([BitConverter]::ToString($signature, 0, 8)) -eq
                '89-50-4E-47-0D-0A-1A-0A'
        )
        $jpeg = (
            $read -ge 3 -and
            $signature[0] -eq 0xff -and
            $signature[1] -eq 0xd8 -and
            $signature[2] -eq 0xff
        )
        if (-not $png -and -not $jpeg) {
            Throw-WriterOperationFailure `
                -Code 'IMAGE_FORMAT_UNSUPPORTED' `
                -Message 'Only PNG and JPEG image bytes are supported.'
        }
        $extension = if ($png) { '.png' } else { '.jpg' }
        $stage = Join-Path ([IO.Path]::GetTempPath()) (
            'wps-skills-image-' + [guid]::NewGuid().ToString('N') + $extension
        )
        $output = [IO.File]::Open(
            $stage,
            [IO.FileMode]::CreateNew,
            [IO.FileAccess]::Write,
            [IO.FileShare]::None
        )
        $input.CopyTo($output)
        $output.Flush()
    }
    catch [InvalidOperationException] {
        throw
    }
    catch {
        Throw-WriterOperationFailure `
            -Code 'IMAGE_SOURCE_UNREADABLE' `
            -Message 'The image source could not be staged safely.'
    }
    finally {
        if ($null -ne $output) { $output.Dispose() }
        if ($null -ne $input) { $input.Dispose() }
    }
    $bytes = [IO.File]::ReadAllBytes($stage)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $hash = ([BitConverter]::ToString(
            $sha.ComputeHash($bytes)
        )).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        $decoded = [Drawing.Image]::FromFile($stage)
        try {
            if ($decoded.Width -lt 1 -or $decoded.Height -lt 1) {
                throw 'Image dimensions are empty.'
            }
        }
        finally {
            $decoded.Dispose()
        }
    }
    catch {
        Remove-Item -LiteralPath $stage -Force -ErrorAction SilentlyContinue
        Throw-WriterOperationFailure `
            -Code 'IMAGE_FORMAT_UNSUPPORTED' `
            -Message 'The staged image could not be decoded.'
    }
    return [ordered]@{
        stage = $stage
        mediaType = if ($png) { 'image/png' } else { 'image/jpeg' }
        byteLength = [long]$bytes.Length
        sha256 = $hash
    }
}

function Set-ImageSize {
    param(
        [Parameter(Mandatory = $true)]$Image,
        [Parameter(Mandatory = $true)]$Size
    )

    switch ([string]$Size.kind) {
        'intrinsic' { }
        'width' {
            $Image.LockAspectRatio = $true
            $Image.Width = Convert-LengthToPoints $Size.width
        }
        'height' {
            $Image.LockAspectRatio = $true
            $Image.Height = Convert-LengthToPoints $Size.height
        }
        'box' {
            $width = Convert-LengthToPoints $Size.width
            $height = Convert-LengthToPoints $Size.height
            if ([string]$Size.fit -eq 'stretch') {
                $Image.LockAspectRatio = $false
                $Image.Width = $width
                $Image.Height = $height
            }
            else {
                $Image.LockAspectRatio = $true
                $scale = [Math]::Min(
                    $width / [double]$Image.Width,
                    $height / [double]$Image.Height
                )
                $Image.Width = [double]$Image.Width * $scale
            }
        }
    }
}

function Invoke-InsertImage {
    param([Parameter(Mandatory = $true)]$Arguments)

    Assert-BoundDocument -DocumentId ([string]$Arguments.documentId)
    if ([bool]$script:Document.ReadOnly) {
        Throw-WriterOperationFailure `
            -Code 'DOCUMENT_READ_ONLY' `
            -Message 'The bound document is read-only.'
    }
    $operationArguments = $Arguments.operationArguments
    $beforeFingerprint = Sync-ContentRevision
    $revisionBefore = $script:Revision
    $position = Resolve-BodyAnchor -Anchor $operationArguments.anchor
    $source = Get-ImageSourceSnapshot `
        -Path ([string]$operationArguments.source.path)
    $anchorRange = $null
    $image = $null
    $imageRange = $null
    $imageAnchor = $null
    $wrapFormat = $null
    try {
        $anchorRange = $script:Document.Range($position, $position)
        if ([string]$operationArguments.placement.kind -eq 'inline') {
            $image = $script:Document.InlineShapes.AddPicture(
                $source.stage,
                $false,
                $true,
                $anchorRange
            )
        }
        else {
            $image = $script:Document.Shapes.AddPicture(
                $source.stage,
                $false,
                $true,
                0,
                0,
                -1,
                -1,
                $anchorRange
            )
        }
        Set-ImageSize -Image $image -Size $operationArguments.size
        $alt = $operationArguments.alternativeText
        if ([string]$alt.kind -eq 'description') {
            $image.Decorative = 0
            $image.AlternativeText = [string]$alt.text
        }
        else {
            $image.Decorative = -1
            $image.AlternativeText = ''
        }
        if ([string]$operationArguments.placement.kind -eq 'floating') {
            $placement = $operationArguments.placement
            $image.WrapFormat.Type = switch ([string]$placement.wrap) {
                'square' { 0 }
                'topBottom' { 4 }
                'behindText' { 5 }
                'inFrontOfText' { 3 }
            }
            $image.RelativeHorizontalPosition = switch (
                [string]$placement.horizontal.relativeTo
            ) {
                'margin' { 0 }
                'page' { 1 }
                'column' { 2 }
            }
            $image.RelativeVerticalPosition = switch (
                [string]$placement.vertical.relativeTo
            ) {
                'margin' { 0 }
                'page' { 1 }
                'paragraph' { 2 }
            }
            $image.Left = Convert-LengthToPoints $placement.horizontal.offset
            $image.Top = Convert-LengthToPoints $placement.vertical.offset
        }
        $observedAlt = [string]$image.AlternativeText
        $observedDecorative = [int]$image.Decorative -ne 0
        # WPS Writer 12 always reports Decorative as true, even after a false
        # assignment.  A non-empty AlternativeText is nevertheless persisted
        # as OOXML descr metadata, so that is the stable descriptive read-back.
        if (
            (
                [string]$alt.kind -eq 'description' -and
                $observedAlt -ne [string]$alt.text
            ) -or
            (
                [string]$alt.kind -eq 'decorative' -and
                (
                    $observedAlt -ne '' -or
                    -not $observedDecorative
                )
            )
        ) {
            Throw-WriterOperationFailure `
                -Code 'IMAGE_VERIFICATION_FAILED' `
                -Message 'Image alternative text did not read back.' `
                -Outcome unknown `
                -BindingDisposition unprovable
        }
        $observedWidth = [double]$image.Width
        $observedHeight = [double]$image.Height
        if ($observedWidth -le 0 -or $observedHeight -le 0) {
            Throw-WriterOperationFailure `
                -Code 'IMAGE_VERIFICATION_FAILED' `
                -Message 'Image dimensions did not read back.' `
                -Outcome unknown `
                -BindingDisposition unprovable
        }
        $requestedSize = $operationArguments.size
        switch ([string]$requestedSize.kind) {
            'width' {
                $expectedWidth = Convert-LengthToPoints $requestedSize.width
                if ([Math]::Abs($observedWidth - $expectedWidth) -gt 0.5) {
                    Throw-WriterOperationFailure `
                        -Code 'IMAGE_VERIFICATION_FAILED' `
                        -Message 'Image width did not read back exactly.' `
                        -Outcome unknown `
                        -BindingDisposition unprovable
                }
            }
            'height' {
                $expectedHeight = Convert-LengthToPoints $requestedSize.height
                if ([Math]::Abs($observedHeight - $expectedHeight) -gt 0.5) {
                    Throw-WriterOperationFailure `
                        -Code 'IMAGE_VERIFICATION_FAILED' `
                        -Message 'Image height did not read back exactly.' `
                        -Outcome unknown `
                        -BindingDisposition unprovable
                }
            }
            'box' {
                $expectedWidth = Convert-LengthToPoints $requestedSize.width
                $expectedHeight = Convert-LengthToPoints $requestedSize.height
                if ([string]$requestedSize.fit -eq 'stretch') {
                    $sizeInvalid = (
                        [Math]::Abs($observedWidth - $expectedWidth) -gt 0.5 -or
                        [Math]::Abs($observedHeight - $expectedHeight) -gt 0.5
                    )
                }
                else {
                    $sizeInvalid = (
                        $observedWidth -gt $expectedWidth + 0.5 -or
                        $observedHeight -gt $expectedHeight + 0.5
                    )
                }
                if ($sizeInvalid) {
                    Throw-WriterOperationFailure `
                        -Code 'IMAGE_VERIFICATION_FAILED' `
                        -Message 'Image box dimensions did not read back exactly.' `
                        -Outcome unknown `
                        -BindingDisposition unprovable
                }
            }
        }
        if ([string]$operationArguments.placement.kind -eq 'inline') {
            $imageRange = $image.Range
            $imageStart = [int]$imageRange.Start
            $imageEnd = [int]$imageRange.End
        }
        else {
            $wrapFormat = $image.WrapFormat
            $observedWrap = switch ([int]$wrapFormat.Type) {
                0 { 'square' }
                4 { 'topBottom' }
                5 { 'behindText' }
                3 { 'inFrontOfText' }
                default { $null }
            }
            $observedHorizontalReference = switch (
                [int]$image.RelativeHorizontalPosition
            ) {
                0 { 'margin' }
                1 { 'page' }
                2 { 'column' }
                default { $null }
            }
            $observedVerticalReference = switch (
                [int]$image.RelativeVerticalPosition
            ) {
                0 { 'margin' }
                1 { 'page' }
                2 { 'paragraph' }
                default { $null }
            }
            $observedLeft = [double]$image.Left
            $observedTop = [double]$image.Top
            $imageAnchor = $image.Anchor
            $imageStart = [int]$imageAnchor.Start
            $imageEnd = [int]$imageAnchor.End
            $placement = $operationArguments.placement
            if (
                $observedWrap -ne [string]$placement.wrap -or
                $observedHorizontalReference -ne
                    [string]$placement.horizontal.relativeTo -or
                $observedVerticalReference -ne
                    [string]$placement.vertical.relativeTo -or
                [Math]::Abs(
                    $observedLeft -
                    (Convert-LengthToPoints $placement.horizontal.offset)
                ) -gt 0.5 -or
                [Math]::Abs(
                    $observedTop -
                    (Convert-LengthToPoints $placement.vertical.offset)
                ) -gt 0.5
            ) {
                Throw-WriterOperationFailure `
                    -Code 'IMAGE_VERIFICATION_FAILED' `
                    -Message 'Floating image placement did not read back exactly.' `
                    -Outcome unknown `
                    -BindingDisposition unprovable
            }
        }
        $resultImage = [ordered]@{
            kind = [string]$operationArguments.placement.kind
            embedded = $true
            source = [ordered]@{
                mediaType = [string]$source.mediaType
                byteLength = [long]$source.byteLength
                sha256 = [string]$source.sha256
            }
            size = [ordered]@{
                width = [ordered]@{ value = $observedWidth; unit = 'pt' }
                height = [ordered]@{ value = $observedHeight; unit = 'pt' }
            }
            alternativeText = if ([string]$alt.kind -eq 'description') {
                [ordered]@{ kind = 'description'; text = [string]$alt.text }
            }
            else {
                [ordered]@{ kind = 'decorative' }
            }
        }
        if ([string]$operationArguments.placement.kind -eq 'inline') {
            $resultImage['range'] = $null
        }
        else {
            $resultImage['anchorRange'] = $null
            $resultImage['wrap'] = $observedWrap
            $resultImage['horizontal'] = [ordered]@{
                relativeTo = $observedHorizontalReference
                offset = [ordered]@{ value = $observedLeft; unit = 'pt' }
            }
            $resultImage['vertical'] = [ordered]@{
                relativeTo = $observedVerticalReference
                offset = [ordered]@{ value = $observedTop; unit = 'pt' }
            }
        }
    }
    finally {
        Release-ComReference -Value $wrapFormat
        Release-ComReference -Value $imageAnchor
        Release-ComReference -Value $imageRange
        Release-ComReference -Value $image
        Release-ComReference -Value $anchorRange
        if ($null -ne $source -and (Test-Path -LiteralPath $source.stage)) {
            Remove-Item -LiteralPath $source.stage -Force -ErrorAction SilentlyContinue
        }
    }
    [void](Complete-ContentMutation `
        -BeforeFingerprint $beforeFingerprint `
        -VerificationCode 'IMAGE_VERIFICATION_FAILED' `
        -VerificationMessage 'The inserted image did not change the document.')
    if ($resultImage.kind -eq 'inline') {
        $resultImage['range'] = New-RangeValue -Start $imageStart -End $imageEnd
    }
    else {
        $resultImage['anchorRange'] = New-RangeValue -Start $imageStart -End $imageEnd
    }
    return [ordered]@{
        revisionBefore = $revisionBefore
        revisionAfter = $script:Revision
        image = $resultImage
    }
}

function Test-PdfArtifact {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    $file = Get-Item -LiteralPath $Path
    if ($file.Length -lt 5) { return $false }
    $stream = $null
    try {
        $stream = [IO.File]::OpenRead($file.FullName)
        $signature = New-Object byte[] 5
        if ($stream.Read($signature, 0, 5) -ne 5) { return $false }
        return [Text.Encoding]::ASCII.GetString($signature) -eq '%PDF-'
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Invoke-ExportPdf {
    param([Parameter(Mandatory = $true)]$Arguments)

    Assert-BoundDocument -DocumentId ([string]$Arguments.documentId)
    $operationArguments = $Arguments.operationArguments
    $outputPath = [string]$operationArguments.outputPath
    if (
        [string]::IsNullOrEmpty($outputPath) -or
        -not [IO.Path]::IsPathRooted($outputPath) -or
        -not $outputPath.EndsWith('.pdf', [StringComparison]::OrdinalIgnoreCase)
    ) {
        Throw-WriterOperationFailure `
            -Code 'OUTPUT_PATH_INVALID' `
            -Message 'The PDF output path must be absolute and end in .pdf.'
    }
    $fullPath = [IO.Path]::GetFullPath($outputPath)
    $parent = [IO.Path]::GetDirectoryName($fullPath)
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        Throw-WriterOperationFailure `
            -Code 'OUTPUT_PARENT_NOT_FOUND' `
            -Message 'The PDF output parent directory does not exist.'
    }
    $replacedExisting = Test-Path -LiteralPath $fullPath -PathType Leaf
    if (
        $replacedExisting -and
        [string]$operationArguments.overwritePolicy -eq 'failIfExists'
    ) {
        Throw-WriterOperationFailure `
            -Code 'OUTPUT_ALREADY_EXISTS' `
            -Message 'The PDF output already exists.'
    }
    $beforeFingerprint = Sync-ContentRevision
    $revisionBefore = $script:Revision
    $stateBefore = [ordered]@{
        persistenceState = Get-DocumentPersistenceState
        readOnly = [bool]$script:Document.ReadOnly
    }
    $temporary = Join-Path $parent (
        '.' + [IO.Path]::GetFileNameWithoutExtension($fullPath) + '.' +
        [guid]::NewGuid().ToString('N') + '.tmp.pdf'
    )
    try {
        $script:Document.ExportAsFixedFormat($temporary, 17)
        if (-not (Test-PdfArtifact -Path $temporary)) {
            Throw-WriterOperationFailure `
                -Code 'OUTPUT_VERIFICATION_FAILED' `
                -Message 'WPS did not produce a readable PDF artifact.' `
                -Outcome unknown `
                -BindingDisposition unchanged
        }
        $afterFingerprint = Get-DocumentFingerprint
        $stateAfter = [ordered]@{
            persistenceState = Get-DocumentPersistenceState
            readOnly = [bool]$script:Document.ReadOnly
        }
        if (
            $afterFingerprint -ne $beforeFingerprint -or
            ($stateAfter | ConvertTo-Json -Compress) -ne
            ($stateBefore | ConvertTo-Json -Compress)
        ) {
            Throw-WriterOperationFailure `
                -Code 'DOCUMENT_CHANGED_DURING_ACTION' `
                -Message 'The document changed while it was being exported.'
        }
        Move-Item -LiteralPath $temporary -Destination $fullPath -Force
        if (-not (Test-PdfArtifact -Path $fullPath)) {
            Throw-WriterOperationFailure `
                -Code 'OUTPUT_VERIFICATION_FAILED' `
                -Message 'The final PDF artifact could not be verified.' `
                -Outcome unknown `
                -BindingDisposition unchanged
        }
        $file = Get-Item -LiteralPath $fullPath
        return [ordered]@{
            revisionBefore = $revisionBefore
            revisionAfter = $revisionBefore
            artifact = [ordered]@{
                path = $outputPath
                format = 'pdf'
                sizeBytes = [long]$file.Length
            }
            documentStateBefore = $stateBefore
            documentStateAfter = $stateAfter
            replacedExisting = [bool]$replacedExisting
        }
    }
    catch [UnauthorizedAccessException] {
        Throw-WriterOperationFailure `
            -Code 'OUTPUT_ACCESS_DENIED' `
            -Message 'Access to the PDF output path was denied.'
    }
    finally {
        if (Test-Path -LiteralPath $temporary) {
            Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        }
    }
}
