param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',

    [ValidateRange(1, 1000)]
    [int]$MaxMessages = 3,

    [string]$StoreContains = '',

    [switch]$OfflineOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$testPrefix = 'RAGSearch-OutlookMapiReader-Smoke-'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ($testPrefix + [Guid]::NewGuid().ToString('N'))
$workspace = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path
$reader = Join-Path $workspace (
    'connectors\outlook_mapi\native\bin\x64\{0}\OutlookMapiReader.exe' -f $Configuration)
$attachmentDirectory = Join-Path $testRoot 'attachments'
$diagnosticsPath = Join-Path $testRoot 'diagnostics.txt'
$maxAttachmentBytes = 64MB
$maxMessageAttachmentBytes = 8MB
$maxTotalAttachmentBytes = 64MB
$maximumAttachmentsPerMessage = 4095

if (-not (Test-Path -LiteralPath $reader -PathType Leaf)) {
    throw "OutlookMapiReader is not built: $reader"
}

New-Item -ItemType Directory -Path $attachmentDirectory -Force | Out-Null
try {
    $helpOutput = (& $reader --help | Out-String)
    if ($LASTEXITCODE -ne 0) {
        throw "OutlookMapiReader --help exited with code $LASTEXITCODE."
    }
    foreach ($expected in @(
            '--attachment-dir',
            '--max-attachment-bytes',
            '--max-message-attachment-bytes',
            '--max-total-attachment-bytes',
            '4095 rows/message')) {
        if (-not $helpOutput.Contains($expected)) {
            throw "OutlookMapiReader --help is missing '$expected'."
        }
    }

    $originalErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $invalidLimitOutput = (& $reader `
            --max-message-attachment-bytes 1099511627777 2>&1 | Out-String)
        $invalidLimitExitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $originalErrorActionPreference
    }
    if ($invalidLimitExitCode -ne 64) {
        throw "Out-of-range message attachment cap returned $invalidLimitExitCode instead of 64.`r`n$invalidLimitOutput"
    }
    Write-Output 'OUTLOOK_MAPI_READER_CLI=PASS'
    if ($OfflineOnly) {
        # The expected invalid-option probe above leaves the native process exit
        # code at 64.  Reset it so a successful offline contract check itself
        # exits successfully when invoked from CI or another PowerShell process.
        $global:LASTEXITCODE = 0
        return
    }

    $readerArguments = @(
        '--jsonl',
        '--max-stores', '16',
        '--max-folders', '100',
        '--max-messages', ([string]$MaxMessages),
        '--body-preview-chars', '200000',
        '--attachment-dir', $attachmentDirectory,
        '--max-attachment-bytes', ([string]$maxAttachmentBytes),
        '--max-message-attachment-bytes', ([string]$maxMessageAttachmentBytes),
        '--max-total-attachment-bytes', ([string]$maxTotalAttachmentBytes)
    )
    if (-not [string]::IsNullOrWhiteSpace($StoreContains)) {
        $readerArguments += @('--store-contains', $StoreContains)
    }

    $originalOutputEncoding = [Console]::OutputEncoding
    $originalErrorActionPreference = $ErrorActionPreference
    try {
        [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
        # Reader diagnostics intentionally use stderr even on a successful run.
        $ErrorActionPreference = 'Continue'
        $lines = @(
            & $reader @readerArguments 2> $diagnosticsPath |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )
        $readerExitCode = $LASTEXITCODE
    }
    finally {
        [Console]::OutputEncoding = $originalOutputEncoding
        $ErrorActionPreference = $originalErrorActionPreference
    }
    if ($readerExitCode -ne 0) {
        $diagnostics = Get-Content -LiteralPath $diagnosticsPath -Raw -ErrorAction SilentlyContinue
        throw "OutlookMapiReader exited with code $readerExitCode.`r`n$diagnostics"
    }
    if ($lines.Count -eq 0) {
        throw 'OutlookMapiReader emitted no JSONL records; use a profile/store containing mail.'
    }
    if ($lines.Count -gt $MaxMessages) {
        throw "OutlookMapiReader emitted $($lines.Count) records for --max-messages $MaxMessages."
    }

    $requiredStringFields = @(
        'store_id',
        'store_name',
        'entry_id',
        'folder_entry_id',
        'folder_path',
        'subject',
        'body',
        'sender_name',
        'sender_email',
        'to',
        'cc',
        'internet_message_id',
        'conversation_id'
    )
    $attachmentRoot = [IO.Path]::GetFullPath($attachmentDirectory).TrimEnd('\') + '\'
    $savedAttachments = 0
    [long]$totalSavedBytes = 0
    $supportsJsonDateKind = (Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')

    for ($lineIndex = 0; $lineIndex -lt $lines.Count; $lineIndex++) {
        try {
            $record = if ($supportsJsonDateKind) {
                $lines[$lineIndex] | ConvertFrom-Json -DateKind String
            }
            else {
                $lines[$lineIndex] | ConvertFrom-Json
            }
        }
        catch {
            throw "JSONL record $($lineIndex + 1) is invalid JSON: $($_.Exception.Message)"
        }

        foreach ($field in $requiredStringFields) {
            if ($record.PSObject.Properties.Name -notcontains $field -or
                $record.$field -isnot [string]) {
                throw "JSONL record $($lineIndex + 1) field '$field' must be a string."
            }
        }
        foreach ($field in @('store_id', 'entry_id', 'folder_entry_id', 'folder_path')) {
            if ([string]::IsNullOrWhiteSpace([string]$record.$field)) {
                throw "JSONL record $($lineIndex + 1) field '$field' must not be empty."
            }
        }
        foreach ($field in @('body_available', 'body_truncated', 'attachments_truncated')) {
            if ($record.PSObject.Properties.Name -notcontains $field -or
                $record.$field -isnot [bool]) {
                throw "JSONL record $($lineIndex + 1) field '$field' must be boolean."
            }
        }
        foreach ($field in @('sent_at', 'received_at', 'modified_at')) {
            if ($record.PSObject.Properties.Name -notcontains $field -or
                ($null -ne $record.$field -and $record.$field -isnot [string])) {
                throw "JSONL record $($lineIndex + 1) field '$field' must be a string or null."
            }
        }
        if ($record.PSObject.Properties.Name -notcontains 'attachments') {
            throw "JSONL record $($lineIndex + 1) has no attachments array."
        }

        $recordAttachments = @($record.attachments)
        if ($recordAttachments.Count -gt $maximumAttachmentsPerMessage) {
            throw "JSONL record $($lineIndex + 1) exceeds the per-message attachment count cap."
        }
        [long]$messageSavedBytes = 0
        foreach ($attachment in $recordAttachments) {
            foreach ($field in @('name', 'content_type', 'temp_path')) {
                if ($attachment.PSObject.Properties.Name -notcontains $field -or
                    $attachment.$field -isnot [string]) {
                    throw "JSONL attachment field '$field' must be a string."
                }
            }
            if ($attachment.PSObject.Properties.Name -notcontains 'size' -or
                [long]$attachment.size -lt 0) {
                throw 'JSONL attachment size must be a non-negative integer.'
            }
            if ([string]::IsNullOrEmpty([string]$attachment.temp_path)) {
                continue
            }

            $savedPath = (Resolve-Path -LiteralPath ([string]$attachment.temp_path)).Path
            $fullSavedPath = [IO.Path]::GetFullPath($savedPath)
            if (-not $fullSavedPath.StartsWith(
                    $attachmentRoot,
                    [StringComparison]::OrdinalIgnoreCase)) {
                throw "Attachment escaped the caller-owned directory: $fullSavedPath"
            }
            $savedFile = Get-Item -LiteralPath $fullSavedPath
            if ($savedFile.PSIsContainer -or $savedFile.Length -gt $maxAttachmentBytes) {
                throw "Attachment file violates the extraction cap: $fullSavedPath"
            }
            if ($savedFile.Length -ne [long]$attachment.size) {
                throw "Attachment size does not match JSONL metadata: $fullSavedPath"
            }
            $savedAttachments++
            $messageSavedBytes += $savedFile.Length
            $totalSavedBytes += $savedFile.Length
        }
        if ($messageSavedBytes -gt $maxMessageAttachmentBytes) {
            throw "Saved attachments for JSONL record $($lineIndex + 1) violate the per-message cap: $messageSavedBytes bytes."
        }
    }
    if ($totalSavedBytes -gt $maxTotalAttachmentBytes) {
        throw "Saved attachments violate the per-process cap: $totalSavedBytes bytes."
    }

    [pscustomobject]@{
        messages = $lines.Count
        saved_attachments = $savedAttachments
        saved_attachment_bytes = $totalSavedBytes
        configuration = $Configuration
    } | ConvertTo-Json -Compress
    Write-Output 'OUTLOOK_MAPI_READER_SMOKE=PASS'
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        $resolved = (Resolve-Path -LiteralPath $testRoot).Path
        $temporaryRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
        $expectedPrefix = $temporaryRoot + '\' + $testPrefix
        if (-not $resolved.StartsWith($expectedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to clean unexpected path: $resolved"
        }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
