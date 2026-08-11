param(
    [string]$LogPath = "$env:TEMP\ragsearch-guard-probe.log"
)

$ErrorActionPreference = 'Stop'
$outlook = $null
$namespace = $null
$inbox = $null
$items = $null
$item = $null

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.Session
    $inbox = $namespace.GetDefaultFolder(6) # olFolderInbox
    $items = $inbox.Items
    $item = $items.GetFirst()

    $lines = @(
        "timestamp=$([DateTime]::UtcNow.ToString('o'))"
        "is_trusted=$($outlook.IsTrusted)"
        "subject=$($item.Subject)"
        # Body is deliberately protected by Outlook Object Model Guard.
        "body_length=$($item.Body.Length)"
        "result=completed"
    )
    $lines | Set-Content -LiteralPath $LogPath -Encoding UTF8
}
catch {
    @(
        "timestamp=$([DateTime]::UtcNow.ToString('o'))"
        "result=failed"
        "exception=$($_.Exception.GetType().FullName)"
        "message=$($_.Exception.Message)"
    ) | Set-Content -LiteralPath $LogPath -Encoding UTF8
    throw
}
finally {
    foreach ($value in @($item, $items, $inbox, $namespace, $outlook)) {
        if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)
        }
    }
}
