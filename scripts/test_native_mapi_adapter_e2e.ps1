$ErrorActionPreference = 'Stop'

function Quote-Path([string]$Value) {
    return '"' + $Value.Replace('"', '""') + '"'
}

$testRoot = Join-Path $env:TEMP ('RAGSearch-MapiAdapter-E2E-' + [Guid]::NewGuid().ToString('N'))
$workspace = Split-Path -Parent $PSScriptRoot
$python = Join-Path $workspace 'service\.venv\Scripts\python.exe'
$service = Join-Path $workspace 'service\run.py'
$adapter = Join-Path $workspace 'service\import_native_mapi.py'
$probe = Join-Path $workspace 'native-mapi-probe\build-direct\NativeMapiProbe.exe'
$data = Join-Path $testRoot 'data'
$spool = Join-Path $testRoot 'spool'
$token = Join-Path $testRoot 'service-token'
$port = 8877
$process = $null

New-Item -ItemType Directory -Path $testRoot | Out-Null
try {
    $arguments = @(
        (Quote-Path $service)
        '--port', $port
        '--data-dir', (Quote-Path $data)
        '--spool-dir', (Quote-Path $spool)
        '--token-path', (Quote-Path $token)
    ) -join ' '
    $process = Start-Process $python -ArgumentList $arguments -WindowStyle Hidden -PassThru

    $ready = $false
    for ($attempt = 0; $attempt -lt 40; $attempt++) {
        try {
            $health = Invoke-RestMethod -Uri "http://127.0.0.1:$port/health" -TimeoutSec 1
            if ($health.status -eq 'ok') {
                $ready = $true
                break
            }
        }
        catch { }
        Start-Sleep -Milliseconds 250
    }
    if (-not $ready) {
        throw 'Temporary service did not become ready.'
    }

    & $python $adapter `
        --executable $probe `
        --service-url "http://127.0.0.1:$port" `
        --token-path $token `
        --spool-dir $spool `
        --max-stores 2 `
        --max-folders 10 `
        --max-messages 3 `
        --body-preview-chars 500 `
        --store-contains Archives
    if ($LASTEXITCODE -ne 0) {
        throw "Adapter exited with code $LASTEXITCODE."
    }

    $secret = (Get-Content -LiteralPath $token -Raw).Trim()
    $stats = Invoke-RestMethod `
        -Uri "http://127.0.0.1:$port/v1/stats" `
        -Headers @{ 'X-RAGSearch-Token' = $secret }
    if ($stats.messages -ne 3 -or $stats.attachments -lt 1 -or $stats.chunks -lt 3) {
        throw "Unexpected temporary service stats: $($stats | ConvertTo-Json -Compress)"
    }
    $spoolLeftovers = @(Get-ChildItem -LiteralPath $spool -Force)
    if ($spoolLeftovers.Count -ne 0) {
        throw "Adapter left $($spoolLeftovers.Count) item(s) in its temporary spool."
    }
    $stats | ConvertTo-Json -Depth 5
    Write-Output 'NATIVE_MAPI_ADAPTER_E2E=PASS'
}
finally {
    if ($null -ne $process -and -not $process.HasExited) {
        Stop-Process -Id $process.Id -Force
        [void]$process.WaitForExit(5000)
    }

    if (Test-Path -LiteralPath $testRoot) {
        $resolved = (Resolve-Path -LiteralPath $testRoot).Path
        $temporaryRoot = (Resolve-Path -LiteralPath $env:TEMP).Path.TrimEnd('\')
        $expectedPrefix = $temporaryRoot + '\RAGSearch-MapiAdapter-E2E-'
        if (-not $resolved.StartsWith($expectedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to clean unexpected path: $resolved"
        }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
