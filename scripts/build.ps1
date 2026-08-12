[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',

    [ValidateSet('Build', 'Rebuild')]
    [string]$Target = 'Rebuild',

    [string]$CertificateThumbprint
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$workspace = Split-Path -Parent $PSScriptRoot
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path -LiteralPath $vswhere -PathType Leaf)) {
    throw 'Visual Studio Installer\vswhere.exe was not found. Install Visual Studio 2022.'
}

$vsRoot = & $vswhere -latest -products * -requires `
    Microsoft.Component.MSBuild `
    Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
    -property installationPath
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($vsRoot)) {
    throw 'Visual Studio 2022 with MSBuild and Desktop development with C++ was not found.'
}

$msbuild = Join-Path $vsRoot.Trim() 'MSBuild\Current\Bin\MSBuild.exe'
if (-not (Test-Path -LiteralPath $msbuild -PathType Leaf)) {
    throw "MSBuild was not found below $vsRoot."
}

function Invoke-MSBuild {
    param([string[]]$Arguments)

    & $msbuild @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "MSBuild failed with exit code $LASTEXITCODE."
    }
}

if ($Configuration -eq 'Release' -and
    [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
    throw 'Release builds require an explicit production CertificateThumbprint; the automatic development certificate is Debug-only.'
}

$certificate = $null
if (-not [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
    $normalizedThumbprint = $CertificateThumbprint.Replace(' ', '').ToUpperInvariant()
    if ($normalizedThumbprint -notmatch '^[0-9A-F]{40}$') {
        throw 'CertificateThumbprint must be a 40-character SHA-1 certificate thumbprint.'
    }
    $certificate = Get-Item -LiteralPath "Cert:\CurrentUser\My\$normalizedThumbprint" -ErrorAction Stop
    if (-not $certificate.HasPrivateKey) {
        throw 'The selected signing certificate has no accessible private key.'
    }
    if ($certificate.NotAfter -le (Get-Date)) {
        throw 'The selected signing certificate has expired.'
    }
    if (@($certificate.EnhancedKeyUsageList | Where-Object {
            $_.ObjectId -eq '1.3.6.1.5.5.7.3.3'
        }).Count -eq 0) {
        throw 'The selected certificate is not valid for code signing.'
    }
    if ($Configuration -eq 'Release' -and
        $certificate.Subject -eq 'CN=RAGSearch Development') {
        throw 'The local RAGSearch Development certificate cannot sign Release builds.'
    }
}
else {
    $certificate = Get-ChildItem Cert:\CurrentUser\My |
        Where-Object {
            $_.Subject -eq 'CN=RAGSearch Development' -and
            $_.HasPrivateKey -and
            $_.NotAfter -gt (Get-Date).AddDays(1) -and
            @($_.EnhancedKeyUsageList | Where-Object {
                $_.ObjectId -eq '1.3.6.1.5.5.7.3.3'
            }).Count -gt 0
        } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1

    if ($null -eq $certificate) {
        Write-Host 'Creating a non-exportable CurrentUser development code-signing certificate...'
        $certificate = New-SelfSignedCertificate `
            -Subject 'CN=RAGSearch Development' `
            -Type CodeSigningCert `
            -CertStoreLocation 'Cert:\CurrentUser\My' `
            -NotAfter (Get-Date).AddYears(2) `
            -KeyAlgorithm RSA `
            -KeyLength 2048 `
            -HashAlgorithm SHA256 `
            -KeyExportPolicy NonExportable
    }
}

Write-Host "Using local VSTO manifest certificate $($certificate.Thumbprint)."
$solution = Join-Path $workspace 'RAGSearch.sln'
Invoke-MSBuild @(
    $solution,
    "/t:$Target",
    '/m',
    "/p:Configuration=$Configuration",
    '/p:Platform=x64',
    '/p:SignManifests=true',
    "/p:ManifestCertificateThumbprint=$($certificate.Thumbprint)"
)

Write-Host 'RAGSearch build completed.'
