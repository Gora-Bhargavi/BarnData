$base = 'C:\Program Files\dotnet\sdk\10.0.103\Sdks'
$sdks = @(
    'Microsoft.NET.SDK.WorkloadAutoImportPropsLocator',
    'Microsoft.NET.SDK.WorkloadManifestTargetsLocator'
)

foreach ($sdk in $sdks) {
    $sdkDir = Join-Path $base "$sdk\Sdk"
    if (-not (Test-Path $sdkDir)) {
        New-Item -ItemType Directory -Path $sdkDir -Force | Out-Null
    }

    $props = Join-Path $sdkDir 'Sdk.props'
    $targets = Join-Path $sdkDir 'Sdk.targets'

    if (-not (Test-Path $props)) {
        Set-Content -Path $props -Value '<Project />' -Encoding UTF8
    }

    if (-not (Test-Path $targets)) {
        Set-Content -Path $targets -Value '<Project />' -Encoding UTF8
    }
}

Write-Output 'SDK_STUBS_READY'
