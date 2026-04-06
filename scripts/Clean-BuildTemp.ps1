$paths = @(
    '.sdkshim',
    'restore-diag.log',
    'web-restore-diag.log',
    'web-restore-diag2.log',
    'msbuild-restore.log',
    'web-restore.binlog',
    'core-restore-diag.log'
)

foreach ($path in $paths) {
    if (Test-Path $path) {
        Remove-Item -Recurse -Force $path -ErrorAction SilentlyContinue
    }
}

Write-Output 'CLEANED_TEMP'
