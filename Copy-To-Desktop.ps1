$ErrorActionPreference = 'Stop'
$desktopPath = [Environment]::GetFolderPath('Desktop')
$target = Join-Path $desktopPath 'Hamrah-Mohasel'
if ([IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\') -eq [IO.Path]::GetFullPath($target).TrimEnd('\')) {
    Write-Host 'The project is already on your Desktop.'
    exit 0
}
if (Test-Path $target) { $target = Join-Path $desktopPath ('Hamrah-Mohasel-' + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
New-Item -ItemType Directory -Path $target | Out-Null
Get-ChildItem -LiteralPath $PSScriptRoot -Force | Where-Object { $_.Name -notin @('node_modules', '.env', 'dist') } | Copy-Item -Destination $target -Recurse
Write-Host "Project copied to: $target"
Start-Process explorer.exe -ArgumentList $target
