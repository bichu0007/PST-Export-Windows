$ErrorActionPreference = 'Stop'
$launcher = Join-Path $PSScriptRoot 'Open PST Export.bat'
if (-not (Test-Path -LiteralPath $launcher)) {
    throw "Launcher not found: $launcher"
}
$desktop = [Environment]::GetFolderPath('Desktop')
$linkPath = Join-Path $desktop 'PST Export Toolkit.lnk'
$shell = New-Object -ComObject WScript.Shell
$link = $shell.CreateShortcut($linkPath)
$link.TargetPath = $launcher
$link.WorkingDirectory = $PSScriptRoot
$link.Description = 'Open PST Export Toolkit'
$link.Save()
Write-Host "Desktop shortcut created: $linkPath" -ForegroundColor Green
Write-Host 'Double-click the PST Export Toolkit shortcut to open the app.'
