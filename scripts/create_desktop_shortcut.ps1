$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$launcherPath = Join-Path $projectRoot 'run_app.bat'
$iconPath = Join-Path $projectRoot 'Assets\KIKAI_PM.ico'
$desktopPath = [Environment]::GetFolderPath('DesktopDirectory')
$shortcutPath = Join-Path $desktopPath 'KIKAI Solution PM.lnk'

if (-not (Test-Path -LiteralPath $launcherPath -PathType Leaf)) {
    throw "Khong tim thay file khoi dong: $launcherPath"
}
if (-not (Test-Path -LiteralPath $iconPath -PathType Leaf)) {
    throw "Khong tim thay icon: $iconPath"
}
if (-not (Test-Path -LiteralPath $desktopPath -PathType Container)) {
    throw "Khong tim thay Desktop: $desktopPath"
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $launcherPath
$shortcut.WorkingDirectory = $projectRoot
$shortcut.IconLocation = "$iconPath,0"
$shortcut.Description = 'Mo KIKAI Solution PM'
$shortcut.Save()

Write-Output "Da tao loi tat: $shortcutPath"
