$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$launcherPath = Join-Path $projectRoot 'launch_pm.vbs'
$wscriptPath = Join-Path $env:SystemRoot 'System32\wscript.exe'
$iconPath = Join-Path $projectRoot 'Assets\KIKAI_PM.ico'
$desktopPath = [Environment]::GetFolderPath('DesktopDirectory')
$shortcutPath = Join-Path $desktopPath 'KIKAI Solution PM.lnk'

if (-not (Test-Path -LiteralPath $launcherPath -PathType Leaf)) {
    throw "Khong tim thay file khoi dong: $launcherPath"
}
if (-not (Test-Path -LiteralPath $wscriptPath -PathType Leaf)) {
    throw "Khong tim thay Windows Script Host: $wscriptPath"
}
if (-not (Test-Path -LiteralPath $iconPath -PathType Leaf)) {
    throw "Khong tim thay icon: $iconPath"
}
if (-not (Test-Path -LiteralPath $desktopPath -PathType Container)) {
    throw "Khong tim thay Desktop: $desktopPath"
}

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $wscriptPath
$shortcut.Arguments = '"' + $launcherPath + '"'
$shortcut.WorkingDirectory = $projectRoot
$shortcut.IconLocation = "$iconPath,0"
$shortcut.Description = 'Mo KIKAI Solution PM'
$shortcut.Save()

Write-Output "Da tao loi tat: $shortcutPath"
