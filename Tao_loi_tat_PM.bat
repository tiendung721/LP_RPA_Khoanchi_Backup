@echo off
setlocal

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\create_desktop_shortcut.ps1"
set "SHORTCUT_EXIT=%ERRORLEVEL%"

echo.
if not "%SHORTCUT_EXIT%"=="0" (
    echo Khong the tao loi tat Desktop. Kiem tra thong bao loi phia tren.
) else (
    echo Da tao loi tat KIKAI Solution PM tren Desktop.
)
pause
exit /b %SHORTCUT_EXIT%
