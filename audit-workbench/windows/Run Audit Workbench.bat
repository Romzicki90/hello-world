@echo off
cd /d "%~dp0"
where pyw >nul 2>nul && (
    start "" pyw -3 -m auditwb.gui
    exit /b 0
)
where pythonw >nul 2>nul && (
    start "" pythonw -m auditwb.gui
    exit /b 0
)
echo.
echo   Python was not found on this computer. Please read "START HERE.txt".
echo.
pause
