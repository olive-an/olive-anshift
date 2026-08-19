@echo off
rem =====================================================================
rem  Merge 6 shift workbooks into one (see README.md)
rem  Double-click this file to run.
rem =====================================================================
setlocal
chcp 932 >nul

set "PS1=%~dp0Merge-ShiftWorkbooks.ps1"

if not exist "%PS1%" (
    echo Merge-ShiftWorkbooks.ps1 not found next to this file.
    echo Path: "%PS1%"
    goto :done
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
if errorlevel 1 (
    echo.
    echo *** FAILED *** See the message above.
)

:done
echo.
pause
endlocal
