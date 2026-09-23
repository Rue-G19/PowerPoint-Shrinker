@echo off
setlocal
title PPTSLauncher v1.2 Rueben Gill
set "SCRIPT=%~dp0RG_PPTShrinker-Public.ps1"
if not exist "%SCRIPT%" (
    echo PowerPoint Bulk Shrinker - Error
    echo The PowerShell script could not be found:
    echo "%SCRIPT%"
    echo Make sure this bat is in the same folder
    echo as the ps1.
    echo.
    pause
    exit /b 1)
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT%"
set "EXITCODE=%ERRORLEVEL%"
echo.
echo PowerPoint Bulk Shrinker has finished.
echo Pretty neat, eh? :)
if not "%EXITCODE%"=="0" echo PowerShell exited with code %EXITCODE%.
echo.
pause
endlocal
