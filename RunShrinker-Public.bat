::[Bat To Exe Converter]
::
::fBE1pAF6MU+EWHreyHcjLQlHcCOgbDPpV+wgxuHo/OuJ4mASQO0tOKzUyKSBMuEH40rqSbMf6VR3v+kzXEsILEL+US4Hlk9XomuINtOVvAHyCgXJwkQzDmhxiXfDsBsLVOtbktEK3Su77nHNuJog4UzMVrsHG2jk0+FYLcsM9An6eFqZ+g==
::fBE1pAF6MU+EWHreyHcjLQlHcCOgbDPpV+wgxuHo/OuJ4mASQO0tOKzUyKSBMuEH40rqSbMf6VR3v+kzXEsILEL+US4Hlk9XomuINtOVvAHyCgXJwkQzDmhxiXfDsBsLVOtbktEK3Su77nHck6AR323vEKsPAQM=
::fBE1pAF6MU+EWHreyHcjLQlHcCOgbDPpV+wgxuHo/OuJ4mASQO0tOKzUyKSBMuEH40rqSbMf6VR3v+kzXEsILEL+US4Hlk9XomuINtOVvAHyCgXJwkQzDmhxiXfDsBsLVOtbktEK3Su77nHWkqQX1FX+WaANHgM=
::YAwzoRdxOk+EWAjk
::fBw5plQjdCyDJGyX8VAjFDx2HFzRbTKGKbsZzPry+e/H7w0zXfEseYGb97uaL/JTyUr2ZZk/125Tl8UwKBRPcB6kbwsnlWhLumG6Z5fM41+xGhjR5x4zQzRy0juD23s6MoY5n8EAi3S6qBWxlqYfsQ==
::YAwzuBVtJxjWCl3EqQJgSA==
::ZR4luwNxJguZRRnk
::Yhs/ulQjdF+5
::cxAkpRVqdFKZSDk=
::cBs/ulQjdF+5
::ZR41oxFsdFKZSDk=
::eBoioBt6dFKZSDk=
::cRo6pxp7LAbNWATEpCI=
::egkzugNsPRvcWATEpCI=
::dAsiuh18IRvcCxnZtBJQ
::cRYluBh/LU+EWAnk
::YxY4rhs+aU+IeA==
::cxY6rQJ7JhzQF1fEqQJiZkk0
::ZQ05rAF9IBncCkqN+0xwdVsFAlbi
::ZQ05rAF9IAHYFVzEqQICDyRkfDCxNHmzCL4Z+og=
::eg0/rx1wNQPfEVWB+kM9LVsJDGQ=
::fBEirQZwNQPfEVWB+kM9LVsJDGQ=
::cRolqwZ3JBvQF1fEqQIAJwxRXjSNNWWuRpMV5O27zOWKsl8YR/Ewau8=
::dhA7uBVwLU+EWDk=
::YQ03rBFzNR3SWATElA==
::dhAmsQZ3MwfNWATElA==
::ZQ0/vhVqMQ3MEVWAtB9wSA==
::Zg8zqx1/OA3MEVWAtB9wSA==
::dhA7pRFwIByZRRnk
::Zh4grVQjdCyDJGyX8VAjFDx2HFzRbTKGKLwP++n1r8eItkIPFMEwap/UyLWaKe8d1mDWSrgA8VhlyJtcXksNQTOYUS4hvWFPt3CMOMmP80KhbkeK80Y1FXFnu1PguBMIaMFhlMgGwRyc6UTzm5kg3mrrX6sCEEHoz616IYcF5Q/U
::YB416Ek+ZG8=
::
::
::978f952a14a936cc963da21a135fa983
@echo off
setlocal
if /I not "%~1"=="__MINIMIZED__" (
    start "" /min "%ComSpec%" /c call "%~f0" __MINIMIZED__
    exit /b
)
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
PowerShell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT%"
set "EXITCODE=%ERRORLEVEL%"
echo.
echo PowerPoint Bulk Shrinker has finished.
echo Pretty neat, eh? :)
if not "%EXITCODE%"=="0" echo PowerShell exited with code %EXITCODE%.
echo.
endlocal & exit /b %EXITCODE%
