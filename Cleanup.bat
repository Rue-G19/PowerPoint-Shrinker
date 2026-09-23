@echo off
setlocal
setlocal EnableExtensions
del /s /q "%~dp0PPT_Shrinking_Report.csv" 2>nul
for /r "%~dp0" %%F in (Small_*.pptx) do (
	for /f "tokens=1,* delims=_" %%A in ("%%~nxF") do (
		ren "%%~fF" "%%B"))
endlocal