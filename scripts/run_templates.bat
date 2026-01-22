@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "OUTPUT_DIR=%SCRIPT_DIR%templates"

if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

skillrunner run budget_matcher --generate-templates "%OUTPUT_DIR%"

echo Vorlagen wurden in %OUTPUT_DIR% erstellt.
endlocal
