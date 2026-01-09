@echo off
setlocal

:: === Configuration ===
set "SOURCE=C:\adaept"
set "DESTINATION=%UserProfile%\OneDrive\Documents"

:: Create destination folder if it doesn't exist
if not exist "%DESTINATION%" (
    mkdir "%DESTINATION%"
)

:: Copy only when file content changes, and show only changed files
robocopy "%SOURCE%" "%DESTINATION%" /E /XC /XN /XO /NFL /NDL /NJH /NJS /NP /NS /NC

echo Sync complete. Only new or updated files were copied.
pause