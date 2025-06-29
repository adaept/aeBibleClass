@echo off
setlocal

:: === Configuration ===
set "SOURCE=C:\adaept"
set "DESTINATION=%UserProfile%\OneDrive\Documents"

:: Create destination folder if it doesn't exist
if not exist "%DESTINATION%" (
    mkdir "%DESTINATION%"
)

:: Copy new files only (no overwrite)
xcopy "%SOURCE%\*" "%DESTINATION%\" /E /I /Y /D /C

echo Sync complete. Only new or updated files were copied.
pause
