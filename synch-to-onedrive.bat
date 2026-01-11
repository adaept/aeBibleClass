@echo off
setlocal

set "SOURCE=C:\adaept"
set "DESTINATION=%UserProfile%\OneDrive\Documents"

if not exist "%DESTINATION%" (
    md "%DESTINATION%"
)

REM Convert Windows paths to WSL format
for /f "usebackq delims=" %%A in (`wsl wslpath -u "%SOURCE%"`) do set "WSL_SRC=%%A"
for /f "usebackq delims=" %%B in (`wsl wslpath -u "%DESTINATION%"`) do set "WSL_DST=%%B"

echo Syncing from WSL source: %WSL_SRC%
echo Syncing to WSL dest:   %WSL_DST%

REM Now call rsync with properly escaped exclude patterns
wsl -- bash -lc "rsync -a --update --info=name1,progress2 --exclude=\"*/venv/**\" \"%WSL_SRC%/\" \"%WSL_DST%/\""

echo Sync complete.
pause



REM xcopy "%SOURCE%\*" "%DESTINATION%\" /E /I /Y /D /C

REM robocopy "%SOURCE%" "%DESTINATION%" /E /M /FFT /NP /NDL /NJH /NJS /NC /NS

REM echo Sync complete.
REM pause
