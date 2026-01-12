@echo off
setlocal

set "SOURCE=C:\adaept\aeBibleClass"
set "DESTINATION=%UserProfile%\OneDrive\Documents"

if not exist "%DESTINATION%" (
    md "%DESTINATION%"
)

REM Convert Windows paths to WSL format
for /f "usebackq delims=" %%A in (`wsl wslpath -u "%SOURCE%"`) do set "WSL_SRC=%%A"
for /f "usebackq delims=" %%B in (`wsl wslpath -u "%DESTINATION%"`) do set "WSL_DST=%%B"

echo Syncing from WSL source: %WSL_SRC%
echo Syncing to WSL dest:   %WSL_DST%

REM Now call rsync with properly escaped exclude patterns using ^ CMD line-continuation character
REM --itemize-changes to show the filenames and rsync changes code:
REM Column	Char	Meaning
REM 1	>	File was sent from source → destination
REM 2	f	It is a regular file
REM 3	.	File type unchanged
REM 4	s	Size changed
REM 5	t	Timestamp changed
REM 6	.	Permissions unchanged
REM 7	.	Owner unchanged
REM 8	.	Group unchanged
REM 9	.	ACL unchanged
REM 10	.	Extended attributes unchanged

wsl -- bash -lc ^
"rsync -a --update ^
  --itemize-changes ^
  --info=stats1 ^
  --exclude='*/venv/**' ^
  --exclude='*/.git/**' ^
  \"%WSL_SRC%/\" ^
  \"%WSL_DST%/\""

REM wsl -- bash -lc "rsync -a --update --info=name1,stats1 --exclude=\"*/venv/**\" \"%WSL_SRC%/\" \"%WSL_DST%/\""

echo Sync complete.
pause



REM xcopy "%SOURCE%\*" "%DESTINATION%\" /E /I /Y /D /C

REM robocopy "%SOURCE%" "%DESTINATION%" /E /M /FFT /NP /NDL /NJH /NJS /NC /NS

REM echo Sync complete.
REM pause
