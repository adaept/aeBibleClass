@echo off
setlocal ENABLEDELAYEDEXPANSION

REM ============================================================
REM  DEFINE YOUR FOLDER OPTIONS HERE
REM ============================================================
REM Option 0 = default (current setup)
set "FOLDER_0=aeBibleClass"

REM Options 1–9 = additional folders under C:\adaept
set "FOLDER_1=aedh"
set "FOLDER_2=aeRWB"
set "FOLDER_3=aeChat"
set "FOLDER_4=Project4"
set "FOLDER_5=Project5"
set "FOLDER_6=Project6"
set "FOLDER_7=Project7"
set "FOLDER_8=Project8"
set "FOLDER_9=Project9"

REM ============================================================
REM  DISPLAY MENU
REM ============================================================
echo.
echo Select a folder to sync:
echo   0. %FOLDER_0%   (default)
echo   1. %FOLDER_1%
echo   2. %FOLDER_2%
echo   3. %FOLDER_3%
echo   4. %FOLDER_4%
echo   5. %FOLDER_5%
echo   6. %FOLDER_6%
echo   7. %FOLDER_7%
echo   8. %FOLDER_8%
echo   9. %FOLDER_9%
echo.

set /p "CHOICE=Enter option (0-9, default=0): "

REM ============================================================
REM  DEFAULT IF ENTER PRESSED
REM ============================================================
if "%CHOICE%"=="" set "CHOICE=0"

REM ============================================================
REM  VALIDATE INPUT
REM ============================================================
echo %CHOICE%| findstr /r "^[0-9]$" >nul
if errorlevel 1 (
    echo Invalid selection: %CHOICE%
    echo Must be a single digit 0-9
    pause
    exit /b 1
)

REM ============================================================
REM  RESOLVE SELECTED FOLDER
REM ============================================================
set "SELECTED=!FOLDER_%CHOICE%!"

echo.
echo Selected option %CHOICE%: %SELECTED%
echo.

REM ============================================================
REM  BUILD SOURCE AND DESTINATION PATHS
REM ============================================================
set "SOURCE=C:\adaept\%SELECTED%"
set "DESTINATION=%UserProfile%\OneDrive\Backups\adaept\%SELECTED%"

echo SOURCE:      %SOURCE%
echo DESTINATION: %DESTINATION%
echo.

REM ============================================================
REM  CREATE DESTINATION IF NEEDED
REM ============================================================
if not exist "%DESTINATION%" (
    echo Creating destination folder...
    md "%DESTINATION%"
)

REM Hydrate OneDrive files to avoid rsync misreads
attrib -P "%DESTINATION%\*" /S /D

REM Convert Windows paths to WSL format
for /f "usebackq delims=" %%A in (`wsl wslpath -u "%SOURCE%"`) do set "WSL_SRC=%%A"
for /f "usebackq delims=" %%B in (`wsl wslpath -u "%DESTINATION%"`) do set "WSL_DST=%%B"

echo Syncing from WSL source: %WSL_SRC%
echo Syncing to WSL dest:     %WSL_DST%
echo.

REM ============================================================
REM  EXISTING RSYNC COMMAND HERE
REM ============================================================
REM Add this to the command line for a dry run
REM --dry-run

wsl -- bash -lc ^
"rsync -a --update ^
  --checksum ^
  --itemize-changes ^
  --info=stats1 ^
  --exclude='.git/' ^
  --exclude='node_modules/' ^
  --exclude='venv/' ^
  --exclude='~[$]*' ^
  --exclude='~*.tmp' ^
  --exclude='*.wbk' ^
  \"%WSL_SRC%/\" \"%WSL_DST%/\""

echo Sync complete.
pause

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

REM xcopy "%SOURCE%\*" "%DESTINATION%\" /E /I /Y /D /C

REM robocopy "%SOURCE%" "%DESTINATION%" /E /M /FFT /NP /NDL /NJH /NJS /NC /NS

REM echo Sync complete.
REM pause
