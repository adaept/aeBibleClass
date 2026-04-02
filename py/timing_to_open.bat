@echo off
setlocal

echo.
echo ============================================================
echo   Word Document Open Time Test
echo ============================================================
echo.

:: Record start time
for /f "tokens=1-4 delims=:.," %%a in ("%time%") do (
    set /a START_H=%%a
    set /a START_M=%%b
    set /a START_S=%%c
    set /a START_CS=%%d
)
set /a START_TOTAL_CS=(START_H*360000)+(START_M*6000)+(START_S*100)+START_CS
echo   Start time : %time%

:: Open the document and wait for Word to finish loading
start /wait "" "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" /t "C:\adaept\aeBibleClass\MyOutfile.docm"

:: Record end time
for /f "tokens=1-4 delims=:.," %%a in ("%time%") do (
    set /a END_H=%%a
    set /a END_M=%%b
    set /a END_S=%%c
    set /a END_CS=%%d
)
set /a END_TOTAL_CS=(END_H*360000)+(END_M*6000)+(END_S*100)+END_CS
echo   End time   : %time%

:: Calculate elapsed
set /a ELAPSED_CS=END_TOTAL_CS-START_TOTAL_CS
set /a ELAPSED_S=ELAPSED_CS/100
set /a ELAPSED_MS=(ELAPSED_CS%%100)*10

echo.
echo   Elapsed    : %ELAPSED_S%.%ELAPSED_MS% seconds
echo.
echo ============================================================

pause
endlocal