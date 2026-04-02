@echo off
setlocal

:: ── Configuration ─────────────────────────────────────────────────────────────
set INPUT=MyInfile.docm
set OUTPUT=MyOutfile.docm
:: ──────────────────────────────────────────────────────────────────────────────

echo.
echo ============================================================
echo   Fix Times Font in styles.xml
echo ============================================================
echo.
echo   Input  : %INPUT%
echo   Output : %OUTPUT%
echo.
echo ============================================================
echo   IMPORTANT: Make sure the document is CLOSED in Word
echo              before continuing.
echo ============================================================
echo.

choice /M "Is the document closed and ready to proceed?" /C YN /D N /T 15
if errorlevel 2 (
    echo.
    echo   Cancelled. Please close the document in Word and run again.
    echo.
    pause
    exit /b 0
)

echo.
echo   Running fix_times_font.py via WSL ...
echo.

wsl python3 fix_times_font.py %INPUT% %OUTPUT%

if %ERRORLEVEL% EQU 0 (
    echo.
    echo   SUCCESS. Open %OUTPUT% in Word to verify.
    echo.
) else (
    echo.
    echo   ERROR: Script failed. Check the output above for details.
    echo.
)

pause
endlocal
