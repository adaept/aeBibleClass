@echo off
echo.
echo ========================================
echo  VBA Casing Normalizer
echo ========================================

wsl python3 /mnt/c/adaept/aeBibleClass/normalize_vba.py /mnt/c/adaept/aeBibleClass/src

if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Normalization failed.
    pause
    exit /b 1
)

echo.
echo Normalization complete. Safe to commit.
echo.
pause