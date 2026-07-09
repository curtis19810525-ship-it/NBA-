@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo Daily Data: Convert old format to new
echo ========================================
echo.
echo Please close Excel and backup the file first.
echo.

python "convert_daily_data_format.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ========================================
    echo Failed!
    echo ========================================
    pause
    exit /b 1
)

echo.
pause
