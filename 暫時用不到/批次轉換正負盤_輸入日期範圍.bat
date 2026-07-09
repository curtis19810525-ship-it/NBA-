@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo 批次轉換正負盤
echo ============================================================
echo.

set /p DATE_RANGE="請輸入日期範圍（格式：YYYYMMDD~YYYYMMDD，例如：20251022~20260214）："

if "%DATE_RANGE%"=="" (
    echo 錯誤：未輸入日期範圍
    pause
    exit /b 1
)

echo.
echo 開始轉換...
echo.

python "批次轉換正負盤.py" "%DATE_RANGE%"

echo.
pause
