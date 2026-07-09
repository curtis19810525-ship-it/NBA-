@echo off
chcp 65001 >nul
echo ========================================
echo NBA 明日賽事爬蟲
echo ========================================
echo.

cd /d "%~dp0"

python "nba_future_fetcher.py"

echo.
echo ========================================
echo 執行完成
echo ========================================
pause






























