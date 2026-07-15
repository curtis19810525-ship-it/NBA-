@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo 抓取今日賽果
echo ============================================================
python run_all.py --preset today_results
echo.
pause
