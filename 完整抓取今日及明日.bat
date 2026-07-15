@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo 完整抓取今日及明日
echo ============================================================
python run_all.py --preset full_range
echo.
pause
