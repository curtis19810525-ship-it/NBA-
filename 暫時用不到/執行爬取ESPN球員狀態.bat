@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo 爬取 ESPN 球員狀態與傷兵
echo ============================================================
echo.
echo 將寫入：球員狀態、球員傷兵 分頁（清空後重寫）
echo.

python "%~dp0espn_player_stats.py"

echo.
pause
