@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo 爬取 Rotowire MLB 先發名單（日期範圍）
echo ============================================================
echo.
echo 輸入日期為台灣日；程式會換算美東日曆後再向 Rotowire 請求。
echo 寫入分頁：先發名單（每次覆蓋）；「日期」欄為台灣日。
echo.

python "%~dp0mlb_lineups.py"

echo.
pause
