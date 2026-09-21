@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo 診斷玖九盤口來源（只讀 Excel，不爬網）
echo ============================================================
echo.
echo 本工具檢查「玖九盤口變化.xlsx」的比賽結果／MLB快照是否齊全。
echo 盤口爬蟲不在本資料夾；sibling 為「抓取玖九盤口動態資訊」或手機 VPS。
echo.
if "%~1"=="" (
  python diagnose_jiujiu_odds.py
) else (
  python diagnose_jiujiu_odds.py %*
)
echo.
pause
