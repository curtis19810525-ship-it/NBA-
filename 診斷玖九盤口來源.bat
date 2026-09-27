@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo 診斷玖九賽果／盤口來源（只讀 Excel，不爬網）
echo ============================================================
echo.
echo 檢查「玖九盤口變化.xlsx」的比賽結果／MLB快照／執行紀錄。
echo 例：診斷玖九盤口來源.bat 20260927
echo     診斷玖九盤口來源.bat 20260926 20260927
echo.
if "%~1"=="" (
  python diagnose_jiujiu_odds.py
) else (
  python diagnose_jiujiu_odds.py %*
)
echo.
pause
