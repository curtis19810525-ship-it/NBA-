@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo 只匯出 NotebookLM（不爬蟲、不重算）
echo ============================================================
echo.
echo 請先關閉 Excel：MLB26-27 / 盤口觀察
echo 建議先暫停 OneDrive 同步
echo.
echo 用法例：賽果日=20260917、關注日=20260918
echo 會產出：
echo   - 20260917賽後結果.txt
echo   - 20260918(1~8)賽前分析.txt
echo   - 20260918(9~N)賽前分析.txt
echo.
if "%~2"=="" (
  python integrated_export.py
) else (
  python integrated_export.py %~1 %~2
)
echo.
pause
