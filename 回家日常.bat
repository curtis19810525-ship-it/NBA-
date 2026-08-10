@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo 回家日常（一鍵）
echo ============================================================
echo.
echo 請先確認：
echo   1. 手機玖九資料已同步到本機
echo   2. 已關閉 Excel：MLB26-27 / 玖九盤口變化 / 盤口觀察
echo   3. 本機有安裝桌面版 Excel（步驟會用 COM 重算）
echo   4. 結束後會匯出紀錄 txt，並把「今天」複製到剪貼簿供貼 NotebookLM
echo.
python home_pipeline.py
echo.
pause
