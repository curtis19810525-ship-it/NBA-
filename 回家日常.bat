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
echo   4. 輸入兩日：賽果日（歸檔）＋ 關注日（總表 A1）
echo      例 8/11 晚間：賽果日=20260811、關注日=20260812
echo   5. 結束後 exports 資料夾會有兩份整合檔
echo      - 賽果日_賽後結果.txt
echo      - 關注日_賽前分析.txt
echo      剪貼簿為賽前分析全文，供貼 NotebookLM
echo.
python home_pipeline.py
echo.
pause
