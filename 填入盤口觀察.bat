@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================================
echo 填入盤口觀察
echo ============================================================
echo.
echo 【重要】請先在 Excel「存檔」並關閉：
echo   - MLB26-27數據.xlsx
echo   - 玖九盤口變化.xlsx
echo   - 盤口觀察.xlsx
echo 本程式讀的是磁碟已存檔內容，未存檔會讀到舊資料。
echo 場次來源：玖九「比賽結果」；金流／量化有「紀錄」才補。
echo.
python fill_observation.py
echo.
pause
