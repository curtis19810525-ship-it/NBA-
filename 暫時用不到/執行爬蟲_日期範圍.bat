@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========================================
echo 執行 Selenium 爬蟲程式（日期範圍模式）
echo ========================================
echo.
echo 請輸入起始日期（格式：YYYYMMDD，例如：20251022）：
set /p START_DATE=
echo.
echo 請輸入結束日期（格式：YYYYMMDD，例如：20251215）：
set /p END_DATE=
echo.
echo.
echo 確認資訊：
echo   Start Date: %START_DATE%
echo   End Date: %END_DATE%
echo.
echo 開始執行爬蟲...
echo ========================================
echo.

REM 使用 Python 來處理日期範圍和執行爬蟲
python "執行範圍爬蟲.py" %START_DATE% %END_DATE%

echo.
echo ========================================
echo 所有日期執行完成！
echo ========================================
pause





































