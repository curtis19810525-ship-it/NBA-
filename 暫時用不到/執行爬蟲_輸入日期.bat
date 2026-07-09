@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在執行 Selenium 爬蟲程式...
echo.
echo 請輸入日期（格式：YYYYMMDD）：
set /p GAME_DATE=
echo.
python "使用Selenium完整版.py" %GAME_DATE%
echo.
echo 執行完成！
pause




















































