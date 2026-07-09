@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在建立 NBA 25-26 賽季統計表...
echo.
python "建立NBA統計表.py"
echo.
echo 執行完成！
pause
















































