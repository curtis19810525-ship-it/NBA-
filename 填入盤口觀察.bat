@echo off
chcp 65001 >nul
cd /d "%~dp0"
python fill_observation.py
echo.
pause
