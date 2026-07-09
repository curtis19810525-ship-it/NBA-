@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo 更新 MLB 分析母版
echo ============================================================
echo.
echo 請先關閉 Excel 中的相關檔案：
echo   MLB26-27數據.xlsx
echo   玖九盤口變化.xlsx
echo   MLB26-分析母版.xlsx
echo.

set /p "FOCUS_DATE=請輸入關注日 YYYYMMDD（直接 Enter = 今天）："

if "%FOCUS_DATE%"=="" (
    python "build_analysis_workbook.py"
) else (
    python "build_analysis_workbook.py" "%FOCUS_DATE%"
)

echo.
pause
