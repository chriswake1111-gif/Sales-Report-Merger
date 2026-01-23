@echo off
chcp 65001 >nul
echo ============================================
echo   批次轉換工具 - .xls 轉 .xlsx
echo ============================================
echo.

set /p FOLDER="請輸入要轉換的資料夾路徑 (可直接拖曳資料夾到此視窗): "

if "%FOLDER%"=="" (
    echo 未輸入路徑，程式結束。
    pause
    exit
)

echo.
echo 開始轉換...
echo.

python "%~dp0batch_convert.py" %FOLDER%

echo.
echo ============================================
echo 轉換完成！按任意鍵關閉視窗...
pause >nul
