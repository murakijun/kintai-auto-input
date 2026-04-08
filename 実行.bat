@echo off
chcp 65001 > nul

REM 初回セットアップが済んでいない場合は自動実行
if not exist config.ini (
    echo [情報] 初回セットアップを開始します...
    call setup.bat
    if not exist config.ini exit /b 1
)

echo ================================================
echo  勤務表自動入力を起動します
echo ================================================
echo.

python 勤務表自動入力.py

echo.
pause
