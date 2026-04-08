@echo off
chcp 65001 > nul
echo ================================================
echo  勤務表自動入力 セットアップ
echo ================================================
echo.

REM Python バージョン確認
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo [エラー] Python が見つかりません。
    echo   https://www.python.org/ からインストールしてください。
    pause
    exit /b 1
)

echo [1/3] Python ライブラリをインストール中...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [エラー] ライブラリのインストールに失敗しました。
    pause
    exit /b 1
)

echo.
echo [2/3] 設定ファイルを作成中...
if not exist config.ini (
    copy config.ini.example config.ini > nul
    echo   config.ini を作成しました。
) else (
    echo   config.ini は既に存在します（スキップ）。
)

echo.
echo [3/3] config.ini をメモ帳で開きます。sender_email を設定してください...
echo.
notepad config.ini

echo ================================================
echo  セットアップ完了！
echo ================================================
echo.
echo 次回からは「実行.bat」をダブルクリックするだけです。
echo.
pause
