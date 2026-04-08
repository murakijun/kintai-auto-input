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
echo [3/3] セットアップ完了！
echo.
echo 次のステップ:
echo   1. config.ini をメモ帳で開いて sender_email を設定する
echo   2. Outlook を起動する
echo   3. python 勤務表自動入力.py を実行する
echo.
pause
