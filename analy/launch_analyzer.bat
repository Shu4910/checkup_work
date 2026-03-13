@echo off
chcp 65001 > nul
cd /d "%~dp0"

:: Python の確認
python --version > nul 2>&1
if errorlevel 1 (
    echo [エラー] Python がインストールされていません。
    echo.
    echo 以下のサイトから Python をインストールしてください:
    echo   https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

:: analyze_activity.py を実行
python analyze_activity.py

if errorlevel 1 (
    echo.
    echo [エラー] 実行中にエラーが発生しました。
    echo 上記のエラーメッセージを確認してください。
)

pause
