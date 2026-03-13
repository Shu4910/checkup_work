@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ============================================================
echo   Activity Analyzer セットアップ
echo ============================================================
echo.

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

:: Pillow のインストール（アイコン生成に必要）
echo 必要なパッケージをインストール中...
python -m pip install --upgrade pillow pandas > nul 2>&1
if errorlevel 1 (
    echo [警告] パッケージのインストールに失敗しました。
    echo       アイコンなしでショートカットを作成します。
    echo.
)

:: ショートカット作成
python setup_shortcut.py
if errorlevel 1 (
    echo.
    echo [エラー] セットアップに失敗しました。
    echo 上記のエラーメッセージを確認してください。
    echo.
    pause
    exit /b 1
)

echo.
pause
