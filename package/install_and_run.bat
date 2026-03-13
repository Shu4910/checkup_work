@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ============================================================
echo   Activity Analyzer
echo ============================================================
echo.

:: ── Python の確認 ──────────────────────────────────────────────
python --version > nul 2>&1
if errorlevel 1 (
    echo [エラー] Python がインストールされていません。
    echo.
    echo 以下のサイトから Python をインストールしてください:
    echo   https://www.python.org/downloads/
    echo.
    echo インストール時に "Add Python to PATH" にチェックを入れてください。
    echo.
    pause
    exit /b 1
)

echo Python バージョン:
python --version
echo.

:: ── pip の確認 ─────────────────────────────────────────────────
python -m pip --version > nul 2>&1
if errorlevel 1 (
    echo [エラー] pip が見つかりません。
    echo.
    echo Python を再インストールするか、以下を実行してください:
    echo   python -m ensurepip --upgrade
    echo.
    pause
    exit /b 1
)

:: ── 依存パッケージのインストール ───────────────────────────────
echo 必要なパッケージを確認・インストール中...
python -m pip install --upgrade pandas pillow 2>&1
if errorlevel 1 (
    echo.
    echo [エラー] pandas のインストールに失敗しました。
    echo.
    echo 考えられる原因:
    echo   - インターネットに接続されていない
    echo   - 管理者権限が必要な場合がある
    echo.
    echo 管理者権限でコマンドプロンプトを開き、以下を実行してください:
    echo   pip install pandas
    echo.
    pause
    exit /b 1
)

:: ── デスクトップショートカット作成（初回のみ） ─────────────────
if not exist "neinei.ico" (
    echo デスクトップショートカットを作成中...
    python setup_shortcut.py
    if errorlevel 1 (
        echo [警告] ショートカット作成に失敗しましたが、ツールは起動できます。
    )
    echo.
)

echo インストール完了。ツールを起動します...
echo.

:: ── ツール実行 ─────────────────────────────────────────────────
python -m activity_analyzer
if errorlevel 1 (
    echo.
    echo [エラー] 実行中にエラーが発生しました。
    echo 上記のエラーメッセージを確認してください。
    echo.
    pause
    exit /b 1
)
