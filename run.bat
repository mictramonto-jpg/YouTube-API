@echo off
chcp 65001 >nul 2>&1
setlocal EnableDelayedExpansion

echo ========================================
echo   YouTube コメント抽出ツール
echo ========================================
echo.

REM ==== Python コマンドの検出 ====
set PYTHON_CMD=
where py >nul 2>&1
if %ERRORLEVEL%==0 (
    set PYTHON_CMD=py
    goto :python_found
)
where python >nul 2>&1
if %ERRORLEVEL%==0 (
    set PYTHON_CMD=python
    goto :python_found
)
where python3 >nul 2>&1
if %ERRORLEVEL%==0 (
    set PYTHON_CMD=python3
    goto :python_found
)

echo [エラー] Python が見つかりません。
echo.
echo 以下の手順でPythonをインストールしてください:
echo   1. https://www.python.org/downloads/ からインストーラーをダウンロード
echo   2. インストーラー起動後、必ず "Add python.exe to PATH" にチェック
echo   3. インストール完了後、PowerShell / コマンドプロンプトを再起動
echo   4. 再度 run.bat をダブルクリック
echo.
pause
exit /b 1

:python_found
echo Python コマンド: %PYTHON_CMD%
%PYTHON_CMD% --version
if %ERRORLEVEL% NEQ 0 (
    echo [エラー] Python の実行に失敗しました。
    echo Microsoft Store のエイリアスが有効になっている可能性があります。
    echo "設定" - "アプリ" - "アプリ実行エイリアス" から
    echo "アプリインストーラー python.exe / python3.exe" をOFFにしてください。
    echo.
    pause
    exit /b 1
)
echo.

REM ==== 依存パッケージのインストール ====
echo 依存パッケージをインストール中...
%PYTHON_CMD% -m pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [エラー] 依存パッケージのインストールに失敗しました。
    echo インターネット接続やプロキシ設定を確認してください。
    echo.
    pause
    exit /b 1
)
echo.

REM ==== アプリ起動 ====
echo アプリを起動中...
echo.
%PYTHON_CMD% main.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [エラー] アプリの実行中にエラーが発生しました。
    echo 上記のメッセージを確認してください。
    echo.
)

pause
