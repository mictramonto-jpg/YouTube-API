@echo off
chcp 65001 >nul 2>&1
echo ========================================
echo   YouTube コメント抽出ツール
echo ========================================
echo.
echo 依存パッケージをインストール中...
py -m pip install -r requirements.txt --quiet 2>nul || python -m pip install -r requirements.txt --quiet 2>nul || pip install -r requirements.txt --quiet
echo.
echo アプリを起動中...
py main.py 2>nul || python main.py 2>nul || python3 main.py
echo.
if errorlevel 1 (
    echo エラーが発生しました。Pythonがインストールされているか確認してください。
    echo https://www.python.org/downloads/
)
pause
