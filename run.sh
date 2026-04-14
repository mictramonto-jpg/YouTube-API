#!/bin/bash
echo "========================================"
echo "  YouTube コメント抽出ツール"
echo "========================================"
echo ""
echo "依存パッケージをインストール中..."
pip3 install -r requirements.txt --quiet 2>/dev/null || pip install -r requirements.txt --quiet
echo ""
echo "アプリを起動中..."
python3 main.py 2>/dev/null || python main.py
