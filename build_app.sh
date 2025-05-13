#!/bin/bash
# m4a_reportアプリケーションをビルドするスクリプト

# 環境設定
export PATH=$PATH:$HOME/Library/Python/3.10/bin

# クリーンアップ
echo "以前のビルドファイルを削除しています..."
rm -rf build/ dist/

# PyInstallerでビルド
echo "アプリケーションをビルドしています..."
pyinstaller --name m4a_report \
  --windowed \
  --onedir \
  --add-data "【※サイト構築用※】音声テンプレ.xlsx:." \
  --add-data "m4a:m4a" \
  --hidden-import PyQt5.QtCore \
  --hidden-import PyQt5.QtWidgets \
  --hidden-import PyQt5.QtGui \
  --hidden-import openpyxl \
  --hidden-import polars \
  --hidden-import gdown \
  --hidden-import mutagen \
  main.py

# ビルド完了メッセージ
echo "ビルドが完了しました！"
echo "アプリケーションは dist/m4a_report.app にあります"
