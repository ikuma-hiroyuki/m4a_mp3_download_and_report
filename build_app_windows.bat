@echo off
REM m4a_reportアプリケーションをWindows用にexe化するためのスクリプト

REM 必要なパッケージのインストール
pip install pyinstaller
pip install PyQt5
pip install openpyxl
pip install polars
pip install gdown
pip install mutagen

REM 以前のビルドファイルを削除
echo "以前のビルドファイルを削除しています..."
rmdir /s /q build
rmdir /s /q dist

REM PyInstallerでビルド
echo "アプリケーションをexe化しています..."
pyinstaller --name m4a_report ^
  --windowed ^
  --onefile ^
  --add-data "【※サイト構築用※】音声テンプレ.xlsx;." ^
  --add-data "m4a;m4a" ^
  --hidden-import PyQt5.QtCore ^
  --hidden-import PyQt5.QtWidgets ^
  --hidden-import PyQt5.QtGui ^
  --hidden-import openpyxl ^
  --hidden-import polars ^
  --hidden-import gdown ^
  --hidden-import mutagen ^
  main.py

REM ビルド完了メッセージ
echo "ビルドが完了しました！"
echo "アプリケーションは dist/m4a_report.exe にあります"
pause
