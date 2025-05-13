# M4A Report

Googleドライブ共有リンクからファイルをダウンロードして分析するアプリケーションです。特にExcelシート内のリンクからm4a/mp3ファイルをダウンロードし、その再生時間を分析して結果を新しいシートに出力します。

## 機能

- ExcelファイルのJ列に含まれるGoogleドライブ共有リンクからファイルをダウンロード
- ダウンロードしたm4aまたはmp3ファイルの再生時間を分析
- 分析結果をExcelファイルの新しいシートに出力
- ダウンロードしたファイル名でセルにハイパーリンクを設定

## 必要条件

- Python 3.8以上
- 必要なパッケージ：
  - PyQt5
  - openpyxl
  - pandas
  - polars
  - gdown
  - mutagen

## インストール方法

```bash
# 仮想環境を作成
python -m venv .venv

# 仮想環境を有効化（Windows）
.venv\Scripts\activate

# 仮想環境を有効化（macOS/Linux）
source .venv/bin/activate

# 必要なパッケージをインストール
pip install -e .
```

## 使い方

1. アプリケーションを起動します。

```bash
python main.py
```

2. 「Excelファイル選択」ボタンをクリックして処理対象のExcelファイルを選択します。
3. 「ダウンロードしたファイルの保存先」ボタンをクリックして保存先フォルダを選択します。
4. 処理対象のシートをチェックボックスで選択します。
5. 「実行」ボタンをクリックして処理を開始します。
6. 処理状況が画面下部に表示されます。
7. 処理が完了するとExcelファイルが自動的に開かれ、結果が確認できます。

## 注意事項

- Excelファイルが開かれている状態では処理を実行できません。
- 処理対象のExcelファイルのJ列には、GoogleドライブのファイルへのURLが含まれている必要があります。
- URLは以下の形式である必要があります：`https://drive.google.com/file/d/[FILE_ID]/view?usp=drive_link`
