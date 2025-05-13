# m4a_reportアプリケーションのWindows用exe化手順

## 前提条件
- Windows 10以上のOS
- Python 3.8以上のインストール
- 必要なライブラリ（PyQt5, openpyxl, polars, gdown, mutagen）のインストール

## 手順

### 1. ソースコードの準備
- プロジェクトフォルダをZIPファイルとしてWindowsマシンに転送します
- Windows上でZIPファイルを展開します

### 2. 必要なパッケージのインストール
コマンドプロンプトを開き、以下のコマンドを実行します：
```
pip install pyinstaller
pip install PyQt5
pip install openpyxl
pip install polars
pip install gdown
pip install mutagen
```

### 3. ビルドスクリプトの実行
- プロジェクトフォルダに含まれる `build_app_windows.bat` を右クリックし、「管理者として実行」を選択します
- または、コマンドプロンプトで以下のコマンドを実行します：
```
cd プロジェクトフォルダのパス
build_app_windows.bat
```

### 4. exeファイルの確認
- ビルド完了後、`dist` フォルダ内に `m4a_report.exe` が生成されます
- このexeファイルをダブルクリックするだけでアプリケーションが起動します

### 5. 配布方法
- `dist` フォルダ内の `m4a_report.exe` ファイルをZIPファイルに圧縮して配布します
- 受け取った側は、ZIPファイルを展開してexeファイルをダブルクリックすることで利用可能です

## トラブルシューティング

### exeファイルが起動しない場合
- Windowsのセキュリティ設定により実行がブロックされている可能性があります
- exeファイルを右クリック→「プロパティ」→「ブロックの解除」にチェックを入れ、「OK」をクリックします

### 依存関係のエラーが発生する場合
- Microsoft Visual C++ Redistributable がインストールされていない可能性があります
- 最新のVisual C++ Redistributableをマイクロソフトの公式サイトからダウンロードしてインストールしてください

### その他のエラー
- Windows上でPythonスクリプトを直接実行して、エラーメッセージを確認してください
- 必要なライブラリが適切にインストールされているか確認してください
