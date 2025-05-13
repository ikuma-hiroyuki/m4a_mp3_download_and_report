# macOSからWindows用exeファイルを作成する方法

macOSからWindows用のexeファイルを作成するには、以下の方法があります：

## 1. 仮想マシンを使用する方法（推奨）

### 必要なもの
- VirtualBox, VMware, Parallelsなどの仮想化ソフトウェア
- Windows 10/11のインストールイメージ

### 手順
1. macOS上に仮想マシンをインストールします
2. 仮想マシン上にWindowsをインストールします
3. Windows VM上にPythonと必要なライブラリをインストールします
4. プロジェクトファイルを仮想マシンに転送します
5. 付属の`build_app_windows.bat`スクリプトを実行してexeファイルを生成します
6. 生成されたexeファイルをmacOSに転送します

## 2. Docker + Wine を使用する方法

### 必要なもの
- Dockerのインストール
- Wine（Windowsエミュレータ）を含むDockerイメージ

### 手順
1. 以下のコマンドでDockerイメージをプルします：
```bash
docker pull cdrx/pyinstaller-windows
```

2. 以下のコマンドでexeファイルをビルドします：
```bash
docker run -v "$(pwd):/src/" cdrx/pyinstaller-windows
```

3. dist/windowsフォルダに生成されたexeファイルを確認します

## 3. クラウドサービスを利用する方法

### 使用可能なサービス
- GitHub Actions
- Azure Pipelines
- CircleCI

### GitHub Actionsの例
1. `.github/workflows/build-windows-exe.yml`ファイルを作成します
2. 以下のようなワークフローを定義します：
```yaml
name: Build Windows EXE

on:
  push:
    branches: [ main ]
  pull_request:
    branches: [ main ]

jobs:
  build:
    runs-on: windows-latest

    steps:
    - uses: actions/checkout@v3
    - name: Set up Python
      uses: actions/setup-python@v4
      with:
        python-version: '3.10'
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install pyinstaller
        pip install -r requirements.txt
    - name: Build with PyInstaller
      run: |
        pyinstaller --name m4a_report --windowed --onefile --add-data "【※サイト構築用※】音声テンプレ.xlsx;." --add-data "m4a;m4a" main.py
    - name: Upload EXE
      uses: actions/upload-artifact@v3
      with:
        name: m4a_report
        path: dist/m4a_report.exe
```

3. GitHubにプッシュしてワークフローを実行します
4. 生成されたexeファイルをダウンロードします

## 注意点

- Windows用のexeファイルは、必ずWindowsでテストしてください
- データパスの区切り文字がWindows（バックスラッシュ）とmacOS（スラッシュ）で異なるため、パス操作には`os.path`モジュールを使用してください
- PyQt5のパスがmacOSとWindowsで異なる場合があります
- アンチウイルスソフトウェアがexeファイルをブロックする場合があります
