import os
import re
import gdown
from mutagen.mp3 import MP3
from mutagen.mp4 import MP4
import pandas as pd
import requests
import urllib.parse
import shutil
import glob


class FileProcessor:
    def __init__(self, download_dir):
        self.download_dir = download_dir
        self.file_info = {}  # ダウンロードした音声ファイルの情報を保持する辞書
        self.sheet_row_map = {}  # シート名と行番号のマッピング

    def extract_file_id(self, url):
        """Google Driveの共有URLからファイルIDを抽出する"""
        pattern = r"https://drive\.google\.com/file/d/([a-zA-Z0-9_-]+)/view"
        match = re.search(pattern, url)
        if match:
            return match.group(1)
        return None

    def get_gdrive_filename(self, file_id):
        """Googleドライブからファイル名を取得する"""
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        session = requests.Session()
        response = session.get(url, stream=True)
        cd = response.headers.get('Content-Disposition', '')
        print(f"Content-Disposition: {cd}")  # デバッグ出力

        # UTF-8エンコードされたfilename*を探す
        filename_star = re.search(r'filename\*=UTF-8\'\'([^;]+)', cd)
        if filename_star:
            try:
                # URLデコードして返す
                decoded = urllib.parse.unquote(filename_star.group(1))
                print(f"Decoded filename*: {decoded}")
                return decoded
            except Exception as e:
                print(f"Decoding error: {e}")

        # 通常のfilename
        filename = re.search(r'filename="([^"]+)"', cd)
        if filename:
            try:
                decoded = filename.group(1)
                print(f"Decoded filename: {decoded}")
                return decoded
            except Exception as e:
                print(f"Decoding error: {e}")

        # 両方見つからない場合はfile_idを返す
        return file_id

    def download_with_requests(self, file_id, output_path):
        """requestsを使用してGoogle Driveからファイルをダウンロードする"""
        url = f"https://drive.google.com/uc?export=download&id={file_id}"
        session = requests.Session()

        # 初回リクエスト - 大きいファイルの場合は確認ページが表示される
        response = session.get(url, stream=True)

        # 確認トークンを取得（大きいファイルの場合）
        token = None
        for key, value in response.cookies.items():
            if key.startswith('download_warning'):
                token = value
                break

        if token:
            # 確認トークンがある場合は、それを使って再リクエスト
            url = f"{url}&confirm={token}"
            response = session.get(url, stream=True)

        # ファイルに保存
        with open(output_path, 'wb') as f:
            shutil.copyfileobj(response.raw, f)

        return output_path

    def download_file(self, url, sheet_name, row_num):
        """Googleドライブからファイルをダウンロードしてローカルストレージに保存"""
        file_id = self.extract_file_id(url)
        if not file_id:
            return None, f"無効なURL: {url}"

        try:
            os.makedirs(self.download_dir, exist_ok=True)

            # ダウンロード前のm4a/mp3ファイル一覧を取得
            before_files = set(glob.glob("*.m4a") + glob.glob("*.mp3"))

            # gdownでダウンロード（output指定なし）
            output_file = gdown.download(
                # f"https://drive.google.com/uc?id={file_id}",
                url,
                quiet=False,
                fuzzy=True
            )
            print(f"gdownダウンロード結果: {output_file}")

            # ダウンロード後のm4a/mp3ファイル一覧を取得
            after_files = set(glob.glob("*.m4a") + glob.glob("*.mp3"))
            new_files = after_files - before_files

            if new_files:
                downloaded_file = list(new_files)[0]
                # 保存先ディレクトリに移動
                dest_path = os.path.join(self.download_dir, os.path.basename(downloaded_file))
                shutil.move(downloaded_file, dest_path)
                file_path = dest_path
                file_name = os.path.basename(file_path)
                print(f"保存されたファイル名: {file_name}")

                # シート名と行番号の情報を保存
                if file_path not in self.sheet_row_map:
                    self.sheet_row_map[file_path] = []
                self.sheet_row_map[file_path].append((sheet_name, row_num))

                # m4aまたはmp3ファイルの場合は分析を実行
                if os.path.splitext(file_name)[1].lower() in ('.m4a', '.mp3'):
                    self.analyze_audio_file(file_path, sheet_name, row_num)

                return file_path, None
            else:
                return None, "ダウンロード失敗"

        except Exception as e:
            print(f"ダウンロードエラー: {str(e)}")
            return None, f"エラー: {str(e)}"

    def analyze_audio_file(self, file_path, sheet_name, row_num):
        """音声ファイルを分析し、再生時間などの情報を取得"""
        file_name = os.path.basename(file_path)
        file_ext = os.path.splitext(file_name)[1].lower()

        try:
            # ファイル形式に応じて適切なライブラリで分析
            if file_ext == '.mp3':
                audio = MP3(file_path)
                duration = audio.info.length  # 秒単位
            elif file_ext == '.m4a':
                audio = MP4(file_path)
                duration = audio.info.length  # 秒単位
            else:
                return

            # 秒数から分:秒形式の文字列に変換
            mins = int(duration // 60)
            secs = int(duration % 60)
            duration_str = f"{mins:02d}:{secs:02d}"

            # ファイル情報を辞書に格納
            key = (sheet_name, row_num)
            self.file_info[key] = {
                'file_name': file_name,
                'file_path': file_path,
                'duration': duration_str,
                'sheet_name': sheet_name,
                'row_num': row_num
            }

        except Exception as e:
            print(f"音声ファイル分析エラー: {str(e)}")

    def get_file_info_dataframe(self):
        """ファイル情報を保持するDataFrameを作成"""
        data = []
        for key, info in self.file_info.items():
            data.append({
                'シート名': info['sheet_name'],
                '行番号': info['row_num'],
                'ファイル名': info['file_name'],
                '再生時間': info['duration'],
                'ファイルパス': info['file_path']
            })

        if data:
            return pd.DataFrame(data)
        else:
            return pd.DataFrame(columns=['シート名', '行番号', 'ファイル名', '再生時間', 'ファイルパス'])
