import os
import platform
import subprocess


def is_excel_file_open(file_path):
    """
    指定されたExcelファイルが開かれているかどうかを確認する
    """
    if platform.system() == "Darwin":  # Mac
        try:
            # lsofコマンドを使用して、ファイルを開いているプロセスを確認
            result = subprocess.run(['lsof', file_path], capture_output=True, text=True)
            return result.returncode == 0
        except Exception:
            return False
    elif platform.system() == "Windows":  # Windows
        try:
            # Windowsの場合、ファイルが読み取り専用モードで開けるかどうかで判断
            with open(file_path, 'r+b'):
                return False
        except PermissionError:
            return True
        except Exception:
            return False
    return False


def open_excel_file(file_path):
    """
    Excelファイルを開く
    """
    if platform.system() == "Darwin":  # Mac
        subprocess.run(['open', file_path], check=False)
    elif platform.system() == "Windows":  # Windows
        os.startfile(file_path)
