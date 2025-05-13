import polars as pl
from openpyxl import load_workbook
from openpyxl.worksheet.hyperlink import Hyperlink
import utils


class ExcelProcessor:
    def __init__(self, excel_file_path):
        self.excel_file_path = excel_file_path
        self.workbook = None
        self.sheets = []

    def load_excel(self):
        """Excelファイルを読み込み、シート情報を取得する"""
        try:
            # ファイルが開かれているかチェック
            if utils.is_excel_file_open(self.excel_file_path):
                return False, "Excelファイルが開かれています。閉じてから処理を実行してください。"

            # Excelファイルを読み込む
            self.workbook = load_workbook(self.excel_file_path)
            self.sheets = self.workbook.sheetnames
            return True, None
        except Exception as e:
            return False, f"Excelファイル読み込みエラー: {str(e)}"

    def get_sheets(self):
        """利用可能なシート名のリストを返す"""
        return self.sheets

    def get_links_from_sheets(self, target_sheets):
        """指定されたシートからリンクを収集する"""
        links_data = []

        for sheet_name in target_sheets:
            if sheet_name in self.sheets:
                sheet = self.workbook[sheet_name]
                for row_idx, row in enumerate(sheet.iter_rows(min_row=2), 2):  # ヘッダー行をスキップ
                    if len(row) >= 10:  # J列（10列目）が存在する場合
                        cell = row[9]  # J列（インデックスは0から始まるので9）
                        cell_value = cell.value

                        # セルの値がURLかどうかチェック
                        if cell_value and isinstance(cell_value, str) and "drive.google.com/file" in cell_value:
                            links_data.append({
                                'sheet_name': sheet_name,
                                'row_num': row_idx,
                                'url': cell_value
                            })
                        # ハイパーリンクがある場合
                        elif cell.hyperlink and cell.hyperlink.target and "drive.google.com/file" in cell.hyperlink.target:
                            links_data.append({
                                'sheet_name': sheet_name,
                                'row_num': row_idx,
                                'url': cell.hyperlink.target
                            })

        return links_data

    def update_cell_with_hyperlink(self, sheet_name, row_num, file_name, url, duration=None):
        """セルをハイパーリンク付きテキストに更新する"""
        sheet = self.workbook[sheet_name]
        cell = sheet.cell(row=row_num, column=10)  # J列

        # ハイパーリンクを設定
        cell.value = file_name
        # Hyperlinkオブジェクトを直接作成
        cell.hyperlink = Hyperlink(target=url, ref=cell.coordinate, tooltip=f"Open {file_name}")

        # 再生時間が提供されている場合はL列に出力
        if duration:
            duration_cell = sheet.cell(row=row_num, column=12)  # L列
            duration_cell.value = duration

    def save_results_to_sheet(self, results_df):
        """結果を新しいシートに保存する"""
        if not results_df.empty:
            # 既存のresultsシートを削除（存在する場合）
            if 'results' in self.workbook.sheetnames:
                del self.workbook['results']

            # 新しいresultsシートを作成
            results_sheet = self.workbook.create_sheet('results')

            # ヘッダーを設定
            headers = results_df.columns.tolist()
            for col_idx, header in enumerate(headers, 1):
                results_sheet.cell(row=1, column=col_idx).value = header

            # データを書き込み
            for row_idx, row in enumerate(results_df.itertuples(), 2):
                for col_idx, value in enumerate(row[1:], 1):
                    results_sheet.cell(row=row_idx, column=col_idx).value = value

            # 列幅を自動調整（簡易的な実装）
            for col_idx in range(1, len(headers) + 1):
                results_sheet.column_dimensions[chr(64 + col_idx)].width = 20

    def save_excel(self):
        """Excelファイルを保存する"""
        try:
            self.workbook.save(self.excel_file_path)
            return True, None
        except Exception as e:
            return False, f"Excelファイル保存エラー: {str(e)}"

    def save_results_to_polars(self, file_info_df):
        """Polarsを使用して結果をresultsシートに保存する"""
        try:
            # pandasデータフレームをpolarsデータフレームに変換
            pl_df = pl.from_pandas(file_info_df)
            print(f"Polarsデータフレーム作成: {pl_df.shape}")

            # 既存のresultsシートを削除
            if 'results' in self.workbook.sheetnames:
                del self.workbook['results']
                print("既存のresultsシートを削除しました")

            # 新しいresultsシートを作成
            results_sheet = self.workbook.create_sheet('results')
            print("新しいresultsシートを作成しました")

            # ヘッダーを設定
            headers = file_info_df.columns.tolist()
            for col_idx, header in enumerate(headers, 1):
                results_sheet.cell(row=1, column=col_idx).value = header

            # Polarsをリストに変換して一気に書き込み
            data_rows = pl_df.to_numpy().tolist()
            for row_idx, row_data in enumerate(data_rows, 2):
                for col_idx, value in enumerate(row_data, 1):
                    results_sheet.cell(row=row_idx, column=col_idx).value = value

            # 列幅を自動調整
            for col_idx in range(1, len(headers) + 1):
                results_sheet.column_dimensions[chr(64 + col_idx)].width = 20

            print(f"{len(data_rows)}行のデータを書き込みました")
            return True, None
        except Exception as e:
            print(f"結果シート作成エラー: {str(e)}")
            return False, f"結果保存エラー: {str(e)}"

    def open_excel(self):
        """処理後にExcelファイルを開く"""
        utils.open_excel_file(self.excel_file_path)
