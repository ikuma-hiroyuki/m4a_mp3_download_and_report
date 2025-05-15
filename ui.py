import os
import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QLabel, QPushButton, QFileDialog,
    QCheckBox, QProgressBar, QTextEdit, QGroupBox,
    QScrollArea, QMessageBox
)
from PyQt5.QtCore import QThread, pyqtSignal

from excel_processor import ExcelProcessor
from file_processor import FileProcessor


class WorkerThread(QThread):
    """バックグラウンドで処理を実行するためのスレッド"""
    update_progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)
    file_processed = pyqtSignal(str, int, str, str)  # シート名、行番号、ファイル名、結果

    def __init__(self, excel_file, download_dir, selected_sheets, check_l_column):
        super().__init__()
        self.excel_file = excel_file
        self.download_dir = download_dir
        self.selected_sheets = selected_sheets
        self.check_l_column = check_l_column
        self.excel_processor = None
        self.file_processor = None

    def run(self):
        # Excelプロセッサの初期化
        self.excel_processor = ExcelProcessor(self.excel_file)
        success, error_msg = self.excel_processor.load_excel()

        if not success:
            self.finished.emit(False, error_msg)
            return

        # ファイルプロセッサの初期化
        self.file_processor = FileProcessor(self.download_dir)

        # 選択されたシートからリンクを取得
        links_data = self.excel_processor.get_links_from_sheets(self.selected_sheets, self.check_l_column)
        total_links = len(links_data)

        if total_links == 0:
            self.finished.emit(False, "処理対象のファイルリンクが見つかりませんでした。")
            return

        # 各リンクを処理
        for i, link_info in enumerate(links_data):
            sheet_name = link_info['sheet_name']
            row_num = link_info['row_num']
            url = link_info['url']

            progress_percent = int((i / total_links) * 100)
            self.update_progress.emit(progress_percent, f"処理中... ({i + 1}/{total_links})")

            # ファイルをダウンロード
            file_path, error = self.file_processor.download_file(url, sheet_name, row_num)

            if file_path:
                file_name = os.path.basename(file_path)

                # 再生時間を取得（ファイル情報から）
                duration = None
                key = (sheet_name, row_num)
                if key in self.file_processor.file_info:
                    duration = self.file_processor.file_info[key].get('duration')

                # ハイパーリンクを更新（再生時間も含めて）
                self.excel_processor.update_cell_with_hyperlink(
                    sheet_name, row_num, file_name, url, duration
                )
                self.file_processed.emit(sheet_name, row_num, file_name, f"成功 (再生時間: {duration or '不明'})")
            else:
                self.file_processed.emit(sheet_name, row_num, url, f"失敗: {error}")
                break

        # 結果をシートに保存
        file_info_df = self.file_processor.get_file_info_dataframe()
        if file_info_df.height > 0:
            success, error_msg = self.excel_processor.save_results(file_info_df)
            if not success:
                self.update_progress.emit(90, f"結果シートの作成に失敗: {error_msg}")
        else:
            self.update_progress.emit(90, "分析結果がありません。結果シートは作成されません。")

        # Excelファイルを保存
        success, error_msg = self.excel_processor.save_excel()
        if success:
            # Excelファイルを開く
            self.excel_processor.open_excel()
            self.finished.emit(True, "処理が完了しました。")
        else:
            self.finished.emit(False, error_msg)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.excel_file = ""
        self.download_dir = ""
        self.selected_sheets = []
        self.sheet_checkboxes = []
        self.excel_processor = None

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("M4A Report")
        self.setGeometry(100, 100, 800, 600)

        # メインウィジェットとレイアウト
        main_widget = QWidget()
        main_layout = QVBoxLayout()

        # ファイル選択エリア
        file_group = QGroupBox("Excelファイル選択")
        file_layout = QHBoxLayout()

        self.file_label = QLabel("ファイルが選択されていません")
        self.file_btn = QPushButton("参照...")
        self.file_btn.clicked.connect(self.select_excel_file)

        file_layout.addWidget(self.file_label, 1)
        file_layout.addWidget(self.file_btn, 0)
        file_group.setLayout(file_layout)

        # 保存先選択エリア
        save_group = QGroupBox("ダウンロードしたファイルの保存先")
        save_layout = QHBoxLayout()

        self.save_label = QLabel("フォルダが選択されていません")
        self.save_btn = QPushButton("参照...")
        self.save_btn.clicked.connect(self.select_download_dir)

        save_layout.addWidget(self.save_label, 1)
        save_layout.addWidget(self.save_btn, 0)
        save_group.setLayout(save_layout)

        # シート選択エリア
        sheet_group = QGroupBox("処理対象シート選択")
        self.sheet_layout = QVBoxLayout()
        self.sheet_layout.addWidget(QLabel("Excelファイルを選択してください"))

        sheet_scroll = QScrollArea()
        sheet_scroll.setWidgetResizable(True)
        sheet_content = QWidget()
        sheet_content.setLayout(self.sheet_layout)
        sheet_scroll.setWidget(sheet_content)

        # L列チェックボックスを追加
        self.check_l_column = QCheckBox("L列(再生時間)に値がある行をスキップする")
        self.check_l_column.setChecked(True)  # デフォルトでオン

        sheet_group_layout = QVBoxLayout()
        sheet_group_layout.addWidget(sheet_scroll)
        sheet_group_layout.addWidget(self.check_l_column)
        sheet_group.setLayout(sheet_group_layout)

        # 処理状況表示エリア
        progress_group = QGroupBox("処理状況")
        progress_layout = QVBoxLayout()

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)

        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)

        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.status_text)
        progress_group.setLayout(progress_layout)

        # 実行ボタン
        self.execute_btn = QPushButton("実行")
        self.execute_btn.clicked.connect(self.execute_process)
        self.execute_btn.setEnabled(False)

        # レイアウトに追加
        main_layout.addWidget(file_group)
        main_layout.addWidget(save_group)
        main_layout.addWidget(sheet_group, 1)  # シート選択エリアを拡大
        main_layout.addWidget(progress_group, 1)  # 処理状況表示エリアを拡大
        main_layout.addWidget(self.execute_btn)

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Excelファイルを選択", "", "Excel Files (*.xlsx *.xlsm)"
        )

        if file_path:
            self.excel_file = file_path
            self.file_label.setText(file_path)

            # シート情報を取得して表示
            self.excel_processor = ExcelProcessor(file_path)
            success, error_msg = self.excel_processor.load_excel()

            if success:
                self.load_sheets()
            else:
                QMessageBox.critical(self, "エラー", error_msg)
                self.excel_file = ""
                self.file_label.setText("ファイルが選択されていません")

            self.check_execute_button()

    def select_download_dir(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "ダウンロード先フォルダを選択", ""
        )

        if dir_path:
            self.download_dir = dir_path
            self.save_label.setText(dir_path)
            self.check_execute_button()

    def load_sheets(self):
        # 既存のチェックボックスをクリア
        for i in reversed(range(self.sheet_layout.count())):
            widget = self.sheet_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()

        self.sheet_checkboxes = []
        sheets = self.excel_processor.get_sheets()

        if sheets:
            for sheet_name in sheets:
                checkbox = QCheckBox(sheet_name)
                checkbox.stateChanged.connect(self.on_sheet_selection_changed)
                self.sheet_layout.addWidget(checkbox)
                self.sheet_checkboxes.append(checkbox)
        else:
            self.sheet_layout.addWidget(QLabel("シートが見つかりませんでした"))

    def on_sheet_selection_changed(self):
        self.selected_sheets = []
        for checkbox in self.sheet_checkboxes:
            if checkbox.isChecked():
                self.selected_sheets.append(checkbox.text())

        self.check_execute_button()

    def check_execute_button(self):
        # 実行ボタンの有効/無効を設定
        self.execute_btn.setEnabled(
            bool(self.excel_file) and
            bool(self.download_dir) and
            bool(self.selected_sheets)
        )

    def execute_process(self):
        # 処理の実行
        self.worker = WorkerThread(
            self.excel_file,
            self.download_dir,
            self.selected_sheets,
            self.check_l_column.isChecked()
        )
        self.worker.update_progress.connect(self.update_progress)
        self.worker.finished.connect(self.process_finished)
        self.worker.file_processed.connect(self.update_file_status)

        # UIの状態を更新
        self.execute_btn.setEnabled(False)
        self.file_btn.setEnabled(False)
        self.save_btn.setEnabled(False)
        for checkbox in self.sheet_checkboxes:
            checkbox.setEnabled(False)

        self.status_text.clear()
        self.status_text.append("処理を開始します...")

        # 処理開始
        self.worker.start()

    def update_progress(self, value, message):
        self.progress_bar.setValue(value)
        self.status_text.append(message)

    def update_file_status(self, sheet_name, row_num, file_name, status):
        self.status_text.append(f"シート「{sheet_name}」の {row_num} 行目: {file_name} - {status}")

    def process_finished(self, success, message):
        # UIの状態を元に戻す
        self.execute_btn.setEnabled(True)
        self.file_btn.setEnabled(True)
        self.save_btn.setEnabled(True)
        for checkbox in self.sheet_checkboxes:
            checkbox.setEnabled(True)

        # 結果を表示
        if success:
            self.progress_bar.setValue(100)
            self.status_text.append("処理が完了しました。")
            QMessageBox.information(self, "完了", message)
        else:
            self.status_text.append(f"エラー: {message}")
            QMessageBox.critical(self, "エラー", message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
