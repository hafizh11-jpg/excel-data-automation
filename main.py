import sys
import os
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QFileDialog, QProgressBar, QTextEdit,
    QCheckBox, QComboBox, QGroupBox, QGridLayout, QMessageBox, QTabWidget
)
from PySide6.QtCore import Qt, QThread, Signal
import processor

CHART_TYPES = ['Bar', 'Histogram', 'Line', 'Stacked', 'Scatter', 'Heatmap', 'Funnel', 'Pie']

class Worker(QThread):
    progress = Signal(str, int)
    finished = Signal(dict)
    error = Signal(str)

    def __init__(self, input_path, output_path, options):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.options = options

    def run(self):
        try:
            summary = processor.process_file(self.input_path, self.output_path, self.options, lambda m, p: self.progress.emit(m, p))
            self.finished.emit(summary)
        except Exception as e:
            import traceback
            self.error.emit(traceback.format_exc())

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Data Automation")
        self.resize(900, 700)
        self.input_path = ""
        self.df = None
        self.header = []
        self.wb = None
        self.sheet_name = None

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # File row
        file_group = QGroupBox("File Excel")
        fg = QHBoxLayout(file_group)
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("Pilih file .xlsx...")
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browse)
        load_btn = QPushButton("Load")
        load_btn.clicked.connect(self.load_file)
        fg.addWidget(self.path_edit)
        fg.addWidget(browse_btn)
        fg.addWidget(load_btn)
        layout.addWidget(file_group)

        self.info_label = QLabel("Belum ada file dimuat.")
        layout.addWidget(self.info_label)

        tabs = QTabWidget()
        layout.addWidget(tabs)

        # Tab options
        opt_widget = QWidget()
        opt_layout = QVBoxLayout(opt_widget)

        # detection options
        det_group = QGroupBox("Deteksi & Sheet")
        dg = QVBoxLayout(det_group)
        self.chk_dup = QCheckBox("Buat sheet Data_Duplikat")
        self.chk_dup.setChecked(True)
        self.chk_empty = QCheckBox("Buat sheet Data_Kosong")
        self.chk_empty.setChecked(True)
        self.chk_clean = QCheckBox("Buat sheet Data_Bersih (tanpa duplikat & kosong, diurutkan)")
        self.chk_clean.setChecked(True)
        self.chk_cat = QCheckBox("Buat sheet per kategori (dari Data_Bersih)")
        self.chk_cat.setChecked(True)
        self.chk_charts = QCheckBox("Buat sheet Charts")
        self.chk_charts.setChecked(True)
        for c in [self.chk_dup, self.chk_empty, self.chk_clean, self.chk_cat, self.chk_charts]:
            dg.addWidget(c)
        dg.addWidget(QLabel("Sumber sheet kategori:"))
        self.cat_source_combo = QComboBox()
        self.cat_source_combo.addItems(["Data Bersih", "Data Original", "Original (skip dup & kosong)"])
        self.cat_source_combo.setCurrentIndex(1)  # default: Data Original (tanpa mengurangi)
        dg.addWidget(self.cat_source_combo)
        opt_layout.addWidget(det_group)

        # empty columns selection
        empty_group = QGroupBox("Kolom untuk deteksi kosong")
        eg = QVBoxLayout(empty_group)
        self.empty_combo_all = QCheckBox("Semua kolom")
        self.empty_combo_all.setChecked(True)
        eg.addWidget(self.empty_combo_all)
        self.empty_list_label = QLabel("(Jika tidak semua, pilih dari dropdown kolom utama)")
        eg.addWidget(self.empty_list_label)
        self.empty_col_combo = QComboBox()
        eg.addWidget(self.empty_col_combo)
        opt_layout.addWidget(empty_group)

        # sort/group
        sort_group = QGroupBox("Pengurutan & Pengelompokan")
        sg = QGridLayout(sort_group)
        sg.addWidget(QLabel("Sort by:"), 0, 0)
        self.sort_combo = QComboBox()
        self.sort_combo.addItem("(none)")
        sg.addWidget(self.sort_combo, 0, 1)
        sg.addWidget(QLabel("Group by (kategori):"), 1, 0)
        self.group_combo = QComboBox()
        self.group_combo.addItem("(auto)")
        sg.addWidget(self.group_combo, 1, 1)
        sg.addWidget(QLabel("Kolom kategori untuk sheet per kategori:"), 2, 0)
        self.cat_col_combo = QComboBox()
        self.cat_col_combo.addItem("(auto)")
        sg.addWidget(self.cat_col_combo, 2, 1)
        sg.addWidget(QLabel("Kolom kategori untuk chart:"), 3, 0)
        self.chart_cat_combo = QComboBox()
        self.chart_cat_combo.addItem("(auto)")
        sg.addWidget(self.chart_cat_combo, 3, 1)
        opt_layout.addWidget(sort_group)

        tabs.addTab(opt_widget, "Opsi")

        # Tab charts
        chart_widget = QWidget()
        chart_layout = QVBoxLayout(chart_widget)
        chart_group = QGroupBox("Jenis Chart")
        cg = QGridLayout(chart_group)
        self.chart_checks = {}
        for i, ct in enumerate(CHART_TYPES):
            cb = QCheckBox(ct)
            cb.setChecked(True)
            self.chart_checks[ct] = cb
            cg.addWidget(cb, i // 4, i % 4)
        chart_layout.addWidget(chart_group)
        tabs.addTab(chart_widget, "Charts")

        # Process
        proc_layout = QHBoxLayout()
        self.process_btn = QPushButton("Proses & Simpan")
        self.process_btn.clicked.connect(self.process)
        proc_layout.addWidget(self.process_btn)
        self.open_btn = QPushButton("Buka Output")
        self.open_btn.clicked.connect(self.open_output)
        self.open_btn.setEnabled(False)
        proc_layout.addWidget(self.open_btn)
        layout.addLayout(proc_layout)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        self.output_path = ""

    def log_msg(self, msg):
        self.log.append(msg)

    def browse(self):
        path, _ = QFileDialog.getOpenFileName(self, "Pilih file Excel", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.path_edit.setText(path)

    def load_file(self):
        path = self.path_edit.text().strip()
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "Error", "File tidak ditemukan.")
            return
        try:
            wb, ws, df, header, sheet_name = processor.load_workbook_df(path)
            self.wb = wb
            self.df = df
            self.header = header
            self.sheet_name = sheet_name
            self.input_path = path
            self.info_label.setText(f"Sheet: {sheet_name} | Baris: {len(df)} | Kolom: {len(header)}")
            # populate combos
            for combo in [self.sort_combo, self.group_combo, self.cat_col_combo, self.chart_cat_combo, self.empty_col_combo]:
                combo.clear()
            self.sort_combo.addItem("(none)")
            self.group_combo.addItem("(auto)")
            self.cat_col_combo.addItem("(auto)")
            self.chart_cat_combo.addItem("(auto)")
            for c in header:
                self.sort_combo.addItem(c)
                self.group_combo.addItem(c)
                self.cat_col_combo.addItem(c)
                self.chart_cat_combo.addItem(c)
                self.empty_col_combo.addItem(c)
            self.log_msg(f"Loaded: {path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def process(self):
        if self.df is None:
            QMessageBox.warning(self, "Error", "Load file dulu.")
            return
        input_path = self.input_path
        base, ext = os.path.splitext(input_path)
        output_path = base + "_processed" + ext
        out, _ = QFileDialog.getSaveFileName(self, "Simpan hasil", output_path, "Excel Files (*.xlsx)")
        if not out:
            return
        self.output_path = out

        empty_cols = None
        if not self.empty_combo_all.isChecked():
            empty_cols = [self.empty_col_combo.currentText()]

        chart_types = [ct for ct, cb in self.chart_checks.items() if cb.isChecked()]

        sort_by = self.sort_combo.currentText()
        if sort_by == "(none)":
            sort_by = None
        group_by = self.group_combo.currentText()
        if group_by == "(auto)":
            group_by = None
        cat_col = self.cat_col_combo.currentText()
        if cat_col == "(auto)":
            cat_col = None
        chart_cat = self.chart_cat_combo.currentText()
        txt = self.cat_source_combo.currentText()
        if txt == "Data Bersih":
            cat_source = 'clean'
        elif txt == "Original (skip dup & kosong)":
            cat_source = 'filtered'
        else:
            cat_source = 'original'
        if chart_cat == "(auto)":
            chart_cat = None

        options = {
            'add_dup_sheet': self.chk_dup.isChecked(),
            'add_empty_sheet': self.chk_empty.isChecked(),
            'add_clean_sheet': self.chk_clean.isChecked(),
            'add_category_sheets': self.chk_cat.isChecked(),
            'add_charts': self.chk_charts.isChecked(),
            'empty_cols': empty_cols,
            'sort_by': sort_by,
            'group_by': group_by,
            'category_col': cat_col,
            'category_source': cat_source,
            'chart_cat_col': chart_cat,
            'chart_types': chart_types,
        }

        self.process_btn.setEnabled(False)
        self.progress.setValue(0)
        self.worker = Worker(input_path, out, options)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_progress(self, msg, pct):
        self.log_msg(msg)
        self.progress.setValue(pct)

    def on_finished(self, summary):
        self.process_btn.setEnabled(True)
        self.open_btn.setEnabled(True)
        self.log_msg(f"Selesai. Baris: {summary['rows']}, Duplikat: {summary['duplicates']}, Kosong: {summary['empty']}, Bersih: {summary['clean']}")
        self.log_msg(f"Output: {summary['output']}")
        QMessageBox.information(self, "Selesai", f"File disimpan:\n{summary['output']}")

    def on_error(self, err):
        self.process_btn.setEnabled(True)
        self.log_msg("ERROR:\n" + err)
        QMessageBox.critical(self, "Error", err[:500])

    def open_output(self):
        if self.output_path and os.path.exists(self.output_path):
            if sys.platform == 'win32':
                os.startfile(self.output_path)
            elif sys.platform == 'darwin':
                os.system(f'open "{self.output_path}"')
            else:
                os.system(f'xdg-open "{self.output_path}"')

def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
