import sys
import os
import json
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QLineEdit,
                             QFileDialog, QMessageBox, QProgressBar, QTextEdit,
                             QGroupBox, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QDragEnterEvent, QDropEvent, QPalette, QColor

from report import parse_pdf, generate_excel, generate_json, generate_html


class WorkerThread(QThread):
    progress = pyqtSignal(int)
    log = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(self, pdf_path, output_dir, gen_xlsx, gen_json, gen_html):
        super().__init__()
        self.pdf_path = pdf_path
        self.output_dir = output_dir
        self.gen_xlsx = gen_xlsx
        self.gen_json = gen_json
        self.gen_html = gen_html

    def run(self):
        try:
            self.log.emit("Обработка PDF...")
            self.progress.emit(10)
            data = parse_pdf(self.pdf_path)
            self.progress.emit(40)
            self.log.emit(f"  Учеников: {len(data['students'])}")
            self.log.emit(f"  Предметов: {len(data['subjects'])}")

            base = os.path.splitext(os.path.basename(self.pdf_path))[0]

            if self.gen_xlsx:
                self.log.emit("Генерация Excel...")
                xlsx_path = os.path.join(self.output_dir, base + '.xlsx')
                generate_excel(data, xlsx_path)
                self.log.emit(f"  Сохранено: {xlsx_path}")
                self.progress.emit(60)

            if self.gen_json:
                self.log.emit("Генерация JSON...")
                json_path = os.path.join(self.output_dir, base + '.json')
                generate_json(data, json_path)
                self.log.emit(f"  Сохранено: {json_path}")
                self.progress.emit(80)

            if self.gen_html:
                self.log.emit("Генерация HTML...")
                html_path = os.path.join(self.output_dir, base + '.html')
                generate_html(data, html_path)
                self.log.emit(f"  Сохранено: {html_path}")
                self.progress.emit(100)

            self.finished.emit(True, "Готово!")
        except Exception as e:
            self.finished.emit(False, str(e))


class DropLabel(QLabel):
    file_dropped = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet(
            "QLabel { border: 2px dashed #aaa; border-radius: 8px; "
            "background: #f9f9f9; color: #888; font-size: 14px; }"
            "QLabel:hover { border-color: #1a237e; background: #e8eaf6; }"
        )
        self.setText("Перетащите PDF сюда\nили нажмите кнопку ниже")
        self.setMinimumHeight(120)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].toLocalFile().lower().endswith('.pdf'):
                event.acceptProposedAction()
                self.setStyleSheet(
                    "QLabel { border: 2px dashed #1a237e; border-radius: 8px; "
                    "background: #e8eaf6; color: #1a237e; font-size: 14px; }"
                )

    def dragLeaveEvent(self, event):
        self.setStyleSheet(
            "QLabel { border: 2px dashed #aaa; border-radius: 8px; "
            "background: #f9f9f9; color: #888; font-size: 14px; }"
            "QLabel:hover { border-color: #1a237e; background: #e8eaf6; }"
        )

    def dropEvent(self, event: QDropEvent):
        self.setStyleSheet(
            "QLabel { border: 2px solid #2e7d32; border-radius: 8px; "
            "background: #e8f5e9; color: #2e7d32; font-size: 14px; }"
        )
        if event.mimeData().hasUrls():
            path = event.mimeData().urls()[0].toLocalFile()
            if path.lower().endswith('.pdf'):
                self.file_dropped.emit(path)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pdf_path = None
        self.worker = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("GradeReport — Генератор отчётов")
        self.setMinimumSize(600, 500)
        self.setStyleSheet("""
            QMainWindow { background: #f0f2f5; }
            QPushButton {
                background: #1a237e; color: white; border: none;
                border-radius: 6px; padding: 8px 16px; font-size: 13px;
            }
            QPushButton:hover { background: #283593; }
            QPushButton:disabled { background: #9fa8da; }
            QLineEdit {
                border: 1px solid #ccc; border-radius: 4px;
                padding: 6px 8px; font-size: 13px;
            }
            QTextEdit {
                border: 1px solid #ddd; border-radius: 4px;
                background: #fafafa; font-family: Consolas, monospace;
                font-size: 12px;
            }
            QCheckBox { font-size: 13px; }
            QProgressBar {
                border: 1px solid #ccc; border-radius: 4px;
                text-align: center; height: 20px;
            }
            QProgressBar::chunk {
                background: #1a237e; border-radius: 3px;
            }
        """)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # Drop zone
        self.drop_label = DropLabel()
        self.drop_label.file_dropped.connect(self.on_file_dropped)
        layout.addWidget(self.drop_label)

        # File selection
        file_box = QHBoxLayout()
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.file_edit.setPlaceholderText("PDF-файл не выбран")
        btn_select = QPushButton("Выбрать PDF")
        btn_select.clicked.connect(self.select_file)
        file_box.addWidget(self.file_edit)
        file_box.addWidget(btn_select)
        layout.addLayout(file_box)

        # Output directory
        dir_box = QHBoxLayout()
        self.dir_edit = QLineEdit()
        self.dir_edit.setPlaceholderText("Папка сохранения")
        btn_dir = QPushButton("Обзор")
        btn_dir.clicked.connect(self.select_dir)
        dir_box.addWidget(self.dir_edit)
        dir_box.addWidget(btn_dir)
        layout.addLayout(dir_box)

        # Format checkboxes
        fmt_group = QGroupBox("Форматы вывода")
        fmt_layout = QHBoxLayout()
        self.cb_xlsx = QCheckBox("Excel (.xlsx)")
        self.cb_xlsx.setChecked(True)
        self.cb_json = QCheckBox("JSON (.json)")
        self.cb_json.setChecked(True)
        self.cb_html = QCheckBox("HTML (.html)")
        self.cb_html.setChecked(True)
        fmt_layout.addWidget(self.cb_xlsx)
        fmt_layout.addWidget(self.cb_json)
        fmt_layout.addWidget(self.cb_html)
        fmt_group.setLayout(fmt_layout)
        layout.addWidget(fmt_group)

        # Progress
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        # Log
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(120)
        layout.addWidget(self.log)

        # Buttons
        btn_layout = QHBoxLayout()
        self.btn_run = QPushButton("Сгенерировать отчёт")
        self.btn_run.clicked.connect(self.run_report)
        btn_open = QPushButton("Открыть папку")
        btn_open.clicked.connect(self.open_folder)
        btn_layout.addWidget(self.btn_run)
        btn_layout.addWidget(btn_open)
        layout.addLayout(btn_layout)

    def select_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Выберите PDF-файл", "", "PDF Files (*.pdf)")
        if path:
            self.set_pdf(path)

    def on_file_dropped(self, path):
        self.set_pdf(path)

    def set_pdf(self, path):
        self.pdf_path = path
        self.file_edit.setText(path)
        self.drop_label.setText(os.path.basename(path))
        if not self.dir_edit.text():
            self.dir_edit.setText(os.path.dirname(path))

    def select_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Выберите папку")
        if d:
            self.dir_edit.setText(d)

    def append_log(self, text):
        self.log.append(text)

    def run_report(self):
        if not self.pdf_path:
            QMessageBox.warning(self, "Внимание", "Выберите PDF-файл")
            return
        if not os.path.exists(self.pdf_path):
            QMessageBox.warning(self, "Ошибка", f"Файл не найден:\n{self.pdf_path}")
            return

        output_dir = self.dir_edit.text().strip()
        if not output_dir:
            output_dir = os.path.dirname(self.pdf_path)
            self.dir_edit.setText(output_dir)

        if not os.path.isdir(output_dir):
            QMessageBox.warning(self, "Ошибка", f"Папка не существует:\n{output_dir}")
            return

        if not any([self.cb_xlsx.isChecked(), self.cb_json.isChecked(), self.cb_html.isChecked()]):
            QMessageBox.warning(self, "Внимание", "Выберите хотя бы один формат")
            return

        self.log.clear()
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.btn_run.setEnabled(False)

        self.worker = WorkerThread(
            self.pdf_path, output_dir,
            self.cb_xlsx.isChecked(),
            self.cb_json.isChecked(),
            self.cb_html.isChecked()
        )
        self.worker.progress.connect(self.progress.setValue)
        self.worker.log.connect(self.append_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def on_finished(self, success, message):
        self.progress.setVisible(False)
        self.btn_run.setEnabled(True)
        if success:
            QMessageBox.information(self, "Готово", message)
        else:
            QMessageBox.critical(self, "Ошибка", message)

    def open_folder(self):
        d = self.dir_edit.text().strip()
        if d and os.path.isdir(d):
            os.startfile(d)


def main():
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
