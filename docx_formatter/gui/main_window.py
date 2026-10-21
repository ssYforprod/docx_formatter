from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from docx_formatter.core.profile import ProfileError, load_profile
from docx_formatter.core.report import FormatReport
from docx_formatter.gui.worker import FormatWorker
from docx_formatter.io.profile_repo import ProfileEntry, list_profiles


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("DOCX Formatter")

        self._worker: FormatWorker | None = None
        self._profiles: list[ProfileEntry] = list_profiles()

        self._file_edit = QLineEdit()
        self._file_edit.setReadOnly(True)
        browse_button = QPushButton("Обзор…")
        browse_button.clicked.connect(self._choose_file)

        self._profile_combo = QComboBox()
        for entry in self._profiles:
            self._profile_combo.addItem(entry.name, entry)

        self._format_button = QPushButton("Форматировать")
        self._format_button.setEnabled(False)
        self._format_button.clicked.connect(self._start_formatting)

        self._result_view = QTextEdit()
        self._result_view.setReadOnly(True)

        file_row = QHBoxLayout()
        file_row.addWidget(self._file_edit)
        file_row.addWidget(browse_button)

        layout = QVBoxLayout()
        layout.addWidget(QLabel("Файл отчёта (.docx):"))
        layout.addLayout(file_row)
        layout.addWidget(QLabel("Профиль оформления:"))
        layout.addWidget(self._profile_combo)
        layout.addWidget(self._format_button)
        layout.addWidget(self._result_view)

        central = QWidget()
        central.setLayout(layout)
        self.setCentralWidget(central)

        if not self._profiles:
            self._show_result("Не найдено ни одного профиля оформления в папке profiles/.")

    def _choose_file(self) -> None:
        path_str, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл отчёта", "", "Документы Word (*.docx)"
        )
        if not path_str:
            return
        self._file_edit.setText(path_str)
        self._format_button.setEnabled(bool(self._profiles))

    def _start_formatting(self) -> None:
        input_path = Path(self._file_edit.text())
        entry = self._profile_combo.currentData()
        if not isinstance(entry, ProfileEntry):
            return

        try:
            profile, warnings = load_profile(entry.path)
        except ProfileError as exc:
            self._show_result("\n".join(f"Ошибка профиля: {error}" for error in exc.errors))
            return

        self._format_button.setEnabled(False)
        lines = ["Форматирование…"]
        lines.extend(f"Предупреждение профиля: {warning}" for warning in warnings)
        self._show_result("\n".join(lines))

        self._worker = FormatWorker(profile, input_path)
        self._worker.succeeded.connect(self._on_succeeded)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(lambda: self._format_button.setEnabled(True))
        self._worker.start()

    def _on_succeeded(self, out_path: object, report: object) -> None:
        if not isinstance(report, FormatReport):
            return
        lines = [f"Готово: {out_path}"]
        for warning in report.warnings:
            lines.append(f"Внимание: {warning.scope}: {warning.what}")
        self._show_result("\n".join(lines))

    def _on_failed(self, message: str) -> None:
        self._show_result(f"Ошибка форматирования: {message}")

    def _show_result(self, text: str) -> None:
        self._result_view.setPlainText(text)
