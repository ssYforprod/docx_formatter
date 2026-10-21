from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread, Signal

from docx_formatter.core.pipeline import Formatter
from docx_formatter.core.profile import Profile


class FormatWorker(QThread):
    """Форматирование в фоновом потоке — Formatter.run() может занять
    заметное время на большом документе, и делать это в UI-потоке
    подвесило бы окно."""

    succeeded = Signal(object, object)  # (Path, FormatReport)
    failed = Signal(str)

    def __init__(self, profile: Profile, input_path: Path) -> None:
        super().__init__()
        self._profile = profile
        self._input_path = input_path

    def run(self) -> None:
        try:
            out_path, report = Formatter(self._profile).run(self._input_path)
        except Exception as exc:  # показываем пользователю любую ошибку форматирования
            self.failed.emit(str(exc))
            return
        self.succeeded.emit(out_path, report)
