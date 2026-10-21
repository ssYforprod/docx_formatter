from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from docx_formatter.gui.main_window import MainWindow


def run() -> None:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(480, 320)
    window.show()
    app.exec()
