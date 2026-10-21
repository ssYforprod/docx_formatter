from __future__ import annotations

import re
from pathlib import Path

CORE_DIR = Path(__file__).resolve().parent.parent / "docx_formatter" / "core"

_QT_IMPORT = re.compile(r"\bPySide6\b")


def test_core_has_no_qt_imports() -> None:
    offenders = [
        path
        for path in CORE_DIR.rglob("*.py")
        if _QT_IMPORT.search(path.read_text(encoding="utf-8"))
    ]
    assert not offenders, f"docx_formatter/core must not import PySide6: {offenders}"
