from __future__ import annotations

from pathlib import Path

from docx import Document

from factories import make_docx

import docx_formatter


def test_package_is_importable() -> None:
    assert docx_formatter.__name__ == "docx_formatter"


def test_factory_creates_openable_docx(tmp_path: Path) -> None:
    path = make_docx(tmp_path, blocks=[("p", "Привет, мир")])

    document = Document(path)

    assert [p.text for p in document.paragraphs] == ["Привет, мир"]
