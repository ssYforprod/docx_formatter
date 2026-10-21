from __future__ import annotations

from pathlib import Path

from docx import Document


def make_docx(tmp_path: Path, blocks: list[tuple[str, str]]) -> Path:
    """Build a minimal .docx from (kind, text) blocks, for use as test input."""
    document = Document()
    for kind, text in blocks:
        if kind == "p":
            document.add_paragraph(text)
        else:
            raise ValueError(f"unknown block kind: {kind!r}")

    path = tmp_path / "input.docx"
    document.save(path)
    return path
