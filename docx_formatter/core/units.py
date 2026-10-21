from __future__ import annotations

from docx.shared import Cm, Mm, Pt


def pt(value: float) -> Pt:
    return Pt(value)


def cm(value: float) -> Cm:
    return Cm(value)


def mm(value: float) -> Mm:
    return Mm(value)
