from __future__ import annotations

from docx.shared import Cm, Length, Mm, Pt


def pt(value: float) -> Pt:
    return Pt(value)


def cm(value: float) -> Cm:
    return Cm(value)


def mm(value: float) -> Mm:
    return Mm(value)


def lengths_equal(a: Length, b: Length) -> bool:
    """Сравнить две длины так, как их на самом деле хранит Word.

    Разведка (scratchpad/explore_oxml.py) показала: Cm(1.25), записанный в
    XML, превращается в твипы (709) и при обратном чтении восстанавливается
    как Twips(709) = 450215 EMU — а не исходные 450000 EMU у Cm(1.25).
    Прямое сравнение Length == Length после записи-чтения не срабатывает
    даже само с собой; .twips — то, что реально лежит в атрибуте XML.
    """
    return a.twips == b.twips
