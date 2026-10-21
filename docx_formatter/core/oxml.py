from __future__ import annotations

from typing import Iterator, Protocol, Union

from docx.document import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

# Прямые атрибуты шрифта на каждый скрипт и их theme-варианты.
# Написание "w:cstheme" (не "w:csTheme") — так атрибут назван в самой
# спецификации OOXML, это не опечатка.
_FONT_ATTRS = ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia")
_THEME_ATTRS = ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme")


class _HasElement(Protocol):
    """И Run, и стиль абзаца несут один и тот же .element (CT_R / CT_Style)
    с методом get_or_add_rPr() — разведкой подтверждено, что это один и тот
    же путь для обоих типов объектов."""

    @property
    def element(self): ...


def set_run_font(run_or_style: _HasElement, font_name: str) -> None:
    """Установить шрифт для всех скриптов (ascii/hAnsi/cs/eastAsia).

    Принимает Run или Style (например, document.styles["Normal"]) — у
    обоих get_or_add_rPr()/get_or_add_rFonts() работают одинаково.

    `run.font.name` пишет только ascii/hAnsi — подтверждено разведкой
    (scratchpad/explore_oxml.py): кириллица и complex-script текст не
    подхватывают шрифт без явного cs/eastAsia. В python-docx 1.2.0 для
    cs/eastAsia нет питоничных дескрипторов (`rFonts.cs = ...` молча
    ничего не делает) — поэтому пишем сырые XML-атрибуты напрямую.
    Также убираем *Theme-атрибуты: при их наличии Word может предпочесть
    тему вместо явно заданного имени шрифта.
    """
    rPr = run_or_style.element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    for attr in _FONT_ATTRS:
        rFonts.set(qn(attr), font_name)
    for attr in _THEME_ATTRS:
        key = qn(attr)
        if rFonts.get(key) is not None:
            del rFonts.attrib[key]


BlockItem = Union[Paragraph, Table]


def iter_block_items(document: Document) -> Iterator[BlockItem]:
    """Обойти параграфы и таблицы в порядке документа.

    `document.paragraphs`/`.tables` не сохраняют взаимный порядок, поэтому
    обходим `document.element.body` напрямую, пропуская непечатаемые узлы
    вроде `w:sectPr` (подтверждено разведкой: он всегда последний в body).
    """
    for child in document.element.body:
        if child.tag == qn("w:p"):
            yield Paragraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield Table(child, document)
