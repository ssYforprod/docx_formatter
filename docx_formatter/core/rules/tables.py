from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from docx_formatter.core.context import DocContext
from docx_formatter.core.oxml import iter_block_items, set_run_font
from docx_formatter.core.profile import Profile, TablesProfile
from docx_formatter.core.report import Change
from docx_formatter.core.units import lengths_equal, pt

_ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

_FONT_ATTRS = ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia")


def _run_matches_font(run: Run, font_name: str, size: Pt) -> bool:
    if run.font.size is None or not lengths_equal(run.font.size, size):
        return False
    rPr = run._element.find(qn("w:rPr"))
    if rPr is None:
        return False
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        return False
    return all(rFonts.get(qn(attr)) == font_name for attr in _FONT_ATTRS)


def _cell_paragraph_matches(paragraph: Paragraph, cfg: TablesProfile, *, want_bold: bool | None) -> bool:
    pf = paragraph.paragraph_format
    if pf.alignment != _ALIGNMENTS[cfg.cell_alignment]:
        return False
    if pf.line_spacing != cfg.line_spacing:
        return False
    size = pt(cfg.size_pt)
    for run in paragraph.runs:
        if not _run_matches_font(run, cfg.font, size):
            return False
        if want_bold is not None and run.bold != want_bold:
            return False
    return True


def _apply_cell_paragraph(paragraph: Paragraph, cfg: TablesProfile, *, want_bold: bool | None) -> bool:
    if _cell_paragraph_matches(paragraph, cfg, want_bold=want_bold):
        return False

    size = pt(cfg.size_pt)
    for run in paragraph.runs:
        set_run_font(run, cfg.font)
        run.font.size = size
        if want_bold is not None:
            run.bold = want_bold

    pf = paragraph.paragraph_format
    pf.alignment = _ALIGNMENTS[cfg.cell_alignment]
    pf.line_spacing = cfg.line_spacing
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    return True


class TablesRule:
    """Шрифт/кегль/выравнивание ячеек таблицы; шапка (первая строка)
    дополнительно выделяется жирным по tables.header_bold. Остальные
    строки жирность не навязывают — в профиле для них такого параметра
    нет."""

    id = "tables"

    def enabled_for(self, profile: Profile) -> bool:
        return True

    def apply(self, context: DocContext) -> list[Change]:
        cfg = context.profile.tables
        changes: list[Change] = []
        table_index = 0

        for item in iter_block_items(context.document):
            if not isinstance(item, Table):
                continue
            table_index += 1
            changed = False
            for row_index, row in enumerate(item.rows):
                want_bold = cfg.header_bold if row_index == 0 else None
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if _apply_cell_paragraph(paragraph, cfg, want_bold=want_bold):
                            changed = True
            if changed:
                changes.append(Change(self.id, "info", f"таблица #{table_index}", "оформлены ячейки таблицы"))

        return changes
