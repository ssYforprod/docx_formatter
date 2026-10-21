from __future__ import annotations

from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from docx_formatter.core.context import DocContext
from docx_formatter.core.oxml import iter_block_items, set_run_font
from docx_formatter.core.profile import BodyProfile, Profile
from docx_formatter.core.report import Change
from docx_formatter.core.units import cm, lengths_equal, pt

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


def _paragraph_matches_body(paragraph: Paragraph, body: BodyProfile) -> bool:
    pf = paragraph.paragraph_format
    if pf.alignment != _ALIGNMENTS[body.alignment]:
        return False
    if pf.line_spacing != body.line_spacing:
        return False
    if pf.first_line_indent is None or not lengths_equal(pf.first_line_indent, cm(body.first_line_indent_cm)):
        return False
    if pf.space_before is None or not lengths_equal(pf.space_before, pt(body.space_before_pt)):
        return False
    if pf.space_after is None or not lengths_equal(pf.space_after, pt(body.space_after_pt)):
        return False
    size = pt(body.size_pt)
    return all(_run_matches_font(run, body.font, size) for run in paragraph.runs)


def _normal_style_matches(document: Document, font_name: str) -> bool:
    rPr = document.styles["Normal"].element.find(qn("w:rPr"))
    if rPr is None:
        return False
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        return False
    return all(rFonts.get(qn(attr)) == font_name for attr in _FONT_ATTRS)


class BodyTextRule:
    """Шрифт/интерлиньяж/отступ/выравнивание для обычного текста.

    Применяется ко всем абзацам тела документа, КРОМЕ опознанных как
    заголовки, подписи к рисункам/таблицам или абзаца с полем оглавления
    (context.heading_matches / context.caption_matches /
    context.toc_paragraph_index, посчитаны один раз в pipeline — см.
    DocContext) — у тех своё оформление или оно не нужно вовсе. Абзацы
    внутри таблиц тоже не трогаем — у них свой шрифт, см. правило tables.
    """

    id = "body_text"

    def enabled_for(self, profile: Profile) -> bool:
        return True

    def apply(self, context: DocContext) -> list[Change]:
        profile = context.profile
        body = profile.body
        threshold = profile.headings.detect_threshold
        size = pt(body.size_pt)
        alignment = _ALIGNMENTS[body.alignment]
        indent = cm(body.first_line_indent_cm)
        space_before = pt(body.space_before_pt)
        space_after = pt(body.space_after_pt)

        changes: list[Change] = []
        for index, item in enumerate(iter_block_items(context.document)):
            if not isinstance(item, Paragraph):
                continue
            match = context.heading_matches.get(index)
            if match is not None and match.confidence >= threshold:
                continue
            if index in context.caption_matches:
                continue
            if index == context.toc_paragraph_index:
                continue
            if _paragraph_matches_body(item, body):
                continue

            for run in item.runs:
                set_run_font(run, body.font)
                run.font.size = size
            pf = item.paragraph_format
            pf.alignment = alignment
            pf.line_spacing = body.line_spacing
            pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
            pf.first_line_indent = indent
            pf.space_before = space_before
            pf.space_after = space_after

            changes.append(
                Change(self.id, "info", f"абзац #{index + 1}", "приведён к профилю (шрифт/интерлиньяж/отступ)")
            )

        if not _normal_style_matches(context.document, body.font):
            set_run_font(context.document.styles["Normal"], body.font)
            changes.append(
                Change(self.id, "info", "стиль Normal", "установлен шрифт по умолчанию для нового текста")
            )

        return changes
