from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph

from docx_formatter.core.context import DocContext
from docx_formatter.core.oxml import iter_block_items, set_run_font
from docx_formatter.core.profile import HeadingLevel, Profile
from docx_formatter.core.report import Change
from docx_formatter.core.units import cm, lengths_equal, pt

_ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# Порог "форматировать" задаётся профилем (headings.detect_threshold).
# Порог "хотя бы предупредить" — нет, он намеренно не вынесен в профиль:
# это внутренний компромисс против шума, а не параметр оформления.
_WARN_FLOOR = 0.4


def _heading_matches(paragraph: Paragraph, level_cfg: HeadingLevel) -> bool:
    pf = paragraph.paragraph_format
    if pf.alignment != _ALIGNMENTS[level_cfg.alignment]:
        return False
    if pf.first_line_indent is None or not lengths_equal(
        pf.first_line_indent, cm(level_cfg.first_line_indent_cm)
    ):
        return False
    if pf.space_after is None or not lengths_equal(pf.space_after, pt(level_cfg.space_after_pt)):
        return False
    if bool(pf.page_break_before) != level_cfg.page_break_before:
        return False
    size = pt(level_cfg.size_pt)
    for run in paragraph.runs:
        if not run.text.strip():
            continue
        if run.bold != level_cfg.bold:
            return False
        if run.font.size is None or not lengths_equal(run.font.size, size):
            return False
        if level_cfg.caps and run.text != run.text.upper():
            return False
    return True


def _apply_heading_level(paragraph: Paragraph, level_cfg: HeadingLevel, font: str) -> None:
    size = pt(level_cfg.size_pt)
    for run in paragraph.runs:
        set_run_font(run, font)
        run.font.size = size
        run.bold = level_cfg.bold
        if level_cfg.caps and run.text.strip():
            run.text = run.text.upper()

    pf = paragraph.paragraph_format
    pf.alignment = _ALIGNMENTS[level_cfg.alignment]
    pf.first_line_indent = cm(level_cfg.first_line_indent_cm)
    pf.space_after = pt(level_cfg.space_after_pt)
    pf.page_break_before = level_cfg.page_break_before


class HeadingsRule:
    """Распознавание и оформление заголовков — в т.ч. набранных руками,
    без стилей Word (частый случай в реальных отчётах).

    Уверенные совпадения (>= profile.headings.detect_threshold)
    оформляются по профилю. Пограничные (между _WARN_FLOOR и порогом) —
    не трогаются, но дают предупреждение «похоже на заголовок, проверьте
    вручную». Остальное остаётся обычным текстом. 100%-ное распознавание
    не обещаем — это осознанный компромисс, см. план.
    """

    id = "headings"

    def enabled_for(self, profile: Profile) -> bool:
        return bool(profile.headings.levels)

    def apply(self, context: DocContext) -> list[Change]:
        profile = context.profile
        levels_by_number = {level.level: level for level in profile.headings.levels}
        max_level = max(levels_by_number)

        changes: list[Change] = []
        for index, item in enumerate(iter_block_items(context.document)):
            if not isinstance(item, Paragraph):
                continue
            match = context.heading_matches.get(index)
            if match is None:
                continue

            if match.confidence >= profile.headings.detect_threshold:
                level_cfg = levels_by_number.get(match.level) or levels_by_number[max_level]
                if _heading_matches(item, level_cfg):
                    continue
                _apply_heading_level(item, level_cfg, profile.body.font)
                source = "по стилю" if match.by_style else f"эвристика, уверенность {match.confidence:.2f}"
                changes.append(
                    Change(
                        self.id,
                        "info",
                        f"абзац #{index + 1}",
                        f"оформлен как заголовок уровня {level_cfg.level} ({source})",
                    )
                )
            elif match.confidence >= _WARN_FLOOR and not match.by_style:
                changes.append(
                    Change(
                        self.id,
                        "warning",
                        f"абзац #{index + 1}",
                        f"похож на заголовок (уверенность {match.confidence:.2f}), но не тронут — проверьте вручную",
                    )
                )

        return changes
