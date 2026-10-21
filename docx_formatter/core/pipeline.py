from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.text.paragraph import Paragraph

from docx_formatter.core.captions import find_captions
from docx_formatter.core.context import DocContext
from docx_formatter.core.detect import HeadingMatch, detect_heading
from docx_formatter.core.oxml import iter_block_items, paragraph_has_toc_field
from docx_formatter.core.profile import Profile
from docx_formatter.core.report import FormatReport
from docx_formatter.core.rules.base import Rule
from docx_formatter.core.rules.body_text import BodyTextRule
from docx_formatter.core.rules.figures import CaptionsRule
from docx_formatter.core.rules.headings import HeadingsRule
from docx_formatter.core.rules.page_numbers import PageNumbersRule
from docx_formatter.core.rules.page_setup import PageSetupRule
from docx_formatter.core.rules.tables import TablesRule
from docx_formatter.core.rules.toc import TocRule

# Порядок важен: между правилами есть зависимости — headings
# переопределяет часть абзацев, уже приведённых к профилю body_text
# (body_text их не трогает благодаря заранее посчитанному
# context.heading_matches, см. DocContext); captions правит текст подписи
# уже после того, как body_text/tables задали её шрифт. page_numbers
# работает с футерами секций — не зависит от остальных правил. toc —
# СТРОГО последним: вставляет абзац в начало документа, что сдвинуло бы
# позиции в context.heading_matches/caption_matches для всех, кто идёт
# после него.
RULES: tuple[Rule, ...] = (
    PageSetupRule(),
    BodyTextRule(),
    HeadingsRule(),
    TablesRule(),
    CaptionsRule(),
    PageNumbersRule(),
    TocRule(),
)


def output_path_for(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_formatted{input_path.suffix}")


def _detect_headings(items: list, profile: Profile) -> dict[int, HeadingMatch]:
    matches: dict[int, HeadingMatch] = {}
    for index, item in enumerate(items):
        if isinstance(item, Paragraph):
            match = detect_heading(item, profile.headings, profile.body.size_pt)
            if match is not None:
                matches[index] = match
    return matches


def _detect_toc_paragraph(items: list) -> int | None:
    for index, item in enumerate(items):
        if isinstance(item, Paragraph) and paragraph_has_toc_field(item):
            return index
    return None


class Formatter:
    def __init__(self, profile: Profile, rules: tuple[Rule, ...] = RULES) -> None:
        self._profile = profile
        self._rules = rules

    def run(self, input_path: Path, output_path: Path | None = None) -> tuple[Path, FormatReport]:
        """Отформатировать input_path и сохранить результат в копию.

        Оригинал не открывается на запись ни разу — документ сохраняется
        только в output_path (по умолчанию output_path_for(input_path)),
        поэтому исходный файл физически не может быть изменён этим вызовом.
        """
        document = Document(input_path)
        report = FormatReport()
        items = list(iter_block_items(document))
        context = DocContext(
            document=document,
            profile=self._profile,
            heading_matches=_detect_headings(items, self._profile),
            caption_matches=find_captions(items, self._profile.captions),
            toc_paragraph_index=_detect_toc_paragraph(items),
        )

        for rule in self._rules:
            if rule.id in self._profile.rules_disabled:
                continue
            if not rule.enabled_for(self._profile):
                continue
            report.changes.extend(rule.apply(context))

        out_path = output_path if output_path is not None else output_path_for(input_path)
        document.save(out_path)
        return out_path, report
