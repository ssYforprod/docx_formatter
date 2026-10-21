from __future__ import annotations

from pathlib import Path

from docx import Document

from docx_formatter.core.profile import Profile
from docx_formatter.core.report import FormatReport
from docx_formatter.core.rules.base import Rule
from docx_formatter.core.rules.body_text import BodyTextRule
from docx_formatter.core.rules.page_setup import PageSetupRule

# Порядок важен: между правилами есть зависимости (например, headings
# в следующих этапах переопределяет часть абзацев, уже приведённых к
# профилю body_text).
RULES: tuple[Rule, ...] = (
    PageSetupRule(),
    BodyTextRule(),
)


def output_path_for(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_formatted{input_path.suffix}")


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

        for rule in self._rules:
            if rule.id in self._profile.rules_disabled:
                continue
            if not rule.enabled_for(self._profile):
                continue
            report.changes.extend(rule.apply(document, self._profile))

        out_path = output_path if output_path is not None else output_path_for(input_path)
        document.save(out_path)
        return out_path, report
