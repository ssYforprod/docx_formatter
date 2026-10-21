from __future__ import annotations

from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

from docx_formatter.core.context import DocContext
from docx_formatter.core.oxml import (
    add_toc_field,
    has_update_fields_on_open,
    iter_block_items,
    paragraph_has_toc_field,
    set_update_fields_on_open,
)
from docx_formatter.core.profile import Profile
from docx_formatter.core.report import Change


class TocRule:
    """Оглавление: поле TOC в начале документа + автообновление полей
    при открытии (w:updateFields в settings.xml).

    python-docx не умеет сам посчитать номера страниц (нет движка
    вёрстки) — здесь только поле-заготовка; реальные записи оглавления
    соберёт Word при обновлении. Выключено по умолчанию
    (profile.toc.enabled=false) — решение из плана.

    ДОЛЖНО идти последним в конвейере: вставляет новый абзац в начало
    документа, что сдвигает все позиции в iter_block_items(). Остальные
    правила уже отработали к этому моменту на индексах, посчитанных в
    начале Formatter.run() (см. DocContext) — если TOC вставить раньше,
    эти индексы станут неверными для правил, которые идут после него.
    """

    id = "toc"

    def enabled_for(self, profile: Profile) -> bool:
        return profile.toc.enabled

    def apply(self, context: DocContext) -> list[Change]:
        document = context.document
        changes: list[Change] = []

        if not has_update_fields_on_open(document):
            set_update_fields_on_open(document)
            changes.append(
                Change(self.id, "info", "параметры документа", "включено автообновление полей при открытии")
            )

        if not self._has_toc_field(document):
            paragraph_element = OxmlElement("w:p")
            document.element.body.insert(0, paragraph_element)
            paragraph = Paragraph(paragraph_element, document)
            add_toc_field(paragraph, context.profile.toc.levels)
            changes.append(
                Change(
                    self.id,
                    "info",
                    "начало документа",
                    "вставлено поле оглавления (TOC) — соберётся при обновлении полей в Word",
                )
            )

        return changes

    def _has_toc_field(self, document) -> bool:
        return any(
            isinstance(item, Paragraph) and paragraph_has_toc_field(item)
            for item in iter_block_items(document)
        )
