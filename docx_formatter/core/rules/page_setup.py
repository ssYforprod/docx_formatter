from __future__ import annotations

from docx_formatter.core.context import DocContext
from docx_formatter.core.profile import Profile
from docx_formatter.core.report import Change
from docx_formatter.core.units import lengths_equal, mm

_MARGIN_ATTRS = (
    ("top_margin", "top_mm"),
    ("right_margin", "right_mm"),
    ("bottom_margin", "bottom_mm"),
    ("left_margin", "left_mm"),
)


class PageSetupRule:
    """Поля страницы — для всех секций документа, не только первой
    (документы с несколькими секциями иначе правятся лишь частично)."""

    id = "page_setup"

    def enabled_for(self, profile: Profile) -> bool:
        return True

    def apply(self, context: DocContext) -> list[Change]:
        changes: list[Change] = []
        margins = context.profile.page.margins_mm

        for index, section in enumerate(context.document.sections):
            changed = False
            for section_attr, margin_attr in _MARGIN_ATTRS:
                desired = mm(getattr(margins, margin_attr))
                current = getattr(section, section_attr)
                if current is None or not lengths_equal(current, desired):
                    setattr(section, section_attr, desired)
                    changed = True
            if changed:
                changes.append(
                    Change(self.id, "info", f"раздел #{index + 1}", "установлены поля страницы")
                )
        return changes
