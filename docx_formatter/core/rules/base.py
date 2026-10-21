from __future__ import annotations

from typing import Protocol

from docx_formatter.core.context import DocContext
from docx_formatter.core.profile import Profile
from docx_formatter.core.report import Change


class Rule(Protocol):
    id: str

    def enabled_for(self, profile: Profile) -> bool:
        """Нужно ли вообще применять правило к этому профилю (не считая
        отключения через profile.rules_disabled — это проверяет конвейер)."""
        ...

    def apply(self, context: DocContext) -> list[Change]:
        """Применить правило, вернув только фактически внесённые изменения —
        конвейер полагается на это для идемпотентности отчёта."""
        ...
