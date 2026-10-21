from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class Change:
    rule_id: str
    level: str  # "info" | "warning"
    scope: str
    what: str


@dataclass
class FormatReport:
    changes: list[Change] = field(default_factory=list)

    @property
    def warnings(self) -> list[Change]:
        return [change for change in self.changes if change.level == "warning"]
