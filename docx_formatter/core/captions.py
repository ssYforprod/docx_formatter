from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Union

from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from docx_formatter.core.profile import CaptionsProfile

_CAPTION_RE = re.compile(
    r"^(?:Рисунок|Рис\.?|Таблица|Табл\.?)\s*(\d+(?:\.\d+)?)\s*[—–-]?\s*(.*)$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class ParsedCaption:
    number: str
    description: str


def parse_caption(text: str) -> ParsedCaption | None:
    """Разобрать подпись вида «Рисунок 1 — текст» / «Таблица 2.1 - текст».

    Слово-префикс (Рисунок/Рис./Таблица/Табл.) не влияет на то, какой
    шаблон применит правило — это решает позиция абзаца (рядом с
    картинкой или с таблицей), а не само слово в тексте.
    """
    match = _CAPTION_RE.match(text.strip())
    if not match:
        return None
    return ParsedCaption(number=match.group(1), description=match.group(2).strip())


def paragraph_has_drawing(paragraph: Paragraph) -> bool:
    """Абзац содержит вставленную картинку (inline-рисунок или VML)."""
    element = paragraph._element
    return (
        element.find(".//" + qn("w:drawing")) is not None
        or element.find(".//" + qn("w:pict")) is not None
    )


@dataclass(frozen=True)
class CaptionMatch:
    style_key: str  # "figure" | "table"
    anchor_index: int
    sequence: int
    parsed: ParsedCaption


BlockItem = Union[Paragraph, Table]


def find_captions(items: list[BlockItem], profile: CaptionsProfile) -> dict[int, CaptionMatch]:
    """Найти подписи к рисункам/таблицам один раз, до запуска любых правил.

    Возвращает {индекс_абзаца_с_подписью: CaptionMatch}. Используется и
    body_text (чтобы не трогать такие абзацы), и CaptionsRule (чтобы их
    оформить) — та же причина, что у heading_matches в DocContext: иначе
    на втором прогоне body_text увидел бы уже отцентрированную подпись,
    счёл бы её "не телом текста" и откатил выравнивание назад, а
    CaptionsRule после этого переставил бы его снова — бесконечный
    пинг-понг изменений вместо идемпотентности.
    """
    matches: dict[int, CaptionMatch] = {}
    figure_count = 0
    table_count = 0

    for index, item in enumerate(items):
        if isinstance(item, Paragraph) and paragraph_has_drawing(item):
            figure_count += 1
            style_key, sequence, style = "figure", figure_count, profile.figure
        elif isinstance(item, Table):
            table_count += 1
            style_key, sequence, style = "table", table_count, profile.table
        else:
            continue

        caption_index = index + 1 if style.position == "below" else index - 1
        if not (0 <= caption_index < len(items)):
            continue
        caption_item = items[caption_index]
        if not isinstance(caption_item, Paragraph):
            continue
        parsed = parse_caption(caption_item.text)
        if parsed is None:
            continue

        matches[caption_index] = CaptionMatch(
            style_key=style_key, anchor_index=index, sequence=sequence, parsed=parsed
        )

    return matches
