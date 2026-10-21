from __future__ import annotations

from dataclasses import dataclass, field

from docx.document import Document

from docx_formatter.core.captions import CaptionMatch
from docx_formatter.core.detect import HeadingMatch
from docx_formatter.core.profile import Profile


@dataclass
class DocContext:
    """Общий контекст одного прогона форматирования.

    heading_matches/caption_matches считаются один раз, до запуска любых
    правил (см. pipeline.py), и используются и body_text (чтобы не
    трогать такие абзацы), и headings/captions (чтобы их оформить) —
    иначе на повторном прогоне body_text увидел бы уже оформленный
    заголовок или подпись, счёл бы их "не телом текста" и откатил
    формат назад, а headings/captions на откаченном абзаце приняли бы
    другое решение — бесконечный пинг-понг изменений вместо
    идемпотентности. toc_paragraph_index — та же защита для абзаца с
    полем оглавления, который TocRule вставляет на предыдущем прогоне:
    без неё body_text реформатировал бы его как обычный текст. Индекс —
    позиция в iter_block_items(document); валиден, пока в пределах
    одного run() ни одно правило не добавляет/удаляет абзацы раньше, чем
    эти индексы использованы (TocRule поэтому идёт в конвейере последним).
    """

    document: Document
    profile: Profile
    heading_matches: dict[int, HeadingMatch] = field(default_factory=dict)
    caption_matches: dict[int, CaptionMatch] = field(default_factory=dict)
    toc_paragraph_index: int | None = None
