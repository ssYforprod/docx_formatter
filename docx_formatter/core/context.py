from __future__ import annotations

from dataclasses import dataclass, field

from docx.document import Document

from docx_formatter.core.detect import HeadingMatch
from docx_formatter.core.profile import Profile


@dataclass
class DocContext:
    """Общий контекст одного прогона форматирования.

    heading_matches считается один раз, до запуска любых правил (см.
    pipeline.py), и используется и body_text (чтобы не трогать абзацы,
    опознанные как заголовки), и headings (чтобы их оформить) — иначе
    на повторном прогоне body_text увидел бы уже оформленный заголовок,
    счёл бы его "не телом текста" и откатил формат назад к обычному
    абзацу, а headings на уже откаченном тексте оценил бы его иначе.
    Индекс — позиция в iter_block_items(document); валиден, пока в
    пределах одного run() ни одно правило не добавляет/удаляет абзацы.
    """

    document: Document
    profile: Profile
    heading_matches: dict[int, HeadingMatch] = field(default_factory=dict)
