from __future__ import annotations

import re
from dataclasses import dataclass

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph

from docx_formatter.core.profile import HeadingsProfile

# "1", "1.2", "1.2.3" в начале абзаца, за которым идёт текст.
_NUMBERING_RE = re.compile(r"^(\d+(?:\.\d+)*)[.)]?\s+\S")
_MAX_HEADING_LENGTH = 120
_STYLE_PREFIXES = ("heading", "заголовок")

# Веса сигналов эвристики — сумма ровно 1.0, чтобы confidence был в [0, 1]
# и сравнивался напрямую с profile.headings.detect_threshold.
_WEIGHT_SHORT = 0.15
_WEIGHT_NO_TRAILING_DOT = 0.10
_WEIGHT_NUMBERED = 0.25
_WEIGHT_EMPHASIZED = 0.20
_WEIGHT_LARGER_FONT = 0.15
_WEIGHT_CENTERED = 0.05
_WEIGHT_KEYWORD = 0.10


@dataclass(frozen=True)
class HeadingMatch:
    level: int
    confidence: float
    by_style: bool


def _is_heading_style(style_name: str) -> bool:
    lowered = style_name.lower()
    return any(lowered.startswith(prefix) for prefix in _STYLE_PREFIXES)


def _style_based_level(style_name: str) -> int:
    digits = "".join(ch for ch in style_name if ch.isdigit())
    return int(digits) if digits else 1


def _numbering_level(text: str) -> int | None:
    match = _NUMBERING_RE.match(text)
    if not match:
        return None
    return match.group(1).count(".") + 1


def _max_run_size_pt(paragraph: Paragraph) -> float | None:
    sizes = [run.font.size.pt for run in paragraph.runs if run.font.size is not None]
    return max(sizes) if sizes else None


def _is_emphasized(paragraph: Paragraph, text: str) -> bool:
    runs_with_text = [run for run in paragraph.runs if run.text.strip()]
    all_bold = bool(runs_with_text) and all(run.bold for run in runs_with_text)
    return all_bold or text == text.upper()


def detect_heading(paragraph: Paragraph, profile: HeadingsProfile, body_size_pt: float) -> HeadingMatch | None:
    """Оценить, похож ли абзац на заголовок.

    None — только для пустых абзацев (там просто нечего оценивать). Для
    всего остального возвращает HeadingMatch с confidence в [0, 1];
    решение, что с этим делать (оформить / предупредить / не трогать),
    принимает вызывающее правило, а не эта функция — у неё нет понятия
    порога.
    """
    text = paragraph.text.strip()
    if not text:
        return None

    style_name = paragraph.style.name if paragraph.style is not None else ""
    if _is_heading_style(style_name):
        return HeadingMatch(level=_style_based_level(style_name), confidence=1.0, by_style=True)

    if profile.trust_styles_only:
        return HeadingMatch(level=1, confidence=0.0, by_style=False)

    numbering_level = _numbering_level(text)
    keyword_hit = text.upper() in {kw.upper() for kw in profile.keywords_level1}
    size = _max_run_size_pt(paragraph)

    score = 0.0
    score += _WEIGHT_SHORT if len(text) <= _MAX_HEADING_LENGTH else 0.0
    score += _WEIGHT_NO_TRAILING_DOT if not text.endswith(".") else 0.0
    score += _WEIGHT_NUMBERED if numbering_level is not None else 0.0
    score += _WEIGHT_EMPHASIZED if _is_emphasized(paragraph, text) else 0.0
    score += _WEIGHT_LARGER_FONT if size is not None and size > body_size_pt else 0.0
    score += _WEIGHT_CENTERED if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER else 0.0
    score += _WEIGHT_KEYWORD if keyword_hit else 0.0

    level = numbering_level if numbering_level is not None else 1
    return HeadingMatch(level=level, confidence=round(score, 4), by_style=False)
