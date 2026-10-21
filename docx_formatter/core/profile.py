from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

SUPPORTED_SCHEMA_VERSIONS = (1,)

ALIGNMENTS = frozenset({"left", "right", "center", "justify"})
CAPTION_POSITIONS = frozenset({"above", "below"})
CAPTION_NUMBERINGS = frozenset({"continuous"})
PAGE_NUMBER_POSITIONS = frozenset({"footer_center"})


class ProfileError(Exception):
    """Профиль не прошёл валидацию. `errors` — все найденные проблемы разом."""

    def __init__(self, errors: list[str]) -> None:
        super().__init__("; ".join(errors))
        self.errors = errors


@dataclass(frozen=True)
class Margins:
    top_mm: float
    right_mm: float
    bottom_mm: float
    left_mm: float


@dataclass(frozen=True)
class PageProfile:
    margins_mm: Margins


@dataclass(frozen=True)
class BodyProfile:
    font: str
    size_pt: float
    line_spacing: float
    alignment: str
    first_line_indent_cm: float
    space_before_pt: float
    space_after_pt: float


@dataclass(frozen=True)
class HeadingLevel:
    level: int
    size_pt: float
    bold: bool
    alignment: str
    page_break_before: bool
    space_after_pt: float
    first_line_indent_cm: float
    caps: bool = False


@dataclass(frozen=True)
class HeadingsProfile:
    trust_styles_only: bool
    detect_threshold: float
    levels: tuple[HeadingLevel, ...]
    keywords_level1: tuple[str, ...]


@dataclass(frozen=True)
class TablesProfile:
    font: str
    size_pt: float
    line_spacing: float
    header_bold: bool
    cell_alignment: str


@dataclass(frozen=True)
class CaptionStyle:
    template: str
    alignment: str
    position: str
    renumber: bool
    numbering: str


@dataclass(frozen=True)
class CaptionsProfile:
    figure: CaptionStyle
    table: CaptionStyle


@dataclass(frozen=True)
class PageNumbersProfile:
    enabled: bool
    position: str
    skip_first_page: bool
    font: str
    size_pt: float


@dataclass(frozen=True)
class TocProfile:
    enabled: bool
    levels: str


@dataclass(frozen=True)
class Profile:
    schema_version: int
    id: str
    name: str
    page: PageProfile
    body: BodyProfile
    headings: HeadingsProfile
    tables: TablesProfile
    captions: CaptionsProfile
    page_numbers: PageNumbersProfile
    toc: TocProfile
    rules_disabled: tuple[str, ...]


class _Errors:
    """Копилка ошибок/предупреждений — чтобы сообщать обо всех проблемах разом,
    а не останавливаться на первой."""

    def __init__(self) -> None:
        self.errors: list[str] = []
        self.warnings: list[str] = []

    def error(self, path: str, message: str) -> None:
        self.errors.append(f"{path}: {message}")

    def warn(self, path: str, message: str) -> None:
        self.warnings.append(f"{path}: {message}")


def _join(path: str, key: str) -> str:
    return key if not path else f"{path}.{key}"


def _repr_value(value: object) -> str:
    return repr(value)


def _dict(obj: dict, key: str, path: str, errs: _Errors) -> dict:
    if key not in obj:
        return {}
    value = obj[key]
    if not isinstance(value, dict):
        errs.error(_join(path, key), f"ожидался объект, получено {_repr_value(value)}")
        return {}
    return value


def _number(
    obj: dict,
    key: str,
    path: str,
    errs: _Errors,
    *,
    minimum: float | None = None,
    maximum: float | None = None,
    default: float = 0.0,
) -> float:
    if key not in obj:
        return default
    value = obj[key]
    full = _join(path, key)
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        hint = f" {minimum}..{maximum}" if minimum is not None and maximum is not None else ""
        errs.error(full, f"ожидалось число{hint}, получено {_repr_value(value)}")
        return default
    numeric = float(value)
    if minimum is not None and numeric < minimum:
        errs.error(full, f"значение {numeric} меньше минимума {minimum}")
    if maximum is not None and numeric > maximum:
        errs.error(full, f"значение {numeric} больше максимума {maximum}")
    return numeric


def _int(
    obj: dict,
    key: str,
    path: str,
    errs: _Errors,
    *,
    minimum: int | None = None,
    default: int = 0,
) -> int:
    if key not in obj:
        return default
    value = obj[key]
    full = _join(path, key)
    if not isinstance(value, int) or isinstance(value, bool):
        errs.error(full, f"ожидалось целое число, получено {_repr_value(value)}")
        return default
    if minimum is not None and value < minimum:
        errs.error(full, f"значение {value} меньше минимума {minimum}")
    return value


def _boolean(obj: dict, key: str, path: str, errs: _Errors, *, default: bool = False) -> bool:
    if key not in obj:
        return default
    value = obj[key]
    if not isinstance(value, bool):
        errs.error(_join(path, key), f"ожидался bool, получено {_repr_value(value)}")
        return default
    return value


def _string(
    obj: dict,
    key: str,
    path: str,
    errs: _Errors,
    *,
    choices: Iterable[str] | None = None,
    default: str = "",
) -> str:
    if key not in obj:
        return default
    value = obj[key]
    full = _join(path, key)
    if not isinstance(value, str):
        errs.error(full, f"ожидалась строка, получено {_repr_value(value)}")
        return default
    if choices is not None and value not in choices:
        errs.error(full, f"значение {value!r} не входит в допустимые {sorted(choices)}")
    return value


def _required_string(obj: dict, key: str, path: str, errs: _Errors) -> str:
    value = obj.get(key)
    if not isinstance(value, str) or not value:
        errs.error(_join(path, key), f"обязательное поле, ожидалась непустая строка, получено {_repr_value(value)}")
        return ""
    return value


def _string_list(
    obj: dict, key: str, path: str, errs: _Errors, *, default: tuple[str, ...] = ()
) -> tuple[str, ...]:
    if key not in obj:
        return default
    value = obj[key]
    full = _join(path, key)
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        errs.error(full, f"ожидался список строк, получено {_repr_value(value)}")
        return default
    return tuple(value)


def _check_unknown(obj: dict, known: frozenset[str], path: str, errs: _Errors) -> None:
    for key in obj:
        if key not in known:
            errs.warn(_join(path, key), "неизвестный ключ, будет проигнорирован")


def _parse_page(data: dict, errs: _Errors) -> PageProfile:
    page = _dict(data, "page", "", errs)
    margins = _dict(page, "margins_mm", "page", errs)
    mpath = "page.margins_mm"
    top = _number(margins, "top", mpath, errs, minimum=0, maximum=100, default=20.0)
    right = _number(margins, "right", mpath, errs, minimum=0, maximum=100, default=10.0)
    bottom = _number(margins, "bottom", mpath, errs, minimum=0, maximum=100, default=20.0)
    left = _number(margins, "left", mpath, errs, minimum=0, maximum=100, default=30.0)
    _check_unknown(margins, frozenset({"top", "right", "bottom", "left"}), mpath, errs)
    _check_unknown(page, frozenset({"margins_mm"}), "page", errs)
    return PageProfile(margins_mm=Margins(top, right, bottom, left))


def _parse_body(data: dict, errs: _Errors) -> BodyProfile:
    body = _dict(data, "body", "", errs)
    font = _string(body, "font", "body", errs, default="Times New Roman")
    size_pt = _number(body, "size_pt", "body", errs, minimum=6, maximum=72, default=14.0)
    line_spacing = _number(body, "line_spacing", "body", errs, minimum=1.0, maximum=3.0, default=1.0)
    alignment = _string(body, "alignment", "body", errs, choices=ALIGNMENTS, default="left")
    first_line_indent_cm = _number(body, "first_line_indent_cm", "body", errs, minimum=0, maximum=5, default=0.0)
    space_before_pt = _number(body, "space_before_pt", "body", errs, minimum=0, maximum=72, default=0.0)
    space_after_pt = _number(body, "space_after_pt", "body", errs, minimum=0, maximum=72, default=0.0)
    _check_unknown(
        body,
        frozenset(
            {
                "font",
                "size_pt",
                "line_spacing",
                "alignment",
                "first_line_indent_cm",
                "space_before_pt",
                "space_after_pt",
            }
        ),
        "body",
        errs,
    )
    return BodyProfile(
        font, size_pt, line_spacing, alignment, first_line_indent_cm, space_before_pt, space_after_pt
    )


def _parse_heading_level(item: object, index: int, errs: _Errors) -> HeadingLevel:
    path = f"headings.levels[{index}]"
    if not isinstance(item, dict):
        errs.error(path, f"ожидался объект, получено {_repr_value(item)}")
        item = {}
    level = _int(item, "level", path, errs, minimum=1, default=index + 1)
    size_pt = _number(item, "size_pt", path, errs, minimum=6, maximum=72, default=14.0)
    bold = _boolean(item, "bold", path, errs, default=True)
    caps = _boolean(item, "caps", path, errs, default=False)
    alignment = _string(item, "alignment", path, errs, choices=ALIGNMENTS, default="left")
    page_break_before = _boolean(item, "page_break_before", path, errs, default=False)
    space_after_pt = _number(item, "space_after_pt", path, errs, minimum=0, maximum=72, default=0.0)
    first_line_indent_cm = _number(item, "first_line_indent_cm", path, errs, minimum=0, maximum=5, default=0.0)
    _check_unknown(
        item,
        frozenset(
            {
                "level",
                "size_pt",
                "bold",
                "caps",
                "alignment",
                "page_break_before",
                "space_after_pt",
                "first_line_indent_cm",
            }
        ),
        path,
        errs,
    )
    return HeadingLevel(
        level=level,
        size_pt=size_pt,
        bold=bold,
        alignment=alignment,
        page_break_before=page_break_before,
        space_after_pt=space_after_pt,
        first_line_indent_cm=first_line_indent_cm,
        caps=caps,
    )


def _parse_headings(data: dict, errs: _Errors) -> HeadingsProfile:
    headings = _dict(data, "headings", "", errs)
    trust_styles_only = _boolean(headings, "trust_styles_only", "headings", errs, default=False)
    detect_threshold = _number(
        headings, "detect_threshold", "headings", errs, minimum=0.0, maximum=1.0, default=0.75
    )
    raw_levels = headings.get("levels", [])
    if not isinstance(raw_levels, list):
        errs.error("headings.levels", f"ожидался список, получено {_repr_value(raw_levels)}")
        raw_levels = []
    levels = tuple(_parse_heading_level(item, i, errs) for i, item in enumerate(raw_levels))
    keywords_level1 = _string_list(headings, "keywords_level1", "headings", errs, default=())
    _check_unknown(
        headings,
        frozenset({"trust_styles_only", "detect_threshold", "levels", "keywords_level1"}),
        "headings",
        errs,
    )
    return HeadingsProfile(trust_styles_only, detect_threshold, levels, keywords_level1)


def _parse_tables(data: dict, errs: _Errors) -> TablesProfile:
    tables = _dict(data, "tables", "", errs)
    font = _string(tables, "font", "tables", errs, default="Times New Roman")
    size_pt = _number(tables, "size_pt", "tables", errs, minimum=6, maximum=72, default=12.0)
    line_spacing = _number(tables, "line_spacing", "tables", errs, minimum=1.0, maximum=3.0, default=1.0)
    header_bold = _boolean(tables, "header_bold", "tables", errs, default=True)
    cell_alignment = _string(tables, "cell_alignment", "tables", errs, choices=ALIGNMENTS, default="left")
    _check_unknown(
        tables,
        frozenset({"font", "size_pt", "line_spacing", "header_bold", "cell_alignment"}),
        "tables",
        errs,
    )
    return TablesProfile(font, size_pt, line_spacing, header_bold, cell_alignment)


def _parse_caption_style(parent: dict, key: str, errs: _Errors, *, default_position: str) -> CaptionStyle:
    path = f"captions.{key}"
    style = _dict(parent, key, "captions", errs)
    template = _string(style, "template", path, errs, default="{n} — {text}")
    alignment = _string(style, "alignment", path, errs, choices=ALIGNMENTS, default="center")
    position = _string(style, "position", path, errs, choices=CAPTION_POSITIONS, default=default_position)
    renumber = _boolean(style, "renumber", path, errs, default=False)
    numbering = _string(style, "numbering", path, errs, choices=CAPTION_NUMBERINGS, default="continuous")
    _check_unknown(
        style, frozenset({"template", "alignment", "position", "renumber", "numbering"}), path, errs
    )
    return CaptionStyle(template, alignment, position, renumber, numbering)


def _parse_captions(data: dict, errs: _Errors) -> CaptionsProfile:
    captions = _dict(data, "captions", "", errs)
    figure = _parse_caption_style(captions, "figure", errs, default_position="below")
    table = _parse_caption_style(captions, "table", errs, default_position="above")
    _check_unknown(captions, frozenset({"figure", "table"}), "captions", errs)
    return CaptionsProfile(figure, table)


def _parse_page_numbers(data: dict, errs: _Errors) -> PageNumbersProfile:
    page_numbers = _dict(data, "page_numbers", "", errs)
    enabled = _boolean(page_numbers, "enabled", "page_numbers", errs, default=True)
    position = _string(
        page_numbers, "position", "page_numbers", errs, choices=PAGE_NUMBER_POSITIONS, default="footer_center"
    )
    skip_first_page = _boolean(page_numbers, "skip_first_page", "page_numbers", errs, default=True)
    font = _string(page_numbers, "font", "page_numbers", errs, default="Times New Roman")
    size_pt = _number(page_numbers, "size_pt", "page_numbers", errs, minimum=6, maximum=72, default=14.0)
    _check_unknown(
        page_numbers,
        frozenset({"enabled", "position", "skip_first_page", "font", "size_pt"}),
        "page_numbers",
        errs,
    )
    return PageNumbersProfile(enabled, position, skip_first_page, font, size_pt)


def _parse_toc(data: dict, errs: _Errors) -> TocProfile:
    toc = _dict(data, "toc", "", errs)
    enabled = _boolean(toc, "enabled", "toc", errs, default=False)
    levels = _string(toc, "levels", "toc", errs, default="1-3")
    _check_unknown(toc, frozenset({"enabled", "levels"}), "toc", errs)
    return TocProfile(enabled, levels)


_TOP_LEVEL_KEYS = frozenset(
    {
        "schema_version",
        "id",
        "name",
        "page",
        "body",
        "headings",
        "tables",
        "captions",
        "page_numbers",
        "toc",
        "rules_disabled",
    }
)


def parse_profile(data: dict) -> tuple[Profile, list[str]]:
    """Разобрать и провалидировать профиль из уже распарсенного JSON (dict).

    Поднимает ProfileError со всеми найденными ошибками сразу, если профиль
    некорректен. Неизвестные ключи не считаются ошибкой — попадают в warnings.
    """
    if not isinstance(data, dict):
        raise ProfileError([f"profile: ожидался объект, получено {_repr_value(data)}"])

    schema_version = data.get("schema_version")
    if schema_version not in SUPPORTED_SCHEMA_VERSIONS:
        raise ProfileError(
            [
                f"schema_version: неизвестная версия схемы {_repr_value(schema_version)}, "
                f"поддерживаются: {list(SUPPORTED_SCHEMA_VERSIONS)}"
            ]
        )

    errs = _Errors()
    profile_id = _required_string(data, "id", "", errs)
    name = _required_string(data, "name", "", errs)
    page = _parse_page(data, errs)
    body = _parse_body(data, errs)
    headings = _parse_headings(data, errs)
    tables = _parse_tables(data, errs)
    captions = _parse_captions(data, errs)
    page_numbers = _parse_page_numbers(data, errs)
    toc = _parse_toc(data, errs)
    rules_disabled = _string_list(data, "rules_disabled", "", errs, default=())

    _check_unknown(data, _TOP_LEVEL_KEYS, "", errs)

    if errs.errors:
        raise ProfileError(errs.errors)

    profile = Profile(
        schema_version=schema_version,
        id=profile_id,
        name=name,
        page=page,
        body=body,
        headings=headings,
        tables=tables,
        captions=captions,
        page_numbers=page_numbers,
        toc=toc,
        rules_disabled=rules_disabled,
    )
    return profile, errs.warnings


def load_profile(path: Path) -> tuple[Profile, list[str]]:
    """Прочитать и провалидировать профиль из JSON-файла."""
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise ProfileError([f"{path}: не удалось прочитать файл: {exc}"]) from exc

    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        raise ProfileError([f"{path}: некорректный JSON: {exc}"]) from exc

    return parse_profile(data)
