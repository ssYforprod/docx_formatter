from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_formatter.core.profile import ProfileError, load_profile, parse_profile

REPO_ROOT = Path(__file__).resolve().parent.parent


def _valid_profile_dict() -> dict:
    return {
        "schema_version": 1,
        "id": "test_profile",
        "name": "Test Profile",
        "page": {
            "margins_mm": {"top": 20, "right": 10, "bottom": 20, "left": 30},
        },
        "body": {
            "font": "Times New Roman",
            "size_pt": 14,
            "line_spacing": 1.5,
            "alignment": "justify",
            "first_line_indent_cm": 1.25,
            "space_before_pt": 0,
            "space_after_pt": 0,
        },
        "headings": {
            "trust_styles_only": False,
            "detect_threshold": 0.75,
            "levels": [
                {
                    "level": 1,
                    "size_pt": 16,
                    "bold": True,
                    "caps": True,
                    "alignment": "center",
                    "page_break_before": True,
                    "space_after_pt": 12,
                    "first_line_indent_cm": 0,
                },
                {
                    "level": 2,
                    "size_pt": 14,
                    "bold": True,
                    "alignment": "left",
                    "page_break_before": False,
                    "space_after_pt": 6,
                    "first_line_indent_cm": 1.25,
                },
            ],
            "keywords_level1": ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ"],
        },
        "tables": {
            "font": "Times New Roman",
            "size_pt": 12,
            "line_spacing": 1.0,
            "header_bold": True,
            "cell_alignment": "left",
        },
        "captions": {
            "figure": {
                "template": "Рисунок {n} — {text}",
                "alignment": "center",
                "position": "below",
                "renumber": False,
                "numbering": "continuous",
            },
            "table": {
                "template": "Таблица {n} — {text}",
                "alignment": "left",
                "position": "above",
                "renumber": False,
                "numbering": "continuous",
            },
        },
        "page_numbers": {
            "enabled": True,
            "position": "footer_center",
            "skip_first_page": True,
            "font": "Times New Roman",
            "size_pt": 14,
        },
        "toc": {"enabled": False, "levels": "1-3"},
        "rules_disabled": [],
    }


def test_parses_valid_profile() -> None:
    profile, warnings = parse_profile(_valid_profile_dict())

    assert warnings == []
    assert profile.schema_version == 1
    assert profile.id == "test_profile"
    assert profile.body.font == "Times New Roman"
    assert profile.body.size_pt == 14
    assert profile.page.margins_mm.left_mm == 30
    assert len(profile.headings.levels) == 2
    assert profile.headings.levels[0].level == 1
    assert profile.headings.levels[0].caps is True
    assert profile.headings.levels[1].caps is False
    assert profile.captions.figure.renumber is False


def test_requires_id() -> None:
    data = _valid_profile_dict()
    del data["id"]

    with pytest.raises(ProfileError) as exc_info:
        parse_profile(data)

    assert any("id" in error for error in exc_info.value.errors)


def test_rejects_unsupported_schema_version() -> None:
    data = _valid_profile_dict()
    data["schema_version"] = 99

    with pytest.raises(ProfileError) as exc_info:
        parse_profile(data)

    assert any("schema_version" in error for error in exc_info.value.errors)


def test_collects_multiple_errors_at_once() -> None:
    data = _valid_profile_dict()
    data["body"]["size_pt"] = "big"
    data["page"]["margins_mm"]["top"] = -5

    with pytest.raises(ProfileError) as exc_info:
        parse_profile(data)

    errors = exc_info.value.errors
    assert len(errors) >= 2
    assert any("body.size_pt" in error for error in errors)
    assert any("page.margins_mm.top" in error for error in errors)


def test_unknown_top_level_key_is_a_warning_not_an_error() -> None:
    data = _valid_profile_dict()
    data["some_future_field"] = True

    profile, warnings = parse_profile(data)

    assert profile.id == "test_profile"
    assert any("some_future_field" in warning for warning in warnings)


def test_load_profile_wraps_invalid_json(tmp_path: Path) -> None:
    broken = tmp_path / "broken.json"
    broken.write_text("{не json", encoding="utf-8")

    with pytest.raises(ProfileError):
        load_profile(broken)


def test_loads_bundled_gost_profile() -> None:
    path = REPO_ROOT / "profiles" / "gost_7_32_2017.json"

    profile, warnings = load_profile(path)

    assert warnings == []
    assert profile.id == "gost_7_32_2017"
    assert profile.body.font == "Times New Roman"
    assert profile.body.size_pt == 14
    assert profile.body.line_spacing == 1.5
    # Решение из плана: перенумерацию подписей не включаем по умолчанию —
    # ломает ручные ссылки вида «см. рисунок 3» в тексте.
    assert profile.captions.figure.renumber is False
    assert profile.captions.table.renumber is False
    # Решение из плана: TOC выключен по умолчанию, т.к. python-docx не
    # умеет сам посчитать номера страниц — только вставляет поле для Word.
    assert profile.toc.enabled is False
