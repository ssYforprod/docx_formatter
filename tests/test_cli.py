from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_formatter.cli import _resolve_profile_path


def _write_profile(path: Path, *, profile_id: str, name: str) -> None:
    data = {
        "schema_version": 1,
        "id": profile_id,
        "name": name,
        "page": {"margins_mm": {"top": 20, "right": 10, "bottom": 20, "left": 30}},
    }
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def test_resolves_explicit_file_path(tmp_path: Path) -> None:
    custom = tmp_path / "custom.json"
    _write_profile(custom, profile_id="x", name="X")

    assert _resolve_profile_path(str(custom)) == custom


def test_resolves_bare_id_via_profile_repo(tmp_path: Path) -> None:
    bundled_dir = tmp_path / "bundled"
    user_dir = tmp_path / "user"
    bundled_dir.mkdir()
    _write_profile(bundled_dir / "gost.json", profile_id="gost", name="Встроенный")

    path = _resolve_profile_path("gost", bundled_dir=bundled_dir, user_dir=user_dir)

    # Не привязываемся к тому, что именно вернулось — bundled_dir/gost.json
    # или затравленная копия в user_dir (list_profiles сам копирует
    # встроенные профили при первом обращении, см. profile_repo). Важно,
    # что по id "gost" находится профиль с нужным содержимым.
    assert json.loads(path.read_text(encoding="utf-8"))["name"] == "Встроенный"


def test_user_override_wins_for_bare_id(tmp_path: Path) -> None:
    # Тот самый сценарий, который раньше ломался: пользователь отредактировал
    # профиль через путь, который GUI показывает (%APPDATA%), но у CLI был
    # свой, не связанный с profile_repo способ резолвить id — правка
    # оставалась невидимой для "--profile gost_7_32_2017".
    bundled_dir = tmp_path / "bundled"
    user_dir = tmp_path / "user"
    bundled_dir.mkdir()
    user_dir.mkdir()
    _write_profile(bundled_dir / "gost.json", profile_id="gost", name="Встроенный")
    _write_profile(user_dir / "gost.json", profile_id="gost", name="Мой ГОСТ")

    path = _resolve_profile_path("gost", bundled_dir=bundled_dir, user_dir=user_dir)

    assert path == user_dir / "gost.json"


def test_unknown_id_raises_system_exit(tmp_path: Path) -> None:
    bundled_dir = tmp_path / "bundled"
    user_dir = tmp_path / "user"
    bundled_dir.mkdir()
    user_dir.mkdir()

    with pytest.raises(SystemExit):
        _resolve_profile_path("не_существует_вообще", bundled_dir=bundled_dir, user_dir=user_dir)
