from __future__ import annotations

import json
from pathlib import Path

from docx_formatter.io.profile_repo import ensure_user_profiles_seeded, list_profiles


def _write_profile(path: Path, *, profile_id: str, name: str) -> None:
    data = {
        "schema_version": 1,
        "id": profile_id,
        "name": name,
        "page": {"margins_mm": {"top": 20, "right": 10, "bottom": 20, "left": 30}},
    }
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def test_seeding_copies_missing_bundled_profiles(tmp_path: Path) -> None:
    bundled_dir = tmp_path / "bundled"
    user_dir = tmp_path / "user"
    bundled_dir.mkdir()
    _write_profile(bundled_dir / "gost.json", profile_id="gost", name="ГОСТ")

    ensure_user_profiles_seeded(bundled_dir=bundled_dir, user_dir=user_dir)

    assert (user_dir / "gost.json").exists()


def test_seeding_does_not_overwrite_existing_user_profile(tmp_path: Path) -> None:
    bundled_dir = tmp_path / "bundled"
    user_dir = tmp_path / "user"
    bundled_dir.mkdir()
    user_dir.mkdir()
    _write_profile(bundled_dir / "gost.json", profile_id="gost", name="Встроенный")
    _write_profile(user_dir / "gost.json", profile_id="gost", name="Отредактированный пользователем")

    ensure_user_profiles_seeded(bundled_dir=bundled_dir, user_dir=user_dir)

    saved = json.loads((user_dir / "gost.json").read_text(encoding="utf-8"))
    assert saved["name"] == "Отредактированный пользователем"


def test_user_profile_overrides_bundled_with_same_id(tmp_path: Path) -> None:
    bundled_dir = tmp_path / "bundled"
    user_dir = tmp_path / "user"
    bundled_dir.mkdir()
    user_dir.mkdir()
    _write_profile(bundled_dir / "gost.json", profile_id="gost", name="Встроенный")
    _write_profile(user_dir / "custom_gost.json", profile_id="gost", name="Мой ГОСТ")

    entries = list_profiles(bundled_dir=bundled_dir, user_dir=user_dir)

    assert len(entries) == 1
    assert entries[0].name == "Мой ГОСТ"


def test_broken_profile_is_skipped_not_crashed(tmp_path: Path) -> None:
    bundled_dir = tmp_path / "bundled"
    user_dir = tmp_path / "user"
    bundled_dir.mkdir()
    user_dir.mkdir()
    (bundled_dir / "broken.json").write_text("{не json", encoding="utf-8")
    _write_profile(bundled_dir / "gost.json", profile_id="gost", name="ГОСТ")

    entries = list_profiles(bundled_dir=bundled_dir, user_dir=user_dir)

    assert [entry.id for entry in entries] == ["gost"]


def test_list_profiles_sorted_by_name(tmp_path: Path) -> None:
    bundled_dir = tmp_path / "bundled"
    user_dir = tmp_path / "user"
    bundled_dir.mkdir()
    user_dir.mkdir()
    _write_profile(bundled_dir / "b.json", profile_id="b", name="Бета")
    _write_profile(bundled_dir / "a.json", profile_id="a", name="Альфа")

    entries = list_profiles(bundled_dir=bundled_dir, user_dir=user_dir)

    assert [entry.name for entry in entries] == ["Альфа", "Бета"]
