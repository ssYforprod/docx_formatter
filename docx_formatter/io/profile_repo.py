from __future__ import annotations

import shutil
from dataclasses import dataclass
from pathlib import Path

from docx_formatter.core.profile import ProfileError, load_profile
from docx_formatter.io.paths import resource_root, user_profiles_dir


@dataclass(frozen=True)
class ProfileEntry:
    id: str
    name: str
    path: Path


def bundled_profiles_dir() -> Path:
    return resource_root() / "profiles"


def ensure_user_profiles_seeded(*, bundled_dir: Path | None = None, user_dir: Path | None = None) -> None:
    """Скопировать встроенные профили в пользовательскую папку, если там
    ещё нет профиля с таким же id (существующие пользовательские правки
    не трогаем).

    Сверяем по id из содержимого, а не по имени файла: пользователь мог
    сохранить отредактированный профиль под другим именем файла — иначе
    рядом завелась бы вторая копия с тем же id, и какая из двух "выиграет"
    в списке зависело бы от порядка обхода каталогов.

    Нужно, чтобы профиль было где редактировать руками в onefile-сборке:
    bundled-файлы живут во временной папке распаковки и стираются при
    выходе из программы.
    """
    bundled_dir = bundled_dir if bundled_dir is not None else bundled_profiles_dir()
    user_dir = user_dir if user_dir is not None else user_profiles_dir()
    user_dir.mkdir(parents=True, exist_ok=True)

    if not bundled_dir.is_dir():
        return

    existing_ids = {
        entry.id
        for path in user_dir.glob("*.json")
        if (entry := _try_load_entry(path)) is not None
    }

    for bundled in bundled_dir.glob("*.json"):
        entry = _try_load_entry(bundled)
        if entry is not None and entry.id in existing_ids:
            continue
        target = user_dir / bundled.name
        if not target.exists():
            shutil.copy2(bundled, target)


def _try_load_entry(path: Path) -> ProfileEntry | None:
    try:
        profile, _ = load_profile(path)
    except ProfileError:
        return None
    return ProfileEntry(id=profile.id, name=profile.name, path=path)


def list_profiles(*, bundled_dir: Path | None = None, user_dir: Path | None = None) -> list[ProfileEntry]:
    """Список доступных профилей: встроенные + пользовательские.

    Пользовательский профиль перекрывает встроенный с тем же id (не
    именем файла) — так правка существующего профиля не плодит дубль
    в списке. Профиль, который не проходит валидацию, тихо пропускается
    (не ломает список остальных) — настоящую ошибку пользователь увидит,
    когда попробует применить именно этот профиль.
    """
    bundled_dir = bundled_dir if bundled_dir is not None else bundled_profiles_dir()
    user_dir = user_dir if user_dir is not None else user_profiles_dir()
    ensure_user_profiles_seeded(bundled_dir=bundled_dir, user_dir=user_dir)

    entries: dict[str, ProfileEntry] = {}
    for directory in (bundled_dir, user_dir):
        if not directory.is_dir():
            continue
        for path in sorted(directory.glob("*.json")):
            entry = _try_load_entry(path)
            if entry is not None:
                entries[entry.id] = entry

    return sorted(entries.values(), key=lambda entry: entry.name)
