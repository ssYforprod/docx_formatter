from __future__ import annotations

import os
import sys
from pathlib import Path


def resource_root() -> Path:
    """Корень для файлов, упакованных вместе с приложением (profiles/).

    В PyInstaller onefile это временная папка распаковки (sys._MEIPASS,
    стирается при выходе из программы); при запуске из исходников —
    корень репозитория.
    """
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass is not None:
        return Path(meipass)
    return Path(__file__).resolve().parent.parent.parent


def user_profiles_dir() -> Path:
    """Папка для профилей, которые правит пользователь руками.

    Переживает обновления и переустановку приложения — bundled-профили
    внутри onefile-сборки стираются при выходе из распакованного temp,
    поэтому редактируемая копия должна жить отдельно (см. profile_repo).
    """
    base = os.environ.get("APPDATA")
    root = Path(base) if base else Path.home() / ".config"
    directory = root / "DocxFormatter" / "profiles"
    directory.mkdir(parents=True, exist_ok=True)
    return directory
