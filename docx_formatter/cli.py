from __future__ import annotations

import argparse
import sys
from pathlib import Path

from docx_formatter.core.pipeline import Formatter
from docx_formatter.core.profile import ProfileError, load_profile

_PROFILES_DIR = Path(__file__).resolve().parent.parent / "profiles"


def _resolve_profile_path(value: str) -> Path:
    path = Path(value)
    if path.is_file():
        return path
    candidate = _PROFILES_DIR / f"{value}.json"
    if candidate.is_file():
        return candidate
    raise SystemExit(f"Профиль не найден: {value!r} (ни как файл, ни как {candidate})")


def run(argv: list[str]) -> None:
    parser = argparse.ArgumentParser(
        prog="docx_formatter", description="Форматирование DOCX-отчёта по профилю оформления"
    )
    parser.add_argument("--file", required=True, type=Path, help="путь к исходному .docx")
    parser.add_argument(
        "--profile", required=True, help="id профиля (например, gost_7_32_2017) или путь к JSON-файлу"
    )
    parser.add_argument("--out", type=Path, default=None, help="путь для результата (по умолчанию *_formatted.docx)")
    args = parser.parse_args(argv)

    profile_path = _resolve_profile_path(args.profile)
    try:
        profile, profile_warnings = load_profile(profile_path)
    except ProfileError as exc:
        for error in exc.errors:
            print(f"Ошибка профиля: {error}", file=sys.stderr)
        raise SystemExit(1) from exc

    for warning in profile_warnings:
        print(f"Предупреждение профиля: {warning}", file=sys.stderr)

    out_path, report = Formatter(profile).run(args.file, output_path=args.out)
    applied = [change for change in report.changes if change.level == "info"]

    print(f"Готово: {out_path}")
    print(f"Изменений: {len(applied)}")
    for warning in report.warnings:
        print(f"Внимание: {warning.scope}: {warning.what}")
