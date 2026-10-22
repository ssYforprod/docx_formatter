import io
import sys

from docx_formatter.cli import run as run_cli


def _guard_frozen_windowed_stdio() -> None:
    """PyInstaller onefile --windowed не выделяет консоль: sys.stdout/
    stderr оказываются None, и обычный print() в cli.py падает с
    AttributeError вместо того, чтобы просто быть не видимым глазу.
    Это не баг CLI-кода — тип сборки решается на уровне exe, а main.py
    всё ещё маршрутизирует туда же логику. Подменяем на no-op writer,
    чтобы --file/--profile через такой exe хотя бы не падал."""
    if sys.stdout is None:
        sys.stdout = io.StringIO()
    if sys.stderr is None:
        sys.stderr = io.StringIO()


def main() -> None:
    if getattr(sys, "frozen", False):
        _guard_frozen_windowed_stdio()

    argv = sys.argv[1:]
    if argv:
        run_cli(argv)
    else:
        # Импорт PySide6 только когда реально нужен GUI — CLI и тесты
        # не должны платить за его загрузку.
        from docx_formatter.gui.app import run as run_gui

        run_gui()


if __name__ == "__main__":
    main()
