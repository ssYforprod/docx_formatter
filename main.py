import sys

from docx_formatter.cli import run as run_cli


def main() -> None:
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
