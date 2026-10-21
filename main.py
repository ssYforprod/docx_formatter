import sys

from docx_formatter.cli import run


def main() -> None:
    run(sys.argv[1:])


if __name__ == "__main__":
    main()
