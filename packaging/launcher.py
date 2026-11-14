"""Точка входу для PyInstaller: без аргументів — GUI, з аргументами — командний рядок."""

from lpnu_docx_validator.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
