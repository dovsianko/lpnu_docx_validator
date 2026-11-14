#!/usr/bin/env python3
"""Сумісний запуск старим способом:

    python docx_format_validator.py робота.docx -c docx_validator_config.yaml -o звіт.txt

Увесь код тепер у пакеті ``lpnu_docx_validator``; див. README.md.
Без аргументів відкривається графічний інтерфейс.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lpnu_docx_validator.cli import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
