"""Командний рядок.

    docx-validator робота.docx                      # як і раніше: звіт поруч із файлом
    docx-validator робота.docx -c config.yaml -o report.txt
    docx-validator папка_з_роботами/ --html --docx --summary зведення.xlsx
    docx-validator init-config -o мій_профіль.yaml
    docx-validator list-rules | list-profiles | validate-config файл.yaml | gui
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional

from . import __version__
from .settings import (ConfigError, GROUP_BY_ID, RULES, default_settings, list_profiles, load_profile,
                       load_settings_file, resolve_settings, rules_markdown, settings_to_yaml)

COMMANDS = {"check", "init-config", "list-rules", "list-profiles", "validate-config", "gui"}


def _console_utf8() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        except Exception:
            pass


def _add_settings_args(p: argparse.ArgumentParser) -> None:
    g = p.add_mutually_exclusive_group()
    g.add_argument("-c", "--config", type=Path, help="YAML/JSON файл налаштувань")
    g.add_argument("-p", "--profile", help="вбудований або збережений профіль (див. list-profiles)")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="docx-validator", description="Перевірка оформлення студентських робіт (.docx)")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = parser.add_subparsers(dest="command")

    c = sub.add_parser("check", help="перевірити файли або папки (команда за замовчуванням)")
    c.add_argument("paths", nargs="+", type=Path, help=".docx файли або папки")
    _add_settings_args(c)
    c.add_argument("-o", "--output", type=Path,
                   help="для одного файлу — шлях до звіту; для кількох — папка для звітів")
    c.add_argument("--html", action="store_true", help="також HTML-звіт із фільтрами")
    c.add_argument("--json", action="store_true", help="також JSON (для інтеграцій)")
    c.add_argument("--docx", action="store_true", help="копія DOCX з примітками на проблемних місцях")
    c.add_argument("--no-txt", action="store_true", help="не створювати TXT-звіт")
    c.add_argument("--summary", type=Path, help="зведена таблиця по всіх файлах (.xlsx або .csv)")
    c.add_argument("-r", "--recursive", action="store_true", help="шукати .docx у вкладених папках")
    c.add_argument("-q", "--quiet", action="store_true", help="менше виводу в консоль")

    i = sub.add_parser("init-config", help="створити файл налаштувань із коментарями")
    _add_settings_args(i)
    i.add_argument("-o", "--output", type=Path, default=Path("docx_validator_config.yaml"))
    i.add_argument("--force", action="store_true", help="перезаписати, якщо файл існує")

    lr = sub.add_parser("list-rules", help="список усіх правил")
    lr.add_argument("--markdown", action="store_true")
    sub.add_parser("list-profiles", help="доступні профілі")
    v = sub.add_parser("validate-config", help="перевірити файл налаштувань на помилки")
    v.add_argument("file", type=Path)
    sub.add_parser("gui", help="графічний інтерфейс")
    return parser


def _normalize_argv(argv: list[str]) -> list[str]:
    """Сумісність зі старим запуском без підкоманди: `робота.docx -c cfg -o out`."""
    if not argv or argv[0] in COMMANDS or argv[0] in {"-h", "--help", "--version"}:
        return argv
    return ["check", *argv]


def cmd_check(args: argparse.Namespace) -> int:
    from .batch import collect_files, run_batch, write_outputs, write_summary
    from .checks import validate

    settings = resolve_settings(args.config, args.profile)
    files = collect_files(args.paths, args.recursive)
    missing = [str(p) for p in args.paths if not p.exists()]
    if missing:
        print("Не знайдено: " + ", ".join(missing), file=sys.stderr)
        if not files or all(not f.exists() for f in files):
            return 2
    files = [f for f in files if f.exists()]
    if not files:
        print("Не знайдено жодного .docx файлу.", file=sys.stderr)
        return 2
    formats = [f for f, on in (("txt", not args.no_txt), ("html", args.html), ("json", args.json), ("docx", args.docx)) if on]
    if not formats:
        formats = ["txt"]
    if not args.quiet:
        print(f"Профіль: {settings.name}  ({settings.source})")

    single = len(files) == 1 and not args.paths[0].is_dir()
    if single:
        explicit = {}
        out_dir = None
        if args.output is not None:
            if args.output.suffix.lower() in {".txt", ".html", ".json", ".docx"}:
                explicit[args.output.suffix.lower().lstrip(".")] = args.output
                out_dir = args.output.parent
            else:
                out_dir = args.output
        outcomes = [write_outputs(validate(files[0], settings), formats, out_dir, explicit)]
    else:
        def progress(k: int, n: int, f: Path) -> None:
            if not args.quiet:
                print(f"[{k}/{n}] {f.name}", flush=True)
        outcomes = run_batch(files, settings, formats, args.output, progress)

    rc = 0
    for o in outcomes:
        r = o.result
        if r.fatal:
            print(f"  ✗ {r.path.name}: {r.fatal}", file=sys.stderr)
            rc = 2
            continue
        if r.errors:
            rc = max(rc, 1)
        if not args.quiet or single:
            print(f"  {r.path.name}: помилок {r.errors}, попереджень {r.warnings}")
            for path in o.outputs.values():
                print(f"    Звіт збережено: {path}")
        for e in o.write_errors:
            print(f"    Не вдалося записати {e}", file=sys.stderr)
    if args.summary:
        path = write_summary(outcomes, args.summary)
        print(f"Зведення: {path}")
    return rc


def cmd_init_config(args: argparse.Namespace) -> int:
    if args.output.exists() and not args.force:
        print(f"Файл {args.output} уже існує. Додайте --force, щоб перезаписати.", file=sys.stderr)
        return 2
    s = load_settings_file(args.config) if args.config else (load_profile(args.profile) if args.profile else default_settings())
    s.extends = None
    args.output.write_text(settings_to_yaml(s), encoding="utf-8")
    print(f"Створено {args.output}. Відредагуйте його або відкрийте в GUI.")
    return 0


def cmd_list_rules(args: argparse.Namespace) -> int:
    if args.markdown:
        print(rules_markdown(), end="")
        return 0
    group = None
    for r in RULES:
        if r.group != group:
            group = r.group
            print(f"\n{GROUP_BY_ID[group].title}  [checks.{group}]")
        print(f"  {r.id:<30} {r.severity:<8} {r.title}")
    return 0


def cmd_list_profiles(_: argparse.Namespace) -> int:
    for p in list_profiles():
        kind = "вбудований" if p.builtin else "користувацький"
        print(f"{p.id:<24} {p.name}  ({kind})\n{'':<24} {p.description}\n{'':<24} {p.path}")
    return 0


def cmd_validate_config(args: argparse.Namespace) -> int:
    s = load_settings_file(args.file)
    print(f"Конфіг коректний: {args.file}\nПрофіль: {s.name}")
    return 0


def main(argv: Optional[list[str]] = None) -> int:
    _console_utf8()
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv:
        return cmd_gui()
    args = build_parser().parse_args(_normalize_argv(argv))
    try:
        return {
            "check": cmd_check,
            "init-config": cmd_init_config,
            "list-rules": cmd_list_rules,
            "list-profiles": cmd_list_profiles,
            "validate-config": cmd_validate_config,
            "gui": lambda a: cmd_gui(),
        }[args.command](args)
    except ConfigError as exc:
        print(str(exc), file=sys.stderr)
        return 2


def cmd_gui() -> int:
    try:
        from .gui import run
    except ImportError as exc:
        print(f"Не вдалося запустити GUI ({exc}). Для командного рядка див. --help.", file=sys.stderr)
        return 2
    run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
