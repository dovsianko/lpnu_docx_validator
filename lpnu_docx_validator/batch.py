"""Пакетна перевірка: багато файлів/папок → звіти + зведена таблиця."""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Optional

from .checks import validate
from .report import ValidationResult, render_html, render_json, render_text, write_annotated_docx
from .settings import RULE_BY_ID, Settings

FORMATS = ("txt", "html", "json", "docx")
SUFFIX = {"txt": ".report.txt", "html": ".report.html", "json": ".report.json", "docx": ".checked.docx"}


def collect_files(paths: Iterable[Path], recursive: bool = False) -> list[Path]:
    out: list[Path] = []
    seen: set[Path] = set()
    for p in paths:
        p = Path(p)
        if p.is_dir():
            found = p.rglob("*.docx") if recursive else p.glob("*.docx")
            items = sorted(f for f in found if not f.name.startswith("~$") and not f.name.endswith(".checked.docx"))
        else:
            items = [p]
        for f in items:
            key = f.resolve()
            if key not in seen:
                seen.add(key)
                out.append(f)
    return out


def output_path(src: Path, fmt: str, out_dir: Optional[Path]) -> Path:
    base = src.name[:-5] if src.name.lower().endswith(".docx") else src.stem
    return (out_dir or src.parent) / f"{base}{SUFFIX[fmt]}"


@dataclass
class FileOutcome:
    result: ValidationResult
    outputs: dict[str, Path] = field(default_factory=dict)
    write_errors: list[str] = field(default_factory=list)


def write_outputs(res: ValidationResult, formats: Iterable[str], out_dir: Optional[Path] = None,
                  explicit: Optional[dict[str, Path]] = None) -> FileOutcome:
    outcome = FileOutcome(res)
    for fmt in formats:
        path = (explicit or {}).get(fmt) or output_path(res.path, fmt, out_dir)
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            if fmt == "txt":
                path.write_text(render_text(res), encoding="utf-8")
            elif fmt == "html":
                path.write_text(render_html(res), encoding="utf-8")
            elif fmt == "json":
                path.write_text(render_json(res), encoding="utf-8")
            elif fmt == "docx":
                if res.fatal:
                    continue
                write_annotated_docx(res, path)
            outcome.outputs[fmt] = path
        except PermissionError:
            outcome.write_errors.append(f"{path}: файл відкрито в іншій програмі або немає прав на запис")
        except Exception as exc:
            outcome.write_errors.append(f"{path}: {exc}")
    return outcome


def run_batch(files: list[Path], settings: Settings, formats: Iterable[str], out_dir: Optional[Path] = None,
              progress: Optional[Callable[[int, int, Path], None]] = None,
              cancelled: Optional[Callable[[], bool]] = None) -> list[FileOutcome]:
    formats = list(formats)
    outcomes = []
    used: set[str] = set()
    for k, f in enumerate(files, start=1):
        if cancelled and cancelled():
            break
        if progress:
            progress(k, len(files), f)
        res = validate(f, settings)
        explicit = None
        if out_dir is not None:
            # однакові назви з різних папок не мають перезаписувати звіти одна одної
            name = f.name[:-5] if f.name.lower().endswith(".docx") else f.stem
            if name.lower() in used:
                name = f"{f.parent.name}_{name}"
            n, base = 2, name
            while name.lower() in used:
                name, n = f"{base}_{n}", n + 1
            used.add(name.lower())
            explicit = {fmt: out_dir / f"{name}{SUFFIX[fmt]}" for fmt in formats}
        outcomes.append(write_outputs(res, formats, out_dir, explicit))
    return outcomes


def write_summary(outcomes: list[FileOutcome], path: Path) -> Path:
    """Зведена таблиця: файл → кількість помилок/попереджень і найчастіші правила."""
    header = ["Файл", "Помилки", "Попередження", "Інформація", "Статус", "Найчастіші зауваження", "Звіт"]
    rows = []
    for o in outcomes:
        r = o.result
        top = sorted(r.rule_counts().items(), key=lambda kv: -kv[1])[:5]
        top_s = "; ".join(f"{RULE_BY_ID[k].title if k in RULE_BY_ID else k} ×{v}" for k, v in top)
        status = r.fatal or ("є помилки" if r.errors else ("є попередження" if r.warnings else "OK"))
        report = o.outputs.get("html") or o.outputs.get("txt") or o.outputs.get("docx")
        rows.append([r.path.name, r.errors, r.warnings, r.infos, status, top_s, str(report or "")])
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.suffix.lower() == ".xlsx":
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font, PatternFill
        except ImportError:
            path = path.with_suffix(".csv")
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Зведення"
            ws.append(header)
            for c in ws[1]:
                c.font = Font(bold=True)
            red = PatternFill("solid", fgColor="FDE2E1")
            for row in rows:
                ws.append(row)
                if row[1]:
                    ws.cell(ws.max_row, 2).fill = red
            for col, width in zip("ABCDEFG", (40, 10, 14, 12, 18, 70, 50)):
                ws.column_dimensions[col].width = width
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            wb.save(str(path))
            return path
    with path.open("w", encoding="utf-8-sig", newline="") as f:  # BOM — щоб Excel правильно відкрив кирилицю
        w = csv.writer(f, delimiter=";")
        w.writerow(header)
        w.writerows(rows)
    return path
