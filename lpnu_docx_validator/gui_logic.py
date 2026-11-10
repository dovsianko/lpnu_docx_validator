"""Логіка GUI без tkinter (щоб її можна було тестувати): форма ↔ налаштування, фільтри, стан."""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable

from .settings import (ALIGN_NAMES_UA, GROUPS, PARAMS, RULES, SEVERITIES, Param, Settings, coerce_severity,
                       coerce_value, user_profiles_dir)

SEV_LABELS = {"ERROR": "Помилка", "WARNING": "Попередження", "INFO": "Інформація", "OFF": "Вимкнено"}
SEV_FROM_LABEL = {v: k for k, v in SEV_LABELS.items()}
POSITION_LABELS = {"TOP": "у верхньому колонтитулі", "BOTTOM": "у нижньому колонтитулі", "ANY": "будь-де"}


def choice_labels(param: Param) -> dict[str, str]:
    if param.key.endswith("position"):
        return {c: POSITION_LABELS.get(c, c) for c in param.choices}
    if "alignment" in param.key:
        return {c: ALIGN_NAMES_UA.get(c, c) for c in param.choices}
    return {c: c for c in param.choices}


def format_value(param: Param, value: Any) -> Any:
    """Значення для віджета: bool → bool, choice → підпис, список → рядки, число → '2.5'."""
    if param.type == "bool":
        return bool(value)
    if param.type == "choice":
        return choice_labels(param).get(value, value)
    if param.type in {"list", "regex_list", "float_list"}:
        return "\n".join(f"{x:g}" if isinstance(x, float) else str(x) for x in (value or []))
    if value is None:
        return ""
    if isinstance(value, float):
        return f"{value:g}"
    return str(value)


def parse_widget_value(param: Param, raw: Any) -> Any:
    if param.type == "choice" and isinstance(raw, str):
        back = {v: k for k, v in choice_labels(param).items()}
        raw = back.get(raw, raw)
    return coerce_value(param, raw)


def apply_form(base: Settings, values: dict[str, Any], severities: dict[str, str],
               checks: dict[str, bool]) -> tuple[Settings, list[tuple[str, str]]]:
    """Збирає Settings із значень форми. Повертає (settings, [(ключ, помилка), ...])."""
    s = base.copy()
    errors: list[tuple[str, str]] = []
    by_key = {p.key: p for p in PARAMS}
    for key, raw in values.items():
        p = by_key[key]
        try:
            s.values[key] = parse_widget_value(p, raw)
        except ValueError as exc:
            errors.append((key, f"{p.label}: {exc}"))
    for rid, raw in severities.items():
        try:
            s.severities[rid] = coerce_severity(SEV_FROM_LABEL.get(raw, raw))
        except ValueError as exc:
            errors.append((rid, str(exc)))
    s.checks.update(checks)
    if s.values.get("page_count.min_pages", 0) > s.values.get("page_count.max_pages", 10**6):
        errors.append(("page_count.min_pages", "мінімальна кількість сторінок більша за максимальну"))
    return s, errors


def params_for(section: str, advanced: bool) -> list[Param]:
    return [p for p in PARAMS if p.section == section and (advanced or not p.advanced)]


def sections_in_order() -> list[str]:
    out: list[str] = []
    for p in PARAMS:
        if p.section not in out:
            out.append(p.section)
    return out


def rules_by_group() -> list[tuple[Any, list[Any]]]:
    return [(g, [r for r in RULES if r.group == g.id]) for g in GROUPS]


def filter_issues(issues: Iterable[Any], severities: set[str], query: str = "") -> list[Any]:
    q = query.strip().lower()
    out = []
    for i in issues:
        if i.severity not in severities:
            continue
        if q and q not in f"{i.message} {i.location} {i.rule} {' '.join(i.details)}".lower():
            continue
        out.append(i)
    return out


def search_text(issue: Any) -> str:
    """Фрагмент для Ctrl+F у Word."""
    m = re.search(r"Пошук у Word: «(.+?)…?»", issue.location or "")
    return m.group(1) if m else ""


_TRANSLIT = dict(zip("абвгґдеєжзиіїйклмнопрстуфхцчшщьюя",
                     ["a", "b", "v", "h", "g", "d", "e", "ie", "zh", "z", "y", "i", "i", "i", "k", "l", "m", "n", "o",
                      "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", "sh", "shch", "", "iu", "ia"]))


def slugify(name: str) -> str:
    s = "".join(_TRANSLIT.get(c, c) for c in name.lower())
    s = re.sub(r"[^a-z0-9]+", "_", s).strip("_")
    return s[:50] or "profile"


def user_profile_path(name: str) -> Path:
    base = user_profiles_dir() / f"{slugify(name)}.yaml"
    path, k = base, 2
    while path.exists():
        path = base.with_name(f"{base.stem}_{k}.yaml")
        k += 1
    return path


# ---------------------------------------------------------------------------
# Стан між запусками
# ---------------------------------------------------------------------------

def state_path() -> Path:
    return user_profiles_dir().parent / "gui_state.json"


def load_state() -> dict[str, Any]:
    try:
        return json.loads(state_path().read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_state(state: dict[str, Any]) -> None:
    try:
        p = state_path()
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def open_path(path: Path) -> None:
    """Відкрити файл/папку програмою за замовчуванням."""
    path = Path(path)
    if sys.platform.startswith("win"):
        os.startfile(str(path))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def test_regex(pattern: str, sample: str) -> str:
    try:
        rx = re.compile(pattern)
    except re.error as exc:
        return f"Помилка у виразі: {exc}"
    m = rx.search(sample)
    if not m:
        return "Не збігається."
    groups = {k: v for k, v in m.groupdict().items() if v is not None}
    extra = "; групи: " + ", ".join(f"{k}={v!r}" for k, v in groups.items()) if groups else ""
    return f"Збігається: «{m.group(0)}»{extra}"


SEVERITY_CHOICES = [SEV_LABELS[s] for s in SEVERITIES]
