"""Збір повідомлень і формування звітів: TXT, HTML, JSON, DOCX із примітками."""

from __future__ import annotations

import html
import json
from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from . import __version__
from .docmodel import Block, block_location
from .settings import GROUP_BY_ID, RULE_BY_ID, Settings

SEV_ORDER = ("ERROR", "WARNING", "INFO")
SEV_UA = {"ERROR": "Помилки", "WARNING": "Попередження", "INFO": "Інформація"}


@dataclass
class Issue:
    rule: str
    severity: str
    message: str
    location: str = ""
    details: list[str] = field(default_factory=list)
    block: Optional[Block] = None

    @property
    def rule_title(self) -> str:
        r = RULE_BY_ID.get(self.rule)
        return r.title if r else self.rule

    @property
    def group(self) -> str:
        r = RULE_BY_ID.get(self.rule)
        return r.group if r else ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "rule": self.rule,
            "rule_title": self.rule_title,
            "group": self.group,
            "severity": self.severity,
            "message": self.message,
            "location": self.location,
            "details": self.details,
            "paragraph_no": self.block.paragraph_no if self.block else None,
            "search_text": self.block.text[:120] if self.block and self.block.text else None,
        }


class Reporter:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.issues: list[Issue] = []

    def add(self, rule: str, message: str, block: Optional[Block] = None, location: Optional[str] = None,
            details: Optional[list[str]] = None, label: str = "", severity: Optional[str] = None) -> None:
        sev = (severity or self.settings.sev(rule)).upper()
        if sev == "OFF":
            return
        if sev not in SEV_ORDER:
            sev = "WARNING"
        if location is None:
            location = block_location(block, label) if block is not None else ""
        self.issues.append(Issue(rule, sev, message, location, list(details or []), block))

    def info(self, message: str, block: Optional[Block] = None, location: Optional[str] = None, rule: str = "info") -> None:
        if location is None:
            location = block_location(block) if block is not None else ""
        self.issues.append(Issue(rule, "INFO", message, location, [], block))


@dataclass
class ValidationResult:
    path: Path
    settings: Settings
    issues: list[Issue]
    fatal: Optional[str] = None
    model: Any = None

    def count(self, sev: str) -> int:
        return sum(1 for i in self.issues if i.severity == sev)

    @property
    def errors(self) -> int:
        return self.count("ERROR")

    @property
    def warnings(self) -> int:
        return self.count("WARNING")

    @property
    def infos(self) -> int:
        return self.count("INFO")

    def visible(self) -> list[Issue]:
        show_info = self.settings["report.show_info"]
        return [i for i in self.issues if show_info or i.severity != "INFO"]

    def grouped(self) -> "OrderedDict[str, list[tuple[Issue, int]]]":
        """По рівнях; у межах рівня — обмеження кількості однакових правил.
        Повертає пари (issue, hidden_after) де hidden_after > 0 означає «і ще N»."""
        cap = int(self.settings["report.max_per_rule"] or 0)
        out: "OrderedDict[str, list[tuple[Issue, int]]]" = OrderedDict()
        for sev in SEV_ORDER:
            items = [i for i in self.visible() if i.severity == sev]
            if not items:
                continue
            by_rule: "OrderedDict[str, list[Issue]]" = OrderedDict()
            for i in items:
                by_rule.setdefault(i.rule, []).append(i)
            rows: list[tuple[Issue, int]] = []
            for _, lst in by_rule.items():
                shown = lst if cap <= 0 else lst[:cap]
                hidden = len(lst) - len(shown)
                for k, i in enumerate(shown):
                    rows.append((i, hidden if k == len(shown) - 1 else 0))
            out[sev] = rows
        return out

    def rule_counts(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for i in self.issues:
            if i.severity != "INFO":
                counts[i.rule] = counts.get(i.rule, 0) + 1
        return counts


# ---------------------------------------------------------------------------
# TXT
# ---------------------------------------------------------------------------

def render_text(res: ValidationResult) -> str:
    lines = [
        "ЗВІТ ПЕРЕВІРКИ ОФОРМЛЕННЯ DOCX",
        "=" * 44,
        f"Версія валідатора: {__version__}",
        f"Файл: {res.path.name}",
        f"Профіль: {res.settings.name}",
        f"Налаштування: {res.settings.source}",
        f"Дата перевірки: {datetime.now():%Y-%m-%d %H:%M}",
        f"Помилки: {res.errors}",
        f"Попередження: {res.warnings}",
        f"Інформаційні повідомлення: {res.infos}",
        "",
    ]
    if res.fatal:
        lines.append(f"[ПОМИЛКА] {res.fatal}")
        return "\n".join(lines) + "\n"
    if not any(i.severity != "INFO" for i in res.issues):
        lines.append("[OK] Формальних проблем не знайдено.")
        lines.append("")
    for sev, rows in res.grouped().items():
        lines.append(f"{sev} — {SEV_UA[sev]}")
        lines.append("-" * 44)
        for issue, hidden in rows:
            tag = f"  ({issue.rule})" if issue.rule != "info" else ""
            lines.append(f"[{issue.severity}] {issue.message}{tag}")
            loc_lines = issue.location.splitlines()
            for k, ln in enumerate(loc_lines):
                lines.append(("    Місце: " if k == 0 else "           ") + ln)
            for d in issue.details:
                lines.append(f"    {d}")
            if hidden:
                lines.append(f"    … і ще {hidden} подібних повідомлень правила «{issue.rule_title}».")
            lines.append("")
    lines.append("Вимкнути правило: у конфігу в секції severities поставте <правило>: OFF")
    lines.append("(або у GUI: вкладка «Правила»).")
    return "\n".join(lines).rstrip() + "\n"


# ---------------------------------------------------------------------------
# JSON
# ---------------------------------------------------------------------------

def render_json(res: ValidationResult) -> str:
    data = {
        "validator_version": __version__,
        "file": str(res.path),
        "profile": res.settings.name,
        "settings_source": res.settings.source,
        "errors": res.errors,
        "warnings": res.warnings,
        "infos": res.infos,
        "fatal": res.fatal,
        "issues": [i.to_dict() for i in res.issues],
    }
    return json.dumps(data, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------

_CSS = """
:root{--err:#b42318;--warn:#b54708;--info:#475467;--bg:#f8f9fb;--card:#fff;--line:#e4e7ec;--ink:#101828;--mute:#667085}
*{box-sizing:border-box}body{margin:0;font:15px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;background:var(--bg);color:var(--ink)}
header{padding:24px 32px;background:var(--card);border-bottom:1px solid var(--line)}
h1{margin:0 0 4px;font-size:20px}.meta{color:var(--mute);font-size:13px}
.summary{display:flex;gap:12px;margin-top:16px;flex-wrap:wrap}
.pill{padding:6px 14px;border-radius:999px;border:1px solid var(--line);background:var(--card);cursor:pointer;font-weight:600;user-select:none}
.pill input{margin-right:6px}.pill.e{color:var(--err)}.pill.w{color:var(--warn)}.pill.i{color:var(--info)}
main{padding:20px 32px;max-width:1100px}
input[type=search]{width:100%;max-width:420px;padding:8px 12px;border:1px solid var(--line);border-radius:8px;font:inherit}
.issue{background:var(--card);border:1px solid var(--line);border-left:4px solid var(--info);border-radius:8px;padding:12px 16px;margin:10px 0}
.issue.ERROR{border-left-color:var(--err)}.issue.WARNING{border-left-color:var(--warn)}
.msg{font-weight:600}.rule{color:var(--mute);font-size:12px;font-family:ui-monospace,Consolas,monospace}
.loc{white-space:pre-wrap;color:var(--mute);font-size:13px;margin-top:6px}
.find{background:#fff7d6;padding:1px 4px;border-radius:4px;cursor:copy}
.more{font-size:13px;color:var(--mute);margin-top:6px}
.ok{padding:16px;background:#ecfdf3;border:1px solid #abefc6;border-radius:8px}
"""

_JS = """
const boxes=[...document.querySelectorAll('.pill input')],q=document.getElementById('q');
function f(){const on=new Set(boxes.filter(b=>b.checked).map(b=>b.value)),t=q.value.toLowerCase();
document.querySelectorAll('.issue').forEach(e=>{e.style.display=on.has(e.dataset.sev)&&(!t||e.textContent.toLowerCase().includes(t))?'':'none'})}
boxes.forEach(b=>b.addEventListener('change',f));q.addEventListener('input',f);
document.querySelectorAll('.find').forEach(e=>e.addEventListener('click',()=>{navigator.clipboard&&navigator.clipboard.writeText(e.dataset.t)}));
"""


def render_html(res: ValidationResult) -> str:
    e = html.escape
    parts = [
        "<!DOCTYPE html><html lang='uk'><head><meta charset='utf-8'>",
        f"<title>Звіт: {e(res.path.name)}</title><meta name='viewport' content='width=device-width,initial-scale=1'>",
        f"<style>{_CSS}</style></head><body><header>",
        f"<h1>Перевірка оформлення: {e(res.path.name)}</h1>",
        f"<div class='meta'>Профіль: {e(res.settings.name)} · Валідатор {e(__version__)} · {datetime.now():%Y-%m-%d %H:%M}</div>",
        "<div class='summary'>",
        f"<label class='pill e'><input type='checkbox' value='ERROR' checked>Помилки: {res.errors}</label>",
        f"<label class='pill w'><input type='checkbox' value='WARNING' checked>Попередження: {res.warnings}</label>",
        f"<label class='pill i'><input type='checkbox' value='INFO'>Інформація: {res.infos}</label>",
        "</div></header><main>",
        "<input id='q' type='search' placeholder='Пошук у звіті…'>",
    ]
    if res.fatal:
        parts.append(f"<div class='issue ERROR' data-sev='ERROR'><div class='msg'>{e(res.fatal)}</div></div>")
    elif not any(i.severity != "INFO" for i in res.issues):
        parts.append("<p class='ok'>Формальних проблем не знайдено.</p>")
    for sev, rows in res.grouped().items():
        for issue, hidden in rows:
            loc = e(issue.location)
            if issue.block is not None and issue.block.text:
                t = issue.block.text[:89]
                loc = loc.replace(e(f"«{t}"), f"«<span class='find' title='Натисніть, щоб скопіювати' data-t='{e(t, quote=True)}'>{e(t)}</span>", 1)
            style = " style='display:none'" if sev == "INFO" else ""
            parts.append(
                f"<div class='issue {sev}' data-sev='{sev}'{style}><div class='msg'>{e(issue.message[:1].upper() + issue.message[1:])}</div>"
                f"<div class='rule'>{sev} · {e(issue.rule)} · {e(GROUP_BY_ID[issue.group].title if issue.group in GROUP_BY_ID else '')}</div>"
                + (f"<div class='loc'>{loc}</div>" if loc else "")
                + "".join(f"<div class='loc'>{e(d)}</div>" for d in issue.details)
                + (f"<div class='more'>… і ще {hidden} подібних повідомлень</div>" if hidden else "")
                + "</div>"
            )
    parts.append(f"</main><script>{_JS}</script></body></html>")
    return "".join(parts)


# ---------------------------------------------------------------------------
# DOCX із примітками
# ---------------------------------------------------------------------------

def write_annotated_docx(res: ValidationResult, out_path: Path) -> int:
    """Копія документа з примітками Word на проблемних місцях. Повертає кількість приміток."""
    from docx import Document
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run

    from .docmodel import iter_runs

    # Відкриваємо документ заново, щоб не змінювати модель, з якою працювали перевірки.
    doc = Document(str(res.path))
    body_ps = list(doc.element.body.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"))
    author = res.settings["report.comment_author"]
    cap = int(res.settings["report.max_per_rule"] or 0)
    per_rule: dict[str, int] = {}
    totals: dict[str, int] = {}
    for i in res.issues:
        if i.severity != "INFO":
            totals[i.rule] = totals.get(i.rule, 0) + 1

    # Індекс абзаців у порядку обходу: в обох документах однаковий.
    orig_ps = []
    if res.model is not None:
        orig_ps = list(res.model.doc.element.body.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"))
    pos_of = {id(p): k for k, p in enumerate(orig_ps)}

    def runs_for(p_el: Any) -> list[Run]:
        para = Paragraph(p_el, doc._body)
        runs = [Run(r, para) for r in iter_runs(p_el) if r.getparent() is p_el]
        if not runs:
            runs = [para.add_run("")]
        return runs

    count = 0
    general: list[Issue] = []
    for issue in res.issues:
        if issue.severity == "INFO":
            continue
        n = per_rule.get(issue.rule, 0)
        if cap and n >= cap:
            continue
        per_rule[issue.rule] = n + 1
        anchor = issue.block.anchor_el if issue.block is not None else None
        k = pos_of.get(id(anchor)) if anchor is not None else None
        if k is None or k >= len(body_ps):
            general.append(issue)
            continue
        text = f"[{issue.severity}] {issue.message}"
        if issue.details:
            text += "\n" + "\n".join(issue.details)
        if cap and n + 1 == cap and totals.get(issue.rule, 0) > cap:
            text += f"\n… у документі ще {totals[issue.rule] - cap} подібних місць."
        runs = runs_for(body_ps[k])
        doc.add_comment(runs, text=text, author=author, initials="ВО")
        count += 1

    if general and body_ps:
        text = "Загальні зауваження до оформлення:\n" + "\n".join(f"• [{i.severity}] {i.message}" for i in general)
        doc.add_comment(runs_for(body_ps[0]), text=text, author=author, initials="ВО")
        count += 1
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return count
