"""Перевірки. Кожна група — функція ``check_<group>(ctx)``; реєстр нижче."""

from __future__ import annotations

import re
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable, Optional

from ..docmodel import Block, DocModel, build_model
from ..report import Reporter, ValidationResult
from ..settings import Settings


@dataclass
class Ctx:
    model: DocModel
    s: Settings
    rep: Reporter
    cache: dict[str, Any] = field(default_factory=dict)

    @property
    def blocks(self) -> list[Block]:
        return self.model.blocks

    @property
    def r(self):  # StyleResolver
        return self.model.resolver

    def in_regions(self, *regions: str) -> Iterable[tuple[int, Block]]:
        for i, b in enumerate(self.model.blocks):
            if b.region in regions:
                yield i, b

    def prev_meaningful(self, pos: int) -> Optional[Block]:
        for j in range(pos - 1, -1, -1):
            b = self.blocks[j]
            if b.kind != "empty":
                return None if b.in_front_matter else b
        return None

    def next_meaningful(self, pos: int) -> Optional[Block]:
        for j in range(pos + 1, len(self.blocks)):
            if self.blocks[j].kind != "empty":
                return self.blocks[j]
        return None

    def block_of(self, p_el: Any) -> Optional[Block]:
        if "el2block" not in self.cache:
            m: dict[int, Block] = {}
            for b in self.blocks:
                if b.el is not None:
                    m[id(b.el)] = b
                    if b.kind == "table":
                        for p in b.el.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
                            m.setdefault(id(p), b)
            self.cache["el2block"] = m
        cur = p_el
        while cur is not None:
            b = self.cache["el2block"].get(id(cur))
            if b is not None:
                return b
            cur = cur.getparent()
        return None


# ---------------------------------------------------------------------------
# Розбір згадок «рис. 1.1–1.3, 1.5 і 1.7»
# ---------------------------------------------------------------------------

_CONT_RE = re.compile(
    r"\s*(?P<sep>[-–—]|,|;|(?<!\w)(?:і|та|й|або|and|to)(?!\w))\s*(?:(?:рис\.|табл\.|рисун\w*|таблиц\w*)\s*)?"
    r"(?P<num>\d+(?:\.\d+)*|[А-ЯҐЄІЇA-Z](?:\.\d+)+)",
    re.IGNORECASE,
)


def _expand_range(a: str, b: str) -> list[str]:
    pa, pb = a.split("."), b.split(".")
    if len(pa) == len(pb) and pa[:-1] == pb[:-1] and pa[-1].isdigit() and pb[-1].isdigit():
        lo, hi = sorted((int(pa[-1]), int(pb[-1])))
        if hi - lo <= 200:
            pre = ".".join(pa[:-1])
            return [f"{pre}.{n}" if pre else str(n) for n in range(lo, hi + 1)]
    return [a, b]


def find_object_refs(pattern: "re.Pattern[str]", text: str) -> list[str]:
    nums: list[str] = []
    for m in pattern.finditer(text):
        first = m.group("num").upper()
        nums.append(first)
        shape = first.count(".")
        pos, prev = m.end(), first
        while True:
            c = _CONT_RE.match(text, pos)
            if not c or c.group("num").count(".") != shape:
                break
            num = c.group("num").upper()
            if c.group("sep") in "-–—":
                nums.extend(_expand_range(prev, num)[1:])
            else:
                nums.append(num)
            pos, prev = c.end(), num
    return nums


_LETTER_CONT = re.compile(r"\s*(?P<sep>[-–—]|,|(?<!\w)(?:і|та|й)(?!\w))\s*(?P<letter>[A-ZА-ЯҐЄІЇ])(?!\w)")


def find_letter_refs(pattern: "re.Pattern[str]", text: str, alphabet: str) -> list[str]:
    out: list[str] = []
    for m in pattern.finditer(text):
        prev = m.group("letter").upper()
        out.append(prev)
        pos = m.end()
        while True:
            c = _LETTER_CONT.match(text, pos)
            if not c:
                break
            letter = c.group("letter").upper()
            if c.group("sep") in "-–—" and prev in alphabet and letter in alphabet:
                a, b = alphabet.index(prev), alphabet.index(letter)
                out.extend(alphabet[min(a, b): max(a, b) + 1])
            else:
                out.append(letter)
            pos, prev = c.end(), letter
    return out


# ---------------------------------------------------------------------------
# Запуск
# ---------------------------------------------------------------------------

def _registry() -> list[tuple[str, Callable[[Ctx], None]]]:
    from . import bibliography, objects, page, structure, text
    return [
        ("structure", structure.check_structure),
        ("page_setup", page.check_page_setup),
        ("page_numbering", page.check_page_numbering),
        ("page_count", page.check_page_count),
        ("paragraph_formatting", text.check_paragraph_formatting),
        ("empty_paragraphs", text.check_empty_paragraphs),
        ("headings", structure.check_headings),
        ("numbering", structure.check_numbering),
        ("captions_and_references", objects.check_captions),
        ("section_endings", structure.check_section_endings),
        ("formulas", objects.check_formulas),
        ("tables", text.check_tables),
        ("appendices", structure.check_appendices),
        ("bibliography", bibliography.check_bibliography),
        ("source_references", bibliography.check_source_references),
        ("typography", text.check_typography),
        ("document_hygiene", text.check_document_hygiene),
        ("images", text.check_images),
    ]


def validate(path: Path, settings: Settings) -> ValidationResult:
    path = Path(path)
    rep = Reporter(settings)
    if not path.exists():
        return ValidationResult(path, settings, [], fatal=f"Файл не знайдено: {path}")
    if path.suffix.lower() != ".docx":
        return ValidationResult(path, settings, [], fatal="Підтримуються лише файли .docx (старий .doc спершу збережіть як .docx).")
    try:
        model = build_model(path, settings)
    except Exception as exc:  # пошкоджений файл, не docx тощо
        return ValidationResult(path, settings, [], fatal=f"Не вдалося відкрити DOCX: {exc}")
    ctx = Ctx(model, settings, rep)
    for group, fn in _registry():
        if not settings.enabled(group):
            continue
        try:
            fn(ctx)
        except Exception as exc:  # одна зламана перевірка не має зупиняти решту
            tb = traceback.extract_tb(exc.__traceback__)[-1]
            rep.add("internal_error", f"внутрішня помилка перевірки «{group}»: {exc}", location="",
                    details=[f"{Path(tb.filename).name}:{tb.lineno}. Будь ласка, повідомте розробникам і додайте файл."],
                    severity="WARNING")
    return ValidationResult(path, settings, rep.issues, model=model)
