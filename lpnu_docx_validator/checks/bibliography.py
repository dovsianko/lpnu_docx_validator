"""Список джерел і посилання на них."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

from ..docmodel import Block
from . import Ctx


@dataclass
class Entry:
    no: Optional[int]
    text: str
    block: Block


def bibliography_entries(ctx: Ctx) -> list[Entry]:
    if "bib" in ctx.cache:
        return ctx.cache["bib"]
    start = ctx.model.bib_start
    entries: list[Entry] = []
    if start is not None:
        start_re = re.compile(ctx.s["bibliography.source_start_pattern"])
        auto = 1
        cur: Optional[Entry] = None
        for b in ctx.blocks[start + 1:]:
            if b.kind == "heading" or b.region == "appendix":
                break
            if b.kind not in {"paragraph", "code"} or not b.text:
                continue
            m = start_re.match(b.text)
            if m or b.is_list:
                if cur:
                    entries.append(cur)
                no = int(m.group("bracket") or m.group("plain")) if m else auto
                auto = no + 1
                cur = Entry(no, b.text, b)
            elif cur is not None and re.match(r"^([a-zа-яіїєґ]|(?i:url|режим\s+доступу|doi|https?://|\(дата))", b.text):
                cur.text += " " + b.text  # продовження опису на наступному абзаці
            else:
                if cur:
                    entries.append(cur)
                cur = Entry(None, b.text, b)
        if cur:
            entries.append(cur)
    ctx.cache["bib"] = entries
    return entries


def check_bibliography(ctx: Ctx) -> None:
    s = ctx.s
    start = ctx.model.bib_start
    if start is None:
        ctx.rep.add("bibliography_missing", "не знайдено розділ «Список використаних джерел»", location="Список джерел")
        return
    entries = bibliography_entries(ctx)
    head = ctx.blocks[start]
    if not entries:
        ctx.rep.add("bibliography_empty", "розділ списку джерел є, але джерел не розпізнано", head)
        return
    year_re = re.compile(s["bibliography.year_pattern"])
    url_re = re.compile(s["bibliography.url_pattern"])
    access_re = re.compile(s["bibliography.access_date_pattern"])
    pages_re = re.compile(s["bibliography.pages_pattern"])
    article_re = re.compile(s["bibliography.article_markers_pattern"])
    expected = 1
    for e in entries:
        b, text = e.block, e.text
        label = f"джерело {e.no}" if e.no else "джерело у списку"
        if e.no is None:
            ctx.rep.add("source_number_missing", "джерело без номера (ні «1.», ні «[1]», ні автонумерації Word)", b, label=label)
        else:
            if e.no != expected:
                ctx.rep.add("source_sequence", f"порушено нумерацію: очікувався номер {expected}, знайдено {e.no}", b, label=label)
            expected = e.no + 1
        body = re.sub(r"^\s*(?:\[\d+\]|\d+[.)])\s+", "", text)
        if re.fullmatch(r"(?i)\s*(https?://|www\.)\S+\s*", body):
            ctx.rep.add("source_bare_url", "голе посилання без бібліографічного опису", b, label=label)
            continue
        if len(body) < s["bibliography.min_length"]:
            ctx.rep.add("source_too_short", "надто короткий опис джерела; ймовірно, він неповний", b, label=label)
        if not year_re.search(text) and (s["bibliography.online_requires_year"] or not url_re.search(text)):
            ctx.rep.add("source_missing_year", "не знайдено рік видання/публікації", b, label=label)
        if url_re.search(text):
            if not access_re.search(text):
                ctx.rep.add("source_missing_access_date", "електронний ресурс без дати звернення", b, label=label,
                            details=["Наприклад: «URL: https://… (дата звернення: 01.10.2026)»."])
            if not re.search(r"(?i)https?://|www\.|\bdoi\b", text):
                ctx.rep.add("source_missing_url", "є «URL/Режим доступу», але немає самого посилання", b, label=label)
        if article_re.search(text) and not pages_re.search(text):
            ctx.rep.add("source_article_pages", "схоже на статтю/тези, але не вказано сторінки («С. 10–15»)", b, label=label)
    ctx.rep.info(f"розпізнано джерел: {len(entries)}; перевірка ДСТУ 8302 виконана евристично", head)


# ---------------------------------------------------------------------------
# Посилання [1], [2–4], [1, с. 25; 3]
# ---------------------------------------------------------------------------

REF_RE = re.compile(r"\[([^\[\]]*\d[^\[\]]*)\]")
_NUM_OR_RANGE = re.compile(r"(\d+)(?:\s*[-–—]\s*(\d+))?")


def parse_source_ref(content: str) -> Optional[list[int]]:
    """Номери джерел із вмісту дужок або None, якщо формат нерозпізнаний.

    «1, с. 25» → [1]; «2–4» → [2, 3, 4]; «1; 3, p. 7» → [1, 3]; «1, 3, 5» → [1, 3, 5].
    Усе, що після коми починається з літери (с., p., т., табл.), — це локатор, а не номер.
    """
    nums: list[int] = []
    for part in content.split(";"):
        part = part.strip()
        if not part:
            continue
        for k, chunk in enumerate(x.strip() for x in part.split(",")):
            m = _NUM_OR_RANGE.fullmatch(chunk)
            if m:
                a, b = int(m.group(1)), int(m.group(2) or m.group(1))
                lo, hi = min(a, b), max(a, b)
                if hi - lo > 300:
                    return None
                nums.extend(range(lo, hi + 1))
                continue
            if k > 0 and re.match(r"[^\W\d_]", chunk):
                break  # локатор: «с. 25», «p. 3–5», «т. 2»
            return None
    if not nums or 0 in nums:
        return None
    return nums


def check_source_references(ctx: Ctx) -> None:
    entries = bibliography_entries(ctx)
    if ctx.model.bib_start is None or not entries:
        return
    order: list[int] = []
    refs: set[int] = set()
    for _, b in ctx.in_regions("main"):
        if b.kind not in {"paragraph", "heading", "figure_caption", "table_caption"} or not b.text:
            continue
        for m in REF_RE.finditer(b.text):
            nums = parse_source_ref(m.group(1))
            if nums is None:
                ctx.rep.add("source_ref_suspicious", f"не вдалося розібрати посилання {m.group(0)}", b,
                            details=["Очікувані формати: [1], [2–4], [1; 3], [1, с. 25]."])
                continue
            for n in nums:
                if n not in refs:
                    order.append(n)
                refs.add(n)
                if n not in {e.no for e in entries}:
                    ctx.rep.add("source_ref_unknown_number", f"посилання [{n}], але джерела з таким номером у списку немає", b)
    numbers = sorted({e.no for e in entries if e.no is not None})
    if not refs:
        ctx.rep.add("source_refs_missing_all", "у тексті немає жодного посилання на джерела виду [1]", location="Посилання на джерела")
        return
    missing = [n for n in numbers if n not in refs]
    if missing:
        ctx.rep.add("source_missing_in_text", "джерела зі списку не згадані в тексті: " + ", ".join(f"[{n}]" for n in missing),
                    location="Посилання на джерела")
    else:
        ctx.rep.info("усі джерела зі списку згадано в тексті", location="Посилання на джерела")
    known = [n for n in order if n in set(numbers)]
    for a, b2 in zip(known, known[1:]):
        if b2 < a:
            ctx.rep.add("source_ref_order", f"джерело [{a}] вперше згадано раніше за [{b2}]; список має йти в порядку першої згадки",
                        location="Посилання на джерела")
            break
