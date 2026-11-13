"""Рисунки, таблиці, підписи, згадки; формули."""

from __future__ import annotations

import re

from ..docmodel import Block
from ..settings import ALIGN_NAMES_UA
from . import Ctx, find_object_refs

REF_KINDS = {"paragraph", "heading", "code"}


def _first_refs(ctx: Ctx, pattern: str) -> dict[str, int]:
    rx = re.compile(pattern)
    first: dict[str, int] = {}
    for pos, b in ctx.in_regions("main", "appendix"):
        if b.kind in REF_KINDS and b.text:
            for num in find_object_refs(rx, b.text):
                first.setdefault(num, pos)
    return first


def check_captions(ctx: Ctx) -> None:
    s = ctx.s
    fig_refs = _first_refs(ctx, s["captions.figure_reference_pattern"])
    tbl_refs = _first_refs(ctx, s["captions.table_reference_pattern"])
    cont_re = re.compile(s["captions.table_continuation_pattern"])
    fmt_re = re.compile(s["captions.format_pattern"])
    seen: dict[tuple[str, str], Block] = {}

    for pos, b in ctx.in_regions("main", "appendix"):
        if b.kind in {"figure_caption", "table_caption"}:
            is_fig = b.kind == "figure_caption"
            num = (b.caption_num or "?").upper()
            ua = "рисунка" if is_fig else "таблиці"
            key = (b.kind, num)
            if key in seen:
                ctx.rep.add("caption_duplicate", f"номер {ua} {num} повторюється", b,
                            details=[f"Попередній підпис: {seen[key].text[:80]}"])
            seen[key] = b
            if not fmt_re.match(b.text):
                ctx.rep.add("caption_format", f"підпис {ua} {num} не за зразком «{'Рис.' if is_fig else 'Таблиця'} {num} — Назва»", b)
            if is_fig:
                prev = ctx.prev_meaningful(pos)
                if not b.extra.get("with_image") and (prev is None or prev.kind != "image_paragraph"):
                    ctx.rep.add("caption_position", f"підпис рисунка {num} не стоїть одразу після рисунка", b)
            else:
                nxt = ctx.next_meaningful(pos)
                if nxt is None or nxt.kind != "table":
                    ctx.rep.add("caption_position", f"підпис таблиці {num} не стоїть одразу перед таблицею", b)
            exp = s["captions.figure_alignment"] if is_fig else s["captions.table_alignment"]
            got = ctx.r.alignment(b.el)
            if exp != "ANY" and got != exp:
                ctx.rep.add("caption_alignment", f"підпис {ua} {num}: вирівнювання {ALIGN_NAMES_UA[got]}, очікується {ALIGN_NAMES_UA[exp]}", b)
            if not s["captions.allow_trailing_dot"] and re.search(r"[^.]\.$", b.text):
                ctx.rep.add("caption_trailing_dot", f"у кінці підпису {ua} {num} не ставлять крапку", b)
            ref = (fig_refs if is_fig else tbl_refs).get(num)
            obj_pos = pos
            if is_fig and not b.extra.get("with_image"):
                prev = ctx.prev_meaningful(pos)
                if prev is not None and prev.kind == "image_paragraph":
                    obj_pos = ctx.blocks.index(prev)
            if ref is None:
                ctx.rep.add("object_reference_missing", f"у тексті немає згадки {ua} {num}", b,
                            details=[f"Перед {'рисунком' if is_fig else 'таблицею'} має бути, наприклад, «({'рис.' if is_fig else 'табл.'} {num})»."])
            elif ref > obj_pos:
                ctx.rep.add("object_reference_missing", f"{ua.capitalize()} {num} згадано лише після {'нього' if is_fig else 'неї'}", b,
                            details=[f"Перша згадка: {ctx.blocks[ref].text[:80]}"])

        elif b.kind == "image_paragraph":
            nxt = ctx.next_meaningful(pos)
            prev = ctx.prev_meaningful(pos)
            if prev is not None and prev.kind == "figure_caption" and not prev.extra.get("with_image"):
                continue  # підпис стоїть над рисунком — про це вже є caption_position
            if nxt is None or nxt.kind != "figure_caption":
                ctx.rep.add("caption_missing", "після рисунка немає підпису виду «Рис. 1.1 — Назва»", b, label="рисунок без підпису")
        elif b.kind == "table":
            prev = ctx.prev_meaningful(pos)
            if prev is None or (prev.kind != "table_caption" and not cont_re.match(prev.text or "")):
                ctx.rep.add("caption_missing", "перед таблицею немає підпису виду «Таблиця 1.1 — Назва»", b, label="таблиця без підпису")


def check_formulas(ctx: Ctx) -> None:
    ref_re = re.compile(ctx.s["formulas.reference_pattern"])
    refs: set[str] = set()
    for _, b in ctx.in_regions("main", "appendix"):
        if b.kind in REF_KINDS and b.text:
            refs.update(m.group("num") for m in ref_re.finditer(b.text))
    for _, b in ctx.in_regions("main", "appendix"):
        if b.kind != "formula":
            continue
        if not b.formula_num:
            ctx.rep.add("formula_number_missing", "формула без номера; номер ставлять праворуч у дужках: (2.1)", b, label="формула")
        elif b.formula_num not in refs:
            ctx.rep.add("formula_reference_missing", f"на формулу ({b.formula_num}) немає посилання в тексті", b, label="формула")
