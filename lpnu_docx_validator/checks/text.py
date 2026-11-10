"""Основний текст, порожні абзаци, таблиці, типографіка, залишки редагування, зображення."""

from __future__ import annotations

import re
from typing import Any

from docx.oxml.ns import qn

from ..docmodel import EMU_PER_CM, EMU_PER_INCH, NS, iter_runs, run_text, short_quote, xp
from ..settings import ALIGN_NAMES_UA
from . import Ctx

W = NS["w"]


def _visible_runs(ctx: Ctx, p_el: Any) -> list[tuple[Any, str]]:
    out = []
    for r in iter_runs(p_el):
        t = run_text(r)
        if t.strip() and not ctx.r.run_flag(r, p_el, "w:vanish"):
            out.append((r, t))
    return out


def _is_script(r_el: Any) -> bool:
    va = r_el.find(f"{qn('w:rPr')}/{qn('w:vertAlign')}")
    return va is not None and va.get(qn("w:val")) in {"superscript", "subscript"}


# ---------------------------------------------------------------------------
# Основний текст
# ---------------------------------------------------------------------------

def check_paragraph_formatting(ctx: Ctx) -> None:
    s, r = ctx.s, ctx.r
    code_fonts = {f.lower() for f in s["body_text.code_fonts"]}
    exp_font = s["body_text.font_name"].strip().lower()
    for _, b in ctx.in_regions("main"):
        if b.kind != "paragraph" or b.in_table or not b.text:
            continue
        p = b.el
        skip_layout = b.is_list and s["body_text.skip_list_items"]
        if not skip_layout:
            exp = s["body_text.alignment"]
            got = r.alignment(p)
            if exp != "ANY" and got != exp:
                ctx.rep.add("paragraph_alignment", f"вирівнювання абзацу: {ALIGN_NAMES_UA[got]}, очікується {ALIGN_NAMES_UA[exp]}", b)
            ind = r.first_line_indent_cm(p)
            if abs(ind - s["body_text.first_line_indent_cm"]) > s["body_text.indent_tolerance_cm"]:
                got_s = "немає" if abs(ind) < 0.01 else f"{ind:.2f} см"
                ctx.rep.add("paragraph_indent", f"абзацний відступ: {got_s}, очікується {s['body_text.first_line_indent_cm']:.2f} см", b)
        rule, val = r.line_spacing(p)
        exp_ls = s["body_text.line_spacing"]
        if rule == "multiple":
            if abs(val - exp_ls) > 0.05:
                ctx.rep.add("paragraph_line_spacing", f"міжрядковий інтервал: {val:g}, очікується {exp_ls:g}", b)
        else:
            kind = "точно" if rule == "exact" else "мінімум"
            ctx.rep.add("paragraph_line_spacing", f"міжрядковий інтервал «{kind} {val:g} пт», очікується множник {exp_ls:g}", b)
        for which, key, ua in (("before", "body_text.space_before_pt", "перед абзацом"), ("after", "body_text.space_after_pt", "після абзацу")):
            exp_sp = s[key]
            if exp_sp is None:
                continue
            got_sp = r.spacing_pt(p, which)
            if abs(got_sp - exp_sp) > 0.5:
                ctx.rep.add("paragraph_spacing", f"інтервал {ua}: {got_sp:g} пт, очікується {exp_sp:g} пт", b)
        wrong_fonts: set[str] = set()
        wrong_sizes: set[float] = set()
        for run, t in _visible_runs(ctx, p):
            font = r.font_name(run, p, t)
            if font and font.lower() in code_fonts:
                continue  # вбудований фрагмент коду
            if font and font.strip().lower() != exp_font:
                wrong_fonts.add(font)
            if not _is_script(run):
                size = r.font_size_pt(run, p)
                if abs(size - s["body_text.font_size_pt"]) > 0.3:
                    wrong_sizes.add(size)
        if wrong_fonts:
            ctx.rep.add("paragraph_font", f"шрифт: {', '.join(sorted(wrong_fonts))}; очікується {s['body_text.font_name']}", b)
        if wrong_sizes:
            sizes = ", ".join(f"{x:g}" for x in sorted(wrong_sizes))
            ctx.rep.add("paragraph_font_size", f"розмір шрифту: {sizes} пт; очікується {s['body_text.font_size_pt']:g} пт", b)


def check_empty_paragraphs(ctx: Ctx) -> None:
    streak: list[int] = []
    blocks = ctx.blocks

    def flush(end_pos: int) -> None:
        if len(streak) >= 2:
            before = ctx.prev_meaningful(streak[0])
            after = blocks[end_pos] if end_pos < len(blocks) else None
            details = []
            if before is not None and before.text:
                details.append(f"Перед пропуском: {short_quote(before.text)}")
            if after is not None and after.text:
                details.append(f"Після пропуску: {short_quote(after.text)}")
            first = blocks[streak[0]]
            nums = [blocks[k].paragraph_no for k in streak if blocks[k].paragraph_no]
            loc = f"Розділ: {first.context}\nПорожні абзаци {nums[0]}–{nums[-1]}" if nums else first.context
            ctx.rep.add("empty_paragraphs", f"порожні абзаци підряд ({len(streak)} шт.); ймовірно, Enter використано для відступу",
                        first, location=loc, details=details + ["Для нової сторінки використовуйте Ctrl+Enter або «з нової сторінки»."])
        streak.clear()

    for i, b in enumerate(blocks):
        if b.in_front_matter:
            continue
        if b.kind == "empty" and not b.extra.get("page_break"):
            streak.append(i)
        else:
            flush(i)
    flush(len(blocks))


# ---------------------------------------------------------------------------
# Таблиці
# ---------------------------------------------------------------------------

def check_tables(ctx: Ctx) -> None:
    allowed = ctx.s["tables.allowed_font_sizes_pt"]
    if not allowed:
        return
    for pos, b in ctx.in_regions("main", "appendix"):
        if b.kind != "table":
            continue
        sizes: set[float] = set()
        for p in b.el.iter(f"{{{W}}}p"):
            for run, _ in _visible_runs(ctx, p):
                if _is_script(run):
                    continue
                size = ctx.r.font_size_pt(run, p)
                if all(abs(size - a) > 0.3 for a in allowed):
                    sizes.add(size)
        if sizes:
            prev = ctx.prev_meaningful(pos)
            label = f"таблиця {prev.caption_num}" if prev is not None and prev.kind == "table_caption" else "таблиця"
            ctx.rep.add("table_font_size", f"у таблиці шрифт {', '.join(f'{x:g}' for x in sorted(sizes))} пт; дозволено "
                        + ", ".join(f"{a:g}" for a in allowed) + " пт", b, label=label)


# ---------------------------------------------------------------------------
# Типографіка
# ---------------------------------------------------------------------------

CYR = "А-Яа-яІіЇїЄєҐґ"
WORD_RE = re.compile(rf"[A-Za-z{CYR}'’ʼ]+")
CYR_RE = re.compile(f"[{CYR}]")
LAT_RE = re.compile("[A-Za-z]")
DOUBLE_SPACE = re.compile(r"(?<=\S) {2,}(?=\S)")
SPACE_PUNCT = re.compile(r"(?<=\w) +[,.;:!?](?=\s|$|[»)\]])")
HYPHEN_DASH = re.compile(r"(?<=\w) - (?=\w)")


def _around(text: str, pos: int, width: int = 25) -> str:
    a, b = max(0, pos - width), min(len(text), pos + width)
    return ("…" if a else "") + text[a:b].replace("\n", " ") + ("…" if b < len(text) else "")


def check_typography(ctx: Ctx) -> None:
    for _, b in ctx.in_regions("main", "appendix"):
        if b.kind not in {"paragraph", "heading", "figure_caption", "table_caption"} or b.is_code or not b.text:
            continue
        raw = b.raw_text
        m = DOUBLE_SPACE.search(raw)
        if m:
            ctx.rep.add("double_space", f"кілька пробілів підряд: «{_around(raw, m.start())}»", b)
        m = SPACE_PUNCT.search(raw)
        if m:
            ctx.rep.add("space_before_punctuation", f"пробіл перед розділовим знаком: «{_around(raw, m.start())}»", b)
        m = HYPHEN_DASH.search(raw)
        if m:
            ctx.rep.add("hyphen_instead_of_dash", f"дефіс замість тире: «{_around(raw, m.start())}»", b)
        if '"' in raw or "“" in raw or "”" in raw:
            ctx.rep.add("straight_quotes", "лапки \"…\" або “…” замість «…»", b)
        mixed = []
        for w in WORD_RE.findall(raw):
            if len(w) >= 2 and CYR_RE.search(w) and LAT_RE.search(w):
                lat = "".join(sorted(set(LAT_RE.findall(w))))
                mixed.append(f"{w} (латиниця: {lat})")
        if mixed:
            ctx.rep.add("mixed_script", "латинські літери в кириличних словах: " + "; ".join(mixed[:5]), b,
                        details=["Наберіть ці слова заново з українською розкладкою."])
        if b.kind == "paragraph" and re.match(r"^( {2,}|\t)", raw):
            ctx.rep.add("leading_spaces", "абзац починається з пробілів або табуляції", b,
                        details=["Абзацний відступ задайте в «Абзац → Перший рядок»."])
        if b.kind == "paragraph" and "\n" in raw.strip("\n"):
            ctx.rep.add("manual_line_break", "розрив рядка (Shift+Enter) всередині абзацу", b)


# ---------------------------------------------------------------------------
# Залишки редагування
# ---------------------------------------------------------------------------

def check_document_hygiene(ctx: Ctx) -> None:
    body = ctx.model.doc.element.body
    changes = xp(body, ".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo | .//w:rPrChange | .//w:pPrChange")
    if changes:
        ctx.rep.add("tracked_changes", f"у документі є невиконані виправлення ({len(changes)} шт.)", ctx.block_of(changes[0]),
                    details=["Рецензування → Прийняти всі виправлення."])
    comments = xp(body, ".//w:commentRangeStart | .//w:commentReference")
    if comments:
        n = len({c.get(qn("w:id")) for c in comments})
        ctx.rep.add("comments_present", f"у документі є примітки ({n} шт.)", ctx.block_of(comments[0]))
    for _, b in ctx.in_regions("main", "bibliography", "appendix"):
        if b.el is None or b.el.tag != f"{{{W}}}p":
            continue
        hidden = colored = highlighted = False
        for run in iter_runs(b.el):
            if not run_text(run).strip():
                continue
            if ctx.r.run_flag(run, b.el, "w:vanish"):
                hidden = True
                continue
            if ctx.r.highlight(run):
                highlighted = True
            in_link = run.getparent().tag == qn("w:hyperlink")
            col = ctx.r.color(run, b.el)
            if col and col not in {"000000", "000"} and not in_link and b.region == "main":
                colored = True
        if hidden:
            ctx.rep.add("hidden_text", "абзац містить прихований текст", b)
        if highlighted:
            ctx.rep.add("highlighted_text", "текст виділено маркером або заливкою", b)
        if colored:
            ctx.rep.add("colored_text", "кольоровий текст в основній частині", b)


# ---------------------------------------------------------------------------
# Зображення
# ---------------------------------------------------------------------------

def check_images(ctx: Ctx) -> None:
    from docx.image.image import Image

    min_dpi = ctx.s["images.min_dpi"]
    for _, b in ctx.in_regions("main", "appendix"):
        if b.kind != "image_paragraph":
            continue
        width_limit = ctx.model.section_text_width_cm(b.section_no)
        for inline in xp(b.el, ".//wp:inline | .//wp:anchor"):
            ext = inline.find(f"{{{NS['wp']}}}extent")
            if ext is None:
                continue
            cx = int(ext.get("cx") or 0)
            if width_limit and cx / EMU_PER_CM > width_limit + 0.1:
                ctx.rep.add("image_too_wide", f"ширина зображення {cx / EMU_PER_CM:.1f} см більша за ширину тексту {width_limit:.1f} см", b,
                            label="рисунок")
            for blip in xp(inline, ".//a:blip"):
                rid = blip.get(f"{{{NS['r']}}}embed")
                try:
                    part = ctx.model.doc.part.related_parts[rid]
                    img = Image.from_blob(part.blob)
                except Exception:
                    continue
                if cx <= 0 or img.content_type not in {"image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff"}:
                    continue
                dpi = img.px_width / (cx / EMU_PER_INCH)
                if dpi < min_dpi:
                    ctx.rep.add("image_low_resolution",
                                f"роздільна здатність ≈ {dpi:.0f} DPI ({img.px_width}×{img.px_height} px); бажано ≥ {min_dpi}", b,
                                label="рисунок")
