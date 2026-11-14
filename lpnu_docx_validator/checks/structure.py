"""Структура роботи, зміст, заголовки, нумерація, завершення підрозділів, додатки."""

from __future__ import annotations

import re
from typing import Optional

from docx.oxml.ns import qn

from ..docmodel import Block, has_page_break, iter_runs, run_text, section_break_type, short_quote
from ..settings import ALIGN_NAMES_UA
from . import Ctx, find_letter_refs

UA_ALPHABET = "АБВГҐДЕЄЖЗИІЇЙКЛМНОПРСТУФХЦЧШЩЬЮЯ"
LATIN_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ"  # без I та O
LATIN_LOOKALIKE = {"A": "А", "B": "В", "C": "С", "E": "Е", "H": "Н", "K": "К", "M": "М", "P": "Р", "T": "Т", "X": "Х"}


def pattern_label(pattern: str) -> str:
    """Людська назва з регулярного виразу: '(?i)^\\s*основна\\s*\\(?\\s*технічна…' → 'ОСНОВНА (ТЕХНІЧНА) ЧАСТИНА'."""
    named = re.search(r"\(\?#([^)]*)\)", pattern)  # явна назва: (?#перший розділ)
    if named:
        return named.group(1).strip()
    t = re.sub(r"\(\?[a-zA-Z]+\)", "", pattern)
    t = re.sub(r"\(\?:[^()]*\)\?", "", t)                      # необов'язкові групи (?:РОЗДІЛ\s+)?
    t = re.sub(r"\\[sS][*+?]?|\\b|\\d[*+?]?|[\^$]", " ", t)
    t = re.sub(r"\\([().])\??", r"\1", t)                     # \(? → (
    t = re.sub(r"\.[*+?]", "", t)
    m = re.search(r"\(([^()|]+)\|[^()]*\)", t)                  # (a|b) → a
    while m:
        t = t[: m.start()] + m.group(1) + t[m.end():]
        m = re.search(r"\(([^()|]+)\|[^()]*\)", t)
    t = re.sub(r"[?*+\[\]\\]", "", t)
    t = re.sub(r"\(\s+", "(", re.sub(r"\s+\)", ")", t))
    return re.sub(r"\s+", " ", t).strip(" .").upper() or pattern


def _norm(text: str) -> str:
    t = text.lower().replace("ё", "е").replace("’", "'").replace("ʼ", "'")
    t = re.sub(r"[–—]", "-", t)
    t = re.sub(r"\s+", " ", t).strip(" .")
    return t


# ---------------------------------------------------------------------------
# Структура і зміст
# ---------------------------------------------------------------------------

def check_structure(ctx: Ctx) -> None:
    m = ctx.model
    if m.main_start is None:
        ctx.rep.add("structure_main_start", "не знайдено початок основної частини (ВСТУП або РОЗДІЛ 1 як заголовок)",
                    location="Структура документа",
                    details=["Титульна частина може потрапити в перевірки. Оформіть ВСТУП стилем «Заголовок 1»."])
    else:
        ctx.rep.info("титульна/службова частина пропущена з перевірок основного тексту", ctx.blocks[m.main_start])

    toc_body = range((m.toc_pos or 0) + 1, m.main_start) if m.toc_pos is not None and m.main_start is not None else range(0)
    candidates = [(i, b) for i, b in enumerate(ctx.blocks)
                  if b.kind in {"heading", "paragraph"} and b.text and len(b.text) <= 200 and i not in toc_body]
    found: list[tuple[int, str]] = []
    for pat in ctx.s["structure.required_sections"]:
        rx = re.compile(pat)
        hit = next((i for i, b in candidates if rx.match(b.text)), None)
        label = pattern_label(pat)
        if hit is None:
            ctx.rep.add("structure_missing", f"не знайдено обов'язкову частину «{label}»", location="Структура документа")
        else:
            found.append((hit, label))
    for (i1, l1), (i2, l2) in zip(found, found[1:]):
        if i2 < i1:
            ctx.rep.add("structure_order", f"«{l2}» розміщено раніше за «{l1}»", ctx.blocks[i2])

    if m.toc_pos is not None:
        toc_block = ctx.blocks[m.toc_pos]
        if not m.has_toc_field:
            ctx.rep.add("toc_manual", "зміст набрано вручну, а не вставлено автоматично", toc_block,
                        details=["Посилання → Зміст → Автоматичний зміст. Тоді номери сторінок оновлюються самі."])
        else:
            _check_toc_matches(ctx)


def _check_toc_matches(ctx: Ctx) -> None:
    entries = [b for b in ctx.blocks if b.kind == "toc_entry"]
    if not entries:
        return
    for pat in ctx.s["structure.toc_forbidden_patterns"]:
        rx = re.compile(pat)
        hit = next((e for e in entries if rx.search(e.text)), None)
        if hit is not None:
            name = re.sub(r"[\s.…]*\d+$", "", hit.text)
            ctx.rep.add("toc_forbidden_entry", f"цей пункт не має бути у змісті: {short_quote(name, 60)}", hit)
    max_lvl = ctx.s["structure.toc_max_level"]
    toc_pos = ctx.model.toc_pos
    heads = [b for i, b in enumerate(ctx.blocks)
             if b.kind == "heading" and i != toc_pos and (b.heading_level or 1) <= max_lvl
             and (toc_pos is None or i > toc_pos or b.text != ctx.blocks[toc_pos].text)]
    entry_keys = [_norm(re.sub(r"[\s.…·_\-]*\d+\s*$", "", e.text)) for e in entries]
    head_keys = [_norm(h.text) for h in heads]

    def matches(a: str, b: str) -> bool:
        return a == b or (min(len(a), len(b)) >= 5 and (a.startswith(b) or b.startswith(a)))

    after_toc = [toc_pos is None or ctx.blocks.index(h) > toc_pos for h in heads]
    missing = [h for h, k, a in zip(heads, head_keys, after_toc) if a and not any(matches(k, e) for e in entry_keys)]
    stale = [e for e, k in zip(entries, entry_keys) if k and not any(matches(k, h) for h in head_keys)]
    if missing:
        names = "; ".join(short_quote(h.text, 50) for h in missing[:8]) + (" …" if len(missing) > 8 else "")
        ctx.rep.add("toc_outdated", f"заголовків немає у змісті ({len(missing)}): {names}", entries[0],
                    details=["Оновіть зміст: клацніть по ньому → «Оновити таблицю» → «Оновити цілком»."])
    if stale:
        names = "; ".join(short_quote(e.text, 50) for e in stale[:8]) + (" …" if len(stale) > 8 else "")
        ctx.rep.add("toc_outdated", f"у змісті є пункти, яких немає серед заголовків ({len(stale)}): {names}", stale[0])


# ---------------------------------------------------------------------------
# Заголовки
# ---------------------------------------------------------------------------

def _mostly_bold(ctx: Ctx, b: Block) -> Optional[bool]:
    bold = plain = 0
    for r in iter_runs(b.el):
        t = run_text(r).strip()
        if t:
            if ctx.r.bold(r, b.el):
                bold += len(t)
            else:
                plain += len(t)
    if bold + plain == 0:
        return None
    return bold >= plain


def _starts_new_page(ctx: Ctx, pos: int) -> bool:
    b = ctx.blocks[pos]
    if ctx.r.page_break_before(b.el):
        return True
    first_text = next((r for r in iter_runs(b.el) if run_text(r).strip()), None)
    for r in iter_runs(b.el):
        if r is first_text:
            break
        if r.find(qn("w:br")) is not None and r.find(qn("w:br")).get(qn("w:type")) == "page":
            return True
    for j in range(pos - 1, -1, -1):
        prev = ctx.blocks[j]
        if prev.el is None or prev.el.tag != qn("w:p"):
            return False
        if section_break_type(prev.el) not in (None, "continuous") or has_page_break(prev.el):
            return True
        if prev.kind not in {"empty", "toc_entry"}:
            return False
    return True  # перший елемент документа


def check_headings(ctx: Ctx) -> None:
    s = ctx.s
    sec_re = re.compile(s["headings.section_heading_pattern"])
    num_re = re.compile(s["headings.numbered_heading_pattern"])
    for pos, b in ctx.in_regions("main", "bibliography", "appendix"):
        if b.kind != "heading":
            continue
        text = b.text
        lvl = b.heading_level or 1
        if re.search(r"[^.]\.$", text):
            ctx.rep.add("heading_trailing_dot", "у кінці заголовка не ставлять крапку", b, label="заголовок")
        m_sec = sec_re.match(text)
        m_num = num_re.match(text)
        if m_sec:
            if m_sec.group("dot"):
                ctx.rep.add("heading_number_dot", "після номера розділу не має бути крапки", b, label="заголовок",
                            details=["Правильно: «РОЗДІЛ 1 НАЗВА», а не «РОЗДІЛ 1. НАЗВА»."])
            tail = text[m_sec.end("num"):]
            letters = re.sub(r"[^A-Za-zА-Яа-яІіЇїЄєҐґ]", "", tail)
            if letters and letters != letters.upper():
                ctx.rep.add("heading_section_case", "назву розділу слід писати великими літерами", b, label="заголовок")
            if _mostly_bold(ctx, b) is False:
                ctx.rep.add("heading_bold", "заголовок розділу не напівжирний", b, label="заголовок")
        elif m_num and lvl >= 2:
            if m_num.group("dot"):
                ctx.rep.add("heading_number_dot", "після номера підрозділу не має бути крапки", b, label="заголовок",
                            details=["Правильно: «1.1 Назва», а не «1.1. Назва»."])
            title = text[m_num.end("num"):].lstrip(". ").strip()
            if title[:1].islower():
                ctx.rep.add("heading_subsection_case", "назва підрозділу має починатися з великої літери", b, label="заголовок")
            letters = re.sub(r"[^A-Za-zА-Яа-яІіЇїЄєҐґ]", "", title)
            if len(letters) > 8 and letters == letters.upper():
                ctx.rep.add("heading_subsection_case", "назва підрозділу написана повністю великими літерами", b, label="заголовок")
        exp = s["headings.section_alignment"] if lvl == 1 else s["headings.subsection_alignment"]
        if exp != "ANY":
            got = ctx.r.alignment(b.el)
            if got != exp:
                ctx.rep.add("heading_alignment", f"вирівнювання заголовка: {ALIGN_NAMES_UA[got]}, очікується {ALIGN_NAMES_UA[exp]}",
                            b, label="заголовок")
        if lvl == 1 and not _starts_new_page(ctx, pos):
            ctx.rep.add("section_new_page", "заголовок має починатися з нової сторінки", b, label="заголовок",
                        details=["Поставте розрив сторінки (Ctrl+Enter) перед заголовком або «Абзац → З нової сторінки»."])


# ---------------------------------------------------------------------------
# Нумерація
# ---------------------------------------------------------------------------

def _num_parts(num: str) -> list[str]:
    return num.split(".")


def check_numbering(ctx: Ctx) -> None:
    s = ctx.s
    sec_re = re.compile(s["headings.section_heading_pattern"])
    num_re = re.compile(s["headings.numbered_heading_pattern"])
    last_child: dict[str, int] = {}
    last_section = 0
    for _, b in ctx.in_regions("main"):
        if b.kind != "heading":
            continue
        m_sec = sec_re.match(b.text)
        if m_sec:
            n = int(m_sec.group("num"))
            if n != last_section + 1:
                ctx.rep.add("heading_numbering_sequence", f"після розділу {last_section} іде розділ {n}", b, label="заголовок")
            last_section = n
            continue
        m = num_re.match(b.text)
        if not m or (b.heading_level or 0) < 2:
            continue
        parts = _num_parts(m.group("num"))
        prefix, last = ".".join(parts[:-1]), int(parts[-1])
        if b.chapter is not None and parts[0] != b.chapter:
            ctx.rep.add("heading_numbering_sequence", f"підрозділ {m.group('num')} розміщено в розділі {b.chapter}", b, label="заголовок")
        expected = last_child.get(prefix, 0) + 1
        if last != expected:
            prev = f"{prefix}.{expected - 1}" if expected > 1 else prefix
            ctx.rep.add("heading_numbering_sequence", f"після {prev} очікувався {prefix}.{expected}, а знайдено {m.group('num')}",
                        b, label="заголовок")
        last_child[prefix] = last
        # скидаємо глибші рівні
        for k in [k for k in last_child if k.startswith(m.group("num") + ".")]:
            del last_child[k]

    for kind, ua in (("figure_caption", "рисунка"), ("table_caption", "таблиці")):
        last: dict[str, int] = {}
        for _, b in ctx.in_regions("main", "appendix"):
            if b.kind != kind or not b.caption_num:
                continue
            parts = _num_parts(b.caption_num)
            if not parts[-1].isdigit():
                continue
            if len(parts) >= 2 and b.chapter is not None and parts[0] != b.chapter:
                where = f"додатку {b.chapter}" if b.region == "appendix" else f"розділу {b.chapter}"
                ctx.rep.add("caption_chapter_mismatch", f"номер {ua} {b.caption_num} не відповідає {where}", b)
            prefix = ".".join(parts[:-1])
            n = int(parts[-1])
            exp = last.get(prefix, 0) + 1
            if n != exp and n != last.get(prefix):  # дублікати ловить інше правило
                want = f"{prefix}.{exp}" if prefix else str(exp)
                ctx.rep.add("caption_numbering_sequence", f"очікувався номер {ua} {want}, а знайдено {b.caption_num}", b)
            last[prefix] = max(n, last.get(prefix, 0))


# ---------------------------------------------------------------------------
# Завершення підрозділів
# ---------------------------------------------------------------------------

def check_section_endings(ctx: Ctx) -> None:
    blocks = ctx.blocks
    heads = [(i, b) for i, b in ctx.in_regions("main") if b.kind == "heading"]
    for k, (pos, h) in enumerate(heads):
        lvl = h.heading_level or 1
        if lvl < 2:
            continue
        end = next((p for p, nh in heads[k + 1:] if (nh.heading_level or 1) <= lvl), None)
        if end is None:
            end = ctx.model.bib_start if ctx.model.bib_start is not None else len(blocks)
        nxt = heads[k + 1][0] if k + 1 < len(heads) else end
        own = [b for b in blocks[pos + 1:min(nxt, end)] if b.kind not in {"empty"}]
        content = [b for b in blocks[pos + 1:end] if b.kind not in {"empty", "heading"}]
        if not own and not content:
            ctx.rep.add("section_empty", "підрозділ не містить тексту після заголовка", h, label="підрозділ")
            continue
        if not own:
            continue  # одразу йде вкладений підрозділ — перевіримо його самого
        last = own[-1]
        if last.kind in {"table", "image_paragraph", "figure_caption", "table_caption"}:
            what = {"table": "таблицею", "image_paragraph": "рисунком", "figure_caption": "підписом рисунка",
                    "table_caption": "підписом таблиці"}[last.kind]
            ctx.rep.add("section_ends_with_object", f"підрозділ закінчується {what}; після нього потрібен пояснювальний текст",
                        last, label="останній елемент підрозділу", details=[f"Підрозділ: {short_quote(h.text)}"])


# ---------------------------------------------------------------------------
# Додатки
# ---------------------------------------------------------------------------

def check_appendices(ctx: Ctx) -> None:
    s = ctx.s
    app_re = re.compile(s["appendices.heading_pattern"])
    forbidden = set(s["appendices.forbidden_letters"].upper())
    ua_allowed = "".join(c for c in UA_ALPHABET if c not in forbidden)
    apps: list[tuple[str, Block]] = []
    cyr_doc = False
    for _, b in ctx.in_regions("main", "bibliography", "appendix"):
        if b.kind == "heading":
            m = app_re.match(b.text)
            if m:
                apps.append((m.group("letter").upper(), b))
                cyr_doc = cyr_doc or bool(re.search("[А-ЯҐЄІЇа-яґєії]", b.text[:m.start("letter")]))
    if not apps:
        return
    letters = [x for x, _ in apps]
    for letter, b in apps:
        if letter in forbidden:
            ctx.rep.add("appendix_invalid_letter", f"додаток не можна позначати літерою «{letter}»", b, label="заголовок додатка",
                        details=[f"Не використовуються: {', '.join(sorted(forbidden))}."])
        elif letter in LATIN_LOOKALIKE and cyr_doc:
            ctx.rep.add("appendix_invalid_letter", f"латинська «{letter}» замість кириличної «{LATIN_LOOKALIKE[letter]}»", b,
                        label="заголовок додатка")
        elif letter in "IO":
            ctx.rep.add("appendix_invalid_letter", f"додаток не можна позначати латинською «{letter}»", b, label="заголовок додатка")
    alphabet = ua_allowed if cyr_doc else LATIN_ALPHABET
    norm = [LATIN_LOOKALIKE.get(x, x) if cyr_doc else x for x in letters]
    for k, (letter, (_, b)) in enumerate(zip(norm, apps)):
        if letter not in alphabet:
            continue
        expected = alphabet[k] if k < len(alphabet) else None
        if expected and letter != expected:
            ctx.rep.add("appendix_sequence", f"очікувався ДОДАТОК {expected}, а знайдено ДОДАТОК {letter}", b, label="заголовок додатка")
            break
    ref_re = re.compile(s["appendices.reference_pattern"])
    refs: set[str] = set()
    for _, b in ctx.in_regions("main"):
        if b.kind in {"paragraph", "figure_caption", "table_caption"}:
            refs.update(LATIN_LOOKALIKE.get(x, x) if cyr_doc else x for x in find_letter_refs(ref_re, b.text, alphabet))
    for letter, (_, b) in zip(norm, apps):
        if letter not in refs:
            ctx.rep.add("appendix_not_referenced", f"на ДОДАТОК {letter} немає посилання в основному тексті", b,
                        label="заголовок додатка", details=["Наприклад: «…наведено в додатку А»."])
