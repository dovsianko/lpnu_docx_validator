"""Сторінка, поля, нумерація сторінок, обсяг."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Optional

from docx.oxml.ns import qn

from ..docmodel import EMU_PER_CM, field_instructions, iter_all_paragraphs, short_quote
from ..settings import ALIGN_NAMES_UA
from . import Ctx


def _cm(v: Any) -> Optional[float]:
    return None if v is None else v / EMU_PER_CM


def _section_desc(ctx: Ctx, sec_no: int) -> str:
    sec_blocks = [b for b in ctx.blocks if b.section_no == sec_no and b.text]
    if not sec_blocks:
        return "порожня або службова Word-секція"
    return f"від {short_quote(sec_blocks[0].text, 50)} до {short_quote(sec_blocks[-1].text, 50)}"


def check_page_setup(ctx: Ctx) -> None:
    s = ctx.s
    tol = s["page.tolerance_cm"]
    for i, section in enumerate(ctx.model.doc.sections, start=1):
        if s["page.skip_first_word_section"] and i == 1 and len(ctx.model.doc.sections) > 1:
            continue
        loc = f"Word-секція {i}\nОхоплює: {_section_desc(ctx, i)}"
        w, h = _cm(section.page_width), _cm(section.page_height)
        ew, eh = s["page.width_cm"], s["page.height_cm"]
        if w is None or h is None:
            continue
        portrait_ok = abs(w - ew) <= tol and abs(h - eh) <= tol
        landscape_ok = abs(w - eh) <= tol and abs(h - ew) <= tol
        if landscape_ok and s["page.allow_landscape_sections"]:
            continue
        if not portrait_ok:
            what = "альбомна орієнтація" if landscape_ok else f"формат {w:.2f}×{h:.2f} см"
            ctx.rep.add("page_setup", f"{what}; очікується книжкова {ew:g}×{eh:g} см", location=loc)
            continue
        for name, actual, key in (
            ("ліве поле", section.left_margin, "page.margin_left_cm"),
            ("праве поле", section.right_margin, "page.margin_right_cm"),
            ("верхнє поле", section.top_margin, "page.margin_top_cm"),
            ("нижнє поле", section.bottom_margin, "page.margin_bottom_cm"),
        ):
            a = _cm(actual)
            at_least = s["page.margin_check"] == "AT_LEAST"
            bad = a is None or (a < s[key] - tol if at_least else abs(a - s[key]) > tol)
            if bad:
                got = "невідомо" if a is None else f"{a:.2f} см"
                want = f"не менше {s[key]:.2f} см" if at_least else f"{s[key]:.2f} см"
                ctx.rep.add("page_setup", f"{name}: {got}, очікується {want}", location=loc)


# ---------------------------------------------------------------------------
# Нумерація сторінок
# ---------------------------------------------------------------------------

def _is_page_field(instr: str) -> bool:
    return instr.split(" ")[0].upper() == "PAGE"


def _tab_stops(ctx: Ctx, p_el: Any) -> list[tuple[int, str]]:
    stops: dict[int, str] = {}
    cleared: set[int] = set()
    for ppr in ctx.r.ppr_chain(p_el):
        tabs = ppr.find(qn("w:tabs"))
        if tabs is None:
            continue
        for t in tabs.findall(qn("w:tab")):
            try:
                pos = int(t.get(qn("w:pos")) or 0)
            except ValueError:
                continue
            val = t.get(qn("w:val")) or "left"
            if pos in stops or pos in cleared:
                continue
            if val == "clear":
                cleared.add(pos)
            else:
                stops[pos] = val
    return sorted(stops.items())


def _tabs_before_page_field(p_el: Any) -> int:
    count = 0
    for el in p_el.iter():
        if el.tag == qn("w:tab") and el.getparent().tag == qn("w:r"):
            count += 1
        elif el.tag == qn("w:instrText") and _is_page_field((el.text or "").strip()):
            return count
        elif el.tag == qn("w:fldSimple") and _is_page_field((el.get(qn("w:instr")) or "").strip()):
            return count
    return count


def page_number_alignment(ctx: Ctx, p_el: Any) -> str:
    frame = p_el.find(f"{qn('w:pPr')}/{qn('w:framePr')}")
    if frame is not None and frame.get(qn("w:xAlign")) in {"left", "center", "right"}:
        return frame.get(qn("w:xAlign")).upper()
    align = ctx.r.alignment(p_el)
    tabs = _tabs_before_page_field(p_el)
    if tabs and align == "LEFT":
        stops = _tab_stops(ctx, p_el)
        if tabs <= len(stops):
            val = stops[tabs - 1][1]
            return {"center": "CENTER", "right": "RIGHT", "end": "RIGHT"}.get(val, "LEFT")
    return align


def _page_paragraphs(container: Any) -> list[Any]:
    try:
        root = container._element
    except Exception:
        return []
    return [p for p in iter_all_paragraphs(root) if any(_is_page_field(i) for i in field_instructions(p))]


def _containers(section: Any, position: str, first_page: bool = False) -> list[tuple[str, Any]]:
    out = []
    if position in {"TOP", "ANY"}:
        out.append(("верхній колонтитул першої сторінки" if first_page else "верхній колонтитул",
                    section.first_page_header if first_page else section.header))
    if position in {"BOTTOM", "ANY"}:
        out.append(("нижній колонтитул першої сторінки" if first_page else "нижній колонтитул",
                    section.first_page_footer if first_page else section.footer))
    return out


def _section_state(ctx: Ctx, section: Any, first_page: bool = False) -> tuple[list[tuple[str, list[Any]]], bool, bool]:
    position = ctx.s["page_numbering.position"]
    expected = ctx.s["page_numbering.alignment"]
    found = [(label, _page_paragraphs(c)) for label, c in _containers(section, position, first_page)]
    has_any = any(ps for _, ps in found)
    ok = any(expected == "ANY" or page_number_alignment(ctx, p) == expected for _, ps in found for p in ps)
    return found, has_any, ok


def _expected_text(ctx: Ctx) -> str:
    pos = {"TOP": "зверху", "BOTTOM": "знизу", "ANY": "зверху або знизу"}[ctx.s["page_numbering.position"]]
    return f"очікується: {pos}, {ALIGN_NAMES_UA[ctx.s['page_numbering.alignment']]}"


def check_page_numbering(ctx: Ctx) -> None:
    doc, s, blocks = ctx.model.doc, ctx.s, ctx.blocks
    sections = list(doc.sections)
    toc_block = blocks[ctx.model.toc_pos] if ctx.model.toc_pos is not None else None
    if s["page_numbering.start_from_toc"] and toc_block is not None:
        start_sec, start_block = toc_block.section_no, toc_block
    elif ctx.model.main_start is not None:
        start_sec, start_block = blocks[ctx.model.main_start].section_no, blocks[ctx.model.main_start]
    else:
        start_sec, start_block = 1, None
        ctx.rep.info("не знайдено ЗМІСТ і початок основної частини — номери сторінок перевіряються в усіх секціях",
                     location="Нумерація сторінок")

    problems = False
    other_position = {"TOP": "BOTTOM", "BOTTOM": "TOP", "ANY": None}[s["page_numbering.position"]]
    for i in range(start_sec, len(sections) + 1):
        section = sections[i - 1]
        loc = f"Word-секція {i}\nОхоплює: {_section_desc(ctx, i)}\n{_expected_text(ctx)}"
        found, has_any, ok = _section_state(ctx, section)
        if not has_any:
            problems = True
            msg = "не знайдено номер сторінки (поле PAGE) у " + ", ".join(lbl for lbl, _ in found)
            if other_position:
                other = [lbl for lbl, c in _containers(section, other_position) if _page_paragraphs(c)]
                if other:
                    ctx.rep.add("page_number_alignment", f"номер сторінки розміщено не там: знайдено в «{other[0]}»", location=loc)
                    continue
            ctx.rep.add("page_number_missing", msg, location=loc,
                        details=["Вставте номер: Вставлення → Номер сторінки."])
            continue
        if not ok:
            problems = True
            got = "; ".join(f"{lbl}: {', '.join(ALIGN_NAMES_UA.get(page_number_alignment(ctx, p), '?') for p in ps)}"
                            for lbl, ps in found if ps)
            ctx.rep.add("page_number_alignment", "номер сторінки вирівняно не так, як вимагається", location=loc,
                        details=[f"Знайдено — {got}."])
        # особливий колонтитул першої сторінки
        title_pg = section.different_first_page_header_footer
        first_hidden = False
        if title_pg:
            _, first_has, _ = _section_state(ctx, section, first_page=True)
            first_hidden = not first_has
        is_start = i == start_sec
        first_text = next((b for b in blocks if b.section_no == i and b.text), None)
        starts_before = (is_start and start_block is not None and first_text is not None
                         and first_text.index < start_block.index)
        if starts_before:
            if not first_hidden:
                ctx.rep.add("page_number_section_start", "Word-секція з номерами сторінок починається раніше за ЗМІСТ",
                            location=loc, details=[
                                f"Перший текст секції: {short_quote(first_text.text)}.",  # type: ignore[union-attr]
                                "Додайте розрив розділу (з наступної сторінки) перед ЗМІСТ або увімкніть «Особливий колонтитул для першої сторінки».",
                            ])
        elif first_hidden:
            ctx.rep.add("page_number_first_page", "на першій сторінці секції номер не показується (особливий колонтитул першої сторінки)",
                        location=loc)

    if not problems:
        ctx.rep.info("номери сторінок знайдено в потрібному місці", location="Нумерація сторінок")


# ---------------------------------------------------------------------------
# Обсяг
# ---------------------------------------------------------------------------

def _pages_from_props(path: Path) -> Optional[int]:
    try:
        with zipfile.ZipFile(path) as zf, zf.open("docProps/app.xml") as f:
            root = ET.parse(f).getroot()
        el = root.find("{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Pages")
        if el is not None and (el.text or "").strip().isdigit():
            return int(el.text.strip())  # type: ignore[union-attr]
    except Exception:
        return None
    return None


def _pages_via_libreoffice(path: Path) -> Optional[int]:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        for cand in (r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
            if Path(cand).exists():
                soffice = cand
                break
    if not soffice:
        return None
    with tempfile.TemporaryDirectory(prefix="docx_validator_") as tmp:
        try:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(path)],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True, timeout=180)
        except Exception:
            return None
        pdfs = list(Path(tmp).glob("*.pdf"))
        if not pdfs:
            return None
        data = pdfs[0].read_bytes()
        try:
            from pypdf import PdfReader  # type: ignore
            return len(PdfReader(str(pdfs[0])).pages)
        except Exception:
            import re
            n = len(re.findall(rb"/Type\s*/Page(?!s)", data))
            return n or None


def check_page_count(ctx: Ctx) -> None:
    pages, source = _pages_via_libreoffice(ctx.model.path), "рендер LibreOffice"
    if pages is None:
        pages, source = _pages_from_props(ctx.model.path), "метадані DOCX (останнє збереження у Word)"
    if pages is None:
        ctx.rep.add("page_count_unknown", "не вдалося визначити кількість сторінок", location="Обсяг роботи",
                    details=["Встановіть LibreOffice для точного підрахунку."])
        return
    lo, hi = ctx.s["page_count.min_pages"], ctx.s["page_count.max_pages"]
    if pages < lo or pages > hi:
        ctx.rep.add("page_count_out_of_range", f"обсяг: {pages} стор., очікується {lo}–{hi}", location="Обсяг роботи",
                    details=[f"Джерело підрахунку: {source}."])
    else:
        ctx.rep.info(f"обсяг: {pages} стор. (норма {lo}–{hi}); джерело: {source}", location="Обсяг роботи")
