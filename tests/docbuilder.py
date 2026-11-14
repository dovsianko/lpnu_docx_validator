"""Генератор тестових документів: «правильна» курсова + способи внести конкретні помилки."""

from __future__ import annotations

import io
import struct
import zlib
from pathlib import Path
from typing import Optional

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

A = WD_ALIGN_PARAGRAPH


def png_bytes(w: int = 1200, h: int = 600) -> bytes:
    raw = b"".join(b"\x00" + b"\xff\xff\xff" * w for _ in range(h))

    def chunk(t: bytes, d: bytes) -> bytes:
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b""))


def _strip_theme_fonts(style) -> None:
    rpr = style.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is not None:
        for a in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            if rf.get(qn(a)) is not None:
                del rf.attrib[qn(a)]


def add_field(paragraph, instr: str, result: str = "1") -> None:
    def run_with(el):
        r = OxmlElement("w:r")
        r.append(el)
        paragraph._p.append(r)

    b = OxmlElement("w:fldChar")
    b.set(qn("w:fldCharType"), "begin")
    run_with(b)
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    run_with(it)
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    run_with(sep)
    paragraph.add_run(result)
    e = OxmlElement("w:fldChar")
    e.set(qn("w:fldCharType"), "end")
    run_with(e)


class CourseworkBuilder:
    """Будує документ крок за кроком; кожен метод можна викликати з «помилкою»."""

    def __init__(self) -> None:
        self.doc = Document()
        st = self.doc.styles["Normal"]
        _strip_theme_fonts(st)
        st.font.name = "Times New Roman"
        st.font.size = Pt(14)
        pf = st.paragraph_format
        pf.alignment = A.JUSTIFY
        pf.first_line_indent = Cm(1.25)
        pf.line_spacing = 1.5
        pf.space_after = Pt(0)
        pf.space_before = Pt(0)
        for name in ("Heading 1", "Heading 2", "Heading 3"):
            h = self.doc.styles[name]
            _strip_theme_fonts(h)
            h.font.name = "Times New Roman"
            h.font.size = Pt(14)
            h.font.bold = True
            h.font.color.rgb = RGBColor(0, 0, 0)
            h.paragraph_format.space_before = Pt(0)
            h.paragraph_format.space_after = Pt(0)
            h.paragraph_format.first_line_indent = Cm(1.25)
        self.doc.styles["Heading 1"].paragraph_format.alignment = A.CENTER
        self.doc.styles["Heading 1"].paragraph_format.first_line_indent = Cm(0)
        toc = self.doc.styles.add_style("toc 1", WD_STYLE_TYPE.PARAGRAPH)
        toc.base_style = self.doc.styles["Normal"]
        sec = self.doc.sections[0]
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        for side, v in (("left", 2.5), ("right", 1.5), ("top", 2), ("bottom", 2)):
            setattr(sec, f"{side}_margin", Cm(v))
        self.fig = 0
        self.tbl = 0

    # --- базові ---
    def p(self, text: str, **fmt):
        para = self.doc.add_paragraph(text)
        for k, v in fmt.items():
            setattr(para.paragraph_format, k, v)
        return para

    def heading(self, text: str, level: int = 1, new_page: bool = True):
        if new_page:
            self.page_break()
        return self.doc.add_heading(text, level)

    def page_break(self) -> None:
        last = self.doc.paragraphs[-1] if self.doc.paragraphs else self.doc.add_paragraph()
        last.add_run().add_break(WD_BREAK.PAGE)

    # --- частини роботи ---
    def title_page(self) -> "CourseworkBuilder":
        self.p("Національний університет «Львівська політехніка»", alignment=A.CENTER, first_line_indent=Cm(0))
        self.p("КУРСОВА РОБОТА", alignment=A.CENTER, first_line_indent=Cm(0))
        self.doc.add_section(WD_SECTION.NEW_PAGE)
        sec = self.doc.sections[-1]
        sec.header.is_linked_to_previous = False
        hp = sec.header.paragraphs[0]
        hp.alignment = A.RIGHT
        add_field(hp, "PAGE", "2")
        self.doc.sections[0].header.is_linked_to_previous = False
        return self

    def toc(self, entries: list[str], auto: bool = True) -> "CourseworkBuilder":
        self.doc.add_heading("ЗМІСТ", 1)
        first = True
        for e in entries:
            para = self.doc.add_paragraph(style="toc 1" if auto else None)
            if auto and first:
                add_field(para, 'TOC \\o "1-3" \\h \\z \\u', "")
                first = False
            para.add_run(f"{e}\t3")
        return self

    def body(self, text: str):
        return self.p(text)

    def figure(self, num: Optional[str] = None, caption: Optional[str] = "Схема системи", width_cm: float = 15,
               caption_align=A.CENTER, px_width: int = 1200):
        if num is None:
            self.fig += 1
            num = f"1.{self.fig}"
        pic = self.doc.add_paragraph()
        pic.paragraph_format.alignment = A.CENTER
        pic.paragraph_format.first_line_indent = Cm(0)
        pic.add_run().add_picture(io.BytesIO(png_bytes(px_width, px_width // 2)), width=Cm(width_cm))
        if caption is not None:
            self.p(f"Рис. {num} — {caption}", alignment=caption_align, first_line_indent=Cm(0))
        return num

    def table(self, num: Optional[str] = None, caption: Optional[str] = "Порівняння методів", font_pt: float = 12):
        if num is None:
            self.tbl += 1
            num = f"1.{self.tbl}"
        if caption is not None:
            self.p(f"Таблиця {num} — {caption}", alignment=A.RIGHT, first_line_indent=Cm(0))
        t = self.doc.add_table(rows=2, cols=2)
        for row in t.rows:
            for cell in row.cells:
                cell.paragraphs[0].add_run("значення").font.size = Pt(font_pt)
        return num

    def formula(self, num: Optional[str] = "1.1") -> None:
        para = self.doc.add_paragraph()
        para.paragraph_format.alignment = A.CENTER
        m = OxmlElement("m:oMath")
        r = OxmlElement("m:r")
        t = OxmlElement("m:t")
        t.text = "E=mc^2"
        r.append(t)
        m.append(r)
        para._p.append(m)
        if num:
            para.add_run(f"\t({num})")

    def bibliography(self, entries: list[str]) -> None:
        self.heading("СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ")
        for i, e in enumerate(entries, start=1):
            self.p(f"{i}. {e}")

    def save(self, path: Path) -> Path:
        self.doc.save(str(path))
        return path


GOOD_SOURCES = [
    "Петренко П. П. Захист інформації в комп'ютерних мережах : навч. посіб. Львів : Вид-во Львівської політехніки, 2021. 320 с.",
    "Іваненко І. І. Методи виявлення вторгнень. Вісник Національного університету «Львівська політехніка». 2022. № 3. С. 15–22.",
    "OWASP Top Ten. URL: https://owasp.org/www-project-top-ten/ (дата звернення: 01.10.2026).",
]


def good_coursework(path: Path, **opts) -> Path:
    b = CourseworkBuilder().title_page()
    heads = ["ВСТУП", "РОЗДІЛ 1 АНАЛІЗ ПРЕДМЕТНОЇ ОБЛАСТІ", "1.1 Огляд методів", "1.2 Постановка задачі",
             "ВИСНОВКИ", "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ", "ДОДАТОК А"]
    b.toc(heads, auto=opts.get("auto_toc", True))
    b.heading("ВСТУП")
    b.body("Курсова робота присвячена аналізу методів захисту інформації [1]. Актуальність теми зумовлена "
           "зростанням кількості атак на вебзастосунки [3].")
    b.heading("РОЗДІЛ 1 АНАЛІЗ ПРЕДМЕТНОЇ ОБЛАСТІ")
    b.heading("1.1 Огляд методів", 2, new_page=False)
    b.body("Загальну схему системи наведено на рис. 1.1, а порівняння методів — у табл. 1.1 [2, с. 17].")
    b.figure()
    b.body("Як видно з рисунка, система складається з трьох модулів. Енергію обчислюють за формулою (1.1):")
    b.formula("1.1")
    b.body("де E — енергія, m — маса. Повний лістинг наведено в додатку А.")
    b.table()
    b.body("Отже, найкращим є другий метод, що підтверджують результати з табл. 1.1.")
    b.heading("1.2 Постановка задачі", 2, new_page=False)
    b.body("Метою роботи є розроблення системи виявлення атак.")
    b.heading("ВИСНОВКИ")
    b.body("У роботі розглянуто методи захисту та сформульовано задачу.")
    b.bibliography(GOOD_SOURCES)
    b.heading("ДОДАТОК А")
    b.p("Лістинг програми", alignment=A.CENTER, first_line_indent=Cm(0))
    return b.save(path)
