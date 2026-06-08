#!/usr/bin/env python3
"""
DOCX Format Validator MVP v3.10

Перевіряє базове оформлення студентської роботи у .docx і формує report.txt.

Встановлення:
    pip install python-docx

Запуск:
    python docx_format_validator.py work.docx
    python docx_format_validator.py work.docx -o report.txt

Обмеження:
- DOCX не завжди містить фактичну розбивку на сторінки, тому перевірку
  кількості сторінок у цій версії вимкнено.
- Перевірка ДСТУ для джерел є евристичною: скрипт ловить типові формальні
  помилки, але не гарантує 100% бібліографічну відповідність.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
import shutil
import subprocess
import tempfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field, fields
from pathlib import Path
from typing import Iterable, Optional, Any

from docx import Document
from docx.document import Document as _Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

VERSION = "MVP v3.10"


# -----------------------------
# Налаштування правил кафедри
# -----------------------------

@dataclass(frozen=True)
class Rules:
    # Сторінка. Поки НЕ міняємо за методичкою: пізніше винесемо в конфіг.
    page_width_cm: float = 21.0
    page_height_cm: float = 29.7
    margin_left_cm: float = 3.0
    margin_right_cm: float = 1.0
    margin_top_cm: float = 2.0
    margin_bottom_cm: float = 2.0
    tolerance_cm: float = 0.08
    skip_first_word_section: bool = True

    # Увімкнення/вимкнення груп перевірок
    check_page_setup: bool = True
    check_page_numbering: bool = True
    check_paragraph_formatting: bool = True
    check_empty_paragraphs: bool = True
    check_headings: bool = True
    check_captions_and_references: bool = True
    check_section_endings: bool = True
    check_bibliography: bool = True
    check_source_references: bool = True

    # Обсяг. Зараз перевірка сторінок вимкнена, але значення лишені для майбутнього конфігу.
    min_pages: int = 35
    max_pages: int = 50

    # Основний текст
    body_font_name: str = "Times New Roman"
    body_font_size_pt: float = 14.0
    body_first_line_indent_cm: float = 1.25
    body_line_spacing: float = 1.5
    body_alignment: int = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Підписи
    figure_caption_alignment: int = WD_ALIGN_PARAGRAPH.CENTER
    table_caption_alignment: int = WD_ALIGN_PARAGRAPH.RIGHT

    # Нумерація сторінок
    page_number_alignment: int = WD_ALIGN_PARAGRAPH.RIGHT
    # TOP — верхній колонтитул, BOTTOM — нижній колонтитул, ANY — будь-де з двох.
    page_number_position: str = "TOP"

    # Підписи. Дозволяємо обидві реальні кафедральні форми:
    # Рис. 1.1 — Назва / Рис. 1.1. Назва / Рисунок 1.1 — Назва
    figure_caption_re: str = r"^\s*(?:Рис\.|Рисунок)\s+(?P<num>\d+(?:\.\d+)*)\s*(?:[-–—.]\s*)\S.+$"
    table_caption_re: str = r"^\s*Таблиця\s+(?P<num>\d+(?:\.\d+)*)\s*(?:[-–—.]\s*)\S.+$"

    # Згадки в тексті до появи об'єкта
    figure_ref_re: str = r"(?i)\b(?:рисунок|рисунку|рисунка|рисунком|рис\.)\s*(?P<num>\d+(?:\.\d+)*)\b"
    table_ref_re: str = r"(?i)\b(?:таблиця|таблиці|таблицю|таблицею|табл\.)\s*(?P<num>\d+(?:\.\d+)*)\b"

    # Заголовки
    section_heading_re: str = r"^\s*РОЗДІЛ\s+(?P<num>\d+)(?P<dot>\.)?\s+\S.+$"
    numbered_heading_re: str = r"^\s*(?P<num>\d+(?:\.\d+){1,5})(?P<dot>\.)?\s+\S.+$"
    intro_like_heading_re: str = r"(?i)^\s*(вступ|висновки|список\s+використаних\s+джерел|додатки?)\s*$"

    # Бібліографія
    bibliography_heading_re: str = (
        r"(?i)^\s*(список\s+використаних\s+джерел|"
        r"перелік\s+використаних\s+джерел|"
        r"список\s+літератури|"
        r"література|"
        r"references)\s*$"
    )
    source_start_re: str = r"^\s*(?:\[(?P<bracket>\d+)\]|(?P<plain>\d+)[\.)])\s+.+"
    year_re: str = r"\b(19\d{2}|20\d{2})\b"
    url_re: str = r"(?i)\b(https?://|www\.|doi:|DOI\b|URL\b|Режим\s+доступу)"
    access_date_re: str = r"(?i)(дата\s+звернення|звернення\s*[:—-]|accessed|retrieved|дата\s+обращения)"

    # Рівні повідомлень. Можливі значення: ERROR, WARNING, INFO, OFF.
    severity_page_setup: str = "ERROR"
    severity_page_number_missing: str = "WARNING"
    severity_page_number_alignment: str = "ERROR"
    severity_page_number_section_start: str = "ERROR"
    severity_paragraph_alignment: str = "WARNING"
    severity_paragraph_indent: str = "WARNING"
    severity_paragraph_line_spacing: str = "WARNING"
    severity_paragraph_font: str = "WARNING"
    severity_paragraph_font_size: str = "WARNING"
    severity_empty_paragraphs: str = "WARNING"
    severity_caption_missing: str = "ERROR"
    severity_caption_position: str = "ERROR"
    severity_caption_alignment: str = "ERROR"
    severity_caption_duplicate: str = "ERROR"
    severity_object_reference_missing: str = "ERROR"
    severity_heading_number_dot: str = "ERROR"
    severity_heading_section_case: str = "WARNING"
    severity_heading_bold: str = "WARNING"
    severity_heading_subsection_case: str = "WARNING"
    severity_section_empty: str = "ERROR"
    severity_section_ends_with_object: str = "ERROR"
    severity_bibliography_missing: str = "ERROR"
    severity_bibliography_empty: str = "ERROR"
    severity_source_number_missing: str = "ERROR"
    severity_source_sequence: str = "WARNING"
    severity_source_bare_url: str = "ERROR"
    severity_source_too_short: str = "WARNING"
    severity_source_missing_year: str = "WARNING"
    severity_source_missing_access_date: str = "WARNING"
    severity_source_missing_url: str = "WARNING"
    severity_source_article_pages: str = "WARNING"
    severity_source_refs_missing_all: str = "ERROR"
    severity_source_ref_unknown_number: str = "ERROR"
    severity_source_missing_in_text: str = "ERROR"
    severity_source_ref_suspicious: str = "WARNING"


RULES = Rules()


# -----------------------------
# Модель звіту
# -----------------------------

@dataclass
class Issue:
    severity: str  # ERROR / WARNING / INFO
    message: str
    location: str = ""
    details: list[str] = field(default_factory=list)


class Reporter:
    def __init__(self) -> None:
        self.issues: list[Issue] = []

    def add(self, severity: str, message: str, location: str = "", details: Optional[list[str]] = None) -> None:
        self.issues.append(Issue(severity, message, location, details or []))

    def error(self, message: str, location: str = "", details: Optional[list[str]] = None) -> None:
        self.add("ERROR", message, location, details)

    def warning(self, message: str, location: str = "", details: Optional[list[str]] = None) -> None:
        self.add("WARNING", message, location, details)

    def info(self, message: str, location: str = "", details: Optional[list[str]] = None) -> None:
        self.add("INFO", message, location, details)

    def render(self, filename: str) -> str:
        errors = sum(1 for i in self.issues if i.severity == "ERROR")
        warnings = sum(1 for i in self.issues if i.severity == "WARNING")
        infos = sum(1 for i in self.issues if i.severity == "INFO")

        lines = [
            "ЗВІТ ПЕРЕВІРКИ ОФОРМЛЕННЯ DOCX",
            "=" * 44,
            f"Версія валідатора: {VERSION}",
            f"Файл: {filename}",
            f"Помилки: {errors}",
            f"Попередження: {warnings}",
            f"Інформаційні повідомлення: {infos}",
            "",
        ]

        if not self.issues:
            lines.append("[OK] Формальних проблем не знайдено.")
            return "\n".join(lines) + "\n"

        for sev in ("ERROR", "WARNING", "INFO"):
            group = [i for i in self.issues if i.severity == sev]
            if not group:
                continue
            lines.append(sev)
            lines.append("-" * len(sev))
            for i in group:
                lines.append(f"[{i.severity}] {i.message}")
                if i.location:
                    for line in i.location.splitlines():
                        lines.append(f"    Місце: {line}" if line == i.location.splitlines()[0] else f"           {line}")
                for d in i.details:
                    lines.append(f"    {d}")
                lines.append("")
        return "\n".join(lines).rstrip() + "\n"



# -----------------------------
# Конфігурація
# -----------------------------

ALIGNMENTS = {
    "LEFT": WD_ALIGN_PARAGRAPH.LEFT,
    "CENTER": WD_ALIGN_PARAGRAPH.CENTER,
    "RIGHT": WD_ALIGN_PARAGRAPH.RIGHT,
    "JUSTIFY": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

CONFIG_KEY_MAP = {
    # checks
    ("checks", "page_setup"): "check_page_setup",
    ("checks", "page_numbering"): "check_page_numbering",
    ("checks", "paragraph_formatting"): "check_paragraph_formatting",
    ("checks", "empty_paragraphs"): "check_empty_paragraphs",
    ("checks", "headings"): "check_headings",
    ("checks", "captions_and_references"): "check_captions_and_references",
    ("checks", "section_endings"): "check_section_endings",
    ("checks", "bibliography"): "check_bibliography",
    ("checks", "source_references"): "check_source_references",

    # page
    ("page", "width_cm"): "page_width_cm",
    ("page", "height_cm"): "page_height_cm",
    ("page", "margin_left_cm"): "margin_left_cm",
    ("page", "margin_right_cm"): "margin_right_cm",
    ("page", "margin_top_cm"): "margin_top_cm",
    ("page", "margin_bottom_cm"): "margin_bottom_cm",
    ("page", "tolerance_cm"): "tolerance_cm",
    ("page", "skip_first_word_section"): "skip_first_word_section",

    # body text
    ("body_text", "font_name"): "body_font_name",
    ("body_text", "font_size_pt"): "body_font_size_pt",
    ("body_text", "first_line_indent_cm"): "body_first_line_indent_cm",
    ("body_text", "line_spacing"): "body_line_spacing",
    ("body_text", "alignment"): "body_alignment",

    # captions
    ("captions", "figure_caption_pattern"): "figure_caption_re",
    ("captions", "table_caption_pattern"): "table_caption_re",
    ("captions", "figure_reference_pattern"): "figure_ref_re",
    ("captions", "table_reference_pattern"): "table_ref_re",
    ("captions", "figure_alignment"): "figure_caption_alignment",
    ("captions", "table_alignment"): "table_caption_alignment",

    # headings
    ("headings", "section_heading_pattern"): "section_heading_re",
    ("headings", "numbered_heading_pattern"): "numbered_heading_re",
    ("headings", "intro_like_heading_pattern"): "intro_like_heading_re",

    # page numbering
    ("page_numbering", "alignment"): "page_number_alignment",
    ("page_numbering", "position"): "page_number_position",

    # bibliography
    ("bibliography", "heading_pattern"): "bibliography_heading_re",
    ("bibliography", "source_start_pattern"): "source_start_re",
    ("bibliography", "year_pattern"): "year_re",
    ("bibliography", "url_pattern"): "url_re",
    ("bibliography", "access_date_pattern"): "access_date_re",
}


def report_issue(reporter: Reporter, severity: str, message: str, location: str = "", details: Optional[list[str]] = None) -> None:
    sev = (severity or "WARNING").upper()
    if sev in {"OFF", "DISABLED", "NONE", "IGNORE"}:
        return
    if sev not in {"ERROR", "WARNING", "INFO"}:
        sev = "WARNING"
    reporter.add(sev, message, location, details)


def _coerce_config_value(field_name: str, value: Any) -> Any:
    if field_name.endswith("alignment") or field_name in {"body_alignment", "figure_caption_alignment", "table_caption_alignment", "page_number_alignment"}:
        if isinstance(value, str):
            key = value.strip().upper()
            if key in ALIGNMENTS:
                return ALIGNMENTS[key]
        return value
    if field_name == "page_number_position" and isinstance(value, str):
        key = value.strip().upper()
        if key in {"TOP", "BOTTOM", "ANY"}:
            return key
        return value.strip().upper()
    if field_name.startswith("severity_") and isinstance(value, str):
        return value.strip().upper()
    return value


def _load_raw_config(config_path: Path) -> dict[str, Any]:
    suffix = config_path.suffix.lower()
    text = config_path.read_text(encoding="utf-8")
    if suffix in {".yaml", ".yml"}:
        try:
            import yaml  # type: ignore
        except Exception as exc:
            raise RuntimeError("Для YAML-конфігу потрібно встановити PyYAML: pip install PyYAML") from exc
        data = yaml.safe_load(text) or {}
        if not isinstance(data, dict):
            raise ValueError("Конфіг має бути словником YAML/JSON.")
        return data
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError("Конфіг має бути словником YAML/JSON.")
    return data


def load_rules(config_path: Optional[Path]) -> Rules:
    if config_path is None:
        return RULES
    raw = _load_raw_config(config_path)
    allowed = {f.name for f in fields(Rules)}
    values = {f.name: getattr(RULES, f.name) for f in fields(Rules)}

    # 1) Плоскі ключі напряму збігаються з назвами полів Rules.
    for key, value in raw.items():
        if key in allowed:
            values[key] = _coerce_config_value(key, value)

    # 2) Людинозрозумілі вкладені секції.
    for path, field_name in CONFIG_KEY_MAP.items():
        current: Any = raw
        found = True
        for part in path:
            if isinstance(current, dict) and part in current:
                current = current[part]
            else:
                found = False
                break
        if found and field_name in allowed:
            values[field_name] = _coerce_config_value(field_name, current)

    # 3) Усі severity.* автоматично мапляться в severity_*
    severities = raw.get("severities")
    if isinstance(severities, dict):
        for key, value in severities.items():
            field_name = f"severity_{key}"
            if field_name in allowed:
                values[field_name] = _coerce_config_value(field_name, value)

    return Rules(**values)

# -----------------------------
# Представлення елементів документа
# -----------------------------

@dataclass
class Block:
    index: int
    kind: str  # paragraph / table / image_paragraph / figure_caption / table_caption / heading / empty
    text: str
    obj: Any
    paragraph_no: Optional[int] = None
    heading_level: Optional[int] = None
    caption_num: Optional[str] = None
    section_no: Optional[int] = None
    in_front_matter: bool = False
    context: str = ""


# -----------------------------
# Утиліти DOCX
# -----------------------------

def cm_value(length: Any) -> Optional[float]:
    if length is None:
        return None
    try:
        return length.cm
    except Exception:
        return None


def pt_value(length: Any) -> Optional[float]:
    if length is None:
        return None
    try:
        return length.pt
    except Exception:
        return None


def approx_equal(a: Optional[float], b: float, tolerance: float) -> bool:
    if a is None:
        return False
    return abs(a - b) <= tolerance


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def short_quote(text: str, limit: int = 90) -> str:
    text = normalize_text(text)
    if len(text) <= limit:
        return f"«{text}»"
    return f"«{text[:limit - 1]}…»"


def iter_block_items(parent: _Document | _Cell) -> Iterable[Paragraph | Table]:
    """Yield paragraphs and tables in document order."""
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise TypeError(f"Unsupported parent type: {type(parent)}")

    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def paragraph_has_image(paragraph: Paragraph) -> bool:
    return bool(paragraph._p.xpath(".//w:drawing | .//w:pict"))


def paragraph_style_name(paragraph: Paragraph) -> str:
    try:
        return paragraph.style.name or ""
    except Exception:
        return ""


def is_heading_style(style_name: str) -> bool:
    style = style_name.lower()
    return style.startswith("heading") or style.startswith("заголовок") or style.startswith("titre")


def heading_level_from_style(style_name: str) -> Optional[int]:
    m = re.search(r"(\d+)$", style_name)
    return int(m.group(1)) if m else None


def is_section_heading(text: str, rules: Rules = RULES) -> bool:
    return bool(re.match(rules.section_heading_re, text))


def is_numbered_subheading(text: str, rules: Rules = RULES) -> bool:
    return bool(re.match(rules.numbered_heading_re, text))


def heading_level_from_number(text: str, rules: Rules = RULES) -> Optional[int]:
    if re.match(rules.section_heading_re, text):
        return 1
    m = re.match(rules.numbered_heading_re, text)
    if m:
        return m.group("num").count(".") + 1
    if re.match(rules.intro_like_heading_re, text):
        return 1
    return None


def classify_heading(paragraph: Paragraph, text: str, rules: Rules = RULES) -> Optional[int]:
    style_name = paragraph_style_name(paragraph)
    if is_heading_style(style_name):
        return heading_level_from_style(style_name) or heading_level_from_number(text, rules) or 1
    level = heading_level_from_number(text, rules)
    if level is not None and len(text) <= 160:
        return level
    return None


def get_xml_attr(element: Any, path: str, attr: str) -> Optional[str]:
    found = element.xpath(path)
    if not found:
        return None
    return found[0].get(qn(attr))


def run_font_name(run: Any, paragraph: Paragraph) -> Optional[str]:
    explicit = run.font.name
    if explicit:
        return explicit
    rpr = run._r.rPr
    if rpr is not None:
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            value = get_xml_attr(rpr, ".//w:rFonts", attr)
            if value:
                return value
    try:
        if paragraph.style and paragraph.style.font and paragraph.style.font.name:
            return paragraph.style.font.name
    except Exception:
        pass
    return None


def run_font_size_pt(run: Any, paragraph: Paragraph) -> Optional[float]:
    if run.font.size is not None:
        return pt_value(run.font.size)
    try:
        if paragraph.style and paragraph.style.font and paragraph.style.font.size:
            return pt_value(paragraph.style.font.size)
    except Exception:
        pass
    return None


def para_first_line_indent_cm(paragraph: Paragraph) -> Optional[float]:
    pf = paragraph.paragraph_format
    if pf.first_line_indent is not None:
        return cm_value(pf.first_line_indent)
    try:
        if paragraph.style and paragraph.style.paragraph_format.first_line_indent is not None:
            return cm_value(paragraph.style.paragraph_format.first_line_indent)
    except Exception:
        pass
    return None


def para_line_spacing(paragraph: Paragraph) -> Optional[float]:
    pf = paragraph.paragraph_format
    if isinstance(pf.line_spacing, float):
        return pf.line_spacing
    try:
        if paragraph.style and isinstance(paragraph.style.paragraph_format.line_spacing, float):
            return paragraph.style.paragraph_format.line_spacing
    except Exception:
        pass
    return None


def para_alignment(paragraph: Paragraph) -> Optional[int]:
    if paragraph.alignment is not None:
        return paragraph.alignment
    try:
        if paragraph.style and paragraph.style.paragraph_format.alignment is not None:
            return paragraph.style.paragraph_format.alignment
    except Exception:
        pass
    return None


def paragraph_has_numbering(paragraph: Paragraph) -> bool:
    return bool(paragraph._p.xpath("./w:pPr/w:numPr"))


def paragraph_has_section_break(paragraph: Paragraph) -> bool:
    return bool(paragraph._p.xpath("./w:pPr/w:sectPr"))


def xml_onoff_value(parent: Any, xpath: str) -> Optional[bool]:
    """Read WordprocessingML on/off property, e.g. w:b.

    Returns:
        True  - property is present and enabled;
        False - property is present and explicitly disabled;
        None  - property is not specified here.
    """
    if parent is None:
        return None
    try:
        found = parent.xpath(xpath)
    except Exception:
        return None
    if not found:
        return None
    val = found[0].get(qn("w:val"))
    if val is None:
        return True
    return val.lower() not in {"0", "false", "off", "no"}


def style_bold_value(style: Any) -> Optional[bool]:
    """Try to resolve bold from a style and its base styles.

    python-docx often returns run.bold == None for headings because boldness is
    inherited from Heading 1/Heading 2 style. This helper follows the style chain.
    """
    seen: set[int] = set()
    current = style
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        try:
            if current.font.bold is not None:
                return bool(current.font.bold)
        except Exception:
            pass
        try:
            value = xml_onoff_value(current._element, ".//w:rPr/w:b")
            if value is not None:
                return value
        except Exception:
            pass
        try:
            current = current.base_style
        except Exception:
            current = None
    return None


def run_effective_bold(run: Any, paragraph: Paragraph) -> Optional[bool]:
    """Best-effort effective bold detection for a run.

    Returns None when boldness is unknown rather than treating inheritance as
    a formatting violation.
    """
    if run.bold is not None:
        return bool(run.bold)

    try:
        value = xml_onoff_value(run._r.rPr, ".//w:b")
        if value is not None:
            return value
    except Exception:
        pass

    try:
        value = style_bold_value(run.style)
        if value is not None:
            return value
    except Exception:
        pass

    try:
        value = style_bold_value(paragraph.style)
        if value is not None:
            return value
    except Exception:
        pass

    return None


def is_mostly_bold(paragraph: Paragraph) -> Optional[bool]:
    """Return True/False if confidently known, None if ambiguous.

    This avoids false warnings for headings whose boldness is inherited from
    Word styles and not written on every run.
    """
    visible_runs = [r for r in paragraph.runs if normalize_text(r.text)]
    if not visible_runs:
        value = style_bold_value(paragraph.style)
        return value

    bold_chars = 0
    nonbold_chars = 0
    unknown_chars = 0
    for run in visible_runs:
        text_len = len(normalize_text(run.text))
        value = run_effective_bold(run, paragraph)
        if value is True:
            bold_chars += text_len
        elif value is False:
            nonbold_chars += text_len
        else:
            unknown_chars += text_len

    known_chars = bold_chars + nonbold_chars
    if known_chars == 0:
        return None
    return bold_chars >= nonbold_chars


def has_page_field(paragraph: Paragraph) -> bool:
    xml = paragraph._p.xml
    return ("PAGE" in xml and ("instrText" in xml or "fldSimple" in xml or "fldChar" in xml))


def find_main_start(blocks: list[Block]) -> Optional[int]:
    # Основна перевірка тексту починається з реального ВСТУПУ або РОЗДІЛУ 1.
    # У документах із Word-TOC записи змісту часто не є звичайними параграфами, але
    # якщо будуть — нижче є обмеження на heading/style/короткий текст.
    candidates: list[int] = []
    for i, b in enumerate(blocks):
        t = b.text.strip()
        if re.fullmatch(r"(?i)ВСТУП", t) and b.kind == "heading":
            candidates.append(i)
        elif re.match(r"^\s*РОЗДІЛ\s+1\b", t) and b.kind == "heading":
            candidates.append(i)
    return candidates[0] if candidates else None


def assign_contexts(blocks: list[Block], main_start: Optional[int]) -> None:
    current_section = ""
    current_subsection = ""
    for i, b in enumerate(blocks):
        b.in_front_matter = main_start is not None and i < main_start
        if b.in_front_matter:
            b.context = "титульна/службова частина"
            continue
        if b.kind == "heading":
            lvl = b.heading_level or 99
            if lvl == 1:
                current_section = b.text
                current_subsection = ""
            elif lvl >= 2:
                current_subsection = b.text
        parts = []
        if current_section:
            parts.append(current_section)
        if current_subsection:
            parts.append(current_subsection)
        b.context = " > ".join(parts) if parts else "основна частина"


def assign_section_numbers(blocks: list[Block]) -> None:
    section_no = 1
    for b in blocks:
        b.section_no = section_no
        if b.kind in {"paragraph", "heading", "figure_caption", "table_caption", "image_paragraph", "empty"}:
            try:
                if paragraph_has_section_break(b.obj):
                    section_no += 1
            except Exception:
                pass


def build_blocks(doc: _Document, rules: Rules = RULES) -> tuple[list[Block], Optional[int]]:
    blocks: list[Block] = []
    paragraph_no = 0

    figure_caption_re = re.compile(rules.figure_caption_re)
    table_caption_re = re.compile(rules.table_caption_re)

    for item in iter_block_items(doc):
        idx = len(blocks) + 1
        if isinstance(item, Paragraph):
            paragraph_no += 1
            text = normalize_text(item.text)
            kind = "paragraph"
            caption_num = None
            heading_level = None

            if not text and not paragraph_has_image(item):
                kind = "empty"
            else:
                fig = figure_caption_re.match(text)
                tbl = table_caption_re.match(text)
                heading_level = classify_heading(item, text, rules)
                if fig:
                    kind = "figure_caption"
                    caption_num = fig.group("num")
                    heading_level = None
                elif tbl:
                    kind = "table_caption"
                    caption_num = tbl.group("num")
                    heading_level = None
                elif heading_level is not None:
                    kind = "heading"
                elif paragraph_has_image(item):
                    kind = "image_paragraph"

            blocks.append(Block(idx, kind, text, item, paragraph_no, heading_level, caption_num))
        elif isinstance(item, Table):
            blocks.append(Block(idx, "table", "", item))

    assign_section_numbers(blocks)
    main_start = find_main_start(blocks)
    assign_contexts(blocks, main_start)
    return blocks, main_start


def previous_meaningful(blocks: list[Block], pos: int) -> Optional[Block]:
    for j in range(pos - 1, -1, -1):
        if blocks[j].kind != "empty" and not blocks[j].in_front_matter:
            return blocks[j]
    return None


def next_meaningful(blocks: list[Block], pos: int) -> Optional[Block]:
    for j in range(pos + 1, len(blocks)):
        if blocks[j].kind != "empty" and not blocks[j].in_front_matter:
            return blocks[j]
    return None


def block_location(block: Block, label: str = "") -> str:
    parts = []
    if block.context:
        parts.append(block.context)
    if label:
        parts.append(label)
    elif block.kind in {"figure_caption", "table_caption"} and block.caption_num:
        parts.append(f"підпис {'рисунка' if block.kind == 'figure_caption' else 'таблиці'} {block.caption_num}")
    if block.text:
        parts.append(f"Пошук у Word: {short_quote(block.text)}")
    tech = []
    if block.paragraph_no is not None:
        tech.append(f"абз. {block.paragraph_no}")
    tech.append(f"елем. {block.index}")
    if block.section_no:
        tech.append(f"Word-секція {block.section_no}")
    parts.append("Технічно: " + ", ".join(tech))
    return "\n".join(parts)


def section_descriptions(blocks: list[Block], sections_count: int) -> dict[int, str]:
    result: dict[int, str] = {}
    for sec in range(1, sections_count + 1):
        sec_blocks = [b for b in blocks if b.section_no == sec and b.kind != "empty"]
        if not sec_blocks:
            result[sec] = "порожня або службова Word-секція"
            continue
        first = next((b for b in sec_blocks if b.text), sec_blocks[0])
        last = next((b for b in reversed(sec_blocks) if b.text), sec_blocks[-1])
        result[sec] = f"від {short_quote(first.text or first.kind)} до {short_quote(last.text or last.kind)}"
    return result


# -----------------------------
# Перевірки
# -----------------------------

def read_docx_page_count_from_props(input_path: Path) -> Optional[int]:
    try:
        with zipfile.ZipFile(input_path) as zf:
            with zf.open("docProps/app.xml") as f:
                root = ET.parse(f).getroot()
        ns = {"ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"}
        pages = root.find("ep:Pages", ns)
        if pages is not None and pages.text and pages.text.strip().isdigit():
            return int(pages.text.strip())
    except Exception:
        return None
    return None


def read_docx_page_count_via_libreoffice(input_path: Path) -> Optional[int]:
    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        return None
    try:
        from pypdf import PdfReader  # optional dependency, listed in requirements from v3.7
    except Exception:
        return None

    with tempfile.TemporaryDirectory(prefix="docx_validator_") as tmp:
        tmp_path = Path(tmp)
        cmd = [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(tmp_path),
            str(input_path),
        ]
        try:
            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, timeout=90)
        except Exception:
            return None

        pdfs = list(tmp_path.glob("*.pdf"))
        if not pdfs:
            return None
        try:
            reader = PdfReader(str(pdfs[0]))
            return len(reader.pages)
        except Exception:
            return None


def read_docx_page_count(input_path: Path) -> tuple[Optional[int], str]:
    pages = read_docx_page_count_from_props(input_path)
    if pages is not None:
        return pages, "метадані DOCX"
    pages = read_docx_page_count_via_libreoffice(input_path)
    if pages is not None:
        return pages, "рендер LibreOffice → PDF"
    return None, "не визначено"


def check_page_count(input_path: Path, reporter: Reporter, rules: Rules = RULES) -> None:
    pages, source = read_docx_page_count(input_path)
    if pages is None:
        reporter.warning(
            "не вдалося визначити кількість сторінок",
            "Обсяг пояснювальної записки",
            [
                "Для автоматичного підрахунку потрібні або метадані DOCX, або встановлений LibreOffice + pypdf.",
                "Без цього DOCX не дає надійної фактичної кількості сторінок.",
            ],
        )
        return
    if pages < rules.min_pages or pages > rules.max_pages:
        reporter.error(
            f"обсяг документа: {pages} стор., очікується {rules.min_pages}–{rules.max_pages} стор.",
            "Обсяг пояснювальної записки",
            [f"Джерело підрахунку: {source}."],
        )
    else:
        reporter.info(
            f"обсяг документа: {pages} стор.; норма {rules.min_pages}–{rules.max_pages} стор.",
            "Обсяг пояснювальної записки",
            [f"Джерело підрахунку: {source}."],
        )


def check_sections(doc: _Document, blocks: list[Block], reporter: Reporter, rules: Rules = RULES) -> None:
    desc = section_descriptions(blocks, len(doc.sections))
    for i, section in enumerate(doc.sections, start=1):
        if rules.skip_first_word_section and i == 1:
            reporter.info(
                "Word-секцію 1 пропущено з перевірки як титульну/службову частину",
                f"Word-секція 1\nОхоплює: {desc.get(1, '')}",
            )
            continue

        location = f"Word-секція {i}\nОхоплює: {desc.get(i, '')}"
        width = cm_value(section.page_width)
        height = cm_value(section.page_height)
        left = cm_value(section.left_margin)
        right = cm_value(section.right_margin)
        top = cm_value(section.top_margin)
        bottom = cm_value(section.bottom_margin)

        if not approx_equal(width, rules.page_width_cm, rules.tolerance_cm) or not approx_equal(height, rules.page_height_cm, rules.tolerance_cm):
            report_issue(
                reporter,
                rules.severity_page_setup,
                f"формат сторінки ≈ {width:.2f}×{height:.2f} см, очікується A4 portrait {rules.page_width_cm}×{rules.page_height_cm} см",
                location,
            )

        expected_margins = [
            ("ліве поле Word-секції", left, rules.margin_left_cm),
            ("праве поле Word-секції", right, rules.margin_right_cm),
            ("верхнє поле Word-секції", top, rules.margin_top_cm),
            ("нижнє поле Word-секції", bottom, rules.margin_bottom_cm),
        ]
        for name, actual, expected in expected_margins:
            if not approx_equal(actual, expected, rules.tolerance_cm):
                actual_str = "невідомо" if actual is None else f"{actual:.2f} см"
                report_issue(reporter, rules.severity_page_setup, f"{name}: {actual_str}, очікується {expected:.2f} см", location)


def find_content_block(blocks: list[Block]) -> Optional[Block]:
    for b in blocks:
        if re.fullmatch(r"(?i)\s*ЗМІСТ\s*", b.text or ""):
            return b
    return None


def alignment_name(alignment: Optional[int]) -> str:
    names = {
        WD_ALIGN_PARAGRAPH.LEFT: "ліворуч",
        WD_ALIGN_PARAGRAPH.CENTER: "по центру",
        WD_ALIGN_PARAGRAPH.RIGHT: "праворуч",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "по ширині",
    }
    return names.get(alignment, "невідомо")


def page_number_position_name(position: str) -> str:
    position = (position or "TOP").upper()
    if position == "BOTTOM":
        return "знизу"
    if position == "ANY":
        return "у верхньому або нижньому колонтитулі"
    return "зверху"


def page_number_containers(section: Any, rules: Rules) -> list[tuple[str, list[Paragraph]]]:
    """Return candidate header/footer containers according to rules.page_number_position."""
    position = (rules.page_number_position or "TOP").upper()
    containers: list[tuple[str, Any]] = []
    if position in {"TOP", "ANY"}:
        containers.append(("верхній колонтитул", section.header))
    if position in {"BOTTOM", "ANY"}:
        containers.append(("нижній колонтитул", section.footer))
    if not containers:
        containers.append(("верхній колонтитул", section.header))

    result: list[tuple[str, list[Paragraph]]] = []
    for label, container in containers:
        try:
            result.append((label, [p for p in container.paragraphs if has_page_field(p)]))
        except Exception:
            result.append((label, []))
    return result


def find_valid_page_number_container(section: Any, rules: Rules) -> tuple[list[tuple[str, list[Paragraph]]], bool, bool]:
    """Return containers, whether a PAGE field exists, and whether alignment matches."""
    containers = page_number_containers(section, rules)
    has_any = any(paragraphs for _, paragraphs in containers)
    has_valid_alignment = any(
        any(para_alignment(p) == rules.page_number_alignment for p in paragraphs)
        for _, paragraphs in containers
    )
    return containers, has_any, has_valid_alignment


def page_number_location(section_no: int, content_block: Optional[Block], rules: Rules, container_label: Optional[str] = None) -> str:
    expected = f"очікується: {page_number_position_name(rules.page_number_position)}, {alignment_name(rules.page_number_alignment)}"
    base = f"Word-секція {section_no}"
    if container_label:
        base += f" > {container_label}"
    if content_block is not None:
        base += f"\nБлок змісту: {short_quote(content_block.text)}"
    base += f"\n{expected}"
    return base


def check_page_numbering(doc: _Document, blocks: list[Block], reporter: Reporter, rules: Rules = RULES) -> None:
    # DOCX не дає повного рендеру сторінок, але можна перевірити поле PAGE у header/footer.
    content_block = find_content_block(blocks)
    if content_block is None:
        report_issue(
            reporter,
            rules.severity_page_number_missing,
            "не знайдено блок 'ЗМІСТ', тому неможливо перевірити старт нумерації сторінок від змісту",
            "Нумерація сторінок",
        )
        return

    content_section = content_block.section_no or 1
    expected_text = f"Вимога: нумерація сторінок починаючи зі змісту, {page_number_position_name(rules.page_number_position)}, {alignment_name(rules.page_number_alignment)}."

    # Секції до секції зі змістом не повинні нас цікавити як титульні/службові.
    if content_section <= len(doc.sections):
        section = doc.sections[content_section - 1]
        containers, has_any, has_valid_alignment = find_valid_page_number_container(section, rules)
        location = page_number_location(content_section, content_block, rules)
        if not has_any:
            checked = ", ".join(label for label, _ in containers)
            report_issue(
                reporter,
                rules.severity_page_number_missing,
                f"не знайдено номер сторінки у потрібному місці секції, де розміщено ЗМІСТ: перевірено {checked}",
                location,
                ["Перевірка приблизна: шукається поле PAGE у відповідному колонтитулі Word-секції."],
            )
        elif not has_valid_alignment:
            found_places = "; ".join(
                f"{label}: {', '.join(alignment_name(para_alignment(p)) for p in paragraphs)}"
                for label, paragraphs in containers if paragraphs
            )
            report_issue(
                reporter,
                rules.severity_page_number_alignment,
                "номер сторінки є, але його розташування/вирівнювання не відповідає конфігу",
                location,
                [expected_text, f"Знайдено: {found_places}."],
            )
        else:
            valid_label = next(
                label for label, paragraphs in containers
                if any(para_alignment(p) == rules.page_number_alignment for p in paragraphs)
            )
            reporter.info(
                "номер сторінки знайдено у потрібному місці та з правильним вирівнюванням",
                page_number_location(content_section, content_block, rules, valid_label),
            )

        # Якщо та сама Word-секція починається раніше за ЗМІСТ і вже має номер сторінки,
        # то зазвичай нумерація застосовується не лише від змісту, а й до службових сторінок перед ним.
        first_in_section = next((b for b in blocks if b.section_no == content_section and b.kind != "empty" and b.text), None)
        if has_any and first_in_section and first_in_section.index < content_block.index:
            report_issue(
                reporter,
                rules.severity_page_number_section_start,
                "Word-секція з нумерацією сторінок починається раніше за ЗМІСТ",
                location,
                [
                    f"Перший текст цієї секції: {short_quote(first_in_section.text)}.",
                    "Щоб нумерація точно починалася зі змісту, зазвичай потрібен розрив секції перед ЗМІСТ і окремий header/footer.",
                ],
            )

    # Для наступних секцій теж перевіряємо номер сторінки у місці, заданому конфігом.
    for i in range(content_section + 1, len(doc.sections) + 1):
        section = doc.sections[i - 1]
        containers, has_any, has_valid_alignment = find_valid_page_number_container(section, rules)
        location = page_number_location(i, None, rules)
        if not has_any:
            checked = ", ".join(label for label, _ in containers)
            report_issue(
                reporter,
                rules.severity_page_number_missing,
                f"не знайдено номер сторінки у потрібному місці: перевірено {checked}",
                location,
            )
        elif not has_valid_alignment:
            found_places = "; ".join(
                f"{label}: {', '.join(alignment_name(para_alignment(p)) for p in paragraphs)}"
                for label, paragraphs in containers if paragraphs
            )
            report_issue(
                reporter,
                rules.severity_page_number_alignment,
                "номер сторінки є, але його розташування/вирівнювання не відповідає конфігу",
                location,
                [expected_text, f"Знайдено: {found_places}."],
            )

def check_paragraph_formatting(blocks: list[Block], reporter: Reporter, rules: Rules = RULES, bibliography_start: Optional[int] = None) -> None:
    for pos, block in enumerate(blocks):
        if block.in_front_matter:
            continue
        if bibliography_start is not None and pos > bibliography_start:
            continue
        if block.kind != "paragraph":
            continue
        paragraph: Paragraph = block.obj
        if not block.text:
            continue

        location = block_location(block)

        actual_alignment = para_alignment(paragraph)
        if actual_alignment != rules.body_alignment:
            report_issue(reporter, rules.severity_paragraph_alignment, "вирівнювання абзацу не схоже на очікуване", location)

        indent = para_first_line_indent_cm(paragraph)
        if not approx_equal(indent, rules.body_first_line_indent_cm, 0.08):
            got = "не задано" if indent is None else f"{indent:.2f} см"
            report_issue(reporter, rules.severity_paragraph_indent, f"абзацний відступ: {got}, очікується {rules.body_first_line_indent_cm:.2f} см", location)

        spacing = para_line_spacing(paragraph)
        if spacing is not None and abs(spacing - rules.body_line_spacing) > 0.05:
            report_issue(reporter, rules.severity_paragraph_line_spacing, f"міжрядковий інтервал: {spacing}, очікується {rules.body_line_spacing}", location)

        checked_runs = 0
        wrong_fonts: set[str] = set()
        wrong_sizes: set[str] = set()
        for run in paragraph.runs:
            run_text = normalize_text(run.text)
            if not run_text:
                continue
            checked_runs += 1
            font = run_font_name(run, paragraph)
            size = run_font_size_pt(run, paragraph)
            # None = успадковано. Не сваримо, бо це створювало багато шуму.
            if font and font.lower() != rules.body_font_name.lower():
                wrong_fonts.add(font)
            if size is not None and abs(size - rules.body_font_size_pt) > 0.3:
                wrong_sizes.add(f"{size:.1f} pt")

        if checked_runs > 0:
            if wrong_fonts:
                report_issue(reporter, rules.severity_paragraph_font, f"шрифт фрагментів: {', '.join(sorted(wrong_fonts))}; очікується {rules.body_font_name}", location)
            if wrong_sizes:
                report_issue(reporter, rules.severity_paragraph_font_size, f"розмір шрифту фрагментів: {', '.join(sorted(wrong_sizes))}; очікується {rules.body_font_size_pt:.0f} pt", location)


def check_empty_paragraphs(blocks: list[Block], reporter: Reporter, rules: Rules = RULES) -> None:
    streak: list[Block] = []
    for block in blocks + [Block(index=-1, kind="paragraph", text="__END__", obj=None)]:
        if block.in_front_matter:
            continue
        if block.kind == "empty":
            streak.append(block)
            continue
        if len(streak) >= 2:
            before = previous_meaningful(blocks, blocks.index(streak[0])) if streak[0] in blocks else None
            after = block if block.index != -1 else None
            details = []
            if before and before.text:
                details.append(f"Перед пропуском: {short_quote(before.text)}")
            if after and after.text != "__END__":
                details.append(f"Після пропуску: {short_quote(after.text)}")
            nums = [s.paragraph_no for s in streak if s.paragraph_no is not None]
            loc_text = f"порожні абзаци {nums[0]}–{nums[-1]}" if nums else "порожні абзаци"
            report_issue(reporter, rules.severity_empty_paragraphs, "декілька порожніх абзаців підряд; можливо, використано Enter для візуального відступу", loc_text, details)
        streak = []


def collect_refs_before(blocks: list[Block], rules: Rules = RULES) -> tuple[dict[str, int], dict[str, int]]:
    figure_refs: dict[str, int] = {}
    table_refs: dict[str, int] = {}
    figure_ref_re = re.compile(rules.figure_ref_re)
    table_ref_re = re.compile(rules.table_ref_re)

    for block in blocks:
        if block.in_front_matter or block.kind in {"figure_caption", "table_caption", "empty", "table"}:
            continue
        text = block.text
        for m in figure_ref_re.finditer(text):
            figure_refs.setdefault(m.group("num"), block.index)
        for m in table_ref_re.finditer(text):
            table_refs.setdefault(m.group("num"), block.index)
    return figure_refs, table_refs


def check_captions_and_references(blocks: list[Block], reporter: Reporter, rules: Rules = RULES) -> None:
    figure_refs, table_refs = collect_refs_before(blocks, rules)
    seen_figure_captions: dict[str, Block] = {}
    seen_table_captions: dict[str, Block] = {}

    for pos, block in enumerate(blocks):
        if block.in_front_matter:
            continue
        if block.kind == "figure_caption":
            num = block.caption_num or "?"
            location = block_location(block)
            if num in seen_figure_captions:
                report_issue(reporter, rules.severity_caption_duplicate, f"дубльований підпис рисунка {num}", location, [f"Попередній: {block_location(seen_figure_captions[num])}"])
            seen_figure_captions[num] = block
            prev_block = previous_meaningful(blocks, pos)
            if not prev_block or prev_block.kind != "image_paragraph":
                report_issue(reporter, rules.severity_caption_position, f"підпис рисунка {num} не стоїть одразу після рисунка", location)
            if para_alignment(block.obj) != rules.figure_caption_alignment:
                report_issue(reporter, rules.severity_caption_alignment, f"підпис рисунка {num} не має очікуваного вирівнювання", location)
            ref_index = figure_refs.get(num)
            if ref_index is None:
                report_issue(reporter, rules.severity_object_reference_missing, f"для рисунка {num} не знайдено згадки в тексті до його появи", location)
            elif ref_index >= block.index:
                report_issue(reporter, rules.severity_object_reference_missing, f"перша згадка рисунка {num} розміщена не до його появи", location)

        elif block.kind == "table_caption":
            num = block.caption_num or "?"
            location = block_location(block)
            if num in seen_table_captions:
                report_issue(reporter, rules.severity_caption_duplicate, f"дубльований підпис таблиці {num}", location, [f"Попередній: {block_location(seen_table_captions[num])}"])
            seen_table_captions[num] = block
            next_block = next_meaningful(blocks, pos)
            if not next_block or next_block.kind != "table":
                report_issue(reporter, rules.severity_caption_position, f"підпис таблиці {num} не стоїть одразу перед таблицею", location)
            if para_alignment(block.obj) != rules.table_caption_alignment:
                report_issue(reporter, rules.severity_caption_alignment, f"підпис таблиці {num} не має очікуваного вирівнювання", location)
            ref_index = table_refs.get(num)
            if ref_index is None:
                report_issue(reporter, rules.severity_object_reference_missing, f"для таблиці {num} не знайдено згадки в тексті до її появи", location)
            elif ref_index >= block.index:
                report_issue(reporter, rules.severity_object_reference_missing, f"перша згадка таблиці {num} розміщена не до її появи", location)

    for pos, block in enumerate(blocks):
        if block.in_front_matter:
            continue
        if block.kind == "image_paragraph":
            next_block = next_meaningful(blocks, pos)
            if not next_block or next_block.kind != "figure_caption":
                report_issue(reporter, rules.severity_caption_missing, "після рисунка не знайдено коректного підпису формату 'Рис. N — Назва'", block_location(block, "рисунок без розпізнаного підпису"))
        if block.kind == "table":
            prev_block = previous_meaningful(blocks, pos)
            if not prev_block or prev_block.kind != "table_caption":
                report_issue(reporter, rules.severity_caption_missing, "перед таблицею не знайдено коректного підпису формату 'Таблиця N — Назва'", block_location(block, "таблиця без розпізнаного підпису"))


def check_heading_rules(blocks: list[Block], reporter: Reporter, rules: Rules = RULES) -> None:
    for block in blocks:
        if block.in_front_matter or block.kind != "heading":
            continue
        text = block.text
        location = block_location(block, "заголовок")

        m_sec = re.match(rules.section_heading_re, text)
        m_num = re.match(rules.numbered_heading_re, text)

        if m_sec:
            if m_sec.group("dot"):
                report_issue(reporter, rules.severity_heading_number_dot, "після номера розділу не має бути крапки", location, ["Очікувано: 'РОЗДІЛ 1 НАЗВА', а не 'РОЗДІЛ 1. НАЗВА'."])
            # Назва розділу має бути великими літерами. Перевіряємо тільки частину після номера.
            tail = re.sub(r"^\s*РОЗДІЛ\s+\d+\.?\s*", "", text)
            letters = re.sub(r"[^A-Za-zА-Яа-яІіЇїЄєҐґ]", "", tail)
            if letters and letters != letters.upper():
                report_issue(reporter, rules.severity_heading_section_case, "назва розділу не оформлена великими літерами", location)
            bold_state = is_mostly_bold(block.obj)
            if bold_state is False:
                report_issue(reporter, rules.severity_heading_bold, "заголовок розділу не виглядає напівжирним", location)
            elif bold_state is None:
                reporter.info("не вдалося надійно визначити напівжирність заголовка розділу; перевірка пропущена", location)
            continue

        if m_num:
            if m_num.group("dot"):
                report_issue(reporter, rules.severity_heading_number_dot, "після номера підрозділу/пункту не має бути крапки", location, ["Очікувано: '1.1 Назва', а не '1.1. Назва'."])
            title = re.sub(r"^\s*\d+(?:\.\d+){1,5}\.?\s*", "", text)
            if title and title[:1].islower():
                report_issue(reporter, rules.severity_heading_subsection_case, "назва підрозділу має починатися з великої літери", location)
            letters = re.sub(r"[^A-Za-zА-Яа-яІіЇїЄєҐґ]", "", title)
            if len(letters) > 8 and letters == letters.upper():
                report_issue(reporter, rules.severity_heading_subsection_case, "назва підрозділу виглядає як повністю написана великими літерами", location)


def check_section_endings(blocks: list[Block], reporter: Reporter, rules: Rules = RULES) -> None:
    headings = [(i, b) for i, b in enumerate(blocks) if not b.in_front_matter and b.kind == "heading" and b.heading_level is not None]
    for i, (pos, heading) in enumerate(headings):
        level = heading.heading_level or 99
        if level < 2:
            continue
        end = len(blocks)
        for next_pos, next_heading in headings[i + 1:]:
            next_level = next_heading.heading_level or 99
            if next_level <= level:
                end = next_pos
                break
        content = [b for b in blocks[pos + 1:end] if b.kind != "empty" and not b.in_front_matter]
        location = block_location(heading, "підрозділ")
        if not content:
            report_issue(reporter, rules.severity_section_empty, "підрозділ не містить тексту після заголовка", location)
            continue
        last = content[-1]
        if last.kind in {"table", "image_paragraph", "figure_caption", "table_caption"}:
            report_issue(
                reporter,
                rules.severity_section_ends_with_object,
                f"підрозділ закінчується елементом типу {last.kind}; після рисунка/таблиці обов’язково має бути пояснювальний текст",
                location,
                [f"Останній елемент підрозділу: {block_location(last)}"],
            )


def find_bibliography_start(blocks: list[Block], rules: Rules = RULES) -> Optional[int]:
    heading_re = re.compile(rules.bibliography_heading_re)
    for i, block in enumerate(blocks):
        if not block.in_front_matter and heading_re.match(block.text):
            return i
    return None


def extract_bibliography_entries(blocks: list[Block], start_pos: int, rules: Rules = RULES) -> list[tuple[Optional[int], str, Block]]:
    source_start_re = re.compile(rules.source_start_re)
    entries: list[tuple[Optional[int], str, Block]] = []
    current_no: Optional[int] = None
    current_text_parts: list[str] = []
    current_block: Optional[Block] = None
    auto_no = 1

    for block in blocks[start_pos + 1:]:
        if block.kind == "heading":
            break
        if block.kind not in {"paragraph", "empty"}:
            continue
        text = block.text
        if not text:
            continue
        m = source_start_re.match(text)
        numbered_by_word = block.kind == "paragraph" and paragraph_has_numbering(block.obj)
        if m or numbered_by_word:
            if current_text_parts and current_block:
                entries.append((current_no, " ".join(current_text_parts), current_block))
            if m:
                no_str = m.group("bracket") or m.group("plain")
                current_no = int(no_str) if no_str else None
            else:
                current_no = auto_no
            auto_no = (current_no or auto_no) + 1
            current_text_parts = [text]
            current_block = block
        else:
            if current_text_parts:
                current_text_parts.append(text)
            else:
                entries.append((None, text, block))
    if current_text_parts and current_block:
        entries.append((current_no, " ".join(current_text_parts), current_block))
    return entries


def expand_reference_token(token: str) -> set[int]:
    nums: set[int] = set()
    parts = re.split(r"[;,]", token)
    for part in parts:
        part = part.strip()
        if not part:
            continue
        m_range = re.fullmatch(r"(\d+)\s*[-–—]\s*(\d+)", part)
        if m_range:
            a, b = int(m_range.group(1)), int(m_range.group(2))
            if a <= b:
                nums.update(range(a, b + 1))
            else:
                nums.update(range(b, a + 1))
            continue
        for n in re.findall(r"\d+", part):
            nums.add(int(n))
    return nums


def collect_source_references(blocks: list[Block], bibliography_start: Optional[int]) -> tuple[set[int], list[tuple[int, Block, str]]]:
    refs: set[int] = set()
    invalid_patterns: list[tuple[int, Block, str]] = []
    end = bibliography_start if bibliography_start is not None else len(blocks)
    ref_re = re.compile(r"\[([^\[\]]*\d[^\[\]]*)\]")
    for block in blocks[:end]:
        if block.in_front_matter or block.kind not in {"paragraph", "heading"}:
            continue
        for m in ref_re.finditer(block.text):
            content = m.group(1)
            found = expand_reference_token(content)
            if found:
                refs.update(found)
            else:
                invalid_patterns.append((block.index, block, m.group(0)))
    return refs, invalid_patterns


def check_bibliography(blocks: list[Block], reporter: Reporter, rules: Rules = RULES) -> tuple[Optional[int], list[tuple[Optional[int], str, Block]]]:
    start = find_bibliography_start(blocks, rules)
    if start is None:
        report_issue(reporter, rules.severity_bibliography_missing, "не знайдено розділ 'Список використаних джерел' / 'Список літератури'", "Список джерел")
        return None, []

    entries = extract_bibliography_entries(blocks, start, rules)
    if not entries:
        report_issue(reporter, rules.severity_bibliography_empty, "розділ списку джерел знайдено, але джерела не розпізнано", block_location(blocks[start]))
        return start, []

    year_re = re.compile(rules.year_re)
    url_re = re.compile(rules.url_re)
    access_re = re.compile(rules.access_date_re)
    expected_no = 1

    for no, text, block in entries:
        location = block_location(block, "джерело у списку літератури")
        clean = normalize_text(text)
        if no is None:
            report_issue(reporter, rules.severity_source_number_missing, "джерело не починається з номера у форматі '1.' або '[1]' і не розпізнане як Word-автонумерація", location)
        else:
            if no != expected_no:
                report_issue(reporter, rules.severity_source_sequence, f"порушена послідовність нумерації джерел: очікувався номер {expected_no}, знайдено {no}", location)
            expected_no = no + 1

        without_number = re.sub(r"^\s*(?:\[\d+\]|\d+[\.)])\s+", "", clean)
        if re.fullmatch(r"https?://\S+|www\.\S+", without_number, flags=re.I):
            report_issue(reporter, rules.severity_source_bare_url, "голий URL без бібліографічного опису", location)
        if len(without_number) < 25:
            report_issue(reporter, rules.severity_source_too_short, "джерело дуже коротке; ймовірно, бібліографічний опис неповний", location)
        if not year_re.search(clean):
            report_issue(reporter, rules.severity_source_missing_year, "не знайдено рік видання/публікації", location)
        has_url = bool(url_re.search(clean))
        if has_url and not access_re.search(clean):
            report_issue(reporter, rules.severity_source_missing_access_date, "схоже на електронне джерело, але не знайдено дату звернення", location)
        if has_url and not re.search(r"(?i)https?://|www\.|doi", clean):
            report_issue(reporter, rules.severity_source_missing_url, "є позначка URL/режим доступу, але не знайдено самого посилання або DOI", location)
        looks_like_article = bool(re.search(r"(?i)(журнал|збірник|conference|конференц|науков|вісник|proceedings)", clean))
        if looks_like_article and not re.search(r"\b[СCсc]\.?", clean):
            report_issue(reporter, rules.severity_source_article_pages, "схоже на статтю/тези, але не знайдено сторінки типу 'С. 10–15'", location)

    reporter.info(f"розпізнано джерел: {len(entries)}; перевірка ДСТУ виконана евристично", block_location(blocks[start]))
    return start, entries


def check_source_references(blocks: list[Block], bibliography_start: Optional[int], entries: list[tuple[Optional[int], str, Block]], reporter: Reporter, rules: Rules = RULES) -> None:
    if bibliography_start is None or not entries:
        return
    refs, invalid = collect_source_references(blocks, bibliography_start)
    for _, block, raw in invalid:
        report_issue(reporter, rules.severity_source_ref_suspicious, f"підозріле посилання на джерело {raw}; не вдалося розібрати номер", block_location(block))
    source_numbers = {no for no, _, _ in entries if no is not None}
    if not refs:
        report_issue(reporter, rules.severity_source_refs_missing_all, "у тексті не знайдено посилань на джерела у форматі [1], [2], [4–6]", "Посилання на джерела")
        return
    for ref in sorted(refs):
        if ref not in source_numbers:
            report_issue(reporter, rules.severity_source_ref_unknown_number, f"у тексті є посилання [{ref}], але такого номера немає у списку джерел", "Посилання на джерела")
    missing = sorted(source_numbers - refs)
    if missing:
        joined = ", ".join(f"[{n}]" for n in missing)
        report_issue(reporter, rules.severity_source_missing_in_text, f"джерела зі списку не згадані в тексті: {joined}", "Посилання на джерела")
    else:
        reporter.info("усі розпізнані джерела мають згадки в тексті", "Посилання на джерела")


def run_validation(input_path: Path, output_path: Path, rules: Rules = RULES) -> int:
    reporter = Reporter()
    if not input_path.exists():
        print(f"Файл не знайдено: {input_path}", file=sys.stderr)
        return 2
    try:
        doc = Document(str(input_path))
    except Exception as exc:
        print(f"Не вдалося відкрити DOCX: {exc}", file=sys.stderr)
        return 2

    blocks, main_start = build_blocks(doc, rules)
    if main_start is not None:
        reporter.info(
            "титульна/службова частина пропущена з перевірок основного тексту",
            block_location(blocks[main_start], "початок основної перевірки"),
        )
    else:
        reporter.warning("не вдалося надійно знайти початок основної частини; титульна частина може потрапити в частину перевірок", "Структура документа")

    # Перевірку кількості сторінок вимкнено з v3.8: це простіше і надійніше
    # контролювати під час ручного перегляду документа.
    if rules.check_page_setup:
        check_sections(doc, blocks, reporter, rules)
    if rules.check_page_numbering:
        check_page_numbering(doc, blocks, reporter, rules)

    bibliography_start: Optional[int] = None
    entries: list[tuple[Optional[int], str, Block]] = []
    if rules.check_bibliography:
        bibliography_start, entries = check_bibliography(blocks, reporter, rules)

    if rules.check_paragraph_formatting:
        check_paragraph_formatting(blocks, reporter, rules, bibliography_start)
    if rules.check_empty_paragraphs:
        check_empty_paragraphs(blocks, reporter, rules)
    if rules.check_headings:
        check_heading_rules(blocks, reporter, rules)
    if rules.check_captions_and_references:
        check_captions_and_references(blocks, reporter, rules)
    if rules.check_section_endings:
        check_section_endings(blocks, reporter, rules)
    if rules.check_source_references:
        check_source_references(blocks, bibliography_start, entries, reporter, rules)

    output_path.write_text(reporter.render(input_path.name), encoding="utf-8")
    return 1 if any(i.severity == "ERROR" for i in reporter.issues) else 0


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="DOCX validator for student report formatting")
    parser.add_argument("docx", type=Path, help="шлях до .docx файлу")
    parser.add_argument("-o", "--output", type=Path, default=None, help="шлях до report.txt")
    parser.add_argument("-c", "--config", type=Path, default=None, help="шлях до YAML/JSON конфігу правил")
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    output = args.output or args.docx.with_suffix(".report.txt")
    try:
        rules = load_rules(args.config)
    except Exception as exc:
        print(f"Не вдалося завантажити конфіг: {exc}", file=sys.stderr)
        return 2
    rc = run_validation(args.docx, output, rules)
    print(f"Звіт збережено: {output}")
    return rc


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
