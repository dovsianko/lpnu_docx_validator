"""Модель документа: блоки в порядку появи + розв'язання стилів Word.

python-docx повертає None для властивостей, успадкованих зі стилю, базового стилю
чи ``docDefaults``. ``StyleResolver`` проходить увесь ланцюжок, як це робить Word,
тому перевірки шрифту/інтервалів бачать фактичні значення.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterator, Optional

from docx import Document
from lxml import etree
from docx.document import Document as _Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from .settings import Settings

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
}
W_P, W_TBL, W_SDT, W_SDTCONTENT = qn("w:p"), qn("w:tbl"), qn("w:sdt"), qn("w:sdtContent")
W_R, W_T, W_TAB, W_BR, W_CR = qn("w:r"), qn("w:t"), qn("w:tab"), qn("w:br"), qn("w:cr")
W_NBH = qn("w:noBreakHyphen")
MC_FALLBACK = f"{{{NS['mc']}}}Fallback"
SKIP_TEXT = {qn("w:txbxContent"), qn("w:del"), qn("w:moveFrom"), MC_FALLBACK, qn("w:pPr"), qn("w:rPr")}

TWIPS_PER_CM = 1440 / 2.54
EMU_PER_CM = 360000
EMU_PER_INCH = 914400


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def short_quote(text: str, limit: int = 90) -> str:
    text = normalize_text(text)
    if len(text) <= limit:
        return f"«{text}»"
    return f"«{text[:limit - 1]}…»"


def _onoff(el: Any) -> Optional[bool]:
    if el is None:
        return None
    val = el.get(qn("w:val"))
    if val is None:
        return True
    return val.lower() not in {"0", "false", "off", "no", "none"}


def _attr(el: Any, attr: str) -> Optional[str]:
    return None if el is None else el.get(qn(attr))


# ---------------------------------------------------------------------------
# Текст і обхід XML
# ---------------------------------------------------------------------------

def paragraph_text(p_el: Any) -> str:
    """Видимий текст абзацу, включно з полями, гіперпосиланнями, вставками рецензування.
    Без тексту з текстових полів (у них власні абзаци) і без видаленого тексту."""
    parts: list[str] = []

    def walk(el: Any) -> None:
        for ch in el:
            tag = ch.tag
            if not isinstance(tag, str) or tag in SKIP_TEXT:
                continue
            if tag == W_T:
                parts.append(ch.text or "")
            elif tag == W_TAB:
                parts.append("\t")
            elif tag == W_BR:
                if ch.get(qn("w:type")) in (None, "textWrapping"):
                    parts.append("\n")
            elif tag == W_CR:
                parts.append("\n")
            elif tag == W_NBH:
                parts.append("-")
            else:
                walk(ch)

    walk(p_el)
    return "".join(parts)


def iter_runs(p_el: Any) -> Iterator[Any]:
    """Усі w:r абзацу (у т.ч. в гіперпосиланнях, полях, вставках), крім текстових полів."""
    def walk(el: Any) -> Iterator[Any]:
        for ch in el:
            tag = ch.tag
            if not isinstance(tag, str) or tag in SKIP_TEXT:
                continue
            if tag == W_R:
                yield ch
            else:
                yield from walk(ch)
    yield from walk(p_el)


def run_text(r_el: Any) -> str:
    out = []
    for ch in r_el:
        if ch.tag == W_T:
            out.append(ch.text or "")
        elif ch.tag == W_TAB:
            out.append("\t")
    return "".join(out)


def iter_body_elements(parent: Any) -> Iterator[Any]:
    """w:p та w:tbl у порядку документа, заходячи всередину w:sdt (елементи керування вмістом)."""
    for ch in parent.iterchildren():
        if ch.tag in (W_P, W_TBL):
            yield ch
        elif ch.tag == W_SDT:
            content = ch.find(W_SDTCONTENT)
            if content is not None:
                yield from iter_body_elements(content)
        elif ch.tag == qn("w:customXml"):
            yield from iter_body_elements(ch)


def iter_all_paragraphs(root: Any) -> Iterator[Any]:
    """Усі w:p під елементом (колонтитули, текстові поля, sdt), без дублів із mc:Fallback."""
    def walk(el: Any) -> Iterator[Any]:
        for ch in el:
            if not isinstance(ch.tag, str) or ch.tag == MC_FALLBACK:
                continue
            if ch.tag == W_P:
                yield ch
            yield from walk(ch)
    yield from walk(root)


def xp(el: Any, expr: str) -> list[Any]:
    """XPath з нашими просторами імен (обходить обгортку python-docx)."""
    return etree._Element.xpath(el, expr, namespaces=NS)


WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


def has_image(p_el: Any) -> bool:
    """Рисунок: картинка, діаграма, SmartArt, група фігур (але не просто текстове поле)."""
    for d in xp(p_el, ".//w:drawing"):
        if xp(d, ".//a:blip") or not xp(d, f".//a:graphicData[@uri='{WPS_URI}']"):
            return True
    return bool(xp(p_el, ".//w:pict//v:imagedata | .//w:object//v:imagedata[not(ancestor::w:object//o:OLEObject[starts-with(@ProgID,'Equation')])]"))


def has_formula(p_el: Any) -> bool:
    if xp(p_el, ".//m:oMathPara | .//m:oMath"):
        return True
    return bool(xp(p_el, ".//o:OLEObject[starts-with(@ProgID,'Equation')]"))


def has_page_break(p_el: Any) -> bool:
    return bool(xp(p_el, ".//w:br[@w:type='page']"))


def section_break_type(p_el: Any) -> Optional[str]:
    """Тип розриву секції в кінці абзацу (nextPage за замовчуванням) або None."""
    sect = p_el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}")
    if sect is None:
        return None
    t = sect.find(qn("w:type"))
    return (t.get(qn("w:val")) if t is not None else None) or "nextPage"


def field_instructions(root: Any) -> list[str]:
    out = [normalize_text(x.text or "") for x in root.iter(qn("w:instrText"))]
    out += [normalize_text(x.get(qn("w:instr")) or "") for x in root.iter(qn("w:fldSimple"))]
    return [x for x in out if x]


# ---------------------------------------------------------------------------
# Розв'язання стилів
# ---------------------------------------------------------------------------

class StyleResolver:
    def __init__(self, doc: _Document) -> None:
        styles_el = doc.styles.element
        self.styles: dict[str, Any] = {}
        self.names: dict[str, str] = {}
        self.default_para: Optional[str] = None
        for s in styles_el.findall(qn("w:style")):
            sid = s.get(qn("w:styleId"))
            if not sid:
                continue
            self.styles[sid] = s
            name_el = s.find(qn("w:name"))
            self.names[sid] = name_el.get(qn("w:val")) if name_el is not None else sid
            if s.get(qn("w:type")) == "paragraph" and _onoff_attr(s.get(qn("w:default"))):
                self.default_para = sid
        dd = styles_el.find(qn("w:docDefaults"))
        self.ppr_default = dd.find(f"{qn('w:pPrDefault')}/{qn('w:pPr')}") if dd is not None else None
        self.rpr_default = dd.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}") if dd is not None else None
        self.theme_fonts = _theme_fonts(doc)
        self._chain_cache: dict[Optional[str], list[Any]] = {}

    # --- ланцюжки ---
    def style_chain(self, style_id: Optional[str]) -> list[Any]:
        if style_id in self._chain_cache:
            return self._chain_cache[style_id]
        chain: list[Any] = []
        seen: set[str] = set()
        cur = style_id
        while cur and cur in self.styles and cur not in seen:
            seen.add(cur)
            st = self.styles[cur]
            chain.append(st)
            based = st.find(qn("w:basedOn"))
            cur = based.get(qn("w:val")) if based is not None else None
        self._chain_cache[style_id] = chain
        return chain

    def para_style_id(self, p_el: Any) -> Optional[str]:
        ps = p_el.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
        if ps is not None and ps.get(qn("w:val")) in self.styles:
            return ps.get(qn("w:val"))
        return self.default_para

    def para_style_name(self, p_el: Any) -> str:
        sid = self.para_style_id(p_el)
        return self.names.get(sid, "") if sid else ""

    def ppr_chain(self, p_el: Any) -> list[Any]:
        chain = [p_el.find(qn("w:pPr"))]
        chain += [s.find(qn("w:pPr")) for s in self.style_chain(self.para_style_id(p_el))]
        chain.append(self.ppr_default)
        return [c for c in chain if c is not None]

    def rpr_chain(self, r_el: Any, p_el: Any) -> list[Any]:
        rpr = r_el.find(qn("w:rPr"))
        chain = [rpr]
        if rpr is not None:
            rs = rpr.find(qn("w:rStyle"))
            if rs is not None:
                chain += [s.find(qn("w:rPr")) for s in self.style_chain(rs.get(qn("w:val")))]
        chain += [s.find(qn("w:rPr")) for s in self.style_chain(self.para_style_id(p_el))]
        chain.append(self.rpr_default)
        return [c for c in chain if c is not None]

    @staticmethod
    def first(chain: list[Any], getter: Callable[[Any], Any]) -> Any:
        for el in chain:
            v = getter(el)
            if v is not None:
                return v
        return None

    # --- властивості абзацу ---
    def alignment(self, p_el: Any) -> str:
        val = self.first(self.ppr_chain(p_el), lambda e: _attr(e.find(qn("w:jc")), "w:val"))
        return {
            "both": "JUSTIFY", "distribute": "JUSTIFY", "lowKashida": "JUSTIFY", "mediumKashida": "JUSTIFY",
            "highKashida": "JUSTIFY", "thaiDistribute": "JUSTIFY", "left": "LEFT", "start": "LEFT",
            "center": "CENTER", "right": "RIGHT", "end": "RIGHT",
        }.get(val or "left", "LEFT")

    def first_line_indent_cm(self, p_el: Any) -> float:
        def get(e: Any) -> Optional[float]:
            ind = e.find(qn("w:ind"))
            if ind is None:
                return None
            if ind.get(qn("w:hanging")) is not None:
                return -int(ind.get(qn("w:hanging"))) / TWIPS_PER_CM
            if ind.get(qn("w:firstLine")) is not None:
                return int(ind.get(qn("w:firstLine"))) / TWIPS_PER_CM
            return None
        v = self.first(self.ppr_chain(p_el), get)
        return 0.0 if v is None else v

    def left_indent_cm(self, p_el: Any) -> float:
        def get(e: Any) -> Optional[float]:
            ind = e.find(qn("w:ind"))
            if ind is None:
                return None
            v = ind.get(qn("w:left")) or ind.get(qn("w:start"))
            return int(v) / TWIPS_PER_CM if v is not None else None
        v = self.first(self.ppr_chain(p_el), get)
        return 0.0 if v is None else v

    def line_spacing(self, p_el: Any) -> tuple[str, float]:
        """("multiple", 1.5) або ("exact"/"atLeast", пункти)."""
        def get(e: Any) -> Optional[tuple[str, float]]:
            sp = e.find(qn("w:spacing"))
            if sp is None or sp.get(qn("w:line")) is None:
                return None
            line = int(sp.get(qn("w:line")))
            rule = sp.get(qn("w:lineRule")) or "auto"
            if rule == "auto":
                return ("multiple", round(line / 240, 3))
            return (rule, line / 20)
        v = self.first(self.ppr_chain(p_el), get)
        return v or ("multiple", 1.0)

    def spacing_pt(self, p_el: Any, which: str) -> float:
        def get(e: Any) -> Optional[float]:
            sp = e.find(qn("w:spacing"))
            if sp is None:
                return None
            if _onoff_attr(sp.get(qn(f"w:{which}Autospacing"))):
                return 14.0
            v = sp.get(qn(f"w:{which}"))
            return int(v) / 20 if v is not None else None
        v = self.first(self.ppr_chain(p_el), get)
        return 0.0 if v is None else v

    def page_break_before(self, p_el: Any) -> bool:
        return bool(self.first(self.ppr_chain(p_el), lambda e: _onoff(e.find(qn("w:pageBreakBefore")))))

    def is_list_item(self, p_el: Any) -> bool:
        def get(e: Any) -> Optional[bool]:
            num = e.find(qn("w:numPr"))
            if num is None:
                return None
            nid = num.find(qn("w:numId"))
            return not (nid is not None and nid.get(qn("w:val")) == "0")
        return bool(self.first(self.ppr_chain(p_el), get))

    def outline_level(self, p_el: Any) -> Optional[int]:
        v = self.first(self.ppr_chain(p_el), lambda e: _attr(e.find(qn("w:outlineLvl")), "w:val"))
        try:
            lvl = int(v) if v is not None else None
        except ValueError:
            return None
        return lvl if lvl is not None and lvl < 9 else None

    # --- властивості фрагмента тексту ---
    def font_name(self, r_el: Any, p_el: Any, text: str = "") -> Optional[str]:
        # Кирилиця у Word використовує hAnsi, базова латиниця — ascii.
        slots = ("hAnsi", "ascii") if not text or re.search(r"[^\x00-\x7f]", text) else ("ascii", "hAnsi")
        chain = self.rpr_chain(r_el, p_el)
        for slot in slots:
            def get(e: Any, slot: str = slot) -> Optional[str]:
                rf = e.find(qn("w:rFonts"))
                if rf is None:
                    return None
                theme = rf.get(qn(f"w:{slot}Theme"))
                if theme:
                    return self.theme_fonts.get("major" if theme.startswith("major") else "minor")
                return rf.get(qn(f"w:{slot}"))
            v = self.first(chain, get)
            if v:
                return v
        return None

    def font_size_pt(self, r_el: Any, p_el: Any) -> float:
        v = self.first(self.rpr_chain(r_el, p_el), lambda e: _attr(e.find(qn("w:sz")), "w:val"))
        try:
            return int(v) / 2 if v is not None else 10.0
        except ValueError:
            return 10.0

    def bold(self, r_el: Any, p_el: Any) -> bool:
        return bool(self.first(self.rpr_chain(r_el, p_el), lambda e: _onoff(e.find(qn("w:b")))))

    def run_flag(self, r_el: Any, p_el: Any, tag: str) -> bool:
        return bool(self.first(self.rpr_chain(r_el, p_el), lambda e: _onoff(e.find(qn(tag)))))

    def color(self, r_el: Any, p_el: Any) -> Optional[str]:
        v = self.first(self.rpr_chain(r_el, p_el), lambda e: _attr(e.find(qn("w:color")), "w:val"))
        return None if v in (None, "auto") else v.upper()

    def highlight(self, r_el: Any) -> Optional[str]:
        rpr = r_el.find(qn("w:rPr"))
        if rpr is None:
            return None
        h = rpr.find(qn("w:highlight"))
        if h is not None and h.get(qn("w:val")) not in (None, "none"):
            return h.get(qn("w:val"))
        shd = rpr.find(qn("w:shd"))
        if shd is not None and shd.get(qn("w:fill")) not in (None, "auto", "FFFFFF", "ffffff"):
            return shd.get(qn("w:fill"))
        return None


def _onoff_attr(v: Optional[str]) -> bool:
    return v is not None and v.lower() in {"1", "true", "on"}


def _theme_fonts(doc: _Document) -> dict[str, str]:
    out: dict[str, str] = {}
    try:
        for rel in doc.part.rels.values():
            if rel.reltype.endswith("/theme"):
                root = etree.fromstring(rel.target_part.blob)
                for kind in ("major", "minor"):
                    el = root.find(f".//a:fontScheme/a:{kind}Font/a:latin", NS)
                    if el is not None and el.get("typeface"):
                        out[kind] = el.get("typeface")
    except Exception:
        pass
    return out


# ---------------------------------------------------------------------------
# Блоки документа
# ---------------------------------------------------------------------------

@dataclass
class Block:
    index: int
    kind: str  # paragraph heading figure_caption table_caption image_paragraph table empty toc_entry formula code
    text: str
    el: Any = None                       # w:p або w:tbl
    paragraph: Optional[Paragraph] = None
    raw_text: str = ""
    paragraph_no: Optional[int] = None
    heading_level: Optional[int] = None
    caption_num: Optional[str] = None
    section_no: int = 1
    in_front_matter: bool = False
    in_table: bool = False
    context: str = ""
    chapter: Optional[str] = None        # номер поточного розділу
    region: str = "main"                 # front / main / bibliography / appendix
    is_list: bool = False
    is_code: bool = False
    formula_num: Optional[str] = None
    extra: dict[str, Any] = field(default_factory=dict)

    @property
    def anchor_el(self) -> Any:
        """Абзац, до якого можна прив'язати примітку Word."""
        if self.el is None:
            return None
        if self.el.tag == W_P:
            return self.el
        return next(iter(self.el.iter(W_P)), None)


TOC_STYLE_RE = re.compile(r"(?i)^(toc|зміст|оглавление|table of contents|contents|index)\s*\d+$")
# Ручні списки: «– пункт», «• пункт», «а) пункт», «1) пункт»
MANUAL_LIST_RE = re.compile(r"^\s*(?:[–—\-•●▪◦·*]|[а-яіїєґa-z]\)|\d{1,2}\))\s+\S")
# Код, набраний звичайним шрифтом: команди, конфіги, коментарі
CODE_START_RE = re.compile(
    r"^\s*(?:\$\s|#|//|--\w|<[/!?\w]|[{}\[\]]|(?:sudo|docker|aws|kubectl|git|npm|pip|python3?|curl|wget|ssh|scp|cd|ls|cat|echo|"
    r"apt(?:-get)?|yum|dnf|terraform|helm|mysql|psql|chmod|chown|export|systemctl|nano|vim|ping|nslookup|dig|"
    r"SELECT|INSERT|UPDATE|DELETE|CREATE|ALTER|DROP|FROM|RUN|COPY|ENV|EXPOSE|CMD|ENTRYPOINT|WORKDIR|"
    r"interface|router|ip|no|enable|configure|hostname|access-list)\b)")
CYRILLIC_RE = re.compile(r"[А-Яа-яІіЇїЄєҐґ]")
CODE_CHARS_RE = re.compile(r"[\\{}=<>$|;]|--|://|\w/\w|\(\)|\"\s*:")


def looks_like_code(text: str) -> bool:
    """Евристика для лістингів, набраних не моноширинним шрифтом."""
    t = text.strip()
    if not t:
        return False
    if t.endswith("\\"):
        return True
    if CYRILLIC_RE.search(t):
        return bool(re.match(r"^\s*(#|//)\s", t)) and len(t) < 120  # коментар у лістингу
    return bool(CODE_START_RE.match(t) or CODE_CHARS_RE.search(t))


def is_heading_style(style_name: str) -> bool:
    s = style_name.lower()
    return s.startswith("heading") or s.startswith("заголовок") or s.startswith("titre") or s == "title"


def is_toc_paragraph(p_el: Any, style_name: str) -> bool:
    if TOC_STYLE_RE.match(style_name.strip()):
        return True
    if xp(p_el, ".//w:hyperlink[starts-with(@w:anchor,'_Toc')]"):
        return True
    return any(i.startswith("PAGEREF _Toc") for i in field_instructions(p_el))


@dataclass
class DocModel:
    path: Path
    doc: _Document
    resolver: StyleResolver
    blocks: list[Block]
    main_start: Optional[int]
    toc_pos: Optional[int]
    bib_start: Optional[int]
    appendix_start: Optional[int]
    has_toc_field: bool

    def section_text_width_cm(self, section_no: int) -> Optional[float]:
        try:
            sec = self.doc.sections[section_no - 1]
            return (sec.page_width - sec.left_margin - sec.right_margin) / EMU_PER_CM
        except Exception:
            return None


class Classifier:
    def __init__(self, settings: Settings, resolver: StyleResolver) -> None:
        s = settings
        self.s = s
        self.r = resolver
        self.figure_caption_re = re.compile(s["captions.figure_caption_pattern"])
        self.table_caption_re = re.compile(s["captions.table_caption_pattern"])
        self.section_re = re.compile(s["headings.section_heading_pattern"])
        self.numbered_re = re.compile(s["headings.numbered_heading_pattern"])
        self.intro_re = re.compile(s["headings.intro_like_heading_pattern"])
        self.appendix_re = re.compile(s["appendices.heading_pattern"])
        self.formula_num_re = re.compile(s["formulas.number_pattern"])
        self.code_fonts = {f.lower() for f in s["body_text.code_fonts"]}

    def heading_level(self, p_el: Any, text: str, style_name: str) -> Optional[int]:
        if not text or len(text) > self.s["headings.max_length"]:
            return None
        by_number = self.level_from_number(text)
        if is_heading_style(style_name):
            m = re.search(r"(\d+)$", style_name)
            return int(m.group(1)) if m else (by_number or 1)
        outline = self.r.outline_level(p_el)
        if outline is not None:
            return outline + 1
        return by_number

    def level_from_number(self, text: str) -> Optional[int]:
        if self.section_re.match(text) or self.intro_re.match(text) or self.appendix_re.match(text):
            return 1
        m = self.numbered_re.match(text)
        if m:
            title = text[m.end("num"):].lstrip(". ")
            if title[:1].isupper() and not re.search(r"[,;:]$", text):
                return m.group("num").count(".") + 1
        return None

    def is_code(self, p_el: Any) -> bool:
        runs = [r for r in iter_runs(p_el) if run_text(r).strip()]
        if not runs:
            return False
        fonts = [(self.r.font_name(r, p_el, run_text(r)) or "").lower() for r in runs]
        return sum(1 for f in fonts if f in self.code_fonts) * 2 > len(fonts)

    def paragraph_block(self, p_el: Any, idx: int, paragraph_no: int, in_table: bool = False) -> Block:
        raw = paragraph_text(p_el)
        text = normalize_text(raw)
        style = self.r.para_style_name(p_el)
        b = Block(idx, "paragraph", text, p_el, raw_text=raw, paragraph_no=paragraph_no, in_table=in_table)
        b.is_list = self.r.is_list_item(p_el) or bool(MANUAL_LIST_RE.match(raw))
        img = has_image(p_el)
        if not text and not img and not has_formula(p_el):
            b.kind = "empty"
            b.extra["page_break"] = has_page_break(p_el) or section_break_type(p_el) not in (None, "continuous")
            return b
        if is_toc_paragraph(p_el, style):
            b.kind = "toc_entry"
            return b
        fig = self.figure_caption_re.match(text)
        tbl = self.table_caption_re.match(text)
        if img and not fig:
            # рисунок і підпис в одному абзаці (підпис після Shift+Enter)
            for line in reversed([normalize_text(x) for x in raw.split("\n") if x.strip()]):
                fig = self.figure_caption_re.match(line)
                if fig:
                    b.text = line
                    break
        if fig:
            b.kind, b.caption_num = "figure_caption", fig.group("num")
            if img:
                b.extra["with_image"] = True
            return b
        if tbl:
            b.kind, b.caption_num = "table_caption", tbl.group("num")
            return b
        m_num = self.formula_num_re.search(text)
        if m_num and not self.formula_num_re.sub("", text).strip() and (img or len(text) <= 12):
            # формула картинкою/об'єктом або окремий рядок «(2.1)» поруч із нею
            b.kind, b.formula_num = "formula", m_num.group("num")
            return b
        if has_formula(p_el):
            m = self.formula_num_re.search(text)
            rest = self.formula_num_re.sub("", text).strip() if m else text
            if not rest or xp(p_el, ".//m:oMathPara"):
                b.kind = "formula"
                b.formula_num = m.group("num") if m else None
                return b
        if img:
            b.kind = "image_paragraph"
            return b
        level = self.heading_level(p_el, text, style)
        if level is not None and not in_table:
            b.kind, b.heading_level = "heading", level
            return b
        if self.is_code(p_el) or looks_like_code(raw):
            b.kind, b.is_code = "code", True
        return b


def build_model(path: Path, settings: Settings, doc: Optional[_Document] = None) -> DocModel:
    doc = doc or Document(str(path))
    resolver = StyleResolver(doc)
    cls = Classifier(settings, resolver)
    blocks: list[Block] = []
    paragraph_no = 0
    section_no = 1
    body = doc.element.body

    def last_meaningful() -> Optional[Block]:
        for b in reversed(blocks):
            if b.kind != "empty":
                return b
        return None

    for el in iter_body_elements(body):
        if el.tag == W_P:
            paragraph_no += 1
            b = cls.paragraph_block(el, len(blocks) + 1, paragraph_no)
            b.section_no = section_no
            b.paragraph = Paragraph(el, doc._body)
            blocks.append(b)
            if section_break_type(el):
                section_no += 1
            continue
        # таблиця
        prev = last_meaningful()
        cell_ps = list(el.iter(W_P))
        has_imgs = any(has_image(p) for p in cell_ps)
        layout_figure = (settings["captions.figures_in_layout_tables"] and has_imgs
                         and not (prev and prev.kind == "table_caption"))
        if layout_figure:
            # таблиця-«рамка» для рисунка: розгортаємо рисунки й підписи як звичайні блоки
            for p in cell_ps:
                pb = cls.paragraph_block(p, len(blocks) + 1, None, in_table=True)  # type: ignore[arg-type]
                if pb.kind in {"image_paragraph", "figure_caption"}:
                    pb.section_no = section_no
                    pb.paragraph = Paragraph(p, doc._body)
                    pb.extra["layout_table"] = True
                    blocks.append(pb)
            continue
        tb = Block(len(blocks) + 1, "table", "", el, section_no=section_no)
        blocks.append(tb)

    # продовження лістингу: рядок після «\\» або без кирилиці між двома рядками коду
    for i, b in enumerate(blocks):
        if b.kind != "paragraph" or not b.text or CYRILLIC_RE.search(b.text):
            continue
        prev = blocks[i - 1] if i > 0 else None
        nxt = blocks[i + 1] if i + 1 < len(blocks) else None
        if prev is not None and prev.kind == "code" and (prev.raw_text.rstrip().endswith("\\") or (nxt is not None and nxt.kind == "code")):
            b.kind, b.is_code = "code", True

    # початок основної частини
    toc_re = re.compile(settings["structure.toc_heading_pattern"])
    toc_pos = next((i for i, b in enumerate(blocks) if b.kind in {"heading", "paragraph"} and toc_re.match(b.text)), None)
    if toc_pos is not None:
        # зміст, набраний вручну: рядки «Назва ……… 5» / «Назва<Tab>5» одразу після «ЗМІСТ»
        for b in blocks[toc_pos + 1:]:
            if b.kind in {"empty", "toc_entry"}:
                continue
            if b.kind in {"heading", "paragraph"} and re.search(r"(\t|\.{3,}|…)\s*\d+\s*$", b.raw_text):
                b.kind, b.heading_level = "toc_entry", None
                continue
            break
    main_re = re.compile(settings["structure.main_start_pattern"])
    main_start = next((i for i, b in enumerate(blocks)
                       if b.kind == "heading" and main_re.match(b.text) and (toc_pos is None or i > toc_pos)), None)
    bib_re = re.compile(settings["bibliography.heading_pattern"])
    app_re = cls.appendix_re
    start = main_start or 0
    bib_start = next((i for i in range(start, len(blocks)) if blocks[i].kind in {"heading", "paragraph"} and bib_re.match(blocks[i].text)), None)
    app_from = bib_start if bib_start is not None else start
    appendix_start = next((i for i in range(app_from, len(blocks))
                           if blocks[i].kind in {"heading", "paragraph"} and app_re.match(blocks[i].text)), None)
    if bib_start is not None and blocks[bib_start].kind == "paragraph":
        blocks[bib_start].kind, blocks[bib_start].heading_level = "heading", 1
    if appendix_start is not None:
        for i in range(appendix_start, len(blocks)):
            if blocks[i].kind == "paragraph" and app_re.match(blocks[i].text):
                blocks[i].kind, blocks[i].heading_level = "heading", 1

    has_toc_field = any(i.split(" ")[0].upper() == "TOC" for i in field_instructions(body)) or bool(
        xp(body, ".//w:sdt[w:sdtPr/w:docPartObj/w:docPartGallery[@w:val='Table of Contents']]"))

    # контекст, регіони, номер розділу
    section_title, sub_title, chapter = "", "", None
    for i, b in enumerate(blocks):
        if main_start is not None and i < main_start:
            b.in_front_matter, b.region, b.context = True, "front", "титульна/службова частина"
            continue
        if appendix_start is not None and i >= appendix_start:
            b.region = "appendix"
        elif bib_start is not None and i >= bib_start:
            b.region = "bibliography"
        if b.kind == "heading":
            m = cls.section_re.match(b.text)
            if m or (b.heading_level or 1) == 1:
                # розділ: «РОЗДІЛ 1 …» / «1 НАЗВА» (навіть якщо стиль «Заголовок 2» під «ОСНОВНА ЧАСТИНА») або інша частина 1-го рівня
                section_title, sub_title = b.text, ""
                ma = cls.appendix_re.match(b.text)
                chapter = m.group("num") if m else (ma.group("letter").upper() if ma else None)
            else:
                sub_title = b.text
        b.chapter = chapter
        b.context = " > ".join(x for x in (section_title, sub_title) if x) or "основна частина"

    return DocModel(path, doc, resolver, blocks, main_start, toc_pos, bib_start, appendix_start, has_toc_field)


def block_location(block: Block, label: str = "") -> str:
    parts = []
    if block.context:
        parts.append(f"Розділ: {block.context}")
    if label:
        parts.append(f"Об'єкт: {label}")
    elif block.kind in {"figure_caption", "table_caption"} and block.caption_num:
        parts.append(f"Об'єкт: підпис {'рисунка' if block.kind == 'figure_caption' else 'таблиці'} {block.caption_num}")
    if block.text:
        parts.append(f"Пошук у Word: {short_quote(block.text)}")
    tech = []
    if block.paragraph_no is not None:
        tech.append(f"абз. {block.paragraph_no}")
    if block.in_table:
        tech.append("у таблиці")
    tech.append(f"елем. {block.index}")
    tech.append(f"Word-секція {block.section_no}")
    parts.append("Технічно: " + ", ".join(tech))
    return "\n".join(parts)
