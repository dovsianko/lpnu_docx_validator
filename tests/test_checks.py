"""Перевірки на згенерованих документах. Кожен тест — одна поведінка (зокрема знайдені баги v3.10)."""


import pytest
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt

from conftest import rules_of
from docbuilder import A, CourseworkBuilder, add_field, good_coursework, png_bytes
from lpnu_docx_validator.checks import find_object_refs
from lpnu_docx_validator.checks.bibliography import parse_source_ref
import re


class Chapter(CourseworkBuilder):
    """Робота «ВСТУП → РОЗДІЛ 1 → 1.1 …». ВИСНОВКИ й джерела дописуються в finish()/save()."""

    finished = False

    def finish(self, sources=("Петренко П. П. Захист інформації : навч. посіб. Львів, 2021. 320 с.",)):
        if not self.finished:
            self.finished = True
            self.heading("ВИСНОВКИ")
            self.body("Висновки.")
            self.bibliography(list(sources))
        return self

    def save(self, path):
        self.finish()
        return super().save(path)


def chapter_doc(*paragraphs):
    b = Chapter().title_page()
    b.toc(["ВСТУП"])
    b.heading("ВСТУП")
    b.body("Вступ до роботи [1].")
    b.heading("РОЗДІЛ 1 АНАЛІЗ")
    b.heading("1.1 Огляд", 2, new_page=False)
    for t in paragraphs:
        b.body(t)
    return b


# ---------------------------------------------------------------------------
# Еталон
# ---------------------------------------------------------------------------

def test_good_coursework_is_clean(tmp_path, check):
    res = check(good_coursework(tmp_path / "good.docx"))
    noisy = [(i.rule, i.message) for i in res.issues if i.severity != "INFO"]
    assert noisy == []
    assert res.model.main_start is not None
    assert res.model.blocks[res.model.main_start].text == "ВСТУП"  # а не запис змісту


# ---------------------------------------------------------------------------
# Баги v3.10
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("content, expected", [
    ("1", [1]), ("1, с. 25", [1]), ("2–4", [2, 3, 4]), ("1; 3", [1, 3]), ("1, 3, 5", [1, 3, 5]),
    ("1, с. 25–27; 3, p. 7", [1, 3]), ("4, т. 2, с. 15", [4]), ("abc1", None), ("0", None), ("с. 5", None),
])
def test_parse_source_ref(content, expected):
    assert parse_source_ref(content) == expected


def test_page_locator_is_not_a_source_number(check):
    res = check(chapter_doc("Як зазначено в [1, с. 25], метод ефективний."))
    assert "source_ref_unknown_number" not in rules_of(res)


def test_article_without_pages_is_reported(tmp_path, check):
    b = chapter_doc("Текст [1], [2].").finish()
    b.p("2. Іваненко І. Методи захисту. Вісник НУЛП. 2022. № 3.")
    res = check(b)
    hits = [i for i in res.issues if i.rule == "source_article_pages"]
    assert len(hits) == 1 and "Іваненко" in hits[0].block.text


@pytest.mark.parametrize("text, nums", [
    ("на рис. 1.1–1.3", ["1.1", "1.2", "1.3"]),
    ("на рисунках 2.1 і 2.2", ["2.1", "2.2"]),
    ("(рис. 1.1, 1.4)", ["1.1", "1.4"]),
    ("рисунок 3.2 та 2 рази", ["3.2"]),
    ("рис. А.1", ["А.1"]),
])
def test_figure_reference_ranges(settings, text, nums):
    assert find_object_refs(re.compile(settings["captions.figure_reference_pattern"]), text) == nums


def test_invalid_config_value_does_not_flood(tmp_path):
    from lpnu_docx_validator.settings import ConfigError, load_settings_file
    p = tmp_path / "c.yaml"
    p.write_text("body_text:\n  alignment: JUSTIFIED\n", encoding="utf-8")
    with pytest.raises(ConfigError):
        load_settings_file(p)


def _header_with_sdt_page_number(doc):
    sec = doc.sections[-1]
    hdr = sec.header._element
    for p in list(hdr):
        hdr.remove(p)
    sdt = parse_xml(
        f'<w:sdt {nsdecls("w")}><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Page Numbers (Top of Page)"/>'
        '</w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:pPr><w:jc w:val="right"/></w:pPr>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '</w:p></w:sdtContent></w:sdt>')
    hdr.append(sdt)


def test_page_number_inside_content_control_is_found(check):
    b = chapter_doc("Текст [1].")
    _header_with_sdt_page_number(b.doc)
    res = check(b)
    assert not {"page_number_missing", "page_number_alignment"} & set(rules_of(res))


def test_page_number_right_via_tab_stops(check):
    """Типовий колонтитул Word: стиль «Header» має табуляції по центру й праворуч; номер після двох Tab."""
    def build(tabs):
        b = chapter_doc("Текст [1].")
        hp = b.doc.sections[-1].header.paragraphs[0]
        hp.style = b.doc.styles["Header"]
        hp.alignment = A.LEFT
        for r in list(hp._p.findall(qn("w:r"))):
            hp._p.remove(r)
        hp.add_run("\t" * tabs)
        add_field(hp, "PAGE", "2")
        return b
    assert "page_number_alignment" not in rules_of(check(build(2)))
    hits = [i for i in check(build(1)).issues if i.rule == "page_number_alignment"]
    assert hits and "по центру" in hits[0].details[0]


def test_page_number_missing_and_wrong_place(check):
    b = chapter_doc("Текст [1].")
    sec = b.doc.sections[-1]
    for p in sec.header.paragraphs:
        p._p.getparent().remove(p._p)
    sec.header._element.append(OxmlElement("w:p"))
    res = check(b)
    assert "page_number_missing" in rules_of(res)
    sec.footer.is_linked_to_previous = False
    add_field(sec.footer.paragraphs[0], "PAGE", "2")
    res = check(b)
    hits = [i for i in res.issues if i.rule == "page_number_alignment"]
    assert hits and "нижній" in hits[0].message


# ---------------------------------------------------------------------------
# Розв'язання стилів
# ---------------------------------------------------------------------------

def test_style_inheritance_through_base_style(check):
    b = chapter_doc()
    st = b.doc.styles.add_style("Мій текст", WD_STYLE_TYPE.PARAGRAPH)
    st.base_style = b.doc.styles["Normal"]
    b.doc.paragraphs[-1]  # noqa: B018
    b.p("Абзац у власному стилі, що успадковує все від «Звичайного» [1].").style = st
    res = check(b)
    assert not {r for r in rules_of(res) if r.startswith("paragraph_")}
    st.font.name = "Arial"
    st.paragraph_format.line_spacing = 1.0
    res = check(b)
    rules = rules_of(res)
    assert "paragraph_font" in rules and "paragraph_line_spacing" in rules


def test_theme_font_is_resolved(check):
    b = chapter_doc("Текст [1].")
    rf = b.doc.styles["Normal"].element.get_or_add_rPr().find(qn("w:rFonts"))
    rf.set(qn("w:hAnsiTheme"), "minorHAnsi")
    res = check(b)
    fonts = [i.message for i in res.issues if i.rule == "paragraph_font"]
    assert fonts and "Times New Roman" not in fonts[0].split(";")[0]


def test_exact_line_spacing_and_spacing_after(check):
    b = chapter_doc()
    p = b.body("Абзац з точним інтервалом [1].")
    p.paragraph_format.line_spacing = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    res = check(b)
    msgs = {i.rule: i.message for i in res.issues}
    assert "точно 18" in msgs["paragraph_line_spacing"]
    assert "8 пт" in msgs["paragraph_spacing"]


def test_space_after_check_can_be_disabled(check, settings):
    b = chapter_doc()
    b.body("Абзац [1].").paragraph_format.space_after = Pt(8)
    settings.values["body_text.space_after_pt"] = None
    assert "paragraph_spacing" not in rules_of(check(b, settings))


def test_list_items_and_code_are_not_body_text(check):
    b = chapter_doc()
    item = b.doc.add_paragraph("пункт списку без відступу", style="List Bullet")
    item.paragraph_format.first_line_indent = Cm(0)
    code = b.p("print('hello')  # лістинг", first_line_indent=Cm(0), alignment=A.LEFT)
    code.runs[0].font.name = "Courier New"
    b.body("Пояснення [1].")
    res = check(b)
    flagged = {i.block.text for i in res.issues if i.rule.startswith("paragraph_") and i.block}
    assert not flagged


def test_inline_code_font_is_allowed(check):
    b = chapter_doc()
    p = b.body("Функція ")
    p.add_run("print()").font.name = "Consolas"
    p.add_run(" виводить текст [1].")
    assert "paragraph_font" not in rules_of(check(b))


# ---------------------------------------------------------------------------
# Рисунки й таблиці
# ---------------------------------------------------------------------------

def test_figure_and_caption_in_layout_table(check):
    b = chapter_doc("Схему наведено на рис. 1.1.")
    t = b.doc.add_table(rows=2, cols=1)
    t.cell(0, 0).paragraphs[0].add_run().add_picture(__import__("io").BytesIO(png_bytes()), width=Cm(10))
    t.cell(1, 0).paragraphs[0].text = "Рис. 1.1 — Схема"
    t.cell(1, 0).paragraphs[0].alignment = A.CENTER
    b.body("Пояснення після рисунка.")
    res = check(b)
    assert not {"caption_missing", "caption_position", "object_reference_missing"} & set(rules_of(res))


def test_missing_caption_and_reference(check):
    b = chapter_doc("Текст без згадок.")
    b.figure(caption=None)
    b.body("Пояснення.")
    b.figure("1.2")
    b.body("Ще пояснення.")
    b.table(caption=None)
    b.body("Після таблиці.")
    rules = rules_of(check(b))
    assert rules.count("caption_missing") == 2
    assert "object_reference_missing" in rules


def test_reference_after_figure(check):
    b = chapter_doc("Вступний текст.")
    b.figure("1.1")
    b.body("На рис. 1.1 показано схему.")
    hits = [i for i in check(b).issues if i.rule == "object_reference_missing"]
    assert hits and "лише після" in hits[0].message


def test_table_continuation_needs_no_caption(check):
    b = chapter_doc("Дані в табл. 1.1.")
    b.table("1.1")
    b.p("Продовження таблиці 1.1", alignment=A.RIGHT, first_line_indent=Cm(0))
    b.table(caption=None)
    b.body("Пояснення.")
    assert "caption_missing" not in rules_of(check(b))


def test_caption_numbering_and_chapter(check):
    b = chapter_doc("Див. рис. 1.1, 1.3 та 2.1.")
    b.figure("1.1"); b.body("Текст.")
    b.figure("1.3"); b.body("Текст.")
    b.figure("2.1"); b.body("Текст.")
    rules = rules_of(check(b))
    assert "caption_numbering_sequence" in rules
    assert "caption_chapter_mismatch" in rules


def test_wide_and_low_res_image(check):
    b = chapter_doc("Див. рис. 1.1.")
    b.figure("1.1", width_cm=19)  # ширина тексту 21 − 2.5 − 1.5 = 17 см
    b.body("Див. рис. 1.2.")
    b.figure("1.2", width_cm=15, px_width=300)  # ≈ 51 DPI
    b.body("Текст.")
    issues = check(b).issues
    assert len([i for i in issues if i.rule == "image_too_wide"]) == 1
    low = [i for i in issues if i.rule == "image_low_resolution"]
    assert len(low) == 1 and "51 DPI" in low[0].message


def test_formula_number_and_reference(check):
    b = chapter_doc("Розрахунок виконують так:")
    b.formula(None)
    b.body("де E — енергія.")
    b.formula("1.2")
    b.body("де m — маса.")
    rules = rules_of(check(b))
    assert "formula_number_missing" in rules
    assert "formula_reference_missing" in rules


# ---------------------------------------------------------------------------
# Структура, заголовки, нумерація
# ---------------------------------------------------------------------------

def test_missing_and_misordered_parts(check):
    b = CourseworkBuilder().title_page()
    b.toc(["ВСТУП"])
    b.heading("ВСТУП")
    b.body("Текст [1].")
    b.bibliography(["Петренко П. П. Захист інформації : навч. посіб. Львів, 2021. 320 с."])
    b.heading("ВИСНОВКИ")
    b.body("Висновки.")
    res = check(b)
    missing = [i.message for i in res.issues if i.rule == "structure_missing"]
    assert any("РОЗДІЛ 1" in m for m in missing)
    assert "structure_order" in rules_of(res)


def test_manual_toc(tmp_path, check):
    res = check(good_coursework(tmp_path / "g.docx", auto_toc=False))
    assert "toc_manual" in rules_of(res)


def test_toc_outdated(check):
    b = chapter_doc("Текст [1].")
    res = check(b)  # зміст містить лише «ВСТУП»
    assert "toc_outdated" in rules_of(res)


def test_heading_rules(check):
    b = Chapter().title_page()
    b.toc(["ВСТУП"])
    b.heading("ВСТУП")
    b.body("Текст [1].")
    b.heading("РОЗДІЛ 1. Аналіз", new_page=False)
    b.heading("1.1. огляд методів.", 2, new_page=False)
    b.body("Текст.")
    b.heading("1.3 Пропущений номер", 2, new_page=False)
    b.body("Текст.")
    rules = rules_of(check(b))
    for r in ("heading_number_dot", "heading_section_case", "heading_subsection_case", "heading_trailing_dot",
              "section_new_page", "heading_numbering_sequence"):
        assert r in rules, r


def test_heading_alignment_optional(check, settings):
    b = chapter_doc("Текст [1].")
    assert "heading_alignment" not in rules_of(check(b))
    settings.values["headings.section_alignment"] = "LEFT"
    assert "heading_alignment" in rules_of(check(b, settings))


def test_section_ends_with_figure(check):
    b = chapter_doc("Див. рис. 1.1.")
    b.figure("1.1")
    b.heading("1.2 Наступний", 2, new_page=False)
    b.body("Текст.")
    assert "section_ends_with_object" in rules_of(check(b))


def test_nested_subsection_is_not_empty(check):
    b = chapter_doc("Текст [1].")
    b.heading("1.2 Методи", 2, new_page=False)
    b.heading("1.2.1 Перший метод", 3, new_page=False)
    b.body("Текст [1].")
    assert "section_empty" not in rules_of(check(b))


def test_empty_paragraphs_but_not_page_breaks(check):
    b = chapter_doc("Текст [1].")
    b.p("")
    b.p("")
    b.body("Ще текст.")
    assert rules_of(check(b)).count("empty_paragraphs") == 1


# ---------------------------------------------------------------------------
# Додатки, джерела, типографіка, залишки
# ---------------------------------------------------------------------------

def test_appendices(check):
    b = chapter_doc("Лістинг наведено в додатках А–Б.").finish()
    b.heading("ДОДАТОК А")
    b.body("Лістинг.")
    b.heading("ДОДАТОК Б")
    b.body("Схема.")
    b.heading("ДОДАТОК Г")
    b.body("Інше.")
    b.heading("ДОДАТОК З")
    msgs = [(i.rule, i.message) for i in check(b).issues if i.rule.startswith("appendix")]
    assert ("appendix_sequence", "очікувався ДОДАТОК В, а знайдено ДОДАТОК Г") in msgs
    assert any(r == "appendix_invalid_letter" and "З" in m for r, m in msgs)
    assert sum(1 for r, _ in msgs if r == "appendix_not_referenced") == 2  # Г і З


def test_latin_appendix_letter(check):
    b = chapter_doc("Див. додаток А.").finish()
    b.heading("ДОДАТОК A")  # латинська A
    assert any("латинська" in i.message for i in check(b).issues if i.rule == "appendix_invalid_letter")


def test_bibliography_checks(check):
    b = chapter_doc("Посилання [1], [2], [3], [4], [9].").finish()
    for t in ("2. https://example.com", "3. OWASP. URL: https://owasp.org", "Без номера взагалі, але довгий опис 2020."):
        b.p(t)
    rules = rules_of(check(b))
    for r in ("source_bare_url", "source_missing_year", "source_missing_access_date", "source_number_missing",
              "source_ref_unknown_number"):
        assert r in rules, r


def test_source_order_rule_is_opt_in(check, settings):
    b = chapter_doc("Спершу [1], потім [3], а тоді [2].").finish()
    b.p("2. Іваненко І. Методи захисту : монографія. Київ, 2020. 200 с.")
    b.p("3. Сидоренко С. Мережі : підручник. Львів, 2019. 150 с.")
    assert "source_ref_order" not in rules_of(check(b))
    settings.severities["source_ref_order"] = "WARNING"
    assert "source_ref_order" in rules_of(check(b, settings))


def test_word_autonumbered_bibliography(check):
    b = chapter_doc("Текст [1] і [2].")
    b.finished = True
    b.heading("ВИСНОВКИ")
    b.body("Висновки.")
    b.heading("СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ")
    for t in ("Петренко П. Захист інформації : навч. посіб. Львів, 2021. 320 с.",
              "Іваненко І. Мережі : підручник. Київ, 2020. 210 с."):
        b.doc.add_paragraph(t, style="List Number")
    res = check(b)
    assert "source_number_missing" not in rules_of(res)


def test_typography(check):
    b = chapter_doc("Текст  з подвійним пробілом , тире - дефіс, \"лапки\" і слово рoзділ [1].")
    b.body("  Відступ пробілами.")
    rules = rules_of(check(b))
    for r in ("double_space", "space_before_punctuation", "hyphen_instead_of_dash", "straight_quotes", "mixed_script",
              "leading_spaces"):
        assert r in rules, r


def test_mixed_script_ignores_hyphenated_terms(check):
    b = chapter_doc("Налаштовано Wi-Fi та IPv4-адресу, SQL-запити [1].")
    assert "mixed_script" not in rules_of(check(b))


def test_document_hygiene(check):
    b = chapter_doc()
    p = b.body("Текст із виправленням [1].")
    ins = parse_xml(f'<w:ins {nsdecls("w")} w:id="1" w:author="x" w:date="2026-01-01T00:00:00Z">'
                    '<w:r><w:t>вставка</w:t></w:r></w:ins>')
    p._p.append(ins)
    p.add_run(" прихований").font.hidden = True
    r = p.add_run(" виділений")
    r.font.highlight_color = 7
    rules = rules_of(check(b))
    for x in ("tracked_changes", "hidden_text", "highlighted_text"):
        assert x in rules, x


def test_disabling_group_and_rule(check, settings):
    b = chapter_doc("Текст  з подвійним пробілом [1].")
    settings.severities["double_space"] = "OFF"
    assert "double_space" not in rules_of(check(b, settings))
    settings.checks["typography"] = False
    settings.severities["double_space"] = "WARNING"
    assert "double_space" not in rules_of(check(b, settings))


def test_not_a_docx(tmp_path, check):
    p = tmp_path / "x.docx"
    p.write_bytes(b"not a zip")
    res = check(p)
    assert res.fatal and "Не вдалося відкрити" in res.fatal
    assert check(tmp_path / "nope.docx").fatal


def test_formula_as_picture_with_number(check):
    import io
    b = chapter_doc("Енергію обчислюють за формулою (1.1):")
    p = b.doc.add_paragraph()
    p.add_run().add_picture(io.BytesIO(png_bytes(400, 60)), width=Cm(6))
    p.add_run("\t(1.1)")
    b.body("де E — енергія.")
    res = check(b)
    assert not {"caption_missing", "formula_number_missing", "formula_reference_missing"} & set(rules_of(res))
    assert any(blk.kind == "formula" for blk in res.model.blocks)


def test_libreoffice_toc_styles_and_sdt(check):
    b = chapter_doc("Текст [1].")
    st = b.doc.styles.add_style("Contents 1", WD_STYLE_TYPE.PARAGRAPH)
    for p in b.doc.paragraphs:
        if p.style.name == "toc 1":
            p.style = st
            for el in list(p._p.iter(qn("w:instrText"))):
                el.text = " HYPERLINK "
    sdt = parse_xml(f'<w:sdt {nsdecls("w")}><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>'
                    '</w:docPartObj></w:sdtPr><w:sdtContent/></w:sdt>')
    b.doc.element.body.insert(0, sdt)
    res = check(b)
    assert "toc_manual" not in rules_of(res)
    assert res.model.blocks[res.model.main_start].text == "ВСТУП"


# ---------------------------------------------------------------------------
# Випадки зі справжніх курсових (листопад 2026)
# ---------------------------------------------------------------------------

def test_caption_without_dash_keeps_full_number(check):
    """«Рис. 1.2 Назва» розпізнавалось як рисунок «1» → хибні дублікати й «немає згадки рисунка 1»."""
    b = chapter_doc("Див. рис. 1.1 і рис. 1.2.")
    b.figure("1.1")
    b.body("Текст.")
    b.figure("1.2")
    b.doc.paragraphs[-1].text = "Рис. 1.2 Ініціалізація бази даних"
    b.body("Текст.")
    res = check(b)
    assert [x.caption_num for x in res.model.blocks if x.kind == "figure_caption"] == ["1.1", "1.2"]
    rules = rules_of(res)
    assert "caption_duplicate" not in rules and "object_reference_missing" not in rules
    assert rules.count("caption_format") == 1  # «Рис. 1.2 Назва» без тире


def test_reference_and_caption_without_dot(check, settings):
    assert find_object_refs(re.compile(settings["captions.figure_reference_pattern"]), "адрес (рис 1.3).") == ["1.3"]
    assert find_object_refs(re.compile(settings["captions.figure_reference_pattern"]), "для користувача 1.3") == []
    b = chapter_doc("Див. рис 1.1.")
    b.figure("1.1")
    b.doc.paragraphs[-1].text = "Рис 1.1. Головна сторінка"
    b.body("Текст.")
    rules = rules_of(check(b))
    assert "caption_missing" not in rules and "caption_format" in rules


def test_toc_heading_style_is_the_toc_title(check):
    b = chapter_doc("Текст [1].")
    st = b.doc.styles["TOC Heading"]
    for p in b.doc.paragraphs:
        if p.text == "ЗМІСТ":
            p.style = st
    res = check(b)
    assert not [i for i in res.issues if i.rule == "structure_missing"]
    assert res.model.toc_pos is not None


def test_image_and_caption_in_one_paragraph(check):
    import io
    b = chapter_doc("Результат показано на рис. 1.1.")
    p = b.doc.add_paragraph()
    p.add_run("<script>alert(1)</script>")
    p.add_run().add_break()
    p.add_run().add_picture(io.BytesIO(png_bytes()), width=Cm(12))
    p.add_run().add_break()
    p.add_run("Рис. 1.1 — Результат виконання XSS")
    p.alignment = A.CENTER
    b.body("Пояснення.")
    rules = rules_of(check(b))
    assert not {"caption_missing", "caption_position", "object_reference_missing"} & set(rules)


def test_caption_above_image_reported_once(check):
    b = chapter_doc("Див. рис. 1.1.")
    b.p("Рис. 1.1 — Схема", alignment=A.CENTER, first_line_indent=Cm(0))
    b.figure(caption=None)
    b.body("Текст.")
    rules = rules_of(check(b))
    assert rules.count("caption_position") == 1 and "caption_missing" not in rules


def test_code_in_regular_font_is_not_body_text(check):
    b = chapter_doc("Образ збирають командою:")
    for line in ("docker login --username AWS --password-stdin \\", "    839070695658.dkr.ecr.eu-west-3.amazonaws.com",
                 "# Отримання офіційного образу", '<script>alert("XSS")</script>'):
        p = b.p(line, first_line_indent=Cm(0), alignment=A.LEFT)
        p.runs[0].font.size = Pt(12)
    b.body("Після цього образ з'являється в репозиторії [1].")
    res = check(b)
    flagged = {i.rule for i in res.issues if i.block is not None and i.block.kind == "code"}
    assert not flagged
    assert not {"leading_spaces", "straight_quotes", "paragraph_font_size"} & set(rules_of(res))


def test_manual_dash_list_items(check):
    b = chapter_doc("Завдання роботи:")
    for t in ("– зібрати образ застосунку;", "– налаштувати мережу.", "а) перший варіант;", "1) другий варіант."):
        p = b.p(t, alignment=A.LEFT)
        p.paragraph_format.first_line_indent = Cm(-0.5)
        p.paragraph_format.left_indent = Cm(1)
    b.body("Текст після списку [1].")
    assert not {"paragraph_indent", "paragraph_alignment"} & set(rules_of(check(b)))


def test_toc_outdated_counts_front_matter_headings(check):
    b = Chapter().title_page()
    b.doc.add_heading("ІНДИВІДУАЛЬНЕ ЗАВДАННЯ", 1)
    b.toc(["ІНДИВІДУАЛЬНЕ ЗАВДАННЯ", "ВСТУП", "РОЗДІЛ 1 АНАЛІЗ", "1.1 Огляд", "ВИСНОВКИ", "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ"])
    b.heading("ВСТУП")
    b.body("Текст [1].")
    b.heading("РОЗДІЛ 1 АНАЛІЗ")
    b.heading("1.1 Огляд", 2, new_page=False)
    b.body("Текст.")
    assert "toc_outdated" not in rules_of(check(b))


# ---------------------------------------------------------------------------
# Профіль «Курсовий проект ІКС» і можливості, які він використовує
# ---------------------------------------------------------------------------

def iks_doc(*, chapter="1 КОНТЕЙНЕРИЗАЦІЯ ЗАСТОСУНКУ", toc_extra=()):
    b = Chapter().title_page()
    task = b.doc.add_heading("ІНДИВІДУАЛЬНЕ ЗАВДАННЯ", 1)  # друга сторінка, ще без номера: переносимо в 1-шу секцію
    sect_p = next(p for p in b.doc.paragraphs if p._p.find(qn("w:pPr")) is not None and p._p.pPr.find(qn("w:sectPr")) is not None)
    sect_p._p.addprevious(task._p)
    b.toc([*toc_extra, "АНОТАЦІЯ", "ВСТУП", "ОСНОВНА (ТЕХНІЧНА) ЧАСТИНА", chapter, "1.1 Збірка образу",
           "ВИСНОВКИ", "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ", "ДОДАТКИ"])
    for st in ("Normal", "Heading 2", "Heading 3"):
        b.doc.styles[st].paragraph_format.first_line_indent = Cm(1)
    b.heading("АНОТАЦІЯ")
    b.body("Анотація роботи [1].")
    b.heading("ВСТУП")
    b.body("Вступ.")
    b.heading("ОСНОВНА (ТЕХНІЧНА) ЧАСТИНА")
    b.heading(chapter, 2, new_page=False)
    b.heading("1.1 Збірка образу", 3, new_page=False)
    b.body("Схему наведено на рис. 1.1.")
    b.figure("1.1")
    b.body("Текст після рисунка.")
    b.finish()
    b.heading("ДОДАТКИ")
    return b


@pytest.fixture
def iks():
    from lpnu_docx_validator.settings import load_profile
    return load_profile("iks_coursework")


def test_iks_profile_accepts_methodology_layout(check, iks):
    res = check(iks_doc(), iks)
    noisy = [(i.rule, i.message) for i in res.issues if i.severity != "INFO" and i.rule not in {"page_count_unknown", "page_count_out_of_range"}]
    assert noisy == []
    ch = next(b for b in res.model.blocks if b.kind == "heading" and b.text.startswith("1 КОНТЕЙНЕРИЗАЦІЯ"))
    assert ch.chapter == "1"  # розділ без слова «РОЗДІЛ», хоча стиль «Заголовок 2»


def test_iks_chapter_dot_and_forbidden_toc_entry(check, iks):
    res = check(iks_doc(chapter="1. КОНТЕЙНЕРИЗАЦІЯ ЗАСТОСУНКУ", toc_extra=["ІНДИВІДУАЛЬНЕ ЗАВДАННЯ"]), iks)
    rules = rules_of(res)
    assert "heading_number_dot" in rules
    assert "toc_forbidden_entry" in rules


def test_margins_at_least(check, settings):
    b = chapter_doc("Текст [1].")
    sec = b.doc.sections[-1]
    sec.right_margin = Cm(1.2)
    settings.values.update({"page.margin_right_cm": 1.0, "page.margin_check": "AT_LEAST"})
    assert "page_setup" not in rules_of(check(b, settings))
    sec.right_margin = Cm(0.8)
    assert "page_setup" in rules_of(check(b, settings))


def test_online_source_without_year(check, settings):
    b = chapter_doc("Текст [1].").finish(sources=["Хмарні обчислення [Електронний ресурс] – Режим доступу до ресурсу: "
                                                  "https://uk.wikipedia.org/wiki/Хмарні_обчислення"])
    assert "source_missing_year" in rules_of(check(b))
    settings.values["bibliography.online_requires_year"] = False
    assert "source_missing_year" not in rules_of(check(b, settings))


def test_manual_toc_body_is_not_a_chapter(check):
    b = Chapter().title_page()
    b.toc(["ВСТУП", "РОЗДІЛ 1 АНАЛІЗ", "ВИСНОВКИ"], auto=False)
    b.heading("ВСТУП")
    b.body("Текст [1].")
    b.heading("РОЗДІЛ 1 АНАЛІЗ")
    b.heading("1.1 Огляд", 2, new_page=False)
    b.body("Текст.")
    assert "structure_order" not in rules_of(check(b))


def test_pattern_labels():
    from lpnu_docx_validator.checks.structure import pattern_label
    assert pattern_label(r"(?i)^\s*основна\s*\(?\s*технічна\s*\)?\s*частина\s*$") == "ОСНОВНА (ТЕХНІЧНА) ЧАСТИНА"
    assert pattern_label(r"^\s*РОЗДІЛ\s+1\b") == "РОЗДІЛ 1"
    assert pattern_label(r"(?#перший розділ)^\s*1\s") == "перший розділ"
