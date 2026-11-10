"""Єдине джерело налаштувань і правил валідатора.

Тут описано *всі* параметри (з типами, значеннями за замовчуванням, підписами
українською) та всі правила (з рівнем серйозності за замовчуванням). З цього
реєстру будуються:

* значення за замовчуванням;
* YAML-конфіг із коментарями (``init-config`` і кнопка «Зберегти» у GUI);
* форма налаштувань у GUI;
* перевірка конфігу з людськими повідомленнями про помилки;
* список правил у документації (``list-rules``).
"""

from __future__ import annotations

import difflib
import os
import re
import sys
import textwrap
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

SEVERITIES = ("ERROR", "WARNING", "INFO", "OFF")
ALIGN_CHOICES = ("LEFT", "CENTER", "RIGHT", "JUSTIFY")
ALIGN_NAMES_UA = {
    "LEFT": "ліворуч",
    "CENTER": "по центру",
    "RIGHT": "праворуч",
    "JUSTIFY": "по ширині",
    "ANY": "будь-яке (не перевіряти)",
}
DEFAULT_PROFILE_ID = "coursework"
DEFAULT_CONFIG_FILENAME = "docx_validator_config.yaml"


# ---------------------------------------------------------------------------
# Опис параметрів і правил
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Param:
    key: str                 # "section.name"
    type: str                # float | int | bool | str | regex | choice | list | regex_list | float_list
    default: Any
    label: str
    help: str = ""
    choices: tuple[str, ...] = ()
    nullable: bool = False   # null = «не перевіряти»
    advanced: bool = False
    min: Optional[float] = None
    max: Optional[float] = None

    @property
    def section(self) -> str:
        return self.key.split(".", 1)[0]

    @property
    def name(self) -> str:
        return self.key.split(".", 1)[1]


@dataclass(frozen=True)
class RuleDef:
    id: str
    group: str
    severity: str
    title: str
    description: str = ""


@dataclass(frozen=True)
class Group:
    id: str
    title: str
    description: str
    enabled: bool = True


GROUPS: tuple[Group, ...] = (
    Group("page_setup", "Сторінка й поля", "Розмір аркуша, орієнтація та поля кожної Word-секції."),
    Group("page_numbering", "Нумерація сторінок", "Поле PAGE у колонтитулах і старт нумерації зі змісту."),
    Group("structure", "Структура роботи", "Обов'язкові частини, їхній порядок і автоматичний зміст."),
    Group("paragraph_formatting", "Основний текст", "Шрифт, розмір, відступ, інтервали та вирівнювання абзаців."),
    Group("empty_paragraphs", "Порожні абзаци", "Кілька Enter підряд замість інтервалів чи розриву сторінки."),
    Group("headings", "Заголовки", "Оформлення заголовків розділів і підрозділів."),
    Group("numbering", "Нумерація заголовків і об'єктів", "Пропуски й невідповідності номерів підрозділів, рисунків, таблиць."),
    Group("captions_and_references", "Рисунки, таблиці, підписи", "Підписи до рисунків і таблиць та згадки про них у тексті."),
    Group("section_endings", "Завершення підрозділів", "Підрозділ не має закінчуватися рисунком/таблицею."),
    Group("formulas", "Формули", "Номери формул і посилання на них."),
    Group("tables", "Таблиці", "Розмір шрифту в таблицях."),
    Group("appendices", "Додатки", "Літери додатків і згадки про них у тексті."),
    Group("bibliography", "Список джерел", "Наявність і формальні ознаки бібліографічних описів."),
    Group("source_references", "Посилання на джерела", "Посилання виду [1], [2–4], [1, с. 25] і їхня відповідність списку."),
    Group("typography", "Типографіка", "Подвійні пробіли, дефіс замість тире, лапки, латинські літери в кириличних словах."),
    Group("document_hygiene", "Залишки редагування", "Невиконані виправлення, примітки, прихований і кольоровий текст."),
    Group("images", "Зображення", "Ширина зображень і їхня роздільна здатність."),
    Group("page_count", "Обсяг роботи", "Кількість сторінок (потрібен LibreOffice).", enabled=False),
)
GROUP_BY_ID = {g.id: g for g in GROUPS}

SECTION_TITLES = {
    "profile": "Профіль",
    "page": "Параметри сторінки",
    "body_text": "Основний текст",
    "page_numbering": "Нумерація сторінок",
    "structure": "Структура роботи",
    "headings": "Заголовки",
    "captions": "Рисунки, таблиці, підписи",
    "formulas": "Формули",
    "tables": "Таблиці",
    "appendices": "Додатки",
    "bibliography": "Список джерел",
    "images": "Зображення",
    "page_count": "Обсяг роботи",
    "report": "Звіт",
}

_BIB_HEADING = (
    r"(?i)^\s*(список\s+використаних\s+джерел|перелік\s+використаних\s+джерел|"
    r"список\s+використаної\s+літератури|список\s+літератури|література|references)\s*$"
)

PARAMS: tuple[Param, ...] = (
    # --- профіль ---
    Param("profile.name", "str", "Курсова робота (типові вимоги)", "Назва профілю",
          "Показується у списку профілів у GUI та у звіті."),
    Param("profile.description", "str", "Поля 2.5/1.5/2/2 см, Times New Roman 14, інтервал 1.5, номер сторінки зверху справа.",
          "Опис профілю", "Коротко: для чого цей профіль."),

    # --- сторінка ---
    Param("page.width_cm", "float", 21.0, "Ширина аркуша, см", "A4 = 21.0", min=5, max=100),
    Param("page.height_cm", "float", 29.7, "Висота аркуша, см", "A4 = 29.7", min=5, max=100),
    Param("page.margin_left_cm", "float", 2.5, "Ліве поле, см", min=0, max=10),
    Param("page.margin_right_cm", "float", 1.5, "Праве поле, см", min=0, max=10),
    Param("page.margin_top_cm", "float", 2.0, "Верхнє поле, см", min=0, max=10),
    Param("page.margin_bottom_cm", "float", 2.0, "Нижнє поле, см", min=0, max=10),
    Param("page.tolerance_cm", "float", 0.08, "Допустима похибка, см",
          "Word зберігає розміри в інших одиницях, тому дрібні розбіжності можливі.", advanced=True, min=0, max=2),
    Param("page.allow_landscape_sections", "bool", True, "Дозволяти окремі секції з альбомною орієнтацією",
          "Для великих таблиць/схем. Поля таких секцій не перевіряються."),
    Param("page.skip_first_word_section", "bool", True, "Не перевіряти першу Word-секцію (титулку)",
          "Титульна сторінка часто має інші поля."),

    # --- основний текст ---
    Param("body_text.font_name", "str", "Times New Roman", "Шрифт основного тексту"),
    Param("body_text.font_size_pt", "float", 14.0, "Розмір шрифту, пт", min=4, max=72),
    Param("body_text.first_line_indent_cm", "float", 1.25, "Абзацний відступ, см", min=0, max=5),
    Param("body_text.indent_tolerance_cm", "float", 0.08, "Допустима похибка відступу, см", advanced=True, min=0, max=2),
    Param("body_text.line_spacing", "float", 1.5, "Міжрядковий інтервал (множник)", "1.0 — одинарний, 1.5 — півтора.", min=0.5, max=5),
    Param("body_text.alignment", "choice", "JUSTIFY", "Вирівнювання основного тексту", choices=ALIGN_CHOICES + ("ANY",)),
    Param("body_text.space_before_pt", "float", 0.0, "Інтервал перед абзацом, пт",
          "Порожнє значення (null) — не перевіряти.", nullable=True, min=0, max=100),
    Param("body_text.space_after_pt", "float", 0.0, "Інтервал після абзацу, пт",
          "Порожнє значення (null) — не перевіряти. У стилі «Звичайний» нового Word стоїть 8 пт.", nullable=True, min=0, max=100),
    Param("body_text.skip_list_items", "bool", True, "Не перевіряти відступ і вирівнювання в пунктах списків",
          "Пункти маркованих/нумерованих списків Word часто мають власні відступи."),
    Param("body_text.code_fonts", "list", ["Courier New", "Consolas", "Cascadia Code", "Cascadia Mono", "Lucida Console",
                                            "JetBrains Mono", "Source Code Pro", "DejaVu Sans Mono", "Menlo", "Monaco"],
          "Моноширинні шрифти (лістинги коду)",
          "Абзаци, набрані цими шрифтами, вважаються кодом і не перевіряються як основний текст.", advanced=True),

    # --- нумерація сторінок ---
    Param("page_numbering.position", "choice", "TOP", "Де має бути номер сторінки",
          "TOP — у верхньому колонтитулі, BOTTOM — у нижньому, ANY — будь-де.", choices=("TOP", "BOTTOM", "ANY")),
    Param("page_numbering.alignment", "choice", "RIGHT", "Вирівнювання номера сторінки", choices=("LEFT", "CENTER", "RIGHT", "ANY")),
    Param("page_numbering.start_from_toc", "bool", True, "Нумерація має починатися зі сторінки ЗМІСТ",
          "Якщо вимкнено — перевіряється лише наявність номера в секціях основної частини."),

    # --- структура ---
    Param("structure.toc_heading_pattern", "regex", r"(?i)^\s*зміст\s*$", "Як розпізнати заголовок змісту", advanced=True),
    Param("structure.main_start_pattern", "regex", r"(?i)^\s*(вступ|розділ\s+1\b.*)\s*$",
          "Звідки починається основна частина",
          "Усе до першого такого заголовка вважається титульною/службовою частиною.", advanced=True),
    Param("structure.required_sections", "regex_list",
          [r"(?i)^\s*зміст\s*$", r"(?i)^\s*вступ\s*$", r"^\s*РОЗДІЛ\s+1\b", r"(?i)^\s*висновки\s*$",
           r"(?i)^\s*(список\s+використаних\s+джерел|перелік\s+використаних\s+джерел|список\s+використаної\s+літератури|список\s+літератури)\s*$"],
          "Обов'язкові частини роботи (по порядку)",
          "Кожен рядок — шаблон заголовка. Порядок у списку = очікуваний порядок у роботі.", advanced=True),
    Param("structure.toc_max_level", "int", 3, "Найглибший рівень заголовків у змісті",
          "Заголовки глибших рівнів не мусять бути в змісті.", advanced=True, min=1, max=9),

    # --- заголовки ---
    Param("headings.section_heading_pattern", "regex", r"^\s*РОЗДІЛ\s+(?P<num>\d+)(?P<dot>\.)?\s+\S.+$",
          "Шаблон заголовка розділу", "Група num — номер, dot — зайва крапка після номера.", advanced=True),
    Param("headings.numbered_heading_pattern", "regex", r"^\s*(?P<num>\d+(?:\.\d+){1,5})(?P<dot>\.)?\s+\S.+$",
          "Шаблон заголовка підрозділу", advanced=True),
    Param("headings.intro_like_heading_pattern", "regex",
          r"(?i)^\s*(зміст|вступ|висновки|висновок|список\s+використаних\s+джерел|перелік\s+використаних\s+джерел|"
          r"список\s+використаної\s+літератури|список\s+літератури|перелік\s+умовних\s+позначень.*|додатки?|додаток\s+\S+.*)\s*$",
          "Шаблон структурних заголовків без номера", advanced=True),
    Param("headings.section_alignment", "choice", "ANY", "Вирівнювання заголовків розділів",
          "ANY — не перевіряти.", choices=ALIGN_CHOICES + ("ANY",)),
    Param("headings.subsection_alignment", "choice", "ANY", "Вирівнювання заголовків підрозділів",
          "ANY — не перевіряти.", choices=ALIGN_CHOICES + ("ANY",)),
    Param("headings.max_length", "int", 160, "Максимальна довжина заголовка, символів",
          "Довші абзаци з номером на початку не вважаються заголовками.", advanced=True, min=20, max=1000),

    # --- підписи ---
    Param("captions.figure_caption_pattern", "regex",
          r"^\s*(?:Рис\.|Рисунок)\s+(?P<num>(?:\d+|[А-ЯҐЄІЇA-Z])(?:\.\d+)*)\s*(?:[-–—.]\s*)\S.*$",
          "Шаблон підпису рисунка", "Дозволяє: «Рис. 1.1 — Назва», «Рис. 1.1. Назва», «Рисунок 1.1 — Назва», у додатках «Рис. А.1 — Назва».", advanced=True),
    Param("captions.table_caption_pattern", "regex",
          r"^\s*Таблиця\s+(?P<num>(?:\d+|[А-ЯҐЄІЇA-Z])(?:\.\d+)*)\s*(?:[-–—.]\s*)\S.*$",
          "Шаблон підпису таблиці", "Дозволяє: «Таблиця 1.1 — Назва», «Таблиця 1.1. Назва».", advanced=True),
    Param("captions.figure_reference_pattern", "regex",
          r"(?i)(?<!\w)(?:рисун\w*|рис\.)\s*(?P<num>\d+(?:\.\d+)*|[А-ЯҐЄІЇA-Z](?:\.\d+)+)",
          "Шаблон згадки рисунка в тексті",
          "Діапазони й переліки («рис. 1.1–1.3», «рисунки 2.1 і 2.2») розбираються автоматично.", advanced=True),
    Param("captions.table_reference_pattern", "regex",
          r"(?i)(?<!\w)(?:таблиц\w*|табл\.)\s*(?P<num>\d+(?:\.\d+)*|[А-ЯҐЄІЇA-Z](?:\.\d+)+)",
          "Шаблон згадки таблиці в тексті", advanced=True),
    Param("captions.table_continuation_pattern", "regex", r"(?i)^\s*(продовження|закінчення)\s+(табл\.|таблиці)",
          "Шаблон напису «Продовження таблиці»", "Перед частиною таблиці з таким написом підпис не вимагається.", advanced=True),
    Param("captions.allow_trailing_dot", "bool", False, "Дозволяти крапку в кінці підпису"),
    Param("captions.figure_alignment", "choice", "CENTER", "Вирівнювання підпису рисунка", choices=ALIGN_CHOICES + ("ANY",)),
    Param("captions.table_alignment", "choice", "RIGHT", "Вирівнювання підпису таблиці", choices=ALIGN_CHOICES + ("ANY",)),
    Param("captions.figures_in_layout_tables", "bool", True, "Дозволяти рисунок з підписом у таблиці без меж",
          "Студенти часто ставлять рисунок і підпис у таблицю для вирівнювання. Якщо увімкнено — це не вважається таблицею."),

    # --- формули ---
    Param("formulas.number_pattern", "regex", r"\(\s*(?P<num>\d+(?:\.\d+)*)\s*\)\s*$", "Шаблон номера формули",
          "Номер праворуч у тому ж абзаці, що й формула: (2.1).", advanced=True),
    Param("formulas.reference_pattern", "regex", r"(?i)(?:формул\w*\s*\(?\s*|\()(?P<num>\d+\.\d+)\s*\)?",
          "Шаблон згадки формули в тексті", advanced=True),

    # --- таблиці ---
    Param("tables.allowed_font_sizes_pt", "float_list", [12.0, 14.0], "Дозволені розміри шрифту в таблицях, пт",
          "По одному значенню на рядок."),

    # --- додатки ---
    Param("appendices.heading_pattern", "regex", r"^\s*ДОДАТОК\s+(?P<letter>[A-ZА-ЯҐЄІЇ])\b", "Шаблон заголовка додатка", advanced=True),
    Param("appendices.forbidden_letters", "str", "ҐЄЗІЇЙОЧЬ", "Літери, якими не можна позначати додатки",
          "За ДСТУ 3008:2015: Ґ, Є, З, І, Ї, Й, О, Ч, Ь."),
    Param("appendices.reference_pattern", "regex", r"(?i)(?<!\w)(?:додат\w*|дод\.)\s+(?P<letter>[A-ZА-ЯҐЄІЇ])(?!\w)",
          "Шаблон згадки додатка в тексті", advanced=True),

    # --- список джерел ---
    Param("bibliography.heading_pattern", "regex", _BIB_HEADING, "Як знайти заголовок списку джерел", advanced=True),
    Param("bibliography.source_start_pattern", "regex", r"^\s*(?:\[(?P<bracket>\d+)\]|(?P<plain>\d+)[\.)])\s+.+",
          "Як починається джерело з номером, набраним вручну", "Word-автонумерація обробляється окремо.", advanced=True),
    Param("bibliography.year_pattern", "regex", r"(?<!\d)(19\d{2}|20\d{2})(?!\d)", "Ознака року", advanced=True),
    Param("bibliography.url_pattern", "regex", r"(?i)(https?://|www\.|doi:|\bDOI\b|\bURL\b|Режим\s+доступу)",
          "Ознака електронного ресурсу", advanced=True),
    Param("bibliography.access_date_pattern", "regex",
          r"(?i)(дата\s+звернення|звернення\s*[:—-]|accessed|retrieved|дата\s+обращения)", "Ознака дати звернення", advanced=True),
    Param("bibliography.pages_pattern", "regex", r"(?<!\w)(?:С|C|с|c|P|p|pp|S|s)\.\s*\d+", "Ознака сторінок («С. 10–15»)", advanced=True),
    Param("bibliography.article_markers_pattern", "regex",
          r"(?i)(журнал|збірник|conference|конференц|вісник|proceedings|journal|тези|матеріали)",
          "Ознаки статті/тез", "Для таких джерел очікуються сторінки.", advanced=True),
    Param("bibliography.min_length", "int", 25, "Мінімальна довжина опису джерела, символів", min=0, max=500),

    # --- зображення ---
    Param("images.min_dpi", "int", 150, "Мінімальна роздільна здатність зображення, DPI",
          "Розраховується з розміру зображення в пікселях і його розміру на сторінці.", min=10, max=1200),

    # --- обсяг ---
    Param("page_count.min_pages", "int", 35, "Мінімальна кількість сторінок", min=0, max=1000),
    Param("page_count.max_pages", "int", 50, "Максимальна кількість сторінок", min=0, max=1000),

    # --- звіт ---
    Param("report.max_per_rule", "int", 10, "Скільки однакових повідомлень показувати",
          "Решта згортається в рядок «… і ще N». 0 — показувати всі.", min=0, max=100000),
    Param("report.show_info", "bool", True, "Показувати інформаційні повідомлення (INFO)"),
    Param("report.comment_author", "str", "Валідатор оформлення", "Автор приміток у DOCX з коментарями"),
)
PARAM_BY_KEY = {p.key: p for p in PARAMS}

R = RuleDef
RULES: tuple[RuleDef, ...] = (
    # сторінка
    R("page_setup", "page_setup", "ERROR", "Формат сторінки, орієнтація або поля",
      "Порівнює розмір аркуша й поля кожної Word-секції з налаштуваннями."),
    # нумерація сторінок
    R("page_number_missing", "page_numbering", "WARNING", "Немає номера сторінки", "У потрібному колонтитулі не знайдено поле PAGE."),
    R("page_number_alignment", "page_numbering", "ERROR", "Номер сторінки не там", "Номер є, але вирівняний не так, як вимагається."),
    R("page_number_section_start", "page_numbering", "ERROR", "Нумерація починається не зі змісту",
      "Секція з номерами сторінок починається раніше за ЗМІСТ."),
    R("page_number_first_page", "page_numbering", "WARNING", "Номер прихований на першій сторінці секції",
      "Увімкнено «Особливий колонтитул для першої сторінки», і в ньому немає номера."),
    # структура
    R("structure_missing", "structure", "ERROR", "Немає обов'язкової частини", "Наприклад, ВСТУП чи ВИСНОВКИ."),
    R("structure_order", "structure", "WARNING", "Порушено порядок частин роботи"),
    R("structure_main_start", "structure", "WARNING", "Не знайдено початок основної частини",
      "Без цього титульна частина може потрапити в перевірки."),
    R("toc_manual", "structure", "WARNING", "Зміст набрано вручну", "Немає автоматичного змісту Word (поле TOC)."),
    R("toc_outdated", "structure", "WARNING", "Зміст не відповідає заголовкам", "Зміст треба оновити (F9) або виправити заголовки."),
    # основний текст
    R("paragraph_alignment", "paragraph_formatting", "WARNING", "Вирівнювання абзацу"),
    R("paragraph_indent", "paragraph_formatting", "WARNING", "Абзацний відступ"),
    R("paragraph_line_spacing", "paragraph_formatting", "WARNING", "Міжрядковий інтервал"),
    R("paragraph_spacing", "paragraph_formatting", "WARNING", "Інтервал перед/після абзацу"),
    R("paragraph_font", "paragraph_formatting", "WARNING", "Шрифт основного тексту"),
    R("paragraph_font_size", "paragraph_formatting", "WARNING", "Розмір шрифту основного тексту"),
    R("empty_paragraphs", "empty_paragraphs", "WARNING", "Кілька порожніх абзаців підряд"),
    # заголовки
    R("heading_number_dot", "headings", "ERROR", "Крапка після номера розділу/підрозділу", "«1.1. Назва» замість «1.1 Назва»."),
    R("heading_trailing_dot", "headings", "WARNING", "Крапка в кінці заголовка"),
    R("heading_section_case", "headings", "WARNING", "Назва розділу не великими літерами"),
    R("heading_subsection_case", "headings", "WARNING", "Регістр назви підрозділу"),
    R("heading_bold", "headings", "WARNING", "Заголовок розділу не напівжирний"),
    R("heading_alignment", "headings", "WARNING", "Вирівнювання заголовка"),
    R("section_new_page", "headings", "WARNING", "Розділ не з нової сторінки",
      "Перед заголовком розділу немає розриву сторінки/секції чи «з нової сторінки»."),
    # нумерація
    R("heading_numbering_sequence", "numbering", "WARNING", "Пропуск у нумерації підрозділів", "Наприклад, після 1.2 іде 1.4."),
    R("caption_numbering_sequence", "numbering", "WARNING", "Пропуск у нумерації рисунків/таблиць"),
    R("caption_chapter_mismatch", "numbering", "WARNING", "Номер рисунка/таблиці не відповідає розділу",
      "Рис. 2.1 розміщено в розділі 1."),
    # підписи
    R("caption_missing", "captions_and_references", "ERROR", "Немає підпису рисунка/таблиці"),
    R("caption_position", "captions_and_references", "ERROR", "Підпис не на своєму місці",
      "Рисунок — підпис знизу, таблиця — підпис зверху."),
    R("caption_alignment", "captions_and_references", "ERROR", "Вирівнювання підпису"),
    R("caption_duplicate", "captions_and_references", "ERROR", "Дубльований номер рисунка/таблиці"),
    R("caption_trailing_dot", "captions_and_references", "WARNING", "Крапка в кінці підпису"),
    R("object_reference_missing", "captions_and_references", "ERROR", "Немає згадки рисунка/таблиці до його появи"),
    # завершення підрозділів
    R("section_empty", "section_endings", "ERROR", "Підрозділ без тексту"),
    R("section_ends_with_object", "section_endings", "ERROR", "Підрозділ закінчується рисунком/таблицею"),
    # формули
    R("formula_number_missing", "formulas", "WARNING", "Формула без номера"),
    R("formula_reference_missing", "formulas", "WARNING", "Немає згадки формули в тексті"),
    # таблиці
    R("table_font_size", "tables", "WARNING", "Недозволений розмір шрифту в таблиці"),
    # додатки
    R("appendix_invalid_letter", "appendices", "ERROR", "Недозволена літера додатка"),
    R("appendix_sequence", "appendices", "WARNING", "Порушено послідовність літер додатків"),
    R("appendix_not_referenced", "appendices", "WARNING", "Немає згадки додатка в тексті"),
    # джерела
    R("bibliography_missing", "bibliography", "ERROR", "Немає списку джерел"),
    R("bibliography_empty", "bibliography", "ERROR", "Список джерел порожній"),
    R("source_number_missing", "bibliography", "ERROR", "Джерело без номера"),
    R("source_sequence", "bibliography", "WARNING", "Порушено нумерацію джерел"),
    R("source_bare_url", "bibliography", "ERROR", "Джерело — голе посилання без опису"),
    R("source_too_short", "bibliography", "WARNING", "Надто короткий опис джерела"),
    R("source_missing_year", "bibliography", "WARNING", "Немає року видання"),
    R("source_missing_access_date", "bibliography", "WARNING", "Немає дати звернення до електронного ресурсу"),
    R("source_missing_url", "bibliography", "WARNING", "«URL/Режим доступу» без самого посилання"),
    R("source_article_pages", "bibliography", "WARNING", "Немає сторінок статті («С. 10–15»)"),
    # посилання на джерела
    R("source_refs_missing_all", "source_references", "ERROR", "У тексті немає жодного посилання на джерела"),
    R("source_ref_unknown_number", "source_references", "ERROR", "Посилання на неіснуюче джерело"),
    R("source_missing_in_text", "source_references", "ERROR", "Джерело не згадано в тексті"),
    R("source_ref_suspicious", "source_references", "WARNING", "Нерозпізнане посилання в квадратних дужках"),
    R("source_ref_order", "source_references", "OFF", "Джерела не в порядку першої згадки",
      "Вмикайте, якщо список має бути впорядкований за появою посилань у тексті (а не за абеткою)."),
    # типографіка
    R("double_space", "typography", "WARNING", "Кілька пробілів підряд"),
    R("space_before_punctuation", "typography", "WARNING", "Пробіл перед розділовим знаком"),
    R("hyphen_instead_of_dash", "typography", "INFO", "Дефіс замість тире", "« - » між словами замість « — » або « – »."),
    R("straight_quotes", "typography", "INFO", "Прямі лапки \"…\" замість «…»"),
    R("mixed_script", "typography", "WARNING", "Латинські літери всередині кириличних слів",
      "Часто це помилка розкладки, але так само обходять перевірку на плагіат."),
    R("leading_spaces", "typography", "WARNING", "Відступ пробілами або табуляцією", "Абзацний відступ має задаватися в параметрах абзацу."),
    R("manual_line_break", "typography", "INFO", "Розрив рядка (Shift+Enter) в абзаці"),
    # залишки редагування
    R("tracked_changes", "document_hygiene", "WARNING", "Невиконані виправлення (режим рецензування)"),
    R("comments_present", "document_hygiene", "INFO", "У документі є примітки"),
    R("hidden_text", "document_hygiene", "WARNING", "Прихований текст"),
    R("colored_text", "document_hygiene", "INFO", "Кольоровий текст в основній частині"),
    R("highlighted_text", "document_hygiene", "WARNING", "Виділений маркером текст"),
    # зображення
    R("image_too_wide", "images", "WARNING", "Зображення ширше за поле тексту"),
    R("image_low_resolution", "images", "INFO", "Низька роздільна здатність зображення"),
    # обсяг
    R("page_count_out_of_range", "page_count", "ERROR", "Обсяг роботи поза нормою"),
    R("page_count_unknown", "page_count", "INFO", "Не вдалося порахувати сторінки"),
)
RULE_BY_ID = {r.id: r for r in RULES}


# ---------------------------------------------------------------------------
# Об'єкт налаштувань
# ---------------------------------------------------------------------------

class ConfigError(Exception):
    def __init__(self, errors: list[str], source: str = "") -> None:
        self.errors = errors
        self.source = source
        head = f"Помилки в конфігу {source}:" if source else "Помилки в конфігу:"
        super().__init__(head + "\n" + "\n".join(f"  • {e}" for e in errors))


@dataclass
class Settings:
    values: dict[str, Any] = field(default_factory=dict)
    severities: dict[str, str] = field(default_factory=dict)
    checks: dict[str, bool] = field(default_factory=dict)
    source: str = "вбудовані значення"
    extends: Optional[str] = None

    def __getitem__(self, key: str) -> Any:
        return self.values[key]

    def sev(self, rule_id: str) -> str:
        return self.severities.get(rule_id, RULE_BY_ID[rule_id].severity if rule_id in RULE_BY_ID else "WARNING")

    def enabled(self, group: str) -> bool:
        return self.checks.get(group, True)

    @property
    def name(self) -> str:
        return str(self.values.get("profile.name") or "")

    def copy(self) -> "Settings":
        return Settings(dict(self.values), dict(self.severities), dict(self.checks), self.source, self.extends)

    def to_nested(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        for p in PARAMS:
            out.setdefault(p.section, {})[p.name] = self.values[p.key]
        out["checks"] = dict(self.checks)
        out["severities"] = dict(self.severities)
        return out


def default_settings() -> Settings:
    return Settings(
        values={p.key: _copy_default(p.default) for p in PARAMS},
        severities={r.id: r.severity for r in RULES},
        checks={g.id: g.enabled for g in GROUPS},
    )


def _copy_default(v: Any) -> Any:
    return list(v) if isinstance(v, list) else v


# ---------------------------------------------------------------------------
# Перетворення й перевірка значень
# ---------------------------------------------------------------------------

def _suggest(word: str, options: list[str]) -> str:
    close = difflib.get_close_matches(word, options, n=1, cutoff=0.6)
    return f" Можливо, ви мали на увазі «{close[0]}»?" if close else ""


def coerce_value(param: Param, value: Any) -> Any:
    """Перетворює значення з YAML/GUI до потрібного типу. Кидає ValueError з поясненням."""
    t = param.type
    if value is None or (isinstance(value, str) and value.strip() == "" and param.nullable):
        if param.nullable:
            return None
        raise ValueError("значення не може бути порожнім")
    if t == "bool":
        if isinstance(value, bool):
            return value
        if isinstance(value, str) and value.strip().lower() in {"true", "так", "yes", "1", "on"}:
            return True
        if isinstance(value, str) and value.strip().lower() in {"false", "ні", "no", "0", "off"}:
            return False
        raise ValueError(f"очікується true або false, а не «{value}»")
    if t in {"float", "int"}:
        if isinstance(value, bool):
            raise ValueError(f"очікується число, а не «{value}»")
        try:
            num = float(str(value).replace(",", ".")) if isinstance(value, str) else float(value)
        except (TypeError, ValueError):
            raise ValueError(f"очікується число, а не «{value}»") from None
        if t == "int":
            if num != int(num):
                raise ValueError(f"очікується ціле число, а не «{value}»")
            num = int(num)
        if param.min is not None and num < param.min:
            raise ValueError(f"значення {num} менше за допустиме {param.min:g}")
        if param.max is not None and num > param.max:
            raise ValueError(f"значення {num} більше за допустиме {param.max:g}")
        return num
    if t == "choice":
        key = str(value).strip().upper()
        if key not in param.choices:
            raise ValueError(f"«{value}» — недопустиме значення; можна: {', '.join(param.choices)}.{_suggest(key, list(param.choices))}")
        return key
    if t == "str":
        return str(value)
    if t == "regex":
        s = str(value)
        try:
            re.compile(s)
        except re.error as exc:
            raise ValueError(f"неправильний регулярний вираз: {exc}") from None
        return s
    if t in {"list", "regex_list", "float_list"}:
        if isinstance(value, str):
            items = [x.strip() for x in value.splitlines() if x.strip()]
        elif isinstance(value, (list, tuple)):
            items = list(value)
        else:
            raise ValueError("очікується список")
        out = []
        for i, item in enumerate(items, start=1):
            if t == "float_list":
                try:
                    out.append(float(str(item).replace(",", ".")))
                except ValueError:
                    raise ValueError(f"елемент {i}: очікується число, а не «{item}»") from None
            elif t == "regex_list":
                try:
                    re.compile(str(item))
                except re.error as exc:
                    raise ValueError(f"елемент {i}: неправильний регулярний вираз: {exc}") from None
                out.append(str(item))
            else:
                out.append(str(item))
        return out
    raise ValueError(f"невідомий тип параметра {t}")  # pragma: no cover


def coerce_severity(value: Any) -> str:
    sev = str(value).strip().upper()
    aliases = {"DISABLED": "OFF", "NONE": "OFF", "IGNORE": "OFF", "FALSE": "OFF", "WARN": "WARNING", "ERR": "ERROR"}
    sev = aliases.get(sev, sev)
    if sev not in SEVERITIES:
        raise ValueError(f"«{value}» — недопустимий рівень; можна: {', '.join(SEVERITIES)}")
    return sev


# ---------------------------------------------------------------------------
# Читання конфігу
# ---------------------------------------------------------------------------

def _yaml():
    try:
        import yaml  # type: ignore
    except Exception as exc:  # pragma: no cover
        raise RuntimeError("Для YAML-конфігу потрібно встановити PyYAML: pip install PyYAML") from exc
    return yaml


def _line_map(text: str) -> dict[tuple[str, ...], int]:
    """Номери рядків для ключів YAML (щоб показати, де помилка)."""
    yaml = _yaml()
    result: dict[tuple[str, ...], int] = {}
    try:
        node = yaml.compose(text)
    except Exception:
        return result

    def walk(n: Any, path: tuple[str, ...]) -> None:
        if isinstance(n, yaml.MappingNode):
            for k, v in n.value:
                key = str(k.value)
                result[path + (key,)] = k.start_mark.line + 1
                walk(v, path + (key,))

    if node is not None:
        walk(node, ())
    return result


def parse_config_text(text: str, source: str = "", suffix: str = ".yaml") -> dict[str, Any]:
    if suffix.lower() == ".json":
        import json
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ConfigError([f"рядок {exc.lineno}: синтаксична помилка JSON: {exc.msg}"], source) from None
    else:
        yaml = _yaml()
        try:
            data = yaml.safe_load(text)
        except yaml.YAMLError as exc:
            mark = getattr(exc, "problem_mark", None)
            where = f"рядок {mark.line + 1}: " if mark is not None else ""
            problem = getattr(exc, "problem", None) or str(exc)
            raise ConfigError([f"{where}синтаксична помилка YAML: {problem}. "
                               "Перевірте відступи (лише пробіли) і лапки навколо регулярних виразів."], source) from None
    if data is None:
        data = {}
    if not isinstance(data, dict):
        raise ConfigError(["конфіг має бути набором секцій «назва: значення»"], source)
    return data


def apply_raw(settings: Settings, raw: dict[str, Any], source: str = "", lines: Optional[dict[tuple[str, ...], int]] = None) -> Settings:
    """Накладає сирий словник з конфігу на settings. Збирає всі помилки разом."""
    lines = lines or {}
    errors: list[str] = []
    s = settings.copy()

    def where(*path: str) -> str:
        ln = lines.get(tuple(path))
        loc = ".".join(path)
        return f"рядок {ln}, «{loc}»" if ln else f"«{loc}»"

    sections = sorted({p.section for p in PARAMS})
    allowed_top = set(sections) | {"checks", "severities", "extends"}
    for top, body in raw.items():
        top = str(top)
        if top not in allowed_top:
            errors.append(f"{where(top)}: невідома секція.{_suggest(top, sorted(allowed_top))}")
            continue
        if top == "extends":
            continue
        if not isinstance(body, dict):
            errors.append(f"{where(top)}: очікується вкладений набір параметрів")
            continue
        if top == "checks":
            for k, v in body.items():
                k = str(k)
                if k not in GROUP_BY_ID:
                    errors.append(f"{where(top, k)}: невідома група перевірок.{_suggest(k, list(GROUP_BY_ID))}")
                elif not isinstance(v, bool):
                    errors.append(f"{where(top, k)}: очікується true або false, а не «{v}»")
                else:
                    s.checks[k] = v
            continue
        if top == "severities":
            for k, v in body.items():
                k = str(k)
                if k not in RULE_BY_ID:
                    errors.append(f"{where(top, k)}: невідоме правило.{_suggest(k, list(RULE_BY_ID))}")
                    continue
                try:
                    s.severities[k] = coerce_severity(v)
                except ValueError as exc:
                    errors.append(f"{where(top, k)}: {exc}")
            continue
        names = [p.name for p in PARAMS if p.section == top]
        for k, v in body.items():
            k = str(k)
            key = f"{top}.{k}"
            param = PARAM_BY_KEY.get(key)
            if param is None:
                errors.append(f"{where(top, k)}: невідомий параметр.{_suggest(k, names)}")
                continue
            try:
                s.values[key] = coerce_value(param, v)
            except ValueError as exc:
                errors.append(f"{where(top, k)}: {exc}")

    if errors:
        raise ConfigError(errors, source)
    return s


def load_settings_file(path: Path, _stack: Optional[list[Path]] = None) -> Settings:
    path = Path(path)
    _stack = list(_stack or [])
    resolved = path.resolve()
    if resolved in _stack:
        raise ConfigError([f"циклічне «extends»: {' → '.join(str(p.name) for p in _stack + [resolved])}"], str(path))
    _stack.append(resolved)
    try:
        text = path.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        raise ConfigError(["файл не знайдено"], str(path)) from None
    raw = parse_config_text(text, str(path), path.suffix)
    base = default_settings()
    ext = raw.get("extends")
    if ext:
        base_path = resolve_profile(str(ext), relative_to=path.parent)
        if base_path is None:
            raise ConfigError([f"«extends: {ext}»: профіль не знайдено. Доступні: {', '.join(p.id for p in list_profiles())}"], str(path))
        base = load_settings_file(base_path, _stack)
    lines = _line_map(text) if path.suffix.lower() != ".json" else {}
    s = apply_raw(base, raw, str(path), lines)
    s.source = str(path)
    s.extends = str(ext) if ext else None
    return s


# ---------------------------------------------------------------------------
# Профілі
# ---------------------------------------------------------------------------

@dataclass
class ProfileInfo:
    id: str
    name: str
    description: str
    path: Path
    builtin: bool


def builtin_profiles_dir() -> Path:
    base = Path(getattr(sys, "_MEIPASS", "")) / "lpnu_docx_validator" if getattr(sys, "frozen", False) else Path(__file__).parent
    return base / "profiles"


def user_profiles_dir() -> Path:
    env = os.environ.get("DOCX_VALIDATOR_PROFILES")
    if env:
        return Path(env)
    if os.name == "nt":
        root = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    else:
        root = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    return root / "lpnu_docx_validator" / "profiles"


def _profile_meta(path: Path) -> tuple[str, str]:
    try:
        raw = parse_config_text(path.read_text(encoding="utf-8-sig"), str(path), path.suffix)
        prof = raw.get("profile") or {}
        return str(prof.get("name") or path.stem), str(prof.get("description") or "")
    except Exception:
        return path.stem, "(не вдалося прочитати файл)"


def list_profiles() -> list[ProfileInfo]:
    result: list[ProfileInfo] = []
    seen: set[str] = set()
    for d, builtin in ((builtin_profiles_dir(), True), (user_profiles_dir(), False)):
        if not d.is_dir():
            continue
        files = sorted(d.glob("*.yaml")) + sorted(d.glob("*.yml"))
        if builtin:
            files.sort(key=lambda p: (p.stem != DEFAULT_PROFILE_ID, p.stem))
        for f in files:
            pid = f.stem if builtin else f"user:{f.stem}"
            if pid in seen:
                continue
            seen.add(pid)
            name, desc = _profile_meta(f)
            result.append(ProfileInfo(pid, name, desc, f, builtin))
    return result


def resolve_profile(ref: str, relative_to: Optional[Path] = None) -> Optional[Path]:
    """Профіль за id (``coursework``, ``user:my``) або шляхом до файлу."""
    for p in list_profiles():
        if ref in (p.id, p.path.stem) or (p.builtin is False and ref == p.id.removeprefix("user:")):
            return p.path
    cand = Path(ref)
    if relative_to is not None and not cand.is_absolute():
        cand = relative_to / cand
    if cand.is_file():
        return cand
    return None


def load_profile(ref: str) -> Settings:
    path = resolve_profile(ref)
    if path is None:
        raise ConfigError([f"профіль «{ref}» не знайдено. Доступні: {', '.join(p.id for p in list_profiles())}"])
    return load_settings_file(path)


def find_default_config(start_dirs: Optional[list[Path]] = None) -> Optional[Path]:
    dirs = start_dirs if start_dirs is not None else [Path.cwd(), Path(sys.argv[0]).resolve().parent]
    for d in dirs:
        cand = d / DEFAULT_CONFIG_FILENAME
        if cand.is_file():
            return cand
    return None


def resolve_settings(config: Optional[Path] = None, profile: Optional[str] = None,
                     search_dirs: Optional[list[Path]] = None) -> Settings:
    """Порядок: явний -c → явний -p → docx_validator_config.yaml поруч → вбудований профіль."""
    if config is not None:
        return load_settings_file(config)
    if profile:
        return load_profile(profile)
    found = find_default_config(search_dirs)
    if found is not None:
        return load_settings_file(found)
    try:
        return load_profile(DEFAULT_PROFILE_ID)
    except ConfigError:
        return default_settings()


# ---------------------------------------------------------------------------
# Запис конфігу (YAML з коментарями)
# ---------------------------------------------------------------------------

def _yaml_scalar(v: Any) -> str:
    if v is None:
        return "null"
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        r = repr(v)
        return r if ("." in r or "e" in r) else r + ".0"
    s = str(v).replace("'", "''")
    return f"'{s}'"


def _comment(text: str, indent: str) -> list[str]:
    out = []
    for para in text.split("\n"):
        wrapped = textwrap.wrap(para, width=88 - len(indent)) or [""]
        out.extend(f"{indent}# {w}".rstrip() for w in wrapped)
    return out


def settings_to_yaml(settings: Settings, only_changed_from: Optional[Settings] = None) -> str:
    """Повний YAML із коментарями. Якщо only_changed_from задано — лише відмінності (для профілю з extends)."""
    lines: list[str] = [
        "# " + "=" * 70,
        "# Конфіг правил для валідатора оформлення DOCX",
        "# " + "=" * 70,
        "# Файл можна редагувати вручну або в GUI (вкладка «Налаштування»).",
        "# Рядки, що починаються з #, — коментарі.",
        "# Регулярні вирази записуйте в одинарних лапках '...'.",
        "# Рівні серйозності: ERROR — помилка, WARNING — попередження,",
        "#   INFO — інформація, OFF — вимкнути правило.",
        "",
    ]
    base = only_changed_from
    if settings.extends:
        lines += [f"extends: {_yaml_scalar(settings.extends)}", ""]

    def changed(key: str) -> bool:
        return base is None or base.values.get(key) != settings.values.get(key)

    sections: list[str] = []
    for p in PARAMS:
        if p.section not in sections:
            sections.append(p.section)

    def emit_section(sec: str) -> None:
        params = [p for p in PARAMS if p.section == sec and (changed(p.key) or (sec == "profile" and base is not None))]
        if not params:
            return
        lines.append("# " + "-" * 60)
        lines.append(f"# {SECTION_TITLES.get(sec, sec)}")
        lines.append("# " + "-" * 60)
        lines.append(f"{sec}:")
        for p in params:
            text = p.label + (f". {p.help}" if p.help else "")
            if p.type == "choice":
                text += " Можливі значення: " + ", ".join(p.choices) + "."
            lines.extend(_comment(text, "  "))
            v = settings.values[p.key]
            if p.type in {"list", "regex_list", "float_list"}:
                if not v:
                    lines.append(f"  {p.name}: []")
                else:
                    lines.append(f"  {p.name}:")
                    lines.extend(f"    - {_yaml_scalar(x)}" for x in v)
            else:
                lines.append(f"  {p.name}: {_yaml_scalar(v)}")
            lines.append("")

    emit_section("profile")
    # checks
    chk = [g for g in GROUPS if base is None or base.checks.get(g.id) != settings.checks.get(g.id)]
    if chk:
        lines += ["# " + "-" * 60, "# Увімкнення груп перевірок (true/false)", "# " + "-" * 60, "checks:"]
        for g in chk:
            lines.extend(_comment(f"{g.title}: {g.description}", "  "))
            lines.append(f"  {g.id}: {_yaml_scalar(settings.checks.get(g.id, g.enabled))}")
        lines.append("")
    for sec in sections:
        if sec != "profile":
            emit_section(sec)
    sev = [r for r in RULES if base is None or base.severities.get(r.id) != settings.severities.get(r.id)]
    if sev:
        lines += ["# " + "-" * 60, "# Рівні серйозності окремих правил (ERROR / WARNING / INFO / OFF)", "# " + "-" * 60, "severities:"]
        current = None
        for r in sev:
            if r.group != current:
                current = r.group
                lines.append(f"  # --- {GROUP_BY_ID[r.group].title} ---")
            lines.append(f"  {r.id}: {settings.severities.get(r.id, r.severity)}".ljust(46) + f"# {r.title}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def save_settings(settings: Settings, path: Path, only_changed_from: Optional[Settings] = None) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    text = settings_to_yaml(settings, only_changed_from)
    # Перевіряємо, що записане читається назад без помилок.
    apply_raw(default_settings(), parse_config_text(text, str(path)), str(path))
    path.write_text(text, encoding="utf-8")


def rules_markdown() -> str:
    lines = ["| Правило | Група | За замовчуванням | Опис |", "| --- | --- | --- | --- |"]
    for r in RULES:
        desc = r.title + (f". {r.description}" if r.description else "")
        lines.append(f"| `{r.id}` | {GROUP_BY_ID[r.group].title} | {r.severity} | {desc} |")
    return "\n".join(lines) + "\n"
