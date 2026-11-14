import re
from pathlib import Path

import pytest

from lpnu_docx_validator import settings as S
from lpnu_docx_validator.settings import (ConfigError, default_settings, list_profiles, load_profile, load_settings_file,
                                          resolve_settings, save_settings, settings_to_yaml)

ROOT = Path(__file__).resolve().parents[1]


def test_defaults_are_consistent():
    for p in S.PARAMS:
        assert S.coerce_value(p, p.default) == p.default or p.default is None, p.key
        if p.type == "regex":
            re.compile(p.default)
    for r in S.RULES:
        assert r.group in S.GROUP_BY_ID, r.id
        assert r.severity in S.SEVERITIES
    assert len({r.id for r in S.RULES}) == len(S.RULES)


def test_every_rule_used_in_code_exists():
    """Кожен ідентифікатор правила, який викликають перевірки, описано в реєстрі (ловить одруківки)."""
    src = "".join(p.read_text(encoding="utf-8") for p in (ROOT / "lpnu_docx_validator" / "checks").glob("*.py"))
    used = set(re.findall(r'rep\.add\(\s*"([a-z_]+)"', src))
    assert used, "не знайдено викликів rep.add"
    unknown = used - set(S.RULE_BY_ID) - {"internal_error"}
    assert not unknown
    unused = set(S.RULE_BY_ID) - used
    assert not unused, f"правила без перевірки: {unused}"


def test_shipped_config_files_match_defaults():
    d = default_settings()
    for path in (ROOT / "docx_validator_config.yaml", ROOT / "lpnu_docx_validator" / "profiles" / "coursework.yaml"):
        s = load_settings_file(path)
        assert s.values == d.values, path
        assert s.severities == d.severities
        assert s.checks == d.checks
        assert path.read_text(encoding="utf-8") == settings_to_yaml(d), \
            f"{path.name} застарів: перегенеруйте через `python -m lpnu_docx_validator init-config --force -o ...`"


def test_builtin_profiles():
    ids = [p.id for p in list_profiles()]
    assert ids[0] == "coursework"
    assert "lab_report" in ids
    lab = load_profile("lab_report")
    assert lab.extends == "coursework"
    assert lab.checks["page_numbering"] is False
    assert lab["page.margin_left_cm"] == 2.5  # успадковано


def write(tmp_path, text, name="c.yaml"):
    p = tmp_path / name
    p.write_text(text, encoding="utf-8")
    return p


def test_config_errors_are_human_readable(tmp_path):
    p = write(tmp_path, "body_text:\n  alignment: JUSTIFIED\n  font_size: 14\nseverities:\n  caption_mising: OFF\n"
                        "  paragraph_font: MAYBE\ncaptions:\n  figure_caption_pattern: '(Рис'\npage:\n  margin_left_cm: -1\n"
                        "unknown_section: {}\nchecks:\n  headngs: true\n")
    with pytest.raises(ConfigError) as ei:
        load_settings_file(p)
    errs = "\n".join(ei.value.errors)
    assert "рядок 2" in errs and "«JUSTIFY»" in errs
    assert "font_size_pt" in errs
    assert "caption_missing" in errs
    assert "MAYBE" in errs
    assert "регулярний вираз" in errs
    assert "менше за допустиме" in errs
    assert "невідома секція" in errs
    assert "headings" in errs
    assert len(ei.value.errors) == 8


def test_yaml_syntax_error(tmp_path):
    p = write(tmp_path, "page:\n  margin_left_cm: 3\n bad: x\n")
    with pytest.raises(ConfigError) as ei:
        load_settings_file(p)
    assert "рядок 3" in ei.value.errors[0]


def test_nullable_and_lowercase(tmp_path):
    p = write(tmp_path, "body_text:\n  space_after_pt: null\n  alignment: justify\nseverities:\n  heading_bold: off\n"
                        "page:\n  margin_left_cm: '3,0'\n")
    s = load_settings_file(p)
    assert s["body_text.space_after_pt"] is None
    assert s["body_text.alignment"] == "JUSTIFY"
    assert s.sev("heading_bold") == "OFF"
    assert s["page.margin_left_cm"] == 3.0


def test_legacy_config_format_still_loads(tmp_path):
    """Старий docx_validator_config.yaml із main-гілки читається без змін."""
    old = (Path(__file__).parent / "data" / "legacy_config.yaml")
    s = load_settings_file(old)
    assert s["page.margin_left_cm"] == 2.5
    assert s["page_numbering.position"] == "TOP"
    assert s.sev("paragraph_alignment") == "WARNING"


def test_roundtrip(tmp_path):
    s = default_settings()
    s.values["page.margin_left_cm"] = 3.0
    s.values["structure.required_sections"] = [r"(?i)^вступ$"]
    s.values["body_text.space_before_pt"] = None
    s.values["profile.name"] = "Мій 'профіль'"
    s.severities["double_space"] = "OFF"
    s.checks["images"] = False
    path = tmp_path / "x.yaml"
    save_settings(s, path)
    t = load_settings_file(path)
    assert t.values == s.values and t.severities == s.severities and t.checks == s.checks


def test_extends_minimal_and_cycle(tmp_path):
    base = default_settings()
    s = base.copy()
    s.extends = "coursework"
    s.values["page.margin_right_cm"] = 1.0
    p = tmp_path / "mine.yaml"
    save_settings(s, p, only_changed_from=base)
    text = p.read_text(encoding="utf-8")
    assert "margin_right_cm: 1.0" in text and "margin_left_cm" not in text
    assert load_settings_file(p)["page.margin_right_cm"] == 1.0

    a = write(tmp_path, "extends: b.yaml\n", "a.yaml")
    write(tmp_path, "extends: a.yaml\n", "b.yaml")
    with pytest.raises(ConfigError, match="циклічне"):
        load_settings_file(a)


def test_resolve_order(tmp_path, monkeypatch):
    write(tmp_path, "page:\n  margin_left_cm: 3.3\n", S.DEFAULT_CONFIG_FILENAME)
    assert resolve_settings(search_dirs=[tmp_path])["page.margin_left_cm"] == 3.3
    assert resolve_settings(profile="coursework", search_dirs=[tmp_path])["page.margin_left_cm"] == 2.5
    assert resolve_settings(search_dirs=[tmp_path / "nope"])["page.margin_left_cm"] == 2.5


def test_rules_doc_is_up_to_date():
    doc = (ROOT / "docs" / "RULES.md").read_text(encoding="utf-8")
    assert S.rules_markdown() in doc, "оновіть docs/RULES.md: python -m lpnu_docx_validator list-rules --markdown"


def test_rules_markdown_lists_every_rule():
    md = S.rules_markdown()
    for r in S.RULES:
        assert f"`{r.id}`" in md
