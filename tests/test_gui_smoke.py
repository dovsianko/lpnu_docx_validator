"""Smoke-тест GUI на підробленому tkinter: збирання вікна, форма ↔ налаштування, результати, збереження профілю.

Справжнє вікно перевіряється в CI під xvfb (див. .github/workflows/ci.yml).
"""

import importlib
import sys

import pytest

import fake_tk
from docbuilder import CourseworkBuilder, good_coursework


@pytest.fixture
def app(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)  # щоб не підхопився docx_validator_config.yaml з кореня репозиторію
    monkeypatch.setattr(sys, "argv", [str(tmp_path / "x")])
    real = {k: sys.modules.get(k) for k in ("tkinter", "tkinter.ttk", "tkinter.messagebox", "tkinter.filedialog")}
    ctl = fake_tk.install()
    gui = importlib.import_module("lpnu_docx_validator.gui")
    root = fake_tk.Tk()
    a = gui.App(root)
    a._ctl = ctl
    a._gui = gui
    yield a
    for k, v in real.items():
        if v is None:
            sys.modules.pop(k, None)
        else:
            sys.modules[k] = v
    sys.modules.pop("lpnu_docx_validator.gui", None)


def test_builds_and_loads_default_profile(app):
    from lpnu_docx_validator.settings import PARAMS, RULES
    assert set(app.param_widgets) == {p.key for p in PARAMS}
    assert set(app.sev_vars) == {r.id for r in RULES}
    assert app.settings.name.startswith("Курсова")
    assert app.settings_builtin and not app.dirty
    s = app.collect_settings()
    assert s is not None and s.values == app.settings.values


def test_edit_value_marks_dirty_and_collects(app):
    _, get, set_, _ = app.param_widgets["page.margin_left_cm"]
    set_("3,0")
    assert app.dirty
    _, _, set_align, _ = app.param_widgets["body_text.alignment"]
    set_align("по центру")
    s = app.collect_settings()
    assert s["page.margin_left_cm"] == 3.0
    assert s["body_text.alignment"] == "CENTER"


def test_invalid_value_shows_error(app):
    app.param_widgets["page.margin_left_cm"][2]("abc")
    assert app.collect_settings() is None
    kind, msg = app._ctl["shown"][-1]
    assert kind == "showerror" and "Ліве поле" in msg


def test_results_filter_and_disable_rule(app, tmp_path):
    from lpnu_docx_validator.batch import write_outputs
    from lpnu_docx_validator.checks import validate
    b = CourseworkBuilder().title_page()
    b.toc(["ВСТУП"])
    b.heading("ВСТУП")
    b.body("Текст  з подвійним пробілом і словом рoзділ.")
    path = b.save(tmp_path / "w.docx")
    outcome = write_outputs(validate(path, app.collect_settings()), ["txt"], tmp_path)
    app.show_outcomes([outcome])
    app.on_file_selected()
    shown = {v.rule for v in app.shown_issues}
    assert "double_space" in shown and "info" not in shown  # INFO приховано за замовчуванням
    app.query_var.set("кілька пробілів")
    assert {v.rule for v in app.shown_issues} == {"double_space"}
    app.issues_tree.selection_set("0")
    app.copy_search()
    assert app.root.clip.startswith("Текст з подвійним")
    app.set_rule_severity("OFF")
    assert app.dirty
    assert app.collect_settings().sev("double_space") == "OFF"
    assert app.shown_issues == []


def test_save_as_user_profile_and_reload(app, monkeypatch):
    app.param_widgets["page.margin_right_cm"][2]("1")
    monkeypatch.setattr(app._gui, "SaveAsDialog", lambda master, name: type("D", (), {"result": ("Кафедра КБ", "profile")})())
    assert app.save_as()
    assert not app.settings_builtin and app.settings_path.name == "kafedra_kb.yaml"
    assert app.settings["page.margin_right_cm"] == 1.0
    assert any(p.name == "Кафедра КБ" for p in app.profiles)
    app.param_widgets["page.margin_top_cm"][2]("2.5")
    assert app.save_current()
    from lpnu_docx_validator.settings import load_settings_file
    assert load_settings_file(app.settings_path)["page.margin_top_cm"] == 2.5
    app.on_close()
    from lpnu_docx_validator.gui_logic import load_state
    assert load_state()["profile_path"].endswith("kafedra_kb.yaml")


def test_add_files_and_run_in_thread(app, tmp_path):
    good_coursework(tmp_path / "a.docx")
    (tmp_path / "a.checked.docx").write_bytes(b"")
    app._add_paths([tmp_path])
    assert [f.name for f in app.files] == ["a.docx"]
    app.fmt_vars["summary"].set(True)
    app.run_check()
    app.worker_done = False
    import time
    for _ in range(100):
        if not app.queue.empty() and app.queue.queue[-1][0] in {"done", "error"}:
            break
        time.sleep(0.05)
    app._poll()
    assert not app.running
    assert app.outcomes and app.outcomes[0].result.errors == 0
    assert app.summary_path is not None and app.summary_path.exists()
