import json
import subprocess
import sys
from pathlib import Path

from docx import Document

from docbuilder import CourseworkBuilder, good_coursework
from lpnu_docx_validator.checks import validate
from lpnu_docx_validator.cli import main
from lpnu_docx_validator.report import render_html, render_json, render_text, write_annotated_docx

ROOT = Path(__file__).resolve().parents[1]


def noisy_doc(path: Path, n: int = 25) -> Path:
    b = CourseworkBuilder().title_page()
    b.toc(["ВСТУП"])
    b.heading("ВСТУП")
    for k in range(n):
        b.body(f"Абзац  номер {k} із подвійним пробілом.")
    return b.save(path)


def test_repeated_issues_are_collapsed(tmp_path, settings):
    res = validate(noisy_doc(tmp_path / "n.docx"), settings)
    assert sum(1 for i in res.issues if i.rule == "double_space") == 25
    text = render_text(res)
    assert text.count("кілька пробілів підряд") == 10
    assert "і ще 15 подібних" in text
    settings.values["report.max_per_rule"] = 0
    assert render_text(validate(tmp_path / "n.docx", settings)).count("кілька пробілів підряд") == 25


def test_html_and_json(tmp_path, settings):
    res = validate(noisy_doc(tmp_path / "n.docx", 3), settings)
    html = render_html(res)
    assert html.startswith("<!DOCTYPE html>") and "class='find'" in html and "<script>" in html
    data = json.loads(render_json(res))
    assert data["warnings"] == res.warnings and data["issues"][0]["rule"]


def test_annotated_docx_has_comments(tmp_path, settings):
    src = noisy_doc(tmp_path / "n.docx", 12)
    res = validate(src, settings)
    out = tmp_path / "n.checked.docx"
    n = write_annotated_docx(res, out)
    comments = list(Document(str(out)).comments)
    assert n == len(comments) > 0
    assert any("ще 2 подібних" in c.text for c in comments)  # 12 повторів, показано 10
    assert len(list(Document(str(src)).comments)) == 0  # оригінал не змінено


def test_legacy_invocation(tmp_path):
    src = good_coursework(tmp_path / "Course_Petriv.docx")
    out = tmp_path / "звіт.txt"
    rc = subprocess.run([sys.executable, str(ROOT / "docx_format_validator.py"), str(src), "-c",
                         str(ROOT / "docx_validator_config.yaml"), "-o", str(out)], capture_output=True, text=True)
    assert rc.returncode == 0, rc.stderr
    assert "Звіт збережено" in rc.stdout
    assert "Формальних проблем не знайдено" in out.read_text(encoding="utf-8")


def test_exit_codes_and_default_output(tmp_path, capsys):
    src = noisy_doc(tmp_path / "w.docx", 2)
    good = good_coursework(tmp_path / "g.docx")
    assert main([str(good)]) == 0
    assert (tmp_path / "g.report.txt").exists()
    assert main([str(src), "--no-txt", "--json"]) == 1  # є помилки (немає обов'язкових частин)
    assert (tmp_path / "w.report.json").exists()
    bad_cfg = tmp_path / "bad.yaml"
    bad_cfg.write_text("page:\n  margin_lef_cm: 3\n", encoding="utf-8")
    assert main([str(src), "-c", str(bad_cfg)]) == 2
    assert "margin_left_cm" in capsys.readouterr().err
    assert main([str(tmp_path / "nope.docx")]) == 2


def test_batch_folder_with_summary(tmp_path):
    folder = tmp_path / "роботи"
    folder.mkdir()
    good_coursework(folder / "a.docx")
    noisy_doc(folder / "b.docx", 2)
    (folder / "~$a.docx").write_bytes(b"lock")
    out = tmp_path / "out"
    rc = main([str(folder), "--html", "--docx", "-o", str(out), "--summary", str(tmp_path / "s.xlsx")])
    assert rc == 1  # у b.docx є помилки
    names = sorted(p.name for p in out.iterdir())
    assert names == ["a.checked.docx", "a.report.html", "a.report.txt", "b.checked.docx", "b.report.html", "b.report.txt"]
    from openpyxl import load_workbook
    ws = load_workbook(tmp_path / "s.xlsx").active
    assert [c.value for c in ws[1]][:3] == ["Файл", "Помилки", "Попередження"]
    assert ws.max_row == 3


def test_init_and_validate_config(tmp_path, capsys):
    p = tmp_path / "my.yaml"
    assert main(["init-config", "-p", "lab_report", "-o", str(p)]) == 0
    assert "Звіт з лабораторної" in p.read_text(encoding="utf-8")
    assert main(["init-config", "-o", str(p)]) == 2  # без --force не перезаписує
    assert main(["validate-config", str(p)]) == 0
    assert main(["list-rules", "--markdown"]) == 0
    assert "`double_space`" in capsys.readouterr().out
    assert main(["list-profiles"]) == 0


def test_same_names_in_batch_do_not_overwrite(tmp_path):
    for d in ("гр1", "гр2"):
        (tmp_path / d).mkdir()
        good_coursework(tmp_path / d / "Іваненко.docx")
    out = tmp_path / "out"
    assert main([str(tmp_path / "гр1"), str(tmp_path / "гр2"), "-o", str(out)]) == 0
    assert sorted(p.name for p in out.iterdir()) == ["Іваненко.report.txt", "гр2_Іваненко.report.txt"]
