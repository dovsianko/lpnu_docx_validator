"""Запуск справжнього вікна tkinter (у CI під xvfb): збирання, перевірка файлу, закриття.

    xvfb-run -a python tests/gui_real_smoke.py
"""

import sys
import tempfile
import tkinter as tk
from pathlib import Path

sys.path[:0] = [str(Path(__file__).resolve().parents[1]), str(Path(__file__).resolve().parent)]

from docbuilder import good_coursework  # noqa: E402
from lpnu_docx_validator.gui import App  # noqa: E402

tmp = Path(tempfile.mkdtemp())
import os  # noqa: E402

os.environ["DOCX_VALIDATOR_PROFILES"] = str(tmp / "profiles")
root = tk.Tk()
app = App(root)
root.update()
app._add_paths([good_coursework(tmp / "a.docx")])
app.adv_var.set(True)
app._toggle_advanced()
app.run_check()
deadline = [200]


def wait() -> None:
    if app.running and deadline[0] > 0:
        deadline[0] -= 1
        root.after(50, wait)
        return
    assert not app.running, "перевірка не завершилась"
    assert app.outcomes and app.outcomes[0].result.errors == 0, app.outcomes
    print("GUI OK:", app.status.get())
    root.destroy()


root.after(50, wait)
root.mainloop()
