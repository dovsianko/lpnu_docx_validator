"""Мінімальна підробка tkinter для smoke-тесту логіки GUI без дисплея.

Віджети — «поглиначі» викликів; змінні (StringVar/BooleanVar), Listbox і Treeview
поводяться як справжні настільки, наскільки цього потребує gui.py.
"""

from __future__ import annotations

import sys
import types
from typing import Any


class _Any:
    """Віджет, що приймає будь-які виклики."""

    def __init__(self, *a: Any, **kw: Any) -> None:
        self._kw = kw
        self._children: list[Any] = []

    def __getattr__(self, name: str) -> Any:
        if name.startswith("__"):
            raise AttributeError(name)
        return lambda *a, **kw: _Any()

    def __getitem__(self, key: str) -> Any:
        return self._kw.get(key, 0)

    def __iter__(self):
        return iter(())

    def __contains__(self, item: Any) -> bool:
        return False

    def __bool__(self) -> bool:
        return True

    def __setitem__(self, key: str, value: Any) -> None:
        self._kw[key] = value

    def configure(self, *a: Any, **kw: Any) -> None:
        self._kw.update(kw)

    config = configure

    def cget(self, key: str) -> Any:
        return self._kw.get(key, "")

    def __str__(self) -> str:
        return f".w{id(self)}"


class Variable:
    def __init__(self, master: Any = None, value: Any = None) -> None:
        self._v = value if value is not None else self._default
        self._traces: list[Any] = []

    _default: Any = ""

    def get(self) -> Any:
        return self._v

    def set(self, v: Any) -> None:
        self._v = v
        for cb in self._traces:
            cb("", "", "write")

    def trace_add(self, mode: str, cb: Any) -> None:
        self._traces.append(cb)


class StringVar(Variable):
    _default = ""


class BooleanVar(Variable):
    _default = False

    def get(self) -> bool:
        return bool(self._v)


class Listbox(_Any):
    def __init__(self, *a: Any, **kw: Any) -> None:
        super().__init__(*a, **kw)
        self.items: list[str] = []
        self.sel: tuple[int, ...] = ()

    def insert(self, idx: Any, *items: str) -> None:
        self.items.extend(items)

    def delete(self, first: Any, last: Any = None) -> None:
        if first == 0 and last == "end":
            self.items.clear()
        else:
            del self.items[int(first)]

    def curselection(self) -> tuple[int, ...]:
        return self.sel


class Treeview(_Any):
    def __init__(self, *a: Any, **kw: Any) -> None:
        super().__init__(*a, **kw)
        self.rows: dict[str, dict[str, Any]] = {}
        self._sel: tuple[str, ...] = ()

    def insert(self, parent: str, idx: Any, iid: str | None = None, **kw: Any) -> str:
        iid = iid or str(len(self.rows))
        self.rows[iid] = kw
        return iid

    def delete(self, *iids: str) -> None:
        for i in iids:
            self.rows.pop(i, None)

    def get_children(self, item: str = "") -> tuple[str, ...]:
        return tuple(self.rows)

    def selection(self) -> tuple[str, ...]:
        return self._sel

    def selection_set(self, iid: str) -> None:
        self._sel = (iid,)

    def focus(self, *a: Any) -> None:
        pass


class Text(_Any):
    def __init__(self, *a: Any, **kw: Any) -> None:
        super().__init__(*a, **kw)
        self.value = ""

    def insert(self, idx: Any, text: str) -> None:
        self.value += text

    def delete(self, a: Any, b: Any = None) -> None:
        self.value = ""

    def get(self, a: Any, b: Any = None) -> str:
        return self.value + "\n"


class Combobox(_Any):
    def __init__(self, *a: Any, **kw: Any) -> None:
        super().__init__(*a, **kw)
        self._current = -1
        self.var = kw.get("textvariable")

    def current(self, idx: int | None = None) -> int:
        if idx is None:
            return self._current
        self._current = idx
        values = self._kw.get("values") or []
        if self.var is not None and idx < len(values):
            self.var.set(values[idx])
        return idx


class Tk(_Any):
    def __init__(self, *a: Any, **kw: Any) -> None:
        super().__init__(*a, **kw)
        self.tk = types.SimpleNamespace(splitlist=lambda s: tuple(s.split()))
        self.after_calls: list[Any] = []
        self.clip = ""

    def geometry(self, *a: Any) -> str:
        return "1180x760"

    def after(self, ms: int, fn: Any) -> None:
        self.after_calls.append(fn)

    def clipboard_clear(self) -> None:
        self.clip = ""

    def clipboard_append(self, s: str) -> None:
        self.clip += s


def install() -> dict[str, Any]:
    """Підставляє фальшивий tkinter у sys.modules. Повертає підмодулі для налаштування в тестах."""
    tk = types.ModuleType("tkinter")
    for name in ("Canvas", "Menu", "Toplevel", "Misc", "Widget", "Frame", "Label"):
        setattr(tk, name, _Any)
    tk.Tk, tk.StringVar, tk.BooleanVar, tk.Variable = Tk, StringVar, BooleanVar, Variable
    tk.Listbox, tk.Text = Listbox, Text
    ttk = types.ModuleType("tkinter.ttk")
    for name in ("Frame", "Label", "Button", "Checkbutton", "Radiobutton", "Entry", "Spinbox", "Scrollbar",
                 "LabelFrame", "Notebook", "Progressbar", "PanedWindow", "Style"):
        setattr(ttk, name, _Any)
    ttk.Treeview, ttk.Combobox = Treeview, Combobox
    answers: dict[str, Any] = {"askyesnocancel": False, "askyesno": True, "askopenfilename": "", "asksaveasfilename": "",
                               "askopenfilenames": (), "askdirectory": ""}
    shown: list[tuple[str, str]] = []
    mb = types.ModuleType("tkinter.messagebox")
    for name in ("showerror", "showinfo", "showwarning"):
        setattr(mb, name, lambda title, msg, name=name, **kw: shown.append((name, msg)))
    for name in ("askyesnocancel", "askyesno"):
        setattr(mb, name, lambda *a, name=name, **kw: answers[name])
    fd = types.ModuleType("tkinter.filedialog")
    for name in ("askopenfilename", "asksaveasfilename", "askopenfilenames", "askdirectory"):
        setattr(fd, name, lambda *a, name=name, **kw: answers[name])
    tk.ttk, tk.messagebox, tk.filedialog = ttk, mb, fd
    sys.modules.update({"tkinter": tk, "tkinter.ttk": ttk, "tkinter.messagebox": mb, "tkinter.filedialog": fd})
    sys.modules.pop("lpnu_docx_validator.gui", None)
    return {"answers": answers, "shown": shown}
