"""Графічний інтерфейс (tkinter — входить у стандартний Python, окремо ставити нічого не треба).

Вкладки: «Перевірка» (файли, запуск, результати), «Налаштування» (параметри профілю),
«Правила» (увімкнення груп і рівні серйозності), «Довідка».
Перетягування файлів працює, якщо встановлено tkinterdnd2 (необов'язково).
"""

from __future__ import annotations

import queue
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .gui_logic import (SEV_LABELS, SEVERITY_CHOICES, apply_form, choice_labels, filter_issues, format_value,
                        load_state, open_path, params_for, rules_by_group, save_state, search_text,
                        sections_in_order, test_regex, user_profile_path)
from .settings import (DEFAULT_PROFILE_ID, RULE_BY_ID, SECTION_TITLES, ConfigError, Param, Settings,
                       default_settings, find_default_config, list_profiles, load_settings_file, save_settings, user_profiles_dir)

try:  # необов'язково: перетягування файлів у вікно
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
except Exception:  # pragma: no cover
    TkinterDnD = None
    DND_FILES = None

PAD = 6
GREY = "#667085"
SEV_COLORS = {"ERROR": "#b42318", "WARNING": "#b54708", "INFO": "#475467"}

HELP_TEXT = """\
ЯК КОРИСТУВАТИСЯ

1. Виберіть профіль угорі (набір вимог до оформлення). «Курсова робота» — типові вимоги кафедри.
2. На вкладці «Перевірка» додайте роботи: кнопки «Додати файли…» / «Додати папку…»
   (або перетягніть файли у список, якщо це підтримується).
3. Позначте, які звіти потрібні, і натисніть «Перевірити».
4. Виберіть роботу в таблиці результатів — нижче з'являться зауваження.
   «Текст для пошуку» скопіюйте й знайдіть у Word через Ctrl+F.
5. «DOCX з примітками» — копія роботи, де зауваження вставлено як примітки Word.
   Її можна одразу надіслати студентові.

ЯК ЗМІНИТИ ВИМОГИ

• Вкладка «Налаштування» — поля, шрифт, інтервали, вирівнювання підписів тощо.
  «Показати розширені» відкриває регулярні вирази (для досвідчених).
• Вкладка «Правила» — вмикайте/вимикайте групи перевірок і змінюйте рівень кожного правила:
  Помилка / Попередження / Інформація / Вимкнено.
• Якщо якесь зауваження заважає — клацніть по ньому правою кнопкою → «Вимкнути правило».
• «Зберегти як…» створює ваш власний профіль. Його можна передати колегам:
  це звичайний текстовий файл .yaml з коментарями, який також можна редагувати вручну.

ВАЖЛИВО

Валідатор перевіряє лише формальне оформлення і не замінює перегляд роботи викладачем.
Документи обробляються лише на вашому комп'ютері й нікуди не надсилаються.
"""


class ScrollFrame(ttk.Frame):
    """Прокручувана область із внутрішнім фреймом ``inner``."""

    def __init__(self, master: Any) -> None:
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        self.bind_all("<MouseWheel>", self._on_wheel, add="+")
        self.bind_all("<Button-4>", lambda e: self._on_wheel(e, -3), add="+")
        self.bind_all("<Button-5>", lambda e: self._on_wheel(e, 3), add="+")

    def _on_wheel(self, e: Any, step: Optional[int] = None) -> None:
        try:
            w = self.winfo_containing(e.x_root, e.y_root)
        except Exception:
            return
        if w is None or not str(w).startswith(str(self)) or not self.winfo_ismapped():
            return
        if step is None:
            step = -(e.delta if sys.platform == "darwin" else e.delta // 120)
        self.canvas.yview_scroll(step, "units")


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.state = load_state()
        self.files: list[Path] = []
        self.outcomes: list[Any] = []
        self.current_issues: list[Any] = []
        self.shown_issues: list[Any] = []
        self.summary_path: Optional[Path] = None
        self.queue: "queue.Queue[tuple]" = queue.Queue()
        self.cancel = threading.Event()
        self.running = False
        self.dirty = False
        self._loading = False
        self.param_widgets: dict[str, tuple[Param, Callable[[], Any], Callable[[Any], None], list[tk.Widget]]] = {}
        self.sev_vars: dict[str, tk.StringVar] = {}
        self.check_vars: dict[str, tk.BooleanVar] = {}
        self.settings: Settings = default_settings()
        self.settings_path: Optional[Path] = None
        self.settings_builtin = True

        root.title(f"Валідатор оформлення DOCX — {__version__}")
        root.geometry(self.state.get("geometry", "1180x760"))
        root.minsize(900, 600)
        self._style()
        self._build_profile_bar()
        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True, padx=PAD, pady=(0, PAD))
        self._build_check_tab()
        self._build_settings_tab()
        self._build_rules_tab()
        self._build_help_tab()
        self.status = tk.StringVar(value="Готово.")
        ttk.Label(root, textvariable=self.status, anchor="w", foreground=GREY).pack(fill="x", padx=PAD, pady=(0, 4))
        self._load_initial_profile()
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ------------------------------------------------------------------
    # Оформлення
    # ------------------------------------------------------------------
    def _style(self) -> None:
        st = ttk.Style(self.root)
        if sys.platform.startswith("win") and "vista" in st.theme_names():
            st.theme_use("vista")
        elif "clam" in st.theme_names():
            st.theme_use("clam")
        st.configure("Big.TButton", padding=(16, 8), font=("TkDefaultFont", 11, "bold"))
        st.configure("Hint.TLabel", foreground=GREY)
        st.configure("Treeview", rowheight=24)
        st.configure("Section.TLabelframe.Label", font=("TkDefaultFont", 10, "bold"))

    # ------------------------------------------------------------------
    # Профіль
    # ------------------------------------------------------------------
    def _build_profile_bar(self) -> None:
        bar = ttk.Frame(self.root, padding=(PAD, PAD, PAD, PAD))
        bar.pack(fill="x")
        ttk.Label(bar, text="Профіль вимог:").pack(side="left")
        self.profile_var = tk.StringVar()
        self.profile_cb = ttk.Combobox(bar, textvariable=self.profile_var, state="readonly", width=46)
        self.profile_cb.pack(side="left", padx=PAD)
        self.profile_cb.bind("<<ComboboxSelected>>", self.on_profile_selected)
        ttk.Button(bar, text="Відкрити файл…", command=self.open_config).pack(side="left")
        self.save_btn = ttk.Button(bar, text="Зберегти", command=self.save_current)
        self.save_btn.pack(side="left", padx=(PAD, 0))
        ttk.Button(bar, text="Зберегти як…", command=self.save_as).pack(side="left", padx=(PAD, 0))
        ttk.Button(bar, text="Папка профілів", command=self.open_profiles_dir).pack(side="left", padx=(PAD, 0))
        self.source_var = tk.StringVar()
        ttk.Label(bar, textvariable=self.source_var, style="Hint.TLabel").pack(side="left", padx=PAD)

    def _refresh_profile_list(self) -> None:
        self.profiles = list_profiles()
        names = [p.name + ("" if p.builtin else "  (мій)") for p in self.profiles]
        self.profile_cb["values"] = names

    def _load_initial_profile(self) -> None:
        self._refresh_profile_list()
        last = self.state.get("profile_path")
        if last and Path(last).is_file():
            if self._load_settings_from(Path(last), quiet=True):
                return
        nearby = find_default_config()  # docx_validator_config.yaml поруч із програмою
        if nearby is not None and self._load_settings_from(nearby, quiet=True):
            return
        builtin = next((p for p in self.profiles if p.id == DEFAULT_PROFILE_ID), None)
        if builtin is not None and self._load_settings_from(builtin.path, quiet=True):
            return
        self._apply_settings_to_form(default_settings(), None, True)

    def _load_settings_from(self, path: Path, quiet: bool = False) -> bool:
        try:
            s = load_settings_file(path)
        except ConfigError as exc:
            if not quiet:
                messagebox.showerror("Помилки в конфігу", str(exc), parent=self.root)
            return False
        builtin = any(p.builtin and p.path.resolve() == path.resolve() for p in self.profiles)
        self._apply_settings_to_form(s, path, builtin)
        return True

    def _apply_settings_to_form(self, s: Settings, path: Optional[Path], builtin: bool) -> None:
        self._loading = True
        self.settings, self.settings_path, self.settings_builtin = s, path, builtin
        for key, (p, _get, set_, _w) in self.param_widgets.items():
            set_(format_value(p, s.values[key]))
        for rid, var in self.sev_vars.items():
            var.set(SEV_LABELS[s.sev(rid)])
        for gid, var in self.check_vars.items():
            var.set(s.enabled(gid))
        self._loading = False
        self.dirty = False
        idx = next((k for k, p in enumerate(self.profiles) if path and p.path.resolve() == path.resolve()), None)
        if idx is not None:
            self.profile_cb.current(idx)
        else:
            self.profile_var.set(s.name + ("  (файл)" if path else ""))
        self._update_source_label()

    def _update_source_label(self) -> None:
        if self.settings_path is None:
            src = "вбудовані значення"
        elif self.settings_builtin:
            src = "вбудований профіль (для змін — «Зберегти як…»)"
        else:
            src = str(self.settings_path)
        self.source_var.set(("● змінено, не збережено · " if self.dirty else "") + src)
        self.save_btn.state(["disabled"] if self.settings_builtin or self.settings_path is None else ["!disabled"])

    def mark_dirty(self, *_: Any) -> None:
        if self._loading:
            return
        if not self.dirty:
            self.dirty = True
            self._update_source_label()

    def _confirm_discard(self) -> bool:
        if not self.dirty:
            return True
        ans = messagebox.askyesnocancel("Незбережені зміни",
                                        "Налаштування змінено. Зберегти зміни перед тим, як продовжити?", parent=self.root)
        if ans is None:
            return False
        if ans:
            return self.save_current() if not self.settings_builtin and self.settings_path else self.save_as()
        return True

    def on_profile_selected(self, _: Any = None) -> None:
        idx = self.profile_cb.current()
        if idx < 0:
            return
        target = self.profiles[idx]
        if self.settings_path and target.path.resolve() == self.settings_path.resolve():
            return
        if not self._confirm_discard() or not self._load_settings_from(target.path):
            self._apply_settings_to_form(self.settings, self.settings_path, self.settings_builtin)
            return
        self.status.set(f"Завантажено профіль «{target.name}».")

    def open_config(self) -> None:
        if not self._confirm_discard():
            return
        path = filedialog.askopenfilename(parent=self.root, title="Відкрити файл налаштувань",
                                          filetypes=[("Налаштування YAML/JSON", "*.yaml *.yml *.json"), ("Усі файли", "*.*")])
        if path and self._load_settings_from(Path(path)):
            self.status.set(f"Завантажено налаштування з {path}.")

    def collect_settings(self, show_errors: bool = True) -> Optional[Settings]:
        values = {key: get() for key, (_p, get, _s, _w) in self.param_widgets.items()}
        sevs = {rid: var.get() for rid, var in self.sev_vars.items()}
        checks = {gid: bool(var.get()) for gid, var in self.check_vars.items()}
        s, errors = apply_form(self.settings, values, sevs, checks)
        if errors:
            if show_errors:
                messagebox.showerror("Неправильні налаштування",
                                     "Виправте, будь ласка:\n\n" + "\n".join(f"• {m}" for _, m in errors[:15]), parent=self.root)
                key = errors[0][0]
                if key in self.param_widgets:
                    self.nb.select(self.settings_tab)
                    widgets = self.param_widgets[key][3]
                    if widgets:
                        widgets[0].focus_set()
            return None
        s.source = str(self.settings_path) if self.settings_path else "GUI"
        return s

    def save_current(self) -> bool:
        if self.settings_builtin or self.settings_path is None:
            return self.save_as()
        s = self.collect_settings()
        if s is None:
            return False
        try:
            save_settings(s, self.settings_path)
        except Exception as exc:
            messagebox.showerror("Не вдалося зберегти", str(exc), parent=self.root)
            return False
        self.settings = s
        self.dirty = False
        self._update_source_label()
        self.status.set(f"Збережено: {self.settings_path}")
        return True

    def save_as(self) -> bool:
        s = self.collect_settings()
        if s is None:
            return False
        dlg = SaveAsDialog(self.root, s.name if not self.settings_builtin else f"{s.name} (моя копія)")
        if dlg.result is None:
            return False
        name, where = dlg.result
        s.values["profile.name"] = name
        if where == "profile":
            path = user_profile_path(name)
        else:
            chosen = filedialog.asksaveasfilename(parent=self.root, title="Зберегти налаштування як файл",
                                                  defaultextension=".yaml", initialfile=f"{name}.yaml",
                                                  filetypes=[("YAML", "*.yaml *.yml")])
            if not chosen:
                return False
            path = Path(chosen)
        try:
            s.extends = None
            save_settings(s, path)
        except Exception as exc:
            messagebox.showerror("Не вдалося зберегти", str(exc), parent=self.root)
            return False
        self._refresh_profile_list()
        self._apply_settings_to_form(load_settings_file(path), path, False)
        self.status.set(f"Збережено профіль «{name}»: {path}")
        return True

    def open_profiles_dir(self) -> None:
        d = user_profiles_dir()
        d.mkdir(parents=True, exist_ok=True)
        try:
            open_path(d)
        except Exception as exc:
            messagebox.showinfo("Папка профілів", f"{d}\n\n({exc})", parent=self.root)

    # ------------------------------------------------------------------
    # Вкладка «Перевірка»
    # ------------------------------------------------------------------
    def _build_check_tab(self) -> None:
        tab = ttk.Frame(self.nb, padding=PAD)
        self.nb.add(tab, text="  Перевірка  ")
        left = ttk.Frame(tab, width=380)
        left.pack(side="left", fill="y", padx=(0, PAD))
        right = ttk.Frame(tab)
        right.pack(side="left", fill="both", expand=True)

        box = ttk.LabelFrame(left, text="Роботи для перевірки", padding=PAD, style="Section.TLabelframe")
        box.pack(fill="both", expand=True)
        lf = ttk.Frame(box)
        lf.pack(fill="both", expand=True)
        self.file_list = tk.Listbox(lf, selectmode="extended", height=12, activestyle="none", width=44)
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=sb.set)
        self.file_list.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        hint = "Перетягніть .docx файли або папки сюди" if TkinterDnD is not None else "Додайте .docx файли кнопками нижче"
        ttk.Label(box, text=hint, style="Hint.TLabel").pack(anchor="w", pady=(4, 0))
        btns = ttk.Frame(box)
        btns.pack(fill="x", pady=(PAD, 0))
        ttk.Button(btns, text="Додати файли…", command=self.add_files).pack(side="left")
        ttk.Button(btns, text="Додати папку…", command=self.add_folder).pack(side="left", padx=4)
        ttk.Button(btns, text="Прибрати", command=self.remove_selected).pack(side="left")
        ttk.Button(btns, text="Очистити", command=self.clear_files).pack(side="left", padx=4)
        if TkinterDnD is not None:
            for w in (self.file_list, box):
                w.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
                w.dnd_bind("<<Drop>>", self.on_drop)  # type: ignore[attr-defined]

        out = ttk.LabelFrame(left, text="Що створити", padding=PAD, style="Section.TLabelframe")
        out.pack(fill="x", pady=PAD)
        fmts = self.state.get("formats", {"txt": True, "html": True, "docx": False, "summary": False})
        self.fmt_vars = {k: tk.BooleanVar(value=fmts.get(k, False)) for k in ("txt", "html", "docx", "summary")}
        for key, label in (("html", "HTML-звіт (зручно переглядати у браузері)"), ("txt", "Текстовий звіт (.txt)"),
                           ("docx", "DOCX з примітками — для студента"), ("summary", "Зведена таблиця Excel по всіх роботах")):
            ttk.Checkbutton(out, text=label, variable=self.fmt_vars[key]).pack(anchor="w")
        row = ttk.Frame(out)
        row.pack(fill="x", pady=(PAD, 0))
        ttk.Label(row, text="Куди зберегти:").pack(side="left")
        self.out_dir_var = tk.StringVar(value=self.state.get("out_dir", ""))
        ttk.Entry(row, textvariable=self.out_dir_var).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row, text="…", width=3, command=self.choose_out_dir).pack(side="left")
        ttk.Label(out, text="Порожньо — поруч із роботами.", style="Hint.TLabel").pack(anchor="w")

        run = ttk.Frame(left)
        run.pack(fill="x")
        self.run_btn = ttk.Button(run, text="▶  Перевірити", style="Big.TButton", command=self.run_check)
        self.run_btn.pack(fill="x")
        self.cancel_btn = ttk.Button(run, text="Зупинити", command=self.cancel.set)
        self.progress = ttk.Progressbar(run, mode="determinate")
        self.progress.pack(fill="x", pady=(PAD, 0))

        # результати
        res = ttk.LabelFrame(right, text="Результати", padding=PAD, style="Section.TLabelframe")
        res.pack(fill="x")
        cols = ("errors", "warnings", "status")
        self.files_tree = ttk.Treeview(res, columns=cols, height=5, selectmode="browse")
        self.files_tree.heading("#0", text="Робота")
        self.files_tree.heading("errors", text="Помилки")
        self.files_tree.heading("warnings", text="Попередження")
        self.files_tree.heading("status", text="Стан")
        self.files_tree.column("#0", width=320)
        for c, w in (("errors", 80), ("warnings", 110), ("status", 160)):
            self.files_tree.column(c, width=w, anchor="center", stretch=False)
        self.files_tree.pack(fill="x")
        self.files_tree.bind("<<TreeviewSelect>>", self.on_file_selected)
        self.files_tree.bind("<Double-1>", lambda e: self.open_report())
        fb = ttk.Frame(res)
        fb.pack(fill="x", pady=(4, 0))
        ttk.Button(fb, text="Відкрити звіт", command=self.open_report).pack(side="left")
        ttk.Button(fb, text="Відкрити DOCX з примітками", command=lambda: self.open_output("docx")).pack(side="left", padx=4)
        ttk.Button(fb, text="Відкрити роботу", command=self.open_source).pack(side="left")
        self.summary_btn = ttk.Button(fb, text="Зведена таблиця", command=self.open_summary)

        iss = ttk.LabelFrame(right, text="Зауваження", padding=PAD, style="Section.TLabelframe")
        iss.pack(fill="both", expand=True, pady=(PAD, 0))
        flt = ttk.Frame(iss)
        flt.pack(fill="x")
        self.sev_filter = {s: tk.BooleanVar(value=s != "INFO") for s in ("ERROR", "WARNING", "INFO")}
        for s, label in (("ERROR", "Помилки"), ("WARNING", "Попередження"), ("INFO", "Інформація")):
            ttk.Checkbutton(flt, text=label, variable=self.sev_filter[s], command=self.refresh_issues).pack(side="left")
        ttk.Label(flt, text="   Пошук:").pack(side="left")
        self.query_var = tk.StringVar()
        self.query_var.trace_add("write", lambda *_: self.refresh_issues())
        ttk.Entry(flt, textvariable=self.query_var, width=28).pack(side="left", padx=4)

        pane = ttk.PanedWindow(iss, orient="vertical")
        pane.pack(fill="both", expand=True, pady=(4, 0))
        tf = ttk.Frame(pane)
        icols = ("sev", "msg", "find", "rule")
        self.issues_tree = ttk.Treeview(tf, columns=icols, show="headings", selectmode="browse")
        for c, t, w in (("sev", "Рівень", 100), ("msg", "Зауваження", 380), ("find", "Текст для пошуку (Ctrl+F у Word)", 300),
                        ("rule", "Правило", 160)):
            self.issues_tree.heading(c, text=t)
            self.issues_tree.column(c, width=w, stretch=c in {"msg", "find"})
        for s, col in SEV_COLORS.items():
            self.issues_tree.tag_configure(s, foreground=col)
        isb = ttk.Scrollbar(tf, orient="vertical", command=self.issues_tree.yview)
        self.issues_tree.configure(yscrollcommand=isb.set)
        self.issues_tree.pack(side="left", fill="both", expand=True)
        isb.pack(side="right", fill="y")
        self.issues_tree.bind("<<TreeviewSelect>>", self.on_issue_selected)
        self.issues_tree.bind("<Button-3>", self.issue_menu)
        self.issues_tree.bind("<Button-2>", self.issue_menu)
        self.issues_tree.bind("<Double-1>", lambda e: self.copy_search())
        pane.add(tf, weight=3)
        df = ttk.Frame(pane)
        self.details = tk.Text(df, height=6, wrap="word", relief="flat", background=self.root.cget("background"))
        self.details.pack(fill="both", expand=True)
        self.details.configure(state="disabled")
        pane.add(df, weight=1)
        ab = ttk.Frame(iss)
        ab.pack(fill="x", pady=(4, 0))
        ttk.Button(ab, text="Копіювати текст для пошуку", command=self.copy_search).pack(side="left")
        ttk.Button(ab, text="Вимкнути це правило", command=lambda: self.set_rule_severity("OFF")).pack(side="left", padx=4)
        ttk.Button(ab, text="Зробити попередженням", command=lambda: self.set_rule_severity("WARNING")).pack(side="left")

    # --- файли ---
    def _add_paths(self, paths: list[Path]) -> None:
        added = 0
        for p in paths:
            items = sorted(p.glob("*.docx")) if p.is_dir() else [p]
            for f in items:
                if f.suffix.lower() != ".docx" or f.name.startswith("~$") or f.name.endswith(".checked.docx"):
                    continue
                if f not in self.files:
                    self.files.append(f)
                    self.file_list.insert("end", f.name)
                    added += 1
        self.status.set(f"Додано файлів: {added}. Усього: {len(self.files)}.")

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(parent=self.root, title="Виберіть роботи",
                                            initialdir=self.state.get("last_dir") or None,
                                            filetypes=[("Документи Word", "*.docx"), ("Усі файли", "*.*")])
        if paths:
            self.state["last_dir"] = str(Path(paths[0]).parent)
            self._add_paths([Path(p) for p in paths])

    def add_folder(self) -> None:
        d = filedialog.askdirectory(parent=self.root, title="Папка з роботами", initialdir=self.state.get("last_dir") or None)
        if d:
            self.state["last_dir"] = d
            self._add_paths([Path(d)])

    def on_drop(self, event: Any) -> None:
        self._add_paths([Path(p) for p in self.root.tk.splitlist(event.data)])

    def remove_selected(self) -> None:
        for idx in reversed(self.file_list.curselection()):
            self.file_list.delete(idx)
            del self.files[idx]

    def clear_files(self) -> None:
        self.file_list.delete(0, "end")
        self.files.clear()

    def choose_out_dir(self) -> None:
        d = filedialog.askdirectory(parent=self.root, title="Куди зберігати звіти")
        if d:
            self.out_dir_var.set(d)

    # --- запуск ---
    def run_check(self) -> None:
        if self.running:
            return
        if not self.files:
            messagebox.showinfo("Немає файлів", "Спершу додайте роботи (.docx) для перевірки.", parent=self.root)
            return
        settings = self.collect_settings()
        if settings is None:
            return
        formats = [f for f in ("txt", "html", "docx") if self.fmt_vars[f].get()]
        out_dir = Path(self.out_dir_var.get()) if self.out_dir_var.get().strip() else None
        want_summary = self.fmt_vars["summary"].get()
        files = list(self.files)
        self.running = True
        self.cancel.clear()
        self.run_btn.state(["disabled"])
        self.cancel_btn.pack(fill="x", pady=(PAD, 0))
        self.progress.configure(maximum=len(files), value=0)
        self.files_tree.delete(*self.files_tree.get_children())
        self._show_issues([])

        def work() -> None:
            try:
                from .batch import run_batch, write_summary
                outcomes = run_batch(files, settings, formats, out_dir,
                                     progress=lambda k, n, f: self.queue.put(("progress", k, n, f)),
                                     cancelled=self.cancel.is_set)
                summary = None
                if want_summary and outcomes:
                    target = (out_dir or files[0].parent) / f"зведення_{datetime.now():%Y%m%d_%H%M}.xlsx"
                    summary = write_summary(outcomes, target)
                self.queue.put(("done", outcomes, summary))
            except Exception:
                self.queue.put(("error", traceback.format_exc()))

        threading.Thread(target=work, daemon=True).start()
        self.root.after(100, self._poll)

    def _poll(self) -> None:
        try:
            while True:
                msg = self.queue.get_nowait()
                if msg[0] == "progress":
                    _, k, n, f = msg
                    self.progress.configure(value=k - 1)
                    self.status.set(f"Перевіряю {k}/{n}: {f.name}…")
                elif msg[0] == "done":
                    self._finish(msg[1], msg[2])
                    return
                elif msg[0] == "error":
                    self._finish([], None)
                    messagebox.showerror("Помилка", msg[1], parent=self.root)
                    return
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _finish(self, outcomes: list[Any], summary: Optional[Path]) -> None:
        self.running = False
        self.run_btn.state(["!disabled"])
        self.cancel_btn.pack_forget()
        self.progress.configure(value=self.progress["maximum"])
        self.show_outcomes(outcomes, summary)
        errs = sum(o.result.errors for o in outcomes)
        note = " (зупинено)" if self.cancel.is_set() else ""
        problems = [e for o in outcomes for e in o.write_errors]
        self.status.set(f"Перевірено робіт: {len(outcomes)}{note}. Помилок усього: {errs}."
                        + (f" Не вдалося записати {len(problems)} файл(и) — можливо, вони відкриті в Word." if problems else ""))
        if problems:
            messagebox.showwarning("Не всі звіти збережено", "\n".join(problems[:10]), parent=self.root)

    def show_outcomes(self, outcomes: list[Any], summary: Optional[Path] = None) -> None:
        self.outcomes = outcomes
        self.summary_path = summary
        self.files_tree.delete(*self.files_tree.get_children())
        for k, o in enumerate(outcomes):
            r = o.result
            status = r.fatal or ("є помилки" if r.errors else ("є попередження" if r.warnings else "OK"))
            tag = "ERROR" if (r.errors or r.fatal) else ("WARNING" if r.warnings else "")
            self.files_tree.insert("", "end", iid=str(k), text=r.path.name, values=(r.errors, r.warnings, status), tags=(tag,))
        for s, col in SEV_COLORS.items():
            self.files_tree.tag_configure(s, foreground=col)
        if summary:
            self.summary_btn.pack(side="left", padx=4)
        else:
            self.summary_btn.pack_forget()
        if outcomes:
            self.files_tree.selection_set("0")
            self.files_tree.focus("0")

    def _selected_outcome(self) -> Optional[Any]:
        sel = self.files_tree.selection()
        return self.outcomes[int(sel[0])] if sel else None

    def on_file_selected(self, _: Any = None) -> None:
        o = self._selected_outcome()
        if o is None:
            return
        if o.result.fatal:
            self._show_issues([])
            self._set_details(o.result.fatal)
            return
        self._show_issues(o.result.issues)

    def _display_sev(self, issue: Any) -> str:
        var = self.sev_vars.get(issue.rule)
        if var is None:
            return issue.severity
        from .gui_logic import SEV_FROM_LABEL
        return SEV_FROM_LABEL.get(var.get(), issue.severity)

    def _show_issues(self, issues: list[Any]) -> None:
        self.current_issues = issues
        self.refresh_issues()

    def refresh_issues(self) -> None:
        sevs = {s for s, v in self.sev_filter.items() if v.get()}

        class _View:
            def __init__(self, i: Any, sev: str) -> None:
                self.issue, self.severity = i, sev
                self.message, self.location, self.rule, self.details = i.message, i.location, i.rule, i.details

        views = [_View(i, self._display_sev(i)) for i in self.current_issues]
        self.shown_issues = filter_issues(views, sevs, self.query_var.get())
        self.issues_tree.delete(*self.issues_tree.get_children())
        for k, v in enumerate(self.shown_issues):
            rule = RULE_BY_ID[v.rule].title if v.rule in RULE_BY_ID else v.rule
            self.issues_tree.insert("", "end", iid=str(k), values=(SEV_LABELS.get(v.severity, v.severity), v.message,
                                                                    search_text(v), rule), tags=(v.severity,))
        self._set_details("")

    def _selected_issue(self) -> Optional[Any]:
        sel = self.issues_tree.selection()
        return self.shown_issues[int(sel[0])] if sel else None

    def on_issue_selected(self, _: Any = None) -> None:
        v = self._selected_issue()
        if v is None:
            return
        rule = RULE_BY_ID.get(v.rule)
        text = f"{v.message}\n\n{v.location}"
        if v.details:
            text += "\n" + "\n".join(v.details)
        if rule:
            text += f"\n\nПравило: {rule.title} ({rule.id})" + (f" — {rule.description}" if rule.description else "")
        self._set_details(text)

    def _set_details(self, text: str) -> None:
        self.details.configure(state="normal")
        self.details.delete("1.0", "end")
        self.details.insert("1.0", text)
        self.details.configure(state="disabled")

    def issue_menu(self, event: Any) -> None:
        row = self.issues_tree.identify_row(event.y)
        if not row:
            return
        self.issues_tree.selection_set(row)
        v = self._selected_issue()
        if v is None:
            return
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label="Копіювати текст для пошуку", command=self.copy_search)
        if v.rule in RULE_BY_ID:
            title = RULE_BY_ID[v.rule].title
            m.add_separator()
            m.add_command(label=f"Вимкнути правило «{title}»", command=lambda: self.set_rule_severity("OFF"))
            m.add_command(label="Зробити помилкою", command=lambda: self.set_rule_severity("ERROR"))
            m.add_command(label="Зробити попередженням", command=lambda: self.set_rule_severity("WARNING"))
            m.add_command(label="Зробити інформацією", command=lambda: self.set_rule_severity("INFO"))
        m.tk_popup(event.x_root, event.y_root)

    def copy_search(self) -> None:
        v = self._selected_issue()
        if v is None:
            return
        text = search_text(v)
        if text:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.status.set(f"Скопійовано: «{text[:60]}». Вставте в пошук Word (Ctrl+F).")

    def set_rule_severity(self, sev: str) -> None:
        v = self._selected_issue()
        if v is None or v.rule not in self.sev_vars:
            return
        self.sev_vars[v.rule].set(SEV_LABELS[sev])
        self.mark_dirty()
        self.refresh_issues()
        self.status.set(f"Правило «{RULE_BY_ID[v.rule].title}»: {SEV_LABELS[sev].lower()}. "
                        "Збережіть профіль, щоб зміна залишилась; повторіть перевірку для оновлених звітів.")

    def _open(self, path: Optional[Path]) -> None:
        if path is None:
            return
        try:
            open_path(path)
        except Exception as exc:
            messagebox.showerror("Не вдалося відкрити", f"{path}\n\n{exc}", parent=self.root)

    def open_output(self, fmt: str) -> None:
        o = self._selected_outcome()
        if o is None:
            return
        path = o.outputs.get(fmt)
        if path is None:
            messagebox.showinfo("Немає файлу", "Цей звіт не створювався. Позначте його в блоці «Що створити» і перевірте ще раз.",
                                parent=self.root)
            return
        self._open(path)

    def open_report(self) -> None:
        o = self._selected_outcome()
        if o is None:
            return
        self.open_output("html" if "html" in o.outputs else ("txt" if "txt" in o.outputs else "docx"))

    def open_source(self) -> None:
        o = self._selected_outcome()
        if o is not None:
            self._open(o.result.path)

    def open_summary(self) -> None:
        self._open(self.summary_path)

    # ------------------------------------------------------------------
    # Вкладка «Налаштування»
    # ------------------------------------------------------------------
    def _build_settings_tab(self) -> None:
        tab = ttk.Frame(self.nb, padding=PAD)
        self.settings_tab = tab
        self.nb.add(tab, text="  Налаштування  ")
        top = ttk.Frame(tab)
        top.pack(fill="x", pady=(0, PAD))
        self.adv_var = tk.BooleanVar(value=bool(self.state.get("advanced", False)))
        ttk.Checkbutton(top, text="Показати розширені налаштування (регулярні вирази тощо)", variable=self.adv_var,
                        command=self._toggle_advanced).pack(side="left")
        ttk.Button(top, text="Скинути зміни", command=self.revert).pack(side="right")
        sf = ScrollFrame(tab)
        sf.pack(fill="both", expand=True)
        self._adv_rows: list[list[tk.Widget]] = []
        for sec in sections_in_order():
            lf = ttk.LabelFrame(sf.inner, text=SECTION_TITLES.get(sec, sec), padding=PAD, style="Section.TLabelframe")
            lf.pack(fill="x", padx=2, pady=(0, PAD))
            lf.columnconfigure(1, weight=1)
            for row, p in enumerate(params_for(sec, advanced=True)):
                widgets = self._param_row(lf, row, p)
                if p.advanced:
                    self._adv_rows.append(widgets)
        self._toggle_advanced()

    def _param_row(self, parent: ttk.Frame, row: int, p: Param) -> list[tk.Widget]:
        lab = ttk.Frame(parent)
        lab.grid(row=row, column=0, sticky="nw", padx=(0, 12), pady=3)
        ttk.Label(lab, text=p.label, wraplength=330).pack(anchor="w")
        if p.help:
            ttk.Label(lab, text=p.help, style="Hint.TLabel", wraplength=330).pack(anchor="w")
        holder = ttk.Frame(parent)
        holder.grid(row=row, column=1, sticky="we", pady=3)
        focusable: tk.Widget
        if p.type == "bool":
            var: Any = tk.BooleanVar()
            w = ttk.Checkbutton(holder, variable=var)
            w.pack(anchor="w")
            var.trace_add("write", self.mark_dirty)
            get, set_ = var.get, var.set
            focusable = w
        elif p.type == "choice":
            var = tk.StringVar()
            w = ttk.Combobox(holder, textvariable=var, values=list(choice_labels(p).values()), state="readonly", width=34)
            w.pack(anchor="w")
            var.trace_add("write", self.mark_dirty)
            get, set_ = var.get, var.set
            focusable = w
        elif p.type in {"list", "regex_list", "float_list"}:
            txt = tk.Text(holder, height=4 if p.type != "regex_list" else 6, width=60, wrap="none", undo=True)
            txt.pack(fill="x")
            txt.bind("<<Modified>>", lambda e, t=txt: (t.edit_modified(False), self.mark_dirty()))
            ttk.Label(holder, text="По одному значенню на рядок.", style="Hint.TLabel").pack(anchor="w")

            def get(t: tk.Text = txt) -> str:
                return t.get("1.0", "end").strip()

            def set_(v: Any, t: tk.Text = txt) -> None:
                t.delete("1.0", "end")
                t.insert("1.0", v)
                t.edit_modified(False)
            focusable = txt
        else:
            var = tk.StringVar()
            inner = ttk.Frame(holder)
            inner.pack(fill="x")
            width = 70 if p.type == "regex" else (40 if p.type == "str" else 12)
            if p.type in {"float", "int"}:
                inc = 1 if p.type == "int" else (0.05 if (p.max or 10) <= 10 else 0.5)
                w = ttk.Spinbox(inner, textvariable=var, width=width, from_=p.min if p.min is not None else -1e9,
                                to=p.max if p.max is not None else 1e9, increment=inc)
            else:
                w = ttk.Entry(inner, textvariable=var, width=width)
            w.pack(side="left", fill="x", expand=p.type in {"regex", "str"})
            var.trace_add("write", self.mark_dirty)
            if p.type == "regex":
                ttk.Button(inner, text="Перевірити…", command=lambda v=var, pp=p: RegexTester(self.root, pp, v)).pack(side="left", padx=4)
            if p.nullable:
                ttk.Label(inner, text="порожньо — не перевіряти", style="Hint.TLabel").pack(side="left", padx=6)
            get, set_ = var.get, var.set
            focusable = w
        self.param_widgets[p.key] = (p, get, set_, [focusable])
        return [lab, holder]

    def _toggle_advanced(self) -> None:
        show = self.adv_var.get()
        self.state["advanced"] = show
        for widgets in self._adv_rows:
            for w in widgets:
                if show:
                    w.grid()
                else:
                    w.grid_remove()

    def revert(self) -> None:
        if self.settings_path is not None:
            self._load_settings_from(self.settings_path)
        else:
            self._apply_settings_to_form(default_settings(), None, True)
        self.status.set("Зміни скасовано.")

    # ------------------------------------------------------------------
    # Вкладка «Правила»
    # ------------------------------------------------------------------
    def _build_rules_tab(self) -> None:
        tab = ttk.Frame(self.nb, padding=PAD)
        self.nb.add(tab, text="  Правила  ")
        top = ttk.Frame(tab)
        top.pack(fill="x", pady=(0, PAD))
        ttk.Label(top, text="Помилка — порушення вимог; Попередження — варто перевірити; "
                            "Інформація — довідково; Вимкнено — не перевіряти.", style="Hint.TLabel").pack(side="left")
        ttk.Button(top, text="Рівні за замовчуванням", command=self.reset_severities).pack(side="right")
        sf = ScrollFrame(tab)
        sf.pack(fill="both", expand=True)
        for g, rules in rules_by_group():
            var = tk.BooleanVar(value=g.enabled)
            var.trace_add("write", self.mark_dirty)
            self.check_vars[g.id] = var
            lf = ttk.LabelFrame(sf.inner, padding=PAD, style="Section.TLabelframe",
                                labelwidget=ttk.Checkbutton(sf.inner, text=f"{g.title} — {g.description}", variable=var))
            lf.pack(fill="x", padx=2, pady=(0, PAD))
            lf.columnconfigure(0, weight=1)
            for row, r in enumerate(rules):
                f = ttk.Frame(lf)
                f.grid(row=row, column=0, sticky="we", pady=2)
                ttk.Label(f, text=r.title).pack(anchor="w")
                if r.description:
                    ttk.Label(f, text=r.description, style="Hint.TLabel", wraplength=650).pack(anchor="w")
                sv = tk.StringVar(value=SEV_LABELS[r.severity])
                sv.trace_add("write", self.mark_dirty)
                self.sev_vars[r.id] = sv
                ttk.Combobox(lf, textvariable=sv, values=SEVERITY_CHOICES, state="readonly", width=16).grid(
                    row=row, column=1, sticky="ne", padx=(8, 0), pady=2)

    def reset_severities(self) -> None:
        for rid, var in self.sev_vars.items():
            var.set(SEV_LABELS[RULE_BY_ID[rid].severity])

    # ------------------------------------------------------------------
    def _build_help_tab(self) -> None:
        tab = ttk.Frame(self.nb, padding=PAD)
        self.nb.add(tab, text="  Довідка  ")
        t = tk.Text(tab, wrap="word", relief="flat", padx=12, pady=12, background=self.root.cget("background"))
        t.insert("1.0", HELP_TEXT + f"\nВерсія: {__version__}\nПрофілі користувача: {user_profiles_dir()}\n")
        t.configure(state="disabled")
        t.pack(fill="both", expand=True)

    def on_close(self) -> None:
        if self.running:
            if not messagebox.askyesno("Перевірка триває", "Перервати перевірку й закрити програму?", parent=self.root):
                return
            self.cancel.set()
        if not self._confirm_discard():
            return
        self.state.update({
            "geometry": self.root.geometry(),
            "out_dir": self.out_dir_var.get(),
            "formats": {k: v.get() for k, v in self.fmt_vars.items()},
            "profile_path": str(self.settings_path) if self.settings_path else "",
        })
        save_state(self.state)
        self.root.destroy()


class SaveAsDialog(tk.Toplevel):
    def __init__(self, master: tk.Misc, name: str) -> None:
        super().__init__(master)
        self.title("Зберегти налаштування")
        self.transient(master)
        self.resizable(False, False)
        self.result: Optional[tuple[str, str]] = None
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Назва профілю:").pack(anchor="w")
        self.name_var = tk.StringVar(value=name)
        e = ttk.Entry(f, textvariable=self.name_var, width=50)
        e.pack(fill="x", pady=(2, 10))
        self.where = tk.StringVar(value="profile")
        ttk.Radiobutton(f, text="Як мій профіль (з'явиться у списку профілів)", variable=self.where, value="profile").pack(anchor="w")
        ttk.Radiobutton(f, text="Як окремий файл… (щоб передати колегам)", variable=self.where, value="file").pack(anchor="w")
        b = ttk.Frame(f)
        b.pack(fill="x", pady=(12, 0))
        ttk.Button(b, text="Зберегти", command=self._ok).pack(side="right")
        ttk.Button(b, text="Скасувати", command=self.destroy).pack(side="right", padx=6)
        self.bind("<Return>", lambda _: self._ok())
        self.bind("<Escape>", lambda _: self.destroy())
        e.focus_set()
        e.select_range(0, "end")
        self.grab_set()
        self.wait_window(self)

    def _ok(self) -> None:
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("Назва", "Вкажіть назву профілю.", parent=self)
            return
        self.result = (name, self.where.get())
        self.destroy()


class RegexTester(tk.Toplevel):
    def __init__(self, master: tk.Misc, param: Param, var: tk.StringVar) -> None:
        super().__init__(master)
        self.title(f"Перевірка шаблону: {param.label}")
        self.transient(master)
        self.var = var
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Шаблон (регулярний вираз):").pack(anchor="w")
        ttk.Entry(f, textvariable=var, width=80).pack(fill="x", pady=(2, 8))
        ttk.Label(f, text="Приклад тексту з документа:").pack(anchor="w")
        self.sample = tk.StringVar(value=_SAMPLES.get(param.key, ""))
        ttk.Entry(f, textvariable=self.sample, width=80).pack(fill="x", pady=(2, 8))
        self.result = tk.StringVar()
        ttk.Label(f, textvariable=self.result, wraplength=600).pack(anchor="w")
        ttk.Button(f, text="Закрити", command=self.destroy).pack(anchor="e", pady=(10, 0))
        for v in (var, self.sample):
            v.trace_add("write", lambda *_: self._update())
        self._update()

    def _update(self) -> None:
        self.result.set(test_regex(self.var.get(), self.sample.get()))


_SAMPLES = {
    "captions.figure_caption_pattern": "Рис. 2.1 — Схема системи",
    "captions.table_caption_pattern": "Таблиця 1.2 — Порівняння методів",
    "captions.figure_reference_pattern": "як показано на рис. 2.1",
    "captions.table_reference_pattern": "дані наведено в табл. 1.2",
    "headings.section_heading_pattern": "РОЗДІЛ 1 АНАЛІЗ ПРЕДМЕТНОЇ ОБЛАСТІ",
    "headings.numbered_heading_pattern": "1.2 Постановка задачі",
    "bibliography.heading_pattern": "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ",
    "appendices.heading_pattern": "ДОДАТОК А",
    "formulas.number_pattern": "E = mc²    (2.1)",
}


def _enable_hidpi() -> None:
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # type: ignore[attr-defined]
        except Exception:
            pass


def run() -> None:
    _enable_hidpi()
    root = TkinterDnD.Tk() if TkinterDnD is not None else tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    run()
