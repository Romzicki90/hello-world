"""Point-and-click window for audit teams: no commands, no files to edit.

    Step 1  Audit details (company, period, where to save results)
    Step 2  Tick the Annexure-I topics
    Step 3  Pick the SAP extract files (the window says which ones the topics need)
    Step 4  Settings: thresholds and criteria, pre-filled with the library defaults
    Run     The working paper and Annexure-II are written to the results folder,
            together with a copy of the settings used (audit_settings.toml).

Everything runs on this computer; nothing is sent anywhere.
"""

from __future__ import annotations

import datetime as dt
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "Audit Analytics Workbench (offline)"
DATE_FMT = "%d.%m.%Y"
DEFAULT_RESULTS = Path.home() / "Documents" / "Audit Workbench Results"
NUMBER_FORMATS = ["1,234,567.89", "1.234.567,89", "1 234 567,89"]
DATE_FORMATS = ["DD.MM.YYYY", "MM/DD/YYYY", "YYYY-MM-DD", "DD/MM/YYYY", "DD-MM-YYYY", "YYYY.MM.DD"]
FILE_TYPES = [("SAP exports", "*.xlsx *.xlsm *.csv *.tsv *.txt *.dat"), ("All files", "*.*")]


def open_folder(path: Path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa: S606 - opens Explorer on the results folder
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except OSError:
        pass


def fmt_date(d) -> str:
    return d.strftime(DATE_FMT) if isinstance(d, dt.date) else (d or "")


def parse_ui_date(text: str, label: str) -> dt.date:
    try:
        return dt.datetime.strptime(text.strip(), DATE_FMT).date()
    except ValueError:
        raise ValueError(f"{label}: write the date as DD.MM.YYYY, for example 31.03.2026") from None


class ScrollFrame(ttk.Frame):
    """A frame with a vertical scrollbar (for long lists)."""

    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        bar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        self.canvas.configure(yscrollcommand=bar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        bar.pack(side="right", fill="y")
        self.canvas.bind_all("<MouseWheel>", self._wheel, add="+")

    def _wheel(self, event):
        widget = self.winfo_containing(event.x_root, event.y_root)
        while widget is not None:
            if widget is self:
                self.canvas.yview_scroll(int(-event.delta / 120) or (-1 if event.delta > 0 else 1), "units")
                return
            widget = widget.master

    def clear(self):
        for w in self.inner.winfo_children():
            w.destroy()


class App:
    def __init__(self, root: tk.Tk, lib):
        self.root = root
        self.lib = lib
        self.log_queue: queue.Queue = queue.Queue()
        self.running = False
        self.last_results: Path | None = None

        self.topic_vars: dict[str, tk.BooleanVar] = {}
        self.files: dict[str, list[str]] = {}
        self.file_labels: dict[str, ttk.Label] = {}
        self.need_labels: dict[str, ttk.Label] = {}
        self.param_values: dict[str, dict] = {}      # test id -> {param: value} overrides
        self.param_widgets: dict[tuple, object] = {}
        self.extra: dict = {}                         # profile sections the window does not edit

        root.title(APP_TITLE)
        root.geometry("1100x760")
        root.minsize(900, 600)
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Big.TButton", font=("Segoe UI", 11, "bold"), padding=8)
        style.configure("Head.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("Help.TLabel", foreground="#444444")
        style.configure("Need.TLabel", foreground="#b00020")
        style.configure("Ok.TLabel", foreground="#1b7a1b")

        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True, padx=8, pady=(8, 4))
        self._build_audit_tab()
        self._build_topics_tab()
        self._build_files_tab()
        self._build_settings_tab()
        self.nb.bind("<<NotebookTabChanged>>", self._tab_changed)

        bottom = ttk.Frame(root)
        bottom.pack(fill="both", padx=8, pady=(0, 8))
        buttons = ttk.Frame(bottom)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Open saved audit...", command=self.open_profile).pack(side="left")
        ttk.Button(buttons, text="Save audit settings...", command=self.save_profile).pack(side="left", padx=4)
        ttk.Button(buttons, text="Try with demo data", command=self.load_demo).pack(side="left", padx=4)
        ttk.Button(buttons, text="Check my files", command=self.check_files).pack(side="left", padx=4)
        self.open_btn = ttk.Button(buttons, text="Open results folder", command=self.open_results,
                                   state="disabled")
        self.open_btn.pack(side="right")
        self.run_btn = ttk.Button(buttons, text="RUN ANALYSIS", style="Big.TButton", command=self.start_run)
        self.run_btn.pack(side="right", padx=8)

        self.progress = ttk.Progressbar(bottom, mode="indeterminate")
        self.progress.pack(fill="x", pady=4)
        logf = ttk.Frame(bottom)
        logf.pack(fill="both")
        self.log = tk.Text(logf, height=9, wrap="word", state="disabled", font=("Consolas", 9))
        sb = ttk.Scrollbar(logf, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.write("Welcome. Work through the four tabs from left to right, then press RUN ANALYSIS.")
        self.write("New to this? Press 'Try with demo data' to see a complete example with made-up data.")
        self.write("All processing happens on this computer. Nothing is uploaded anywhere.")
        self.root.after(150, self._drain_log)
        self._update_needed()

    # ------------------------------------------------------------------ tabs

    def _tab(self, title, help_text):
        frame = ttk.Frame(self.nb, padding=10)
        self.nb.add(frame, text=title)
        ttk.Label(frame, text=help_text, style="Help.TLabel", wraplength=1000, justify="left").pack(
            anchor="w", pady=(0, 8))
        return frame

    def _build_audit_tab(self):
        f = self._tab("1. Audit details", "Fill in the company and the audit period. The cut-off date is the date "
                      "on which stock, open items and CWIP are judged (usually the last day of the period).")
        grid = ttk.Frame(f)
        grid.pack(anchor="w")
        today = dt.date.today()
        fy_start = dt.date(today.year - 1 if today.month < 4 else today.year, 4, 1)
        prev_from = dt.date(fy_start.year - 1, 4, 1)
        prev_to = dt.date(fy_start.year, 3, 31)
        self.v = {
            "company": tk.StringVar(), "unit": tk.StringVar(), "lead_team": tk.StringVar(),
            "period_from": tk.StringVar(value=fmt_date(prev_from)),
            "period_to": tk.StringVar(value=fmt_date(prev_to)),
            "cutoff_date": tk.StringVar(value=fmt_date(prev_to)),
            "decimal_notation": tk.StringVar(value=NUMBER_FORMATS[0]),
            "date_format": tk.StringVar(value=DATE_FORMATS[0]),
            "results": tk.StringVar(value=str(DEFAULT_RESULTS)),
            "fewer_reason": tk.StringVar(),
        }
        rows = [
            ("Company", "company", "e.g. MOIL Limited"),
            ("Unit / plant (optional)", "unit", "e.g. Bokaro Steel Plant, for SAIL unit-wise plans"),
            ("Lead Audit Team", "lead_team", "names as in Appendix-B"),
            ("Audit period from (DD.MM.YYYY)", "period_from", ""),
            ("Audit period to (DD.MM.YYYY)", "period_to", ""),
            ("Cut-off date (DD.MM.YYYY)", "cutoff_date", "usually the same as 'period to'"),
        ]
        for r, (label, key, hint) in enumerate(rows):
            ttk.Label(grid, text=label).grid(row=r, column=0, sticky="w", pady=3)
            ttk.Entry(grid, textvariable=self.v[key], width=48).grid(row=r, column=1, sticky="w", padx=8)
            ttk.Label(grid, text=hint, style="Help.TLabel").grid(row=r, column=2, sticky="w")
        r = len(rows)
        ttk.Label(grid, text="Numbers in the SAP files look like").grid(row=r, column=0, sticky="w", pady=3)
        ttk.Combobox(grid, textvariable=self.v["decimal_notation"], values=NUMBER_FORMATS, state="readonly",
                     width=20).grid(row=r, column=1, sticky="w", padx=8)
        ttk.Label(grid, text="as shown in SAP for the user who took the extract", style="Help.TLabel").grid(
            row=r, column=2, sticky="w")
        ttk.Label(grid, text="Dates in the SAP files look like").grid(row=r + 1, column=0, sticky="w", pady=3)
        ttk.Combobox(grid, textvariable=self.v["date_format"], values=DATE_FORMATS, state="readonly",
                     width=20).grid(row=r + 1, column=1, sticky="w", padx=8)
        ttk.Label(grid, text="Save results in").grid(row=r + 2, column=0, sticky="w", pady=3)
        ttk.Entry(grid, textvariable=self.v["results"], width=48).grid(row=r + 2, column=1, sticky="w", padx=8)
        ttk.Button(grid, text="Choose folder...", command=self._choose_results).grid(row=r + 2, column=2, sticky="w")
        ttk.Label(grid, text="If fewer than four business\nprocesses, reason (Annexure-II)").grid(
            row=r + 3, column=0, sticky="w", pady=3)
        ttk.Entry(grid, textvariable=self.v["fewer_reason"], width=48).grid(row=r + 3, column=1, sticky="w", padx=8)

    def _build_topics_tab(self):
        f = self._tab("2. Topics (Annexure-I)", "Tick the audit topics for this company: ordinarily four topics in "
                      "each of four business processes, sixteen in all (OO-53 para 6). Topics marked 'on-site' have no "
                      "remote test; they will appear in Annexure-II as 'To be analysed on-site'.")
        top = ttk.Frame(f)
        top.pack(fill="x")
        self.topic_count = ttk.Label(top, text="", style="Head.TLabel")
        self.topic_count.pack(side="left")
        ttk.Button(top, text="Clear all", command=lambda: self._set_all_topics(False)).pack(side="right")
        sf = ScrollFrame(f)
        sf.pack(fill="both", expand=True, pady=(6, 0))
        r = 0
        for bp, name in self.lib.business_processes.items():
            topics = [t for t in self.lib.topics.values() if t.bp == bp and t.priority != "CS"]
            ttk.Label(sf.inner, text=f"BP{bp}  {name}", style="Head.TLabel").grid(row=r, column=0, columnspan=2,
                                                                                   sticky="w", pady=(10, 2))
            r += 1
            for t in topics + [t for t in self.lib.topics.values() if t.bp == bp and t.priority == "CS"]:
                r = self._topic_row(sf.inner, t, r)
        self._update_topic_count()

    def _topic_row(self, parent, t, r):
        var = tk.BooleanVar(value=False)
        var.trace_add("write", lambda *_: self._topics_changed())
        self.topic_vars[t.id] = var
        tests = self.lib.tests_for_topic(t.id)
        prio = {"I": "P-I", "II": "P-II"}.get(t.priority, "Company-specific")
        status = "test ready" if tests else "on-site only"
        ttk.Label(parent, text=status, style="Ok.TLabel" if tests else "Help.TLabel", width=12).grid(
            row=r, column=0, sticky="w", padx=(12, 4))
        ttk.Checkbutton(parent, variable=var, text=f"{t.id}   {t.title}   ({prio})").grid(
            row=r, column=1, sticky="w")
        return r + 1

    def _build_files_tab(self):
        f = self._tab("3. SAP files", "For each SAP report, choose the exported file(s): Excel (.xlsx) from the "
                      "list's 'Spreadsheet' export, or text / CSV. You may pick several files for one report (for "
                      "example one per plant). Reports marked NEEDED are required by the topics you ticked. The "
                      "files are only read, never changed.")
        sf = ScrollFrame(f)
        sf.pack(fill="both", expand=True)
        for r, (ds_id, ds) in enumerate(self.lib.datasets.items()):
            box = ttk.Frame(sf.inner, padding=(0, 4))
            box.grid(row=r, column=0, sticky="ew")
            sf.inner.columnconfigure(0, weight=1)
            head = ttk.Frame(box)
            head.pack(fill="x")
            ttk.Label(head, text=ds.title, style="Head.TLabel").pack(side="left")
            self.need_labels[ds_id] = ttk.Label(head, text="")
            self.need_labels[ds_id].pack(side="left", padx=10)
            ttk.Button(head, text="Clear", command=lambda d=ds_id: self._set_files(d, [])).pack(side="right")
            ttk.Button(head, text="Choose file(s)...", command=lambda d=ds_id: self._choose_files(d)).pack(
                side="right", padx=4)
            ttk.Label(box, text="SAP: " + "; ".join(ds.sap_sources), style="Help.TLabel").pack(anchor="w")
            self.file_labels[ds_id] = ttk.Label(box, text="(no file chosen)", wraplength=950)
            self.file_labels[ds_id].pack(anchor="w")
            self.files[ds_id] = []

    def _build_settings_tab(self):
        f = self._tab("4. Settings", "Thresholds and criteria for the tests of the ticked topics, pre-filled with "
                      "standard values. Change them only where the company's rules differ (for example the DoP / "
                      "tender limit, the company's document types, or the sensitive G/L accounts). Lists are written "
                      "with commas: SA, AB. Every change is recorded on the Parameters sheet of the working paper.")
        self.settings_frame = ScrollFrame(f)
        self.settings_frame.pack(fill="both", expand=True)

    # ------------------------------------------------------------------ topic / file state

    def _set_all_topics(self, value: bool):
        for var in self.topic_vars.values():
            var.set(value)

    def selected_topics(self) -> list[str]:
        return [tid for tid, var in self.topic_vars.items() if var.get()]

    def selected_tests(self):
        chosen = set(self.selected_topics())
        order = {tid: i for i, tid in enumerate(self.lib.topics)}
        tests = [t for t in self.lib.tests.values() if chosen & set(t.topics)]
        return sorted(tests, key=lambda t: min(order[x] for x in t.topics))

    def _topics_changed(self):
        self._update_topic_count()
        self._update_needed()

    def _update_topic_count(self):
        topics = self.selected_topics()
        bps = {self.lib.topics[t].bp for t in topics}
        self.topic_count.configure(text=f"Selected: {len(topics)} topics in {len(bps)} business processes "
                                        f"(OO-53: sixteen topics, four processes)")

    def _update_needed(self):
        needed = {ds for t in self.selected_tests() for ds in t.datasets}
        for ds_id, label in self.need_labels.items():
            if ds_id in needed and not self.files.get(ds_id):
                label.configure(text="NEEDED", style="Need.TLabel")
            elif ds_id in needed:
                label.configure(text="needed - file chosen", style="Ok.TLabel")
            else:
                label.configure(text="not needed for the ticked topics", style="Help.TLabel")

    def _choose_files(self, ds_id):
        paths = filedialog.askopenfilenames(title=f"Choose {self.lib.datasets[ds_id].title}", filetypes=FILE_TYPES)
        if paths:
            self._set_files(ds_id, list(paths))

    def _set_files(self, ds_id, paths):
        self.files[ds_id] = [str(p) for p in paths]
        self.file_labels[ds_id].configure(text="\n".join(self.files[ds_id]) or "(no file chosen)")
        self._update_needed()

    def _choose_results(self):
        path = filedialog.askdirectory(title="Save results in")
        if path:
            self.v["results"].set(path)

    # ------------------------------------------------------------------ settings tab

    def _tab_changed(self, _event=None):
        if self.nb.index(self.nb.select()) == 3:
            self._build_settings_rows()

    def _collect_params(self):
        """Read edited values from the settings tab back into self.param_values."""
        for (test_id, name), (widget_var, default) in self.param_widgets.items():
            raw = widget_var.get()
            try:
                value = self._convert(raw, default)
            except ValueError:
                raise ValueError(f"Settings, {test_id}: '{raw}' is not a valid value for "
                                 f"{self.lib.tests[test_id].params[name].label}") from None
            if value == default:
                self.param_values.get(test_id, {}).pop(name, None)
            else:
                self.param_values.setdefault(test_id, {})[name] = value

    @staticmethod
    def _convert(raw, default):
        if isinstance(default, bool):
            return bool(raw)
        if isinstance(default, list):
            items = [x.strip() for x in str(raw).split(",") if x.strip()]
            if default and all(isinstance(x, int) for x in default):
                return [int(x) for x in items]
            return items
        if isinstance(default, int):
            text = str(raw).replace(",", "").strip()
            return int(float(text)) if text else default
        if isinstance(default, float):
            return float(str(raw).replace(",", "").strip())
        return str(raw).strip()

    def _build_settings_rows(self):
        if self.param_widgets:
            self._collect_params()
        sf = self.settings_frame
        sf.clear()
        self.param_widgets = {}
        tests = self.selected_tests()
        if not tests:
            ttk.Label(sf.inner, text="Tick some topics on tab 2 first.").grid(row=0, column=0, sticky="w")
            return
        r = 0
        for t in tests:
            ttk.Label(sf.inner, text=f"{t.id}   {t.title}", style="Head.TLabel").grid(
                row=r, column=0, columnspan=3, sticky="w", pady=(10, 2))
            r += 1
            for name, p in t.params.items():
                current = self.param_values.get(t.id, {}).get(name, p.default)
                label = p.label + ("  (REQUIRED: to be supplied by audit)" if p.required else "")
                ttk.Label(sf.inner, text=label, wraplength=400).grid(row=r, column=0, sticky="w", padx=(12, 8))
                if isinstance(p.default, bool):
                    var = tk.BooleanVar(value=bool(current))
                    ttk.Checkbutton(sf.inner, variable=var).grid(row=r, column=1, sticky="w")
                else:
                    shown = ", ".join(str(x) for x in current) if isinstance(current, list) else str(current)
                    var = tk.StringVar(value=shown)
                    ttk.Entry(sf.inner, textvariable=var, width=36).grid(row=r, column=1, sticky="w")
                if p.help:
                    ttk.Label(sf.inner, text=p.help, style="Help.TLabel", wraplength=280).grid(
                        row=r, column=2, sticky="w", padx=8)
                self.param_widgets[(t.id, name)] = (var, p.default)
                r += 1

    # ------------------------------------------------------------------ profile <-> window

    def to_raw(self) -> dict:
        if self.param_widgets:
            self._collect_params()
        v = {k: var.get().strip() for k, var in self.v.items()}
        if not v["company"]:
            raise ValueError("Tab 1: enter the company name.")
        period_from = parse_ui_date(v["period_from"], "Audit period from")
        period_to = parse_ui_date(v["period_to"], "Audit period to")
        cutoff = parse_ui_date(v["cutoff_date"] or v["period_to"], "Cut-off date")
        if period_to < period_from:
            raise ValueError("Tab 1: 'period to' is earlier than 'period from'.")
        raw = {
            "audit": {"company": v["company"], "unit": v["unit"], "tier": "A", "lead_team": v["lead_team"],
                      "period_from": period_from, "period_to": period_to, "cutoff_date": cutoff},
            "sap": {"decimal_notation": v["decimal_notation"], "date_format": v["date_format"]},
            "datasets": {},
            "annexure_ii": {"topics": self.selected_topics(),
                            "fewer_business_processes_reason": v["fewer_reason"]},
            "params": {k: dict(p) for k, p in self.param_values.items() if p},
        }
        for ds_id, files in self.files.items():
            if files:
                spec = dict(self.extra.get("datasets", {}).get(ds_id, {}))
                spec["files"] = list(files)
                raw["datasets"][ds_id] = spec
        if self.extra.get("topic_plans"):
            raw["annexure_ii"]["topic"] = self.extra["topic_plans"]
        if self.extra.get("common_params"):
            raw["params"]["common"] = self.extra["common_params"]
        return raw

    def from_raw(self, raw: dict):
        audit, sap, annex = raw.get("audit", {}), raw.get("sap", {}), raw.get("annexure_ii", {})
        for key in ("company", "unit", "lead_team"):
            self.v[key].set(audit.get(key, ""))
        for key in ("period_from", "period_to", "cutoff_date"):
            self.v[key].set(fmt_date(audit.get(key)))
        self.v["decimal_notation"].set(sap.get("decimal_notation", NUMBER_FORMATS[0]))
        self.v["date_format"].set(sap.get("date_format", DATE_FORMATS[0]))
        self.v["fewer_reason"].set(annex.get("fewer_business_processes_reason", ""))
        wanted = set(annex.get("topics", []))
        for tid, var in self.topic_vars.items():
            var.set(tid in wanted)
        self.extra = {"datasets": {}, "topic_plans": annex.get("topic", {}),
                      "common_params": raw.get("params", {}).get("common", {})}
        for ds_id in self.files:
            spec = raw.get("datasets", {}).get(ds_id, {})
            self.extra["datasets"][ds_id] = {k: v for k, v in spec.items() if k != "files"}
            self._set_files(ds_id, spec.get("files", []))
        self.param_values = {k: dict(v) for k, v in raw.get("params", {}).items()
                             if k != "common" and k in self.lib.tests}
        self.param_widgets = {}
        self._update_topic_count()
        self._update_needed()

    def open_profile(self):
        from .profile import read_raw

        path = filedialog.askopenfilename(title="Open saved audit settings",
                                          filetypes=[("Audit settings", "*.toml"), ("All files", "*.*")])
        if not path:
            return
        try:
            self.from_raw(read_raw(Path(path)))
            self.v["results"].set(str(Path(path).parent / "results"))
            self.write(f"Opened {path}")
        except Exception as exc:  # show any problem to the user in plain words
            messagebox.showerror(APP_TITLE, f"Could not open these settings:\n\n{exc}")

    def save_profile(self, path: Path | None = None) -> Path | None:
        from .profile import dump_toml

        try:
            raw = self.to_raw()
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return None
        if path is None:
            chosen = filedialog.asksaveasfilename(title="Save audit settings", defaultextension=".toml",
                                                  filetypes=[("Audit settings", "*.toml")],
                                                  initialfile=f"{raw['audit']['company']}.toml")
            if not chosen:
                return None
            path = Path(chosen)
        path.write_text(dump_toml(raw, "# Audit Analytics Workbench settings (saved from the window)"),
                        encoding="utf-8")
        self.write(f"Settings saved to {path}")
        return path

    def load_demo(self):
        from .profile import read_raw
        from .synthetic import generate

        target = Path.home() / "Documents" / "Audit Workbench Demo"
        try:
            self.write(f"Creating made-up SAP files in {target} ...")
            profile = generate(target)
            self.from_raw(read_raw(profile))
            self.v["results"].set(str(target / "results"))
            self.write("Demo loaded: 'Demo Steel Ltd' is fictitious. Look through the tabs, then press RUN ANALYSIS.")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not create the demo:\n\n{exc}")

    # ------------------------------------------------------------------ checking and running

    def check_files(self):
        from .engine import DatasetStatus, check_executable, resolve_params
        from .ingest import load_dataset

        try:
            raw = self.to_raw()
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return
        self.write("Checking files (reading only)...")
        profile = self._profile(raw, Path(self.v["results"].get()) / "unsaved.toml")
        statuses = {}
        for ds_id, src in profile.datasets.items():
            st = DatasetStatus(id=ds_id, title=self.lib.datasets[ds_id].title)
            statuses[ds_id] = st
            try:
                st.load = load_dataset(self.lib.datasets[ds_id], src.files, explicit=src.columns,
                                       decimal_notation=profile.decimal_notation, date_format=profile.date_format)
                rows = sum(f.rows_loaded for f in st.load.files)
                self.write(f"  {st.title}: {rows} rows read; {len(st.load.mapped_fields)} columns recognised")
                for w in st.load.warnings:
                    self.write(f"     WARNING: {w}")
            except Exception as exc:
                self.write(f"  {st.title}: PROBLEM: {exc}")
        ready = 0
        for t in self.selected_tests():
            missing = check_executable(t, statuses, resolve_params(t, profile))
            if missing:
                self.write(f"  {t.id}: cannot run yet: {'; '.join(missing)}")
            else:
                ready += 1
        self.write(f"Check finished: {ready} of {len(self.selected_tests())} tests ready to run.")

    def _profile(self, raw, path):
        from .profile import profile_from_dict

        return profile_from_dict(raw, path)

    def start_run(self):
        if self.running:
            return
        try:
            raw = self.to_raw()
        except ValueError as exc:
            messagebox.showwarning(APP_TITLE, str(exc))
            return
        if not raw["annexure_ii"]["topics"]:
            messagebox.showwarning(APP_TITLE, "Tab 2: tick at least one topic.")
            return
        if not raw["datasets"]:
            messagebox.showwarning(APP_TITLE, "Tab 3: choose at least one SAP file.")
            return
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        out = Path(self.v["results"].get()) / f"{raw['audit']['company']} {stamp}"
        try:
            out.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Cannot create the results folder:\n{exc}")
            return
        settings = self.save_profile(out / "audit_settings.toml")
        if settings is None:
            return
        self.running = True
        self.run_btn.configure(state="disabled")
        self.progress.start(12)
        threading.Thread(target=self._run_worker, args=(settings, out), daemon=True).start()

    def _run_worker(self, settings: Path, out: Path):
        from .engine import EXCEPTIONS, NIL, run
        from .profile import load_profile
        from .report import write_annexure_ii, write_working_paper

        try:
            profile = load_profile(settings)
            result = run(profile, self.lib, progress=lambda m: self.log_queue.put(m))
            self.log_queue.put("Writing Excel files...")
            wp = write_working_paper(result, out)
            ax = write_annexure_ii(result, out)
            lines = []
            for r in result.results:
                detail = f"{r.n_exceptions} {r.test.exception_unit}" if r.executed else r.message
                lines.append(f"  {r.test.id:<18} {r.status:<15} {detail}")
            summary = (f"Finished. {sum(r.status == EXCEPTIONS for r in result.results)} tests found exceptions, "
                       f"{sum(r.status == NIL for r in result.results)} found none, "
                       f"{sum(not r.executed for r in result.results)} could not run.")
            self.log_queue.put(("done", out, lines, summary, wp, ax))
        except Exception as exc:
            self.log_queue.put(("error", exc, traceback.format_exc()))

    def _drain_log(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if isinstance(item, tuple) and item[0] == "done":
                    _, out, lines, summary, wp, ax = item
                    for line in lines:
                        self.write(line)
                    self.write(summary)
                    self.write(f"Results: {wp.name} and {ax.name} in {out}")
                    self._finish(ok=True, out=out)
                    if messagebox.askyesno(APP_TITLE, f"{summary}\n\nOpen the results folder now?"):
                        open_folder(out)
                elif isinstance(item, tuple) and item[0] == "error":
                    _, exc, tb = item
                    self.write("ERROR: " + str(exc))
                    self.write(tb)
                    self._finish(ok=False)
                    messagebox.showerror(APP_TITLE, f"The analysis stopped with a problem:\n\n{exc}\n\n"
                                                    "Details are in the message box at the bottom of the window.")
                else:
                    self.write(str(item))
        except queue.Empty:
            pass
        self.root.after(150, self._drain_log)

    def _finish(self, ok: bool, out: Path | None = None):
        self.running = False
        self.progress.stop()
        self.run_btn.configure(state="normal")
        if ok and out:
            self.last_results = out
            self.open_btn.configure(state="normal")

    def open_results(self):
        if self.last_results:
            open_folder(self.last_results)

    def write(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")


def main():
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = tk.Tk()
    try:
        from .loader import load_library
        import duckdb  # noqa: F401  (checked here so a missing install gives a clear message)
        import openpyxl  # noqa: F401
        lib = load_library()
    except ImportError as exc:
        root.withdraw()
        messagebox.showerror(APP_TITLE, f"The workbench is not installed completely ({exc.name} is missing).\n\n"
                                        "Double-click 'Install' in the workbench folder first.")
        return
    except Exception as exc:
        root.withdraw()
        messagebox.showerror(APP_TITLE, f"The test library could not be loaded:\n\n{exc}")
        return
    App(root, lib)
    root.mainloop()


if __name__ == "__main__":
    main()
