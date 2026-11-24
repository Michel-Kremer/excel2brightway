"""
gui.py
======
Fenster fuer excel2brightway (tkinter, gehoert zu Python - keine Installation
noetig). Start: `ex2bw-gui` (unter Windows ohne Konsolenfenster, siehe
[project.gui-scripts] in pyproject.toml) oder `python -m excel2brightway.gui`.

Das Fenster ruft nur die Kernfunktionen aus core/ auf; lange Schritte laufen
in einem eigenen Thread, Meldungen kommen ueber eine Queue ins Log.
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from .core.apply import DataChangedError, apply_review, summarize
from .core.backend import list_projects
from .core.checks import CheckOptions, expand_selection, run_checks
from .core.importer import ImportProblem
from .core.model import FEHLER, HINWEIS, TEST_FLUSS, TEST_FORMAT, TEST_VERWEIS, WARNUNG
from .core.resolve import ResolveConfig
from .core.review_sheet import WorkbookLockedError
from .core.settings import Settings

MODE_CHECK = "pruefen"
MODE_WRITE = "schreiben"
MODE_MIGRATE = "umstellen"


def open_in_excel(path):
    path = str(path)
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa: S606 - oeffnet die Datei mit dem Standardprogramm
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except OSError as exc:
        messagebox.showwarning("Oeffnen fehlgeschlagen", f"{path} konnte nicht geoeffnet werden:\n{exc}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("excel2brightway")
        self.minsize(760, 560)
        self.settings = Settings.load()
        self.queue = queue.Queue()
        self._is_busy = False
        self.backends = {}
        self.last_report = None
        self.last_report_paths = None

        self.paths_var = tk.StringVar(value="; ".join(self.settings.last_paths))
        self.project_var = tk.StringVar(value=self.settings.project)
        self.locations_var = tk.StringVar(value=", ".join(self.settings.location_priority))
        self.test_vars = {t: tk.BooleanVar(value=t in self.settings.tests) for t in (TEST_FORMAT, TEST_FLUSS, TEST_VERWEIS)}
        self.bypass_var = tk.BooleanVar(value=self.settings.bypass_format)
        self.open_after_check = tk.BooleanVar(value=True)
        self.mode_var = tk.StringVar(value=MODE_CHECK)
        self.status_var = tk.StringVar(value="Noch nicht geprueft.")

        self._build()
        self._refresh_projects()
        self.after(100, self._poll)
        self.protocol("WM_DELETE_WINDOW", self._close)

    # ------------------------------------------------------------------ Aufbau
    def _build(self):
        pad = {"padx": 6, "pady": 4}
        frame = ttk.Frame(self, padding=8)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Datei(en)/Ordner:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.paths_var).grid(row=0, column=1, sticky="ew", **pad)
        buttons = ttk.Frame(frame)
        buttons.grid(row=0, column=2, sticky="e")
        ttk.Button(buttons, text="Datei(en) ...", command=self._choose_files).pack(side="left", padx=2)
        ttk.Button(buttons, text="Ordner ...", command=self._choose_folder).pack(side="left", padx=2)

        ttk.Label(frame, text="Projekt:").grid(row=1, column=0, sticky="w", **pad)
        self.project_box = ttk.Combobox(frame, textvariable=self.project_var, state="readonly")
        self.project_box.grid(row=1, column=1, sticky="ew", **pad)
        self.project_box.bind("<<ComboboxSelected>>", lambda e: self._load_databases())
        ttk.Button(frame, text="Neu laden", command=self._refresh_projects).grid(row=1, column=2, sticky="e", **pad)

        ttk.Label(frame, text="Verknuepfungs-\nDatenbanken:").grid(row=2, column=0, sticky="nw", **pad)
        db_frame = ttk.Frame(frame)
        db_frame.grid(row=2, column=1, columnspan=2, sticky="ew", **pad)
        self.db_list = tk.Listbox(db_frame, selectmode="multiple", height=5, exportselection=False)
        self.db_list.pack(side="left", fill="x", expand=True)
        db_side = ttk.Frame(db_frame)
        db_side.pack(side="left", padx=6, fill="y")
        ttk.Label(db_side, text="keine Auswahl =\nalle Datenbanken\ndes Projekts", foreground="gray").pack(anchor="w")
        self.import_button = ttk.Button(db_side, text="Datenbank importieren ...", command=self._import_dialog)
        self.import_button.pack(anchor="w", pady=(6, 0))

        ttk.Label(frame, text="Standort-Rangfolge:").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.locations_var).grid(row=3, column=1, sticky="ew", **pad)

        tests = ttk.LabelFrame(frame, text="Tests", padding=6)
        tests.grid(row=4, column=0, columnspan=3, sticky="ew", **pad)
        ttk.Checkbutton(tests, text="Format", variable=self.test_vars[TEST_FORMAT]).pack(side="left", padx=4)
        ttk.Checkbutton(tests, text="Fluss", variable=self.test_vars[TEST_FLUSS]).pack(side="left", padx=4)
        ttk.Checkbutton(tests, text="Verweis", variable=self.test_vars[TEST_VERWEIS]).pack(side="left", padx=4)
        ttk.Separator(tests, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Checkbutton(tests, text="Format-Fehler ueberbruecken", variable=self.bypass_var).pack(side="left", padx=4)
        ttk.Checkbutton(tests, text="nach dem Pruefen in Excel oeffnen", variable=self.open_after_check).pack(side="left", padx=4)

        status = ttk.Frame(frame)
        status.grid(row=5, column=0, columnspan=3, sticky="ew", **pad)
        ttk.Label(status, text="Status:").pack(side="left")
        self.status_label = ttk.Label(status, textvariable=self.status_var, font=("TkDefaultFont", 10, "bold"))
        self.status_label.pack(side="left", padx=6)

        actions = ttk.LabelFrame(frame, text="Was soll mit der Datei passieren?", padding=6)
        actions.grid(row=6, column=0, columnspan=3, sticky="ew", **pad)
        actions.columnconfigure(1, weight=1)

        ttk.Radiobutton(actions, text="Pruefen und korrigieren", value=MODE_CHECK, variable=self.mode_var,
                        command=self._update_buttons).grid(row=0, column=0, sticky="w")
        row0 = ttk.Frame(actions)
        row0.grid(row=0, column=1, sticky="e")
        self.check_button = ttk.Button(row0, text="Pruefen", command=self._check)
        self.check_button.pack(side="left", padx=2)
        self.apply_button = ttk.Button(row0, text="Auswahl uebernehmen", command=self._apply)
        self.apply_button.pack(side="left", padx=2)
        self.open_button = ttk.Button(row0, text="In Excel oeffnen", command=self._open_selected)
        self.open_button.pack(side="left", padx=2)

        ttk.Radiobutton(actions, text="Nach Brightway schreiben", value=MODE_WRITE, variable=self.mode_var,
                        command=self._update_buttons).grid(row=1, column=0, sticky="w")
        self.write_button = ttk.Button(actions, text="Schreiben ...", command=self._write)
        self.write_button.grid(row=1, column=1, sticky="e", padx=2)

        ttk.Radiobutton(actions, text="Datenbank/Version umstellen", value=MODE_MIGRATE, variable=self.mode_var,
                        command=self._update_buttons).grid(row=2, column=0, sticky="w")
        self.migrate_button = ttk.Button(actions, text="Umstellen ...", command=self._migrate_dialog)
        self.migrate_button.grid(row=2, column=1, sticky="e", padx=2)

        log_frame = ttk.Frame(frame)
        log_frame.grid(row=7, column=0, columnspan=3, sticky="nsew", **pad)
        frame.rowconfigure(7, weight=1)
        self.log_toggle = ttk.Button(log_frame, text="Log ausblenden", command=self._toggle_log)
        self.log_toggle.pack(anchor="w")
        self.log = ScrolledText(log_frame, height=12, state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True)
        self._update_buttons()

    # ----------------------------------------------------------------- Helfer
    def _paths(self):
        return [Path(p.strip()) for p in self.paths_var.get().split(";") if p.strip()]

    def _selected_databases(self):
        return [self.db_list.get(i) for i in self.db_list.curselection()]

    def _check_options(self, tests=None):
        selected = tests or tuple(t for t, var in self.test_vars.items() if var.get())
        locations = tuple(x.strip() for x in self.locations_var.get().split(",") if x.strip())
        return CheckOptions(
            tests=selected,
            bypass_format=self.bypass_var.get(),
            resolve_config=ResolveConfig(
                link_databases=self._selected_databases() or None,
                location_priority=locations or ("GLO", "RER", "RoW"),
            ),
        )

    def _save_settings(self):
        self.settings.project = self.project_var.get()
        self.settings.link_databases = self._selected_databases()
        self.settings.location_priority = [x.strip() for x in self.locations_var.get().split(",") if x.strip()]
        self.settings.last_paths = [str(p) for p in self._paths()]
        self.settings.tests = [t for t, var in self.test_vars.items() if var.get()]
        self.settings.bypass_format = self.bypass_var.get()
        self.settings.save()

    def _log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text.rstrip() + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _toggle_log(self):
        if self.log.winfo_ismapped():
            self.log.pack_forget()
            self.log_toggle.configure(text="Log einblenden")
        else:
            self.log.pack(fill="both", expand=True)
            self.log_toggle.configure(text="Log ausblenden")

    def _ready_to_write(self):
        return (
            self.last_report is not None and self.last_report.ok
            and self.last_report_paths == self._paths()
            and bool(self.project_var.get())
        )

    def _update_buttons(self):
        busy = self._is_busy
        state = lambda enabled: "normal" if enabled and not busy else "disabled"  # noqa: E731
        mode = self.mode_var.get()
        self.check_button.configure(state=state(mode == MODE_CHECK))
        self.apply_button.configure(state=state(mode == MODE_CHECK))
        self.open_button.configure(state=state(True))
        self.import_button.configure(state=state(True))
        self.write_button.configure(state=state(mode == MODE_WRITE and self._ready_to_write()))
        self.migrate_button.configure(state=state(mode == MODE_MIGRATE and self._ready_to_write()))

    def _set_status(self, report):
        counts = report.counts()
        text = f"{counts[FEHLER]} Fehler, {counts[WARNUNG]} Warnung(en), {counts[HINWEIS]} Hinweis(e)"
        if report.errors:
            text += f", {len(report.errors)} sonstige Fehler (Log)"
        text += " - bereit zum Schreiben" if report.ok else " - nicht bereit"
        self.status_var.set(text)
        self.status_label.configure(foreground="darkgreen" if report.ok else "firebrick")

    def _backend(self, project=None):
        project = project or self.project_var.get()
        if not project:
            return None
        if project not in self.backends:
            from .core.backend import BrightwayBackend

            self.backends[project] = BrightwayBackend(project)
        return self.backends[project]

    # ------------------------------------------------------------- Threading
    def _run(self, label, work, done=None):
        if self._is_busy:
            return
        self._is_busy = True
        self._update_buttons()
        self._log(f"--- {label} ...")
        self.configure(cursor="watch")

        def target():
            try:
                result = work()
                self.queue.put(("done", done, result))
            except Exception as exc:  # Fehler im Fenster anzeigen statt abstuerzen
                self.queue.put(("error", label, (exc, traceback.format_exc())))

        threading.Thread(target=target, daemon=True).start()

    def _poll(self):
        try:
            while True:
                kind, callback, payload = self.queue.get_nowait()
                if kind == "log":  # Zwischenmeldung aus dem Arbeits-Thread
                    self._log(payload)
                    continue
                self._is_busy = False
                self.configure(cursor="")
                if kind == "done":
                    if callback is not None:
                        callback(payload)
                elif kind == "error":
                    exc, trace = payload
                    self._log(trace)
                    if isinstance(exc, (WorkbookLockedError, DataChangedError, ImportProblem)):
                        messagebox.showwarning("excel2brightway", str(exc))
                    else:
                        messagebox.showerror("excel2brightway", f"{callback} fehlgeschlagen:\n{exc}")
                self._update_buttons()
        except queue.Empty:
            pass
        self.after(100, self._poll)

    # --------------------------------------------------------------- Aktionen
    def _choose_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx")])
        if files:
            self.paths_var.set("; ".join(files))
            self._update_buttons()

    def _choose_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.paths_var.set(folder)
            self._update_buttons()

    def _refresh_projects(self):
        def work():
            return list_projects()

        def done(projects):
            self.project_box.configure(values=projects)
            if not projects:
                self._log("Keine Brightway-Projekte gefunden (bw2data installiert?). Ohne Projekt "
                          "werden nur Format und Verweise innerhalb der Dateien geprueft.")
            if self.project_var.get() in projects:
                self._load_databases()

        self._run("Projekte laden", work, done)

    def _load_databases(self):
        project = self.project_var.get()
        if not project:
            return

        def work():
            self.backends.pop(project, None)
            return self._backend(project).databases()

        def done(databases):
            self.db_list.delete(0, "end")
            for name in databases:
                self.db_list.insert("end", name)
                if name in self.settings.link_databases:
                    self.db_list.selection_set("end")
            self._log(f"Projekt '{project}': {len(databases)} Datenbank(en).")

        self._run("Datenbanken laden", work, done)

    def _check(self):
        paths = self._paths()
        if not paths:
            messagebox.showinfo("excel2brightway", "Bitte zuerst Datei(en) oder einen Ordner waehlen.")
            return
        options = self._check_options()
        if not options.tests:
            messagebox.showinfo("excel2brightway", "Bitte mindestens einen Test waehlen.")
            return
        self._save_settings()
        project = self.project_var.get()

        def work():
            return run_checks(paths, backend=self._backend(project), options=options)

        def done(report):
            self.last_report = report
            self.last_report_paths = paths if set(options.tests) == {TEST_FORMAT, TEST_FLUSS, TEST_VERWEIS} else None
            self._set_status(report)
            self._log(report.summary())
            for issue in report.issues:
                self._log(f"  [{issue.test}] {issue.severity} {issue.location_label()}: {issue.message}")
            for error in report.errors:
                self._log(f"  FEHLER: {error}")
            if report.log_path:
                self._log(f"Fehler-Log: {report.log_path}")
            with_issues = [p for p in report.selected if report.issues_for(p)]
            if with_issues and self.open_after_check.get():
                open_in_excel(with_issues[0])
                if len(with_issues) > 1:
                    self._log(f"Pruefblatt auch in: {', '.join(p.name for p in with_issues[1:])}")

        self._run("Pruefen", work, done)

    def _apply(self):
        files = [p for p in expand_selection(self._paths()) if p.suffix.lower() == ".xlsx"]
        if not files:
            return

        def work():
            return [apply_review(p) for p in files]

        def done(results):
            changed = False
            for result in results:
                self._log(summarize(result))
                changed = changed or result.changed
            if changed:
                self._log("Uebernommen - es wird erneut geprueft.")
                self._check()
            else:
                messagebox.showinfo("excel2brightway", "Keine Auswahl im Pruefblatt gefunden (Datei gespeichert?).")

        self._run("Auswahl uebernehmen", work, done)

    def _open_selected(self):
        files = expand_selection(self._paths())
        if files:
            open_in_excel(files[0])

    def _write(self):
        from .core.write import describe_plan, execute_write, plan_write

        paths = self._paths()
        project = self.project_var.get()
        options = self._check_options()

        def plan():
            return plan_write(paths, self._backend(project), options, allow_broken_references=True)

        def planned(result):
            report, plans = result
            self.last_report = report
            self._set_status(report)
            if not plans:
                self._log(report.summary())
                messagebox.showwarning("excel2brightway", "Nicht geschrieben - bitte die Meldungen im Pruefblatt beheben.")
                return
            text = "\n\n".join(describe_plan(p) for p in plans)
            allow_delete = allow_broken = False
            if any(p.broken_references for p in plans):
                if not messagebox.askyesno(
                    "Verweise wuerden brechen",
                    text + "\n\nAndere Datenbanken zeigen auf Aktivitaeten, die geloescht wuerden. "
                           "Trotzdem schreiben?", icon="warning",
                ):
                    return
                allow_broken = True
            if any(p.to_delete for p in plans):
                if not messagebox.askyesno(
                    "Aktivitaeten loeschen?",
                    text + "\n\nAktivitaeten, die in keiner Datei mehr stehen, werden in Brightway geloescht. "
                           "Fortfahren?", icon="warning",
                ):
                    return
                allow_delete = True
            if not (allow_delete or allow_broken) and not messagebox.askyesno(
                "Schreiben", f"{text}\n\nIn Projekt '{project}' schreiben?",
            ):
                return

            def work():
                return execute_write(plans, project, allow_delete=allow_delete, allow_broken_references=allow_broken)

            def written(result_list):
                self.backends.pop(project, None)  # Projekt hat sich geaendert
                for database, count in result_list:
                    self._log(f"{database}: {count} Aktivitaet(en) geschrieben.")
                messagebox.showinfo("excel2brightway", "Geschrieben.")

            self._run("Schreiben", work, written)

        self._run("Erneut pruefen", plan, planned)

    def _migrate_dialog(self):
        MigrateDialog(self)

    def _import_dialog(self):
        ImportDialog(self)

    def _thread_log(self, text):
        """_log fuer Arbeits-Threads (tkinter darf nur im Haupt-Thread angefasst werden)."""
        self.queue.put(("log", None, str(text)))

    def _close(self):
        self._save_settings()
        self.destroy()


class MigrateDialog(tk.Toplevel):
    """Dialog 'Datenbank/Version umstellen'."""

    def __init__(self, app: App):
        super().__init__(app)
        self.app = app
        self.title("Datenbank/Version umstellen")
        self.transient(app)
        self.source_var = tk.StringVar()
        self.target_project_var = tk.StringVar(value=app.project_var.get())
        self.target_var = tk.StringVar()
        self.source_bio_var = tk.StringVar()
        self.target_bio_var = tk.StringVar()

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill="both", expand=True)
        source_dbs = list(app.db_list.get(0, "end"))
        rows = [
            ("Quelle (technosphere):", ttk.Combobox(frame, textvariable=self.source_var, values=source_dbs, state="readonly")),
            ("Quelle (biosphere, optional):", ttk.Combobox(frame, textvariable=self.source_bio_var, values=[""] + source_dbs, state="readonly")),
            ("Zielprojekt:", ttk.Combobox(frame, textvariable=self.target_project_var, values=app.project_box.cget("values"), state="readonly")),
            ("Ziel (technosphere):", None),
            ("Ziel (biosphere, optional):", None),
        ]
        self.target_box = ttk.Combobox(frame, textvariable=self.target_var, state="readonly")
        self.target_bio_box = ttk.Combobox(frame, textvariable=self.target_bio_var, state="readonly")
        rows[3] = (rows[3][0], self.target_box)
        rows[4] = (rows[4][0], self.target_bio_box)
        for index, (label, widget) in enumerate(rows):
            ttk.Label(frame, text=label).grid(row=index, column=0, sticky="w", padx=4, pady=3)
            widget.grid(row=index, column=1, sticky="ew", padx=4, pady=3)
        frame.columnconfigure(1, weight=1)
        rows[2][1].bind("<<ComboboxSelected>>", lambda e: self._load_targets())
        ttk.Label(frame, foreground="gray", wraplength=420, text=(
            "Die umgestellten Dateien landen in einem Unterordner mit dem Namen der Ziel-Datenbank; die "
            "Originale bleiben unveraendert. Fluesse ohne eindeutigen Treffer im Ziel erscheinen danach im Pruefblatt."
        )).grid(row=len(rows), column=0, columnspan=2, sticky="w", pady=6)
        buttons = ttk.Frame(frame)
        buttons.grid(row=len(rows) + 1, column=0, columnspan=2, sticky="e")
        ttk.Button(buttons, text="Abbrechen", command=self.destroy).pack(side="right", padx=2)
        ttk.Button(buttons, text="Umstellen", command=self._run).pack(side="right", padx=2)
        self._load_targets()

    def _load_targets(self):
        project = self.target_project_var.get()
        if not project:
            return
        try:
            databases = self.app._backend(project).databases()
        except Exception as exc:
            messagebox.showerror("excel2brightway", f"Projekt '{project}' nicht lesbar:\n{exc}", parent=self)
            return
        self.target_box.configure(values=databases)
        self.target_bio_box.configure(values=[""] + databases)

    def _run(self):
        if not self.source_var.get() or not self.target_var.get():
            messagebox.showinfo("excel2brightway", "Bitte Quelle und Ziel waehlen.", parent=self)
            return
        from .core.migrate import migrate_files

        app = self.app
        paths = app._paths()
        options = app._check_options()
        source, target = self.source_var.get(), self.target_var.get()
        source_bio, target_bio = self.source_bio_var.get() or None, self.target_bio_var.get() or None
        target_project = self.target_project_var.get()
        source_project = app.project_var.get()
        self.destroy()

        def work():
            return migrate_files(
                paths, app._backend(source_project), source, app._backend(target_project), target,
                source_biosphere=source_bio, target_biosphere=target_bio, options=options,
            )

        def done(result):
            app._log(result.summary())
            for path in result.files:
                app._log(f"  neue Datei: {path}")
            if result.report is not None:
                app._log(result.report.summary())
            if result.files:
                app.paths_var.set("; ".join(str(p) for p in result.files))
                app.project_var.set(target_project)
                app.last_report = result.report
                app.last_report_paths = list(result.files)
                if result.report is not None:
                    app._set_status(result.report)
                app._load_databases()
                open_in_excel(result.files[0])

        app._run("Umstellen", work, done)


class ImportDialog(tk.Toplevel):
    """Dialog 'Datenbank importieren': ecoinvent-Archiv (.7z/.zip) oder Ordner in ein Projekt einlesen."""

    def __init__(self, app: App):
        super().__init__(app)
        self.app = app
        self.title("Datenbank importieren")
        self.transient(app)
        self.source_var = tk.StringVar()
        self.project_choice_var = tk.StringVar(value=app.project_var.get())
        self.database_var = tk.StringVar()
        self.biosphere_var = tk.StringVar()
        self.version = None

        frame = ttk.Frame(self, padding=10)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Quelle:").grid(row=0, column=0, sticky="w", padx=4, pady=3)
        ttk.Entry(frame, textvariable=self.source_var, width=50).grid(row=0, column=1, sticky="ew", padx=4, pady=3)
        source_buttons = ttk.Frame(frame)
        source_buttons.grid(row=0, column=2, sticky="e")
        ttk.Button(source_buttons, text="Archiv ...", command=self._choose_archive).pack(side="left", padx=2)
        ttk.Button(source_buttons, text="Ordner ...", command=self._choose_folder).pack(side="left", padx=2)

        ttk.Label(frame, text="Projekt:").grid(row=1, column=0, sticky="w", padx=4, pady=3)
        project_box = ttk.Combobox(frame, textvariable=self.project_choice_var, values=app.project_box.cget("values"))
        project_box.grid(row=1, column=1, columnspan=2, sticky="ew", padx=4, pady=3)
        project_box.bind("<<ComboboxSelected>>", lambda e: self._load_biospheres())
        project_box.bind("<FocusOut>", lambda e: self._load_biospheres())

        ttk.Label(frame, text="Datenbankname:").grid(row=2, column=0, sticky="w", padx=4, pady=3)
        ttk.Entry(frame, textvariable=self.database_var).grid(row=2, column=1, columnspan=2, sticky="ew", padx=4, pady=3)

        ttk.Label(frame, text="Biosphere:").grid(row=3, column=0, sticky="w", padx=4, pady=3)
        self.biosphere_box = ttk.Combobox(frame, textvariable=self.biosphere_var)
        self.biosphere_box.grid(row=3, column=1, columnspan=2, sticky="ew", padx=4, pady=3)

        ttk.Label(frame, foreground="gray", wraplength=480, text=(
            "Projekt: vorhandenes waehlen oder neuen Namen eintippen (wird angelegt). Biosphere: eine vorhandene "
            "Datenbank wird verwendet, ein neuer Name wird aus MasterData des Archivs angelegt. Archive werden "
            "neben der Archivdatei entpackt. Der Import dauert je nach Datenbank 20-60 Minuten; bei ungelinkten "
            "Exchanges wird nichts geschrieben."
        )).grid(row=4, column=0, columnspan=3, sticky="w", pady=6)
        buttons = ttk.Frame(frame)
        buttons.grid(row=5, column=0, columnspan=3, sticky="e")
        ttk.Button(buttons, text="Abbrechen", command=self.destroy).pack(side="right", padx=2)
        ttk.Button(buttons, text="Importieren", command=self._run).pack(side="right", padx=2)
        self._load_biospheres()

    def _set_source(self, path):
        from .core.importer import suggest_names

        self.source_var.set(path)
        database, biosphere, self.version = suggest_names(path)
        self.database_var.set(database)
        existing = list(self.biosphere_box.cget("values"))
        if not self.biosphere_var.get() or self.biosphere_var.get() not in existing:
            self.biosphere_var.set(biosphere)

    def _choose_archive(self):
        path = filedialog.askopenfilename(
            parent=self, filetypes=[("ecoinvent-Archiv", "*.7z *.zip"), ("Alle Dateien", "*.*")],
        )
        if path:
            self._set_source(path)

    def _choose_folder(self):
        path = filedialog.askdirectory(parent=self)
        if path:
            self._set_source(path)

    def _load_biospheres(self):
        from .core.backend import is_biosphere_like

        project = self.project_choice_var.get().strip()
        if project not in self.app.project_box.cget("values"):
            self.biosphere_box.configure(values=[])
            return
        try:
            databases = self.app._backend(project).databases()
        except Exception as exc:
            messagebox.showerror("excel2brightway", f"Projekt '{project}' nicht lesbar:\n{exc}", parent=self)
            return
        self.biosphere_box.configure(values=[d for d in databases if is_biosphere_like(d)])

    def _run(self):
        from .core.importer import import_ecospold

        source =self.source_var.get().strip()
        project = self.project_choice_var.get().strip()
        database = self.database_var.get().strip()
        biosphere = self.biosphere_var.get().strip()
        if not (source and project and database and biosphere):
            messagebox.showinfo("excel2brightway", "Bitte Quelle, Projekt, Datenbankname und Biosphere angeben.", parent=self)
            return
        is_new_project = project not in self.app.project_box.cget("values")
        question = (
            f"'{Path(source).name}' als Datenbank '{database}' in "
            f"{'das NEUE ' if is_new_project else ''}Projekt '{project}' importieren?\n"
            f"Biosphere: '{biosphere}'\n\nDas kann 20-60 Minuten dauern."
        )
        if not messagebox.askyesno("Datenbank importieren", question, parent=self):
            return
        app, version = self.app, self.version
        self.destroy()

        def work():
            return import_ecospold(source, project, database, biosphere, version=version, log=app._thread_log)

        def done(result):
            app._log(result.summary())
            app.backends.pop(project, None)
            app.project_var.set(project)
            app._refresh_projects()
            messagebox.showinfo("excel2brightway", result.summary())

        app._run("Datenbank importieren", work, done)


def ensure_std_streams():
    """Ohne Konsole (ex2bw-gui.exe / pythonw) sind sys.stdout und sys.stderr None.
    bw2data richtet beim Import einen structlog-Logger auf sys.stdout ein und
    stuerzt dann ab - deshalb fehlende Kanaele auf os.devnull umleiten."""
    for name in ("stdout", "stderr"):
        if getattr(sys, name) is None:
            setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))  # noqa: SIM115 - bleibt offen


def main():
    ensure_std_streams()
    app = App()
    app.mainloop()
    return 0


if __name__ == "__main__":  # pragma: no cover
    main()
