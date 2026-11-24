"""
Tests fuer gui.py ohne Bildschirm.

Die Fensterklassen erben von tk.Tk bzw. tk.Toplevel. Ein eigenes Attribut
oder eine eigene Methode mit dem Namen einer tkinter-Methode (z. B. `_options`)
ueberdeckt diese und laesst das Fenster beim Start abstuerzen.
"""

import ast
import sys
import tkinter as tk
from pathlib import Path

GUI_PY = Path(__file__).resolve().parents[1] / "src" / "excel2brightway" / "gui.py"
BASES = {"tk.Tk": tk.Tk, "tk.Toplevel": tk.Toplevel}


def _own_names(cls: ast.ClassDef) -> set[str]:
    """Methoden und self-Attribute, die die Klasse selbst definiert."""
    names = {n.name for n in cls.body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))}
    for node in ast.walk(cls):
        if (
            isinstance(node, ast.Attribute)
            and isinstance(node.value, ast.Name)
            and node.value.id == "self"
            and isinstance(node.ctx, ast.Store)
        ):
            names.add(node.attr)
    return names - {"__init__"}


def test_ohne_konsole_werden_stdout_und_stderr_ersetzt(monkeypatch):
    from excel2brightway import gui

    monkeypatch.setattr(sys, "stdout", None)
    monkeypatch.setattr(sys, "stderr", None)
    gui.ensure_std_streams()
    assert sys.stdout is not None and sys.stderr is not None
    print("geht")  # darf nicht abstuerzen


def test_gui_klassen_ueberdecken_keine_tkinter_namen():
    tree = ast.parse(GUI_PY.read_text(encoding="utf-8"))
    checked = 0
    for cls in (n for n in tree.body if isinstance(n, ast.ClassDef)):
        for base in cls.bases:
            tk_base = BASES.get(ast.unparse(base))
            if tk_base is None:
                continue
            checked += 1
            clashes = sorted(name for name in _own_names(cls) if hasattr(tk_base, name))
            assert not clashes, f"{cls.name} ueberdeckt tkinter-Namen: {clashes}"
    assert checked >= 1
