"""
config.py
=========
Loest den Workspace auf, in dem excel2brightway arbeitet. Ein Workspace ist
ein beliebiger Ordner:

    <workspace>/
      *.xlsx                        Eingabe: LCI-Cluster-Excel-Dateien (von Hand
                                     gepflegt), nicht rekursiv gelesen.
      Irgendein_Unterordner/        wird ignoriert (z.B. Entwuerfe) - nur die
                                     .xlsx-Dateien direkt im Root zaehlen.
      output_excel2brightway/       Ausgabe: alles, was excel2brightway selbst
        flow_registry.yaml          schreibt/pflegt - getrennt von der
        unresolved.yaml             handgepflegten Eingabe.
        load_warnings.yaml          Warnungen beim Einlesen (fehlende
                                     Pflichtfelder, unbekannter Exchange-
                                     Typ, fehlende amount/name, ...).
        resolved/

Der Workspace ist NICHT Teil des Package-Codes (`src/excel2brightway/`) - das
Package selbst enthaelt keine Nutzdaten und ist so auch fuer andere Projekte
nutzbar. Aufloesung, in dieser Reihenfolge:

  1. explizit uebergeben (--workspace PATH in der CLI, siehe cli.py),
  2. Umgebungsvariable EX2BW_WORKSPACE,
  3. aktuelles Arbeitsverzeichnis (Path.cwd()) - z.B. `cd <workspace> && ex2bw-check`
     funktioniert dann ohne weitere Angabe, wie bei git/terraform.
"""

import os
from dataclasses import dataclass
from pathlib import Path

OUTPUT_DIRNAME = "output_excel2brightway"
WORKSPACE_ENV_VAR = "EX2BW_WORKSPACE"

REGISTRY_TEMPLATE = """\
# ============================================================
# flow_registry.yaml - einmal pflegen, in allen Cluster-Dateien wiederverwenden
# Nur Flows eintragen, die NICHT ueber den internen Namensabgleich
# (name-Feld einer anderen Aktivitaet) gefunden werden.
# ============================================================

technosphere: {}
biosphere: {}
"""


@dataclass
class Workspace:
    """Aufgeloester Workspace: `root` fuer die Eingabe-Excel, `output_dir`
    (+ Unterpfade) fuer alles, was excel2brightway selbst erzeugt."""

    root: Path

    @property
    def output_dir(self) -> Path:
        return self.root / OUTPUT_DIRNAME

    @property
    def registry(self) -> Path:
        return self.output_dir / "flow_registry.yaml"

    @property
    def unresolved(self) -> Path:
        return self.output_dir / "unresolved.yaml"

    @property
    def load_warnings(self) -> Path:
        return self.output_dir / "load_warnings.yaml"

    @property
    def resolved_dir(self) -> Path:
        return self.output_dir / "resolved"

    def ensure(self) -> None:
        """
        Bereitet einen neuen/unvollstaendigen Workspace vor: legt
        `output_dir`/`resolved_dir` an und schreibt eine leere
        `flow_registry.yaml`, falls noch keine existiert. Ruehrt die
        Eingabe (`root`, `*.xlsx`) nicht an.
        """
        self.resolved_dir.mkdir(parents=True, exist_ok=True)
        if not self.registry.exists():
            self.registry.write_text(REGISTRY_TEMPLATE, encoding="utf-8")


def resolve_workspace(explicit: Path = None) -> Workspace:
    """
    Loest den Workspace auf (siehe Modul-Docstring fuer die Prioritaet).
    `explicit` ist typischerweise der --workspace-Wert aus der CLI (oder
    None).
    """
    if explicit is not None:
        root = Path(explicit)
    elif os.environ.get(WORKSPACE_ENV_VAR):
        root = Path(os.environ[WORKSPACE_ENV_VAR])
    else:
        root = Path.cwd()

    return Workspace(root=root.resolve())
