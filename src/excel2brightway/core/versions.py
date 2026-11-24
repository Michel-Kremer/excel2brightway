"""
versions.py
===========
Erkennt mehrere Fassungen derselben Excel-Datei im Ordner und legt fest,
welche zaehlt. Ohne diese Regel wuerden alte Fassungen als "weitere Datei
derselben Datenbank" mitgeschrieben.

Zwei Dateien derselben Datenbank gehoeren zusammen ("Familie"), wenn
  * ihr Name ohne Versionsendung gleich ist
    (LCI_X_v0.1.xlsx, LCI_X_v0.3.xlsx, LCI_X.xlsx), oder
  * eine eine Versionsendung hat und der Name der anderen mit demselben Stamm
    beginnt (LCI_X_v0.3.xlsx und LCI_X_final.xlsx / "LCI_X_v0.3 - Kopie.xlsx").
Dateien ohne Versionsnummer, die zu einer Familie gehoeren, sind "Varianten".

Welche Datei zaehlt:
  * genau eine Datei der Familie ausgewaehlt  -> diese
  * sonst (Ordner, Mehrfachauswahl, nur Kontext) -> die neueste Version,
    sofern eindeutig (alle Dateien mit Versionsnummer, hoechste nur einmal)

Meldungen (Verweis-Test, an der gezaehlten bzw. ersten ausgewaehlten Datei):
  HINWEIS 'andere_version'        andere Version wird ignoriert
  WARNUNG 'variante_ohne_version' Variante ohne Versionsnummer wird ignoriert
  FEHLER  'version_unklar'        nicht eindeutig, welche Datei zaehlt (nur bei
                                  Ordner-/Mehrfachauswahl; bei Einzelauswahl
                                  WARNUNG) - alle Dateien bleiben drin
Ignorierte Dateien fliessen in keine Pruefung und nicht ins Schreiben ein.
Beim FEHLER 'version_unklar' blendet checks.py alle weiteren Meldungen zu den
Dateien dieser Datenbank aus (Folgefehler wie doppelte codes).
"""

import re
from pathlib import Path

from .model import FEHLER, HINWEIS, TEST_VERWEIS, WARNUNG, Issue

_VERSION_RE = re.compile(r"^(?P<base>.+?)[ _\-]+v(?P<version>\d+(?:[._]\d+)*)$", re.IGNORECASE)
_SEPARATORS = " _-."


def split_version(stem):
    """'LCI_X_v0.3' -> ('lci_x', (0, 3)); ohne Versionsendung -> ('lci_x_final', None)."""
    match = _VERSION_RE.match(stem.strip())
    if not match:
        return stem.strip().lower(), None
    numbers = tuple(int(n) for n in re.split(r"[._]", match.group("version")))
    while len(numbers) > 1 and numbers[-1] == 0:  # v1.0 == v1
        numbers = numbers[:-1]
    return match.group("base").rstrip(_SEPARATORS).lower(), numbers


def _related(a, b):
    (base_a, ver_a), (base_b, ver_b) = split_version(a.stem), split_version(b.stem)
    if base_a == base_b:
        return True
    for (base, ver), other in (((base_a, ver_a), b), ((base_b, ver_b), a)):
        stem = other.stem.lower()
        if ver is not None and stem.startswith(base) and stem[len(base):len(base) + 1] in tuple(_SEPARATORS):
            return True
    return False


def families(models):
    """Gruppen (Listen von Pfaden, >= 2) zusammengehoeriger Dateien je Datenbank."""
    groups = []
    by_database = {}
    for model in models:
        if model.readable and model.database:
            by_database.setdefault(model.database, []).append(Path(model.path))
    for database in sorted(by_database):
        paths = sorted(by_database[database])
        parent = {p: p for p in paths}

        def find(p):
            while parent[p] != p:
                p = parent[p]
            return p

        for i, a in enumerate(paths):
            for b in paths[i + 1:]:
                if _related(a, b):
                    parent[find(b)] = find(a)
        clusters = {}
        for p in paths:
            clusters.setdefault(find(p), []).append(p)
        groups.extend((database, members) for members in clusters.values() if len(members) > 1)
    return groups


def _newest(members):
    """Pfad der eindeutig neuesten Version oder None."""
    versions = {p: split_version(p.stem)[1] for p in members}
    if any(v is None for v in versions.values()):
        return None
    top = max(versions.values())
    winners = [p for p, v in versions.items() if v == top]
    return winners[0] if len(winners) == 1 else None


def resolve_versions(models, selected, single_selection):
    """
    -> (issues, ignored) ; `ignored`: {Pfad: gezaehlter Pfad}.
    `models`: alle gelesenen Dateien, `selected`: ausgewaehlte Pfade,
    `single_selection`: genau eine Datei (kein Ordner) wurde gewaehlt.
    """
    by_path = {Path(m.path): m for m in models}
    selected = [Path(p) for p in selected]
    # Meldungen nur fuer Datenbanken der Auswahl; sonst (reiner Kontext) still die neueste nehmen
    relevant = {by_path[p].database for p in selected if p in by_path}
    issues, ignored = [], {}
    for database, members in families(models):
        chosen_selected = [p for p in members if p in selected]
        chosen = chosen_selected[0] if len(chosen_selected) == 1 else _newest(members)
        if database not in relevant:
            ignored.update({p: chosen for p in members if chosen is not None and p != chosen})
            continue
        anchor = chosen or (chosen_selected[0] if chosen_selected else members[0])
        sheet, row = by_path[anchor].database_location or (None, None)
        base = dict(test=TEST_VERWEIS, file=anchor, sheet=sheet, row=row, column=2)
        names = ", ".join(p.name for p in members)

        if chosen is None:
            blocking = not single_selection
            issues.append(Issue(
                severity=FEHLER if blocking else WARNUNG, kind="version_unklar",
                message=f"Mehrere Fassungen der Datenbank '{database}' ({names}) und es ist nicht eindeutig, "
                        f"welche zaehlt (fehlende oder gleiche Versionsnummer). "
                        + ("Bitte nur eine davon auswaehlen oder die anderen aus dem Ordner verschieben. "
                           "Weitere Meldungen zu diesen Dateien kommen erst danach."
                           if blocking else "Alle werden zusammen geprueft und geschrieben."),
                **base,
            ))
            continue

        newest = _newest(members)
        for path in members:
            if path == chosen:
                continue
            ignored[path] = chosen
            if split_version(path.stem)[1] is None:
                issues.append(Issue(
                    severity=WARNUNG, kind="variante_ohne_version",
                    message=f"'{path.name}' sieht nach einer Fassung derselben Datei aus, hat aber keine "
                            f"Versionsnummer; sie wird ignoriert, es zaehlt '{chosen.name}'.",
                    **base,
                ))
            else:
                newer = " (das ist eine neuere Version!)" if path == newest else ""
                issues.append(Issue(
                    severity=HINWEIS, kind="andere_version",
                    message=f"Andere Version '{path.name}'{newer} wird ignoriert; es zaehlt '{chosen.name}'.",
                    **base,
                ))
    return issues, ignored
