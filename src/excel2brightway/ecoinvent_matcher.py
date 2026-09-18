"""
ecoinvent_matcher.py
=====================
Loest Flows, die weder intern (gegen andere geladene Aktivitaeten) noch
in flow_registry.yaml gefunden werden, automatisch gegen eine lokal
importierte ecoinvent-Datenbank (Technosphere) bzw. die dazugehoerige
Biosphere-Datenbank auf - per bw2data.

Welches Brightway-Projekt und welche Datenbanken darin verwendet werden,
ist NICHT hardcodiert (auch der Biosphere-Datenbankname nicht - der heisst
je nach ecoinvent-Version/Import z.B. 'biosphere3' ODER etwas wie
'ecoinvent-3.10-biosphere'):

  * Das Projekt kommt von aussen (siehe cli.py: der Nutzer waehlt es einmal
    am Anfang von 'ex2bw-check', bevor der Abgleich laeuft).
  * Welche der Datenbanken im Projekt die Technosphere-Quelle ist, wird pro
    Lauf aus der 'database'-Spalte der Exchange-Zeile abgeleitet (z.B.
    'ecoinvent 3.12' -> Datenbank 'ecoinvent-3.12-cutoff'): normalisiert
    verglichen (Gross-/Kleinschreibung, Leer-/Sonderzeichen egal), bei
    eindeutigem Treffer ohne Rueckfrage uebernommen.
  * Fuer die Biosphere-Quelle wird zuerst der Kandidatenkreis auf
    Datenbanken eingeschraenkt, deren Name 'biosphere' enthaelt (deckt
    sowohl 'biosphere3' als auch versionierte Namen wie
    'ecoinvent-3.10-biosphere' ab) - das Excel-Label beschreibt i.d.R. nur
    die *technosphere* Herkunft und waere sonst ein unzuverlaessiges Signal
    fuer die Biosphere-Datenbank. Gibt es davon genau eine, wird sie ohne
    Rueckfrage verwendet.
  * Ist die Zuordnung in keinem der beiden Faelle eindeutig (kein oder
    mehrere Kandidaten), wird - sofern der Aufrufer einen `prompt`-Callback
    mitgibt (siehe cli.py) - einmal PRO LABEL nachgefragt; ohne Callback
    (z.B. in Tests/Bibliotheksnutzung) bleibt der Flow dann unresolved statt
    geraten zu werden.

Baut beim ersten Zugriff je (Projekt, Datenbank) einen Namensindex
(normalisiertes 'reference product' -> Aktivitaeten bzw. normalisierter
Flow-'name' -> Flows) und haelt ihn fuer die Laufzeit im Speicher, damit
nicht bei jedem einzelnen Flow neu durch die gesamte Datenbank gesucht
wird. Datenbank-Zuordnungen (Label -> Datenbankname, inkl. Rueckfragen-
Antworten) werden ebenfalls pro (Projekt, Art, Label) gecacht.

Standortwahl bei mehreren Kandidaten (z.B. gleicher Flow in DE/FR/GLO):
  1. der von der Cluster-Datei mitgegebene 'location'-Hinweis (Excel-
     Spalte 'location' der Exchange-Zeile), falls vorhanden
     UND unter den Kandidaten
  2. sonst 'GLO'
  3. sonst 'RER'
  4. sonst der (alphabetisch nach Ort) erste Kandidat

WICHTIG: bw2data MUSS exakt in der Version installiert sein, die auch
das Zielprojekt zuletzt geschrieben hat (siehe uv.lock/pyproject.toml).
Eine neuere Major-Version loest beim ersten Projektzugriff eine
automatische, irreversible Migration aus.
"""

import re

from .matcher import normalize

BIOSPHERE_NAME_HINT = "biosphere"
LOCATION_PRIORITY = ("GLO", "RER")

_COMPARTMENT_SUFFIX_RE = re.compile(r"^(.*?)\s*\(([^()]+)\)\s*$")
_DB_NAME_NORMALIZE_RE = re.compile(r"[^a-z0-9]")

_technosphere_indexes_cache = {}  # (project, db_name) -> (by_product, by_name)
_biosphere_indexes_cache = {}  # (project, db_name) -> index
_database_resolution_cache = {}  # (project, kind, label) -> db_name oder None


def _ensure_project(project_name: str):
    import bw2data as bd

    if bd.projects.current != project_name:
        bd.projects.set_current(project_name)
    return bd


def _normalize_db_name(value: str) -> str:
    return _DB_NAME_NORMALIZE_RE.sub("", str(value).lower())


def _resolve_database(project_name: str, label, kind: str, prompt=None):
    """
    Ordnet ein Excel-'database'-Label (z.B. 'ecoinvent 3.12', kann auch
    None sein, wenn die Exchange-Zeile keine 'database' gesetzt hat) einer
    tatsaechlichen Datenbank im Projekt zu. `kind` ist 'technosphere' oder
    'biosphere'. Ergebnis (auch ein "kein Treffer") wird pro (project,
    kind, label) gecacht, damit bei vielen Flows mit demselben Label nur
    einmal verglichen bzw. gefragt wird.

    Fuer 'biosphere' wird der Kandidatenkreis zuerst auf Datenbanken mit
    'biosphere' im Namen eingeschraenkt (siehe Moduldoc) - das Label dient
    dort nur noch zur Verfeinerung INNERHALB dieses Kreises, nicht zum
    direkten Vergleich gegen alle Datenbanken (sonst koennte ein
    technosphere-artiges Label wie 'ecoinvent 3.12' versehentlich eine
    Nicht-Biosphere-Datenbank treffen).
    """
    cache_key = (project_name, kind, label)
    if cache_key in _database_resolution_cache:
        return _database_resolution_cache[cache_key]

    bd = _ensure_project(project_name)
    databases = sorted(bd.databases)

    if kind == "biosphere":
        biosphere_like = [name for name in databases if BIOSPHERE_NAME_HINT in _normalize_db_name(name)]
        pool = biosphere_like or databases
    else:
        pool = databases

    resolved = None
    candidates = pool
    if label:
        norm_label = _normalize_db_name(label)
        matches = [name for name in pool if norm_label in _normalize_db_name(name)]
        if len(matches) == 1:
            resolved = matches[0]
        elif matches:
            candidates = matches
    elif len(pool) == 1:
        # Kein Label gesetzt (z.B. Excel-Zeile ohne 'database'-Angabe): bei
        # nur einem plausiblen Kandidaten (z.B. der einzigen biosphere-
        # aehnlichen Datenbank im Projekt) ist das eindeutig genug, um ohne
        # Rueckfrage zu uebernehmen. Bei einem gesetzten, aber nicht
        # treffenden Label bleibt es dagegen bewusst mehrdeutig (siehe
        # Rueckfrage/None unten) statt zu raten.
        resolved = pool[0]

    if resolved is None and prompt is not None:
        resolved = prompt(label, candidates)

    _database_resolution_cache[cache_key] = resolved
    return resolved


def _build_technosphere_indexes(project_name: str, db_name: str):
    """
    Baut zwei Indizes ueber die Aktivitaeten einer Technosphere-Datenbank:
    einen ueber das 'reference product' (der uebliche Schluessel bei
    ausfuehrlichen bw2io-Exporten) und einen ueber den Aktivitaets-'name'
    (Fallback fuer Exporte, die nur den Prozessnamen mitliefern, z.B.
    'cast iron production' statt 'cast iron').
    """
    bd = _ensure_project(project_name)
    by_product, by_name = {}, {}
    for act in bd.Database(db_name):
        by_product.setdefault(normalize(act.get("reference product") or act.get("name")), []).append(act)
        by_name.setdefault(normalize(act.get("name")), []).append(act)
    return by_product, by_name


def _build_biosphere_index(project_name: str, db_name: str) -> dict:
    bd = _ensure_project(project_name)
    index = {}
    for flow in bd.Database(db_name):
        key = normalize(flow.get("name"))
        index.setdefault(key, []).append(flow)
    return index


def _get_technosphere_indexes(project_name: str, db_name: str):
    cache_key = (project_name, db_name)
    if cache_key not in _technosphere_indexes_cache:
        _technosphere_indexes_cache[cache_key] = _build_technosphere_indexes(project_name, db_name)
    return _technosphere_indexes_cache[cache_key]


def _get_biosphere_index(project_name: str, db_name: str) -> dict:
    cache_key = (project_name, db_name)
    if cache_key not in _biosphere_indexes_cache:
        _biosphere_indexes_cache[cache_key] = _build_biosphere_index(project_name, db_name)
    return _biosphere_indexes_cache[cache_key]


def _pick_by_location(candidates: list, location_hint: str):
    by_location = {c.get("location"): c for c in candidates}
    if location_hint and location_hint in by_location:
        return by_location[location_hint]
    for loc in LOCATION_PRIORITY:
        if loc in by_location:
            return by_location[loc]
    return sorted(candidates, key=lambda c: c.get("location") or "")[0]


def _split_biosphere_name(flow_name: str):
    """Trennt einen optionalen Kompartiment-Suffix ab, z.B. 'Methan (air)' -> ('Methan', 'air')."""
    match = _COMPARTMENT_SUFFIX_RE.match(flow_name.strip())
    if match:
        return match.group(1).strip(), match.group(2).strip().lower()
    return flow_name.strip(), None


def match_technosphere(project_name: str, flow_name: str, location_hint: str = None, database_hint: str = None, prompt=None):
    """
    Sucht einen Technosphere-Flow in `project_name`: zuerst ueber
    'reference product' (Standardfall bei ausfuehrlichen bw2io-Exporten),
    dann ueber den Aktivitaets-'name' (Fallback fuer Exporte, die nur den
    Prozessnamen liefern, z.B. 'cast iron production' statt 'cast iron').
    `database_hint` ist die 'database'-Spalte der Exchange-Zeile (siehe
    _resolve_database) und bestimmt, welche Datenbank im Projekt durchsucht
    wird. Gibt einen flow_registry-kompatiblen Eintrag ({'database':...,
    'code':..., 'location':..., 'unit':...}) oder None zurueck. 'code'
    erlaubt matcher.py, den Flow direkt (ohne erneuten Namens-Abgleich) auf
    eine konkrete Aktivitaet zu verweisen; 'unit' erlaubt den Abgleich mit
    der in der Cluster-Datei angegebenen Einheit (siehe matcher.normalize_unit).
    """
    db_name = _resolve_database(project_name, database_hint, "technosphere", prompt=prompt)
    if not db_name:
        return None

    by_product, by_name = _get_technosphere_indexes(project_name, db_name)
    key = normalize(flow_name)
    candidates = by_product.get(key) or by_name.get(key)
    if not candidates:
        return None

    act = _pick_by_location(candidates, location_hint)
    return {
        "database": db_name,
        "code": act.get("code"),
        "location": act.get("location"),
        "unit": act.get("unit"),
    }


def match_biosphere(project_name: str, flow_name: str, database_hint: str = None, prompt=None):
    """
    Sucht einen Biosphere-Flow (per Name, optional mit Kompartiment-Suffix
    wie '(air)') in `project_name`. `database_hint` siehe match_technosphere
    (hier meist leer, da die einzige Datenbank mit 'biosphere' im Namen
    automatisch verwendet wird, siehe _resolve_database).
    Gibt einen flow_registry-kompatiblen Eintrag ({'database':...,
    'code':..., 'categories':..., 'unit':...}) oder None zurueck.
    """
    db_name = _resolve_database(project_name, database_hint, "biosphere", prompt=prompt)
    if not db_name:
        return None

    base_name, compartment_hint = _split_biosphere_name(flow_name)
    candidates = _get_biosphere_index(project_name, db_name).get(normalize(base_name))
    if not candidates:
        return None

    if compartment_hint:
        filtered = [
            c for c in candidates
            if (c.get("categories") or (None,))[0]
            and c["categories"][0].lower() == compartment_hint
        ]
        if filtered:
            candidates = filtered

    flow = sorted(candidates, key=lambda c: len(c.get("categories") or ()))[0]
    return {
        "database": db_name,
        "code": flow.get("code"),
        "categories": list(flow.get("categories") or []),
        "unit": flow.get("unit"),
    }
