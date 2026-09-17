"""
ecoinvent_matcher.py
=====================
Loest Flows, die weder intern (gegen andere geladene Aktivitaeten) noch
in flow_registry.yaml gefunden werden, automatisch gegen eine lokal
importierte ecoinvent-Datenbank (Technosphere) bzw. biosphere3
(Biosphere) auf - per bw2data.

Baut beim ersten Zugriff einen Namensindex (normalisiertes 'reference
product' -> Aktivitaeten bzw. normalisierter Flow-'name' -> Flows) und
haelt ihn fuer die Laufzeit im Speicher, damit nicht bei jedem
einzelnen Flow neu durch die gesamte Datenbank gesucht wird.

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

Benoetigt ein bereits importiertes Brightway-Projekt mit den unten
angegebenen Datenbanken (siehe PROJECT_NAME).
"""

import re

from .matcher import normalize

PROJECT_NAME = "ecoinvent-3.12-cutoff"
TECHNOSPHERE_DB = "ecoinvent-3.12-cutoff"
BIOSPHERE_DB = "biosphere3"

LOCATION_PRIORITY = ("GLO", "RER")

_COMPARTMENT_SUFFIX_RE = re.compile(r"^(.*?)\s*\(([^()]+)\)\s*$")

_technosphere_index = None
_technosphere_by_name_index = None
_biosphere_index = None


def _ensure_project():
    import bw2data as bd

    if bd.projects.current != PROJECT_NAME:
        bd.projects.set_current(PROJECT_NAME)
    return bd


def _build_technosphere_indexes():
    """
    Baut zwei Indizes ueber die ecoinvent-Aktivitaeten: einen ueber das
    'reference product' (der uebliche Schluessel bei ausfuehrlichen
    bw2io-Exports) und einen ueber den Aktivitaets-'name' (Fallback fuer
    Exporte, die nur den Prozessnamen mitliefern, z.B. 'cast iron
    production' statt 'cast iron').
    """
    bd = _ensure_project()
    by_product, by_name = {}, {}
    for act in bd.Database(TECHNOSPHERE_DB):
        by_product.setdefault(normalize(act.get("reference product") or act.get("name")), []).append(act)
        by_name.setdefault(normalize(act.get("name")), []).append(act)
    return by_product, by_name


def _build_biosphere_index() -> dict:
    bd = _ensure_project()
    index = {}
    for flow in bd.Database(BIOSPHERE_DB):
        key = normalize(flow.get("name"))
        index.setdefault(key, []).append(flow)
    return index


def _get_technosphere_indexes():
    global _technosphere_index, _technosphere_by_name_index
    if _technosphere_index is None:
        _technosphere_index, _technosphere_by_name_index = _build_technosphere_indexes()
    return _technosphere_index, _technosphere_by_name_index


def _get_biosphere_index() -> dict:
    global _biosphere_index
    if _biosphere_index is None:
        _biosphere_index = _build_biosphere_index()
    return _biosphere_index


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


def match_technosphere(flow_name: str, location_hint: str = None):
    """
    Sucht einen Technosphere-Flow in der ecoinvent-Datenbank: zuerst
    ueber 'reference product' (Standardfall bei ausfuehrlichen bw2io-
    Exporten), dann ueber den Aktivitaets-'name' (Fallback fuer Exporte,
    die nur den Prozessnamen liefern, z.B. 'cast iron production' statt
    'cast iron'). Gibt einen flow_registry-kompatiblen Eintrag
    ({'database':..., 'code':..., 'location':..., 'unit':...}) oder None
    zurueck. 'code' erlaubt matcher.py, den Flow direkt (ohne erneuten
    Namens-Abgleich) auf eine konkrete Aktivitaet zu verweisen; 'unit'
    erlaubt den Abgleich mit der in der Cluster-Datei angegebenen Einheit
    (siehe matcher.normalize_unit).
    """
    by_product, by_name = _get_technosphere_indexes()
    key = normalize(flow_name)
    candidates = by_product.get(key) or by_name.get(key)
    if not candidates:
        return None

    act = _pick_by_location(candidates, location_hint)
    return {
        "database": TECHNOSPHERE_DB,
        "code": act.get("code"),
        "location": act.get("location"),
        "unit": act.get("unit"),
    }


def match_biosphere(flow_name: str):
    """
    Sucht einen Biosphere-Flow (per Name, optional mit Kompartiment-
    Suffix wie '(air)') in biosphere3. Gibt einen flow_registry-
    kompatiblen Eintrag ({'database':..., 'code':..., 'categories':...,
    'unit':...}) oder None zurueck.
    """
    base_name, compartment_hint = _split_biosphere_name(flow_name)
    candidates = _get_biosphere_index().get(normalize(base_name))
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
        "database": BIOSPHERE_DB,
        "code": flow.get("code"),
        "categories": list(flow.get("categories") or []),
        "unit": flow.get("unit"),
    }
