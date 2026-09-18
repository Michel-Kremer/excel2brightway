"""
Tests fuer excel2brightway.ecoinvent_matcher._resolve_database: ordnet ein
Excel-'database'-Label einer echten Datenbank im (zuvor gewaehlten)
Brightway-Projekt zu - normalisiert verglichen, bei Mehrdeutigkeit per
Callback nachgefragt. Nutzt Monkeypatch von _ensure_project, damit KEIN
echter bw2data-Projektzugriff stattfindet (bd.projects.set_current() legt
bei einem unbekannten Namen sonst ein echtes, leeres Projekt an).
"""

from excel2brightway import ecoinvent_matcher


class _FakeBD:
    def __init__(self, databases):
        self.databases = list(databases)


def _patch_project(monkeypatch, databases):
    monkeypatch.setattr(ecoinvent_matcher, "_ensure_project", lambda project_name: _FakeBD(databases))
    ecoinvent_matcher._database_resolution_cache.clear()


def test_unique_label_match_is_used_without_prompt(monkeypatch):
    _patch_project(monkeypatch, ["ecoinvent-3.12-cutoff", "biosphere3"])
    calls = []

    result = ecoinvent_matcher._resolve_database(
        "proj", "ecoinvent 3.12", "technosphere", prompt=lambda *a: calls.append(a),
    )

    assert result == "ecoinvent-3.12-cutoff"
    assert calls == []


def test_biosphere_resolves_to_only_biosphere_like_db_without_label(monkeypatch):
    _patch_project(monkeypatch, ["ecoinvent-3.12-cutoff", "biosphere3"])

    result = ecoinvent_matcher._resolve_database("proj", None, "biosphere")

    assert result == "biosphere3"


def test_biosphere_resolves_versioned_name_not_just_literal_biosphere3(monkeypatch):
    # Manche ecoinvent-Importe heissen nicht 'biosphere3', sondern z.B.
    # 'ecoinvent-3.10-biosphere' - das darf kein hardcodierter Sonderfall sein.
    _patch_project(monkeypatch, ["ecoinvent-3.10-cutoff", "ecoinvent-3.10-biosphere"])

    result = ecoinvent_matcher._resolve_database("proj", None, "biosphere")

    assert result == "ecoinvent-3.10-biosphere"


def test_technosphere_like_label_does_not_hijack_biosphere_resolution(monkeypatch):
    # Das Excel-'database'-Label beschreibt die technosphere-Herkunft
    # ('ecoinvent 3.12') und matcht damit normalisiert auch die technosphere-
    # Datenbank - fuer 'biosphere' darf trotzdem nur die biosphere-aehnliche
    # Datenbank in Frage kommen.
    _patch_project(monkeypatch, ["ecoinvent-3.12-cutoff", "ecoinvent-3.12-biosphere"])

    result = ecoinvent_matcher._resolve_database("proj", "ecoinvent 3.12", "biosphere")

    assert result == "ecoinvent-3.12-biosphere"


def test_ambiguous_biosphere_candidates_are_narrowed_in_prompt(monkeypatch):
    _patch_project(monkeypatch, ["ecoinvent-3.9-biosphere", "ecoinvent-3.12-biosphere", "ecoinvent-3.12-cutoff"])
    calls = []

    def prompt(label, candidates):
        calls.append(tuple(sorted(candidates)))
        return "ecoinvent-3.12-biosphere"

    result = ecoinvent_matcher._resolve_database("proj", None, "biosphere", prompt=prompt)

    assert result == "ecoinvent-3.12-biosphere"
    assert calls == [("ecoinvent-3.12-biosphere", "ecoinvent-3.9-biosphere")]


def test_ambiguous_label_asks_prompt_once_then_uses_cache(monkeypatch):
    _patch_project(monkeypatch, ["ecoinvent-3.12-cutoff", "ecoinvent-3.12-consequential"])
    calls = []

    def prompt(label, candidates):
        calls.append((label, tuple(sorted(candidates))))
        return "ecoinvent-3.12-cutoff"

    first = ecoinvent_matcher._resolve_database("proj", "ecoinvent 3.12", "technosphere", prompt=prompt)
    second = ecoinvent_matcher._resolve_database("proj", "ecoinvent 3.12", "technosphere", prompt=prompt)

    assert first == "ecoinvent-3.12-cutoff"
    assert second == "ecoinvent-3.12-cutoff"
    assert len(calls) == 1
    assert calls[0] == ("ecoinvent 3.12", ("ecoinvent-3.12-consequential", "ecoinvent-3.12-cutoff"))


def test_no_match_and_no_prompt_returns_none(monkeypatch):
    _patch_project(monkeypatch, ["some-other-db"])

    result = ecoinvent_matcher._resolve_database("proj", "ecoinvent 3.12", "technosphere")

    assert result is None
