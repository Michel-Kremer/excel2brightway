"""Tests fuer den Fluss-Test (core/resolve.py) und die Kandidatensuche (core/search.py)."""

from conftest import BIO, BLANK, EI, Record, activity_rows, make_backend, write_workbook
from excel2brightway.core.resolve import ResolveConfig, Resolver
from excel2brightway.core.search import SearchIndex
from excel2brightway.core.workbook import read_workbook


def _resolve(tmp_path, exchanges, backend=None, config=None, extra_files=None):
    rows = activity_rows("SOEC operation", "op", exchanges, unit="kWh", database_row="SOEC")
    path = write_workbook(tmp_path / "soec.xlsx", {"Data": rows})
    models = [read_workbook(path)]
    for name, sheet_rows in (extra_files or {}).items():
        models.append(read_workbook(write_workbook(tmp_path / name, {"Data": sheet_rows})))
    resolver = Resolver(models, backend=backend, config=config)
    flows = [e for e in models[0].activities[0].exchanges if e.type in ("technosphere", "biosphere")]
    return [resolver.resolve(e) for e in flows], resolver


def test_market_preferred_over_transformation_at_same_location(tmp_path):
    (res,), _ = _resolve(tmp_path, [["market for electricity, medium voltage", 2, "kWh", EI, "DE", "technosphere",
                                     "electricity, medium voltage"]], backend=make_backend())
    assert res.ok and res.record.code == "e-de-market"


def test_same_location_without_market_is_ambiguous_not_guessed(tmp_path):
    # Frueher gewann still der letzte Kandidat je Standort.
    (res,), _ = _resolve(tmp_path, [["steel, low-alloyed", 1, "kg", EI, "RER", "technosphere"]], backend=make_backend())
    assert res.status == "mehrdeutig"
    assert {c.record.code for c in res.candidates} >= {"s-1", "s-2"}


def test_location_priority_includes_row(tmp_path):
    # Nur Laenderdatensaetze + RoW: frueher wurde alphabetisch 'AT' genommen.
    (res,), _ = _resolve(tmp_path, [["tap water", 1, "kg", EI, None, "technosphere"]], backend=make_backend())
    assert res.ok and res.record.location == "RoW"


def test_without_matching_priority_location_it_is_ambiguous(tmp_path):
    (res,), _ = _resolve(
        tmp_path, [["tap water", 1, "kg", EI, None, "technosphere"]], backend=make_backend(),
        config=ResolveConfig(location_priority=("GLO",)),
    )
    assert res.status == "mehrdeutig"


def test_location_not_found_offers_other_locations(tmp_path):
    (res,), _ = _resolve(tmp_path, [["tap water", 1, "kg", EI, "DE", "technosphere"]], backend=make_backend())
    assert res.status == "standort_nicht_gefunden"
    assert "AT" in res.message and {c.record.location for c in res.candidates} >= {"AT", "CH", "RoW"}


def test_unit_mismatch(tmp_path):
    (res,), _ = _resolve(tmp_path, [["nickel, class 1", 1, "kWh", EI, None, "technosphere"]], backend=make_backend())
    assert res.status == "einheit_abweichend"
    assert res.candidates[0].record.code == "n-glo"


def test_unit_aliases_are_compatible(tmp_path):
    (res,), _ = _resolve(tmp_path, [["nickel, class 1", 1, "kg", EI, None, "technosphere"]], backend=make_backend())
    assert res.ok


def test_not_found_gets_search_candidates_from_ecoinvent(tmp_path):
    (res,), _ = _resolve(tmp_path, [["electricty medium voltage", 2, "kWh", EI, "DE", "technosphere"]], backend=make_backend())
    assert res.status == "nicht_gefunden"
    top = res.candidates[0].record
    assert top.code == "e-de-market"  # Tippfehler, Wortstellung, Einheit, Standort, Markt


def test_database_label_is_normalized(tmp_path):
    (res,), _ = _resolve(tmp_path, [["nickel, class 1", 1, "kg", "ecoinvent 3.10", None, "technosphere"]], backend=make_backend())
    assert res.ok and res.record.database == EI


def test_technosphere_label_does_not_pick_biosphere_db(tmp_path):
    resolver = _resolve(tmp_path, [], backend=make_backend())[1]
    stages, problem = resolver.target_databases("ecoinvent 3.10", "technosphere")
    assert problem is None and stages == [[EI]]
    stages, problem = resolver.target_databases("ecoinvent 3.10", "biosphere")
    assert stages == [[BIO]]


def test_biosphere_label_biosphere3_falls_back_to_project_biosphere(tmp_path):
    (res,), _ = _resolve(tmp_path, [["Carbon dioxide, fossil", 0.1, "kg", "biosphere3", None, "biosphere", None, "air"]],
                         backend=make_backend())
    assert res.ok and res.record.code == "co2-air"


def test_biosphere_categories_exact_and_suffix(tmp_path):
    results, _ = _resolve(tmp_path, [
        ["Carbon dioxide, fossil", 0.1, "kg", None, None, "biosphere", None, "air::urban air close to ground"],
        ["Carbon dioxide, fossil (air)", 0.1, "kg", None, None, "biosphere"],
        ["Water", 1, "m3", None, None, "biosphere", None, "water"],
    ], backend=make_backend())
    assert results[0].record.code == "co2-urban"
    assert results[1].record.code == "co2-air"  # kuerzeste Kategorie innerhalb 'air'
    assert results[2].status == "mehrdeutig"  # zwei Unterkategorien, keine unspezifische


def test_unknown_subcategory_is_reported(tmp_path):
    (res,), _ = _resolve(tmp_path, [["Carbon dioxide, fossil", 0.1, "kg", None, None, "biosphere", None, "air::indoor"]],
                         backend=make_backend())
    assert res.status == "kategorie_nicht_gefunden"


def test_flow_defined_in_same_file_is_resolved_internally(tmp_path):
    stack = activity_rows("stack production, SOEC", "stack", [], unit="unit")
    rows = activity_rows("SOEC operation", "op", [["stack production, SOEC", 0.001, "unit", None, None, "technosphere"]],
                         unit="kWh", database_row="SOEC") + stack
    path = write_workbook(tmp_path / "soec.xlsx", {"Data": rows})
    model = read_workbook(path)
    resolver = Resolver([model], backend=None)
    res = resolver.resolve(model.activities[0].exchanges[1])
    assert res.ok and res.record.key == ("SOEC", "stack") and res.record.origin == "datei"


def test_flow_defined_in_other_file_of_folder(tmp_path):
    other = activity_rows("stack production, SOEC", "stack", [], unit="unit", database_row="SOEC")
    (res,), _ = _resolve(
        tmp_path, [["stack production, SOEC", 0.001, "unit", "SOEC", None, "technosphere"]],
        extra_files={"stack.xlsx": other},
    )
    assert res.ok and res.record.code == "stack"


def test_own_database_in_project_is_used_when_not_in_files(tmp_path):
    own = {"SOEC_Process": [Record("SOEC_Process", "bop", "balance of plant, SOEC", None, "DE", "unit")]}
    (res,), _ = _resolve(tmp_path, [["balance of plant, SOEC", 1, "unit", "SOEC_Process", "DE", "technosphere"]],
                         backend=make_backend(extra=own))
    assert res.ok and res.record.database == "SOEC_Process"


def test_file_wins_over_same_database_in_project(tmp_path):
    old = {"SOEC": [Record("SOEC", "stack", "stack production, SOEC (alt)", None, "DE", "unit")]}
    stack = activity_rows("stack production, SOEC", "stack", [], unit="unit")
    rows = activity_rows("SOEC operation", "op", [["stack production, SOEC", 0.001, "unit", "SOEC", None, "technosphere"]],
                         unit="kWh", database_row="SOEC") + stack
    model = read_workbook(write_workbook(tmp_path / "soec.xlsx", {"Data": rows}))
    res = Resolver([model], backend=make_backend(extra=old)).resolve(model.activities[0].exchanges[1])
    assert res.ok and res.record.origin == "datei"


def test_changed_code_is_not_ambiguous_with_old_activity_in_project(tmp_path):
    # in Brightway noch mit altem code, in der Datei mit neuem -> nur die Datei zaehlt
    old = {"SOEC": [Record("SOEC", "stack", "stack production, SOEC", None, "DE", "unit")]}
    stack = activity_rows("stack production, SOEC", "stack_neu", [], unit="unit", location="DE")
    rows = activity_rows("SOEC operation", "op", [["stack production, SOEC", 0.001, "unit", "SOEC", None, "technosphere"]],
                         unit="kWh", database_row="SOEC") + stack
    model = read_workbook(write_workbook(tmp_path / "soec.xlsx", {"Data": rows}))
    res = Resolver([model], backend=make_backend(extra=old)).resolve(model.activities[0].exchanges[1])
    assert res.ok and res.record.key == ("SOEC", "stack_neu")


def test_activity_only_left_in_project_is_no_match_for_file_database(tmp_path):
    # steht nur noch in Brightway und wird beim Schreiben geloescht -> kein Treffer
    old = {"SOEC": [Record("SOEC", "bop", "balance of plant, SOEC", None, "DE", "unit")]}
    rows = activity_rows("SOEC operation", "op", [["balance of plant, SOEC", 1, "unit", "SOEC", "DE", "technosphere"]],
                         unit="kWh", database_row="SOEC")
    model = read_workbook(write_workbook(tmp_path / "soec.xlsx", {"Data": rows}))
    res = Resolver([model], backend=make_backend(extra=old)).resolve(model.activities[0].exchanges[1])
    assert not res.ok and res.status == "nicht_gefunden"


def test_explicit_database_is_the_only_place_searched(tmp_path):
    own = {"SOEC_Process": [Record("SOEC_Process", "x", "something else", None, "DE", "kg")]}
    (res,), _ = _resolve(tmp_path, [["nickel, class 1", 1, "kg", "SOEC_Process", None, "technosphere"]],
                         backend=make_backend(extra=own))
    assert res.status == "nicht_gefunden"
    assert "SOEC_Process" in res.message and EI not in res.message


def test_unknown_database_label(tmp_path):
    (res,), _ = _resolve(tmp_path, [["nickel, class 1", 1, "kg", "gabi 2024", None, "technosphere"]], backend=make_backend())
    assert res.status == "datenbank_unbekannt"
    assert res.candidates and res.candidates[0].record.code == "n-glo"


def test_without_project_external_flows_are_reported(tmp_path):
    (res,), _ = _resolve(tmp_path, [["nickel, class 1", 1, "kg", EI, None, "technosphere"]], backend=None)
    assert res.status == "kein_projekt"


def test_link_databases_restrict_unlabelled_search(tmp_path):
    own = {"SOEC_Process": [Record("SOEC_Process", "n", "nickel, class 1", None, "GLO", "kg")]}
    (res,), _ = _resolve(tmp_path, [["nickel, class 1", 1, "kg", None, None, "technosphere"]],
                         backend=make_backend(extra=own), config=ResolveConfig(link_databases=["SOEC_Process"]))
    assert res.ok and res.record.database == "SOEC_Process"
    (res,), _ = _resolve(tmp_path, [["nickel, class 1", 1, "kg", None, None, "technosphere"]],
                         backend=make_backend(extra=own))
    assert res.status == "mehrdeutig"  # gleicher Fluss in zwei Datenbanken


def test_search_text_overrides_query(tmp_path):
    rows = activity_rows("SOEC operation", "op", [["Strom", 2, "kWh", EI, "DE", "technosphere"]],
                         unit="kWh", database_row="SOEC")
    model = read_workbook(write_workbook(tmp_path / "s.xlsx", {"Data": rows}))
    resolver = Resolver([model], backend=make_backend())
    exchange = model.activities[0].exchanges[1]
    assert resolver.resolve(exchange).candidates == [] or resolver.resolve(exchange).candidates[0].score < 0.5
    res = resolver.resolve(exchange, search_text="electricity medium voltage")
    assert res.candidates[0].record.code == "e-de-market"


def test_search_index_unit_penalty_and_typos():
    records = [
        Record(EI, "a", "market for heat, district", "heat, district", "RER", "megajoule"),
        Record(EI, "b", "market for heat, district", "heat, district", "RER", "kilogram"),
    ]
    found = SearchIndex(records).search("heta district", unit="MJ")
    assert [c.record.code for c in found] == ["a", "b"]
    assert found[0].score > found[1].score + 0.3
