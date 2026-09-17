# excel2brightway

Prueft LCI-/ecoinvent-Excel-Daten im nativen Brightway2-Excel-Format
(`bw2io.ExcelImporter`) auf Korrektheit - Flow-Aufloesung (intern /
`flow_registry.yaml` / ecoinvent, read-only) und Excel-Format - und schreibt
sie bei vollstaendiger Aufloesung in eine Brightway database, inklusive der definierten Parametern und
Formeln.

## Installation

Dieses Projekt nutzt [uv](https://docs.astral.sh/uv/) fuer die
Paketverwaltung. `pyproject.toml` legt die kompatiblen Versionsbereiche fest,
`uv.lock` die daraus konkret aufgeloesten, reproduzierbaren Versionen
(inklusive `bw2data`).

Eigenstaendige, von uv verwaltete Umgebung (Standardfall):

```
cd <pfad-zum-projekt>
uv sync
uv run ex2bw-check
```

`uv sync` legt bei Bedarf eine `.venv` im Projektverzeichnis an und
installiert exakt die in `uv.lock` festgehaltenen Versionen.

Installation in eine bereits bestehende Ziel-Umgebung (z. B. eine, die schon
die zum Ziel-Brightway-Projekt passende `bw2data`-Version enthaelt):

```
uv pip install --python <pfad-zur-ziel-umgebung>/python.exe -e .
```

`bw2data` MUSS in derselben Major-Version installiert sein, die auch das
Ziel-Brightway-Projekt zuletzt geschrieben hat (aktuell 3.6.x, siehe
`pyproject.toml`) - eine neuere Major-Version loest beim ersten
Projektzugriff eine automatische, irreversible Migration am geteilten
Brightway-Projektordner aus.

Die editierbare Installation verankert den Paketpfad als absoluten Pfad in
der Umgebung. Wird `<pfad-zum-projekt>` spaeter verschoben oder umbenannt,
bricht die Installation und muss von `<pfad-zum-projekt>` aus erneut
durchgefuehrt werden.

## Workspace

`excel2brightway` operiert auf einem **Workspace**-Ordner, der getrennt vom
Package-Code liegt:

```
<workspace>/
  *.xlsx                     Eingabe: LCI-Cluster-Excel-Dateien (von Hand gepflegt)
  output_excel2brightway/    Ausgabe: alles, was excel2brightway selbst schreibt
    flow_registry.yaml
    unresolved.yaml
    resolved/
```

Workspace-Aufloesung, in dieser Reihenfolge: `--workspace PATH` > Umgebungs-
variable `EX2BW_WORKSPACE` > aktuelles Arbeitsverzeichnis.

## Nutzung

```
conda activate <env-name>
cd <workspace>
ex2bw-check              # Stufe 1: Testen & Abgleichen -> output_excel2brightway/resolved/*.yaml
ex2bw-load                # Stufe 2: resolved/*.yaml -> Brightway (fragt vor dem Schreiben nach)
ex2bw-run                 # interaktives Menue fuer beide Stufen
ex2bw-tidy-registry        # flow_registry.yaml aufraeumen (Dedupe + unit-Backfill)
```

Jeder Befehl akzeptiert `--workspace PATH`, um nicht aus dem Workspace heraus
aufgerufen zu werden.

`ex2bw-load`/`ex2bw-run` akzeptieren zusaetzlich `--project NAME` fuer das
Ziel-Brightway-Projekt. Ohne `--project` wird vor dem Schreiben interaktiv
aus den vorhandenen bw2data-Projekten gewaehlt (Default: das aktuell
aktivierte Projekt). Am Ende des Schreibvorgangs wird das tatsaechlich
verwendete Projekt in der Konsole ausgegeben.

## Entwicklung

```
uv sync --extra test
uv run pytest
```

`tests/fixtures/excel_formats/` enthaelt bw2io-Beispiel-Excel-Dateien fuer
die Tests (Skip/Cutoff/Parameter/Formel-Handling). Kein Test schreibt nach
Brightway - das beruehrt ein echtes, ggf. geteiltes Brightway-Projekt und
bleibt manuelle Verifikation.
