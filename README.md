# excel2brightway

Prueft von Hand erstellte LCI-Excel-Dateien im nativen Brightway-Excel-Format
(wie `bw2io.ExcelImporter`), hilft beim Korrigieren direkt in der Excel und
schreibt sie danach in eine Brightway-Datenbank - inklusive Parametern und
Formeln.

**Die Excel-Dateien sind der einzige Stand.** Es gibt keine Registry und keine
Zwischendateien: Vor dem Schreiben wird alles erneut geprueft.

## Installation

Dieses Projekt nutzt [uv](https://docs.astral.sh/uv/) fuer die
Paketverwaltung. `pyproject.toml` legt die kompatiblen Versionsbereiche fest,
`uv.lock` die konkret aufgeloesten Versionen (inklusive `bw2data`).

```
cd <pfad-zum-projekt>
uv sync
uv run ex2bw-gui
```

Installation in eine bestehende Umgebung (z. B. die mit der zum
Brightway-Projekt passenden `bw2data`-Version), auch ohne Admin-Rechte:

```
uv pip install --python <pfad-zur-ziel-umgebung>/python.exe -e .
```

`bw2data` MUSS in derselben Major-Version installiert sein, die auch das
Ziel-Brightway-Projekt zuletzt geschrieben hat (aktuell 4.x) - eine andere
Major-Version loest beim ersten Projektzugriff eine automatische, irreversible
Migration des geteilten Brightway-Projektordners aus. excel2brightway prueft
das deshalb, bevor es bw2data laedt, und verweigert mit einer anderen
Major-Version jeden Zugriff auf Brightway (`core/backend.py`,
`require_bw2data`).

## Bedienung: das Fenster

```
ex2bw-gui
```

Unter Windows legt die Installation dafuer eine `ex2bw-gui.exe` ohne
Konsolenfenster an (`[project.gui-scripts]`), die sich auch als Verknuepfung
auf den Desktop legen laesst.

1. **Datei(en) oder Ordner** waehlen. Bei einem Ordner werden alle `*.xlsx`
   darin geprueft.
2. **Projekt** waehlen und optional die **Verknuepfungs-Datenbanken**
   markieren (keine Auswahl = alle Datenbanken des Projekts). Dort wird
   gesucht, wenn eine exchange-Zeile keine `database` angibt.
3. **Tests** waehlen (einzeln oder zusammen) und **Pruefen**. Die erste Datei
   mit Meldungen oeffnet sich in Excel.
4. Im Blatt **`_ex2bw_Pruefung`** je Fluss einen Kandidaten unter **Auswahl**
   waehlen - oder unter **Eigene Suche** einen Suchbegriff eintragen und erneut
   pruefen. Datei speichern und schliessen.
5. **Auswahl uebernehmen**: die Auswahl wird in die exchange-Zeilen
   geschrieben, danach wird automatisch erneut geprueft.
6. Ist alles gruen: **Nach Brightway schreiben** oder **Datenbank/Version
   umstellen**.

## Die Tests

| Test | prueft | Schweregrade |
|---|---|---|
| **Format** | Brightway-Excel-Format: Pflichtfelder und -spalten, gueltige `type`/`amount`, doppelte `code`s, eine production-Zeile, cutoff verdeckt Spalten | ueberbrueckbar: Fehler auf Zeilenebene blockieren dann nicht, solche Zeilen werden beim Schreiben ausgelassen und gemeldet. Strukturfehler (Datei unlesbar, keine/zwei `Database`-Zeilen, Aktivitaet ohne Name oder `unit`, doppelte `code`s, mehrere production-Zeilen) blockieren immer |
| **Fluss** | jede technosphere-/biosphere-Zeile laesst sich genau einem Datensatz zuordnen | nicht ueberbrueckbar |
| **Verweis** | was das Schreiben in Brightway loeschen wuerde und ob andere Datenbanken darauf zeigen | Loeschen und gebrochene Verweise nur nach ausdruecklicher Bestaetigung |

Was Pflicht und was optional ist, steht in `core/format_check.py`
(`FORMAT_SCHEMA`).

### Wo ein Fluss gesucht wird

Die Spalte `database` der exchange-Zeile hat Vorrang:

* nennt sie eine Datenbank, die in den Dateien des Ordners definiert ist ->
  nur diese Dateien. Eine gleichnamige Datenbank im Projekt wird beim
  Schreiben ersetzt; was nur noch dort steht, wird geloescht und ist kein
  Treffer (auch nicht, wenn sich der `code` einer Aktivitaet geaendert hat);
* nennt sie eine andere Datenbank (z. B. `SOEC_Process` oder `ecoinvent 3.10`,
  normalisiert verglichen) -> nur dort;
* ist sie leer -> erst alle Dateien des Ordners, dann die
  Verknuepfungs-Datenbanken; biosphere: die biosphere-Datenbank(en).

### Wann ein Treffer eindeutig ist

* Abgleich ueber `reference product` (und `name`, falls beide angegeben),
  sonst ueber `name`.
* Ist `location` angegeben, muss sie passen. Ohne `location` gilt die
  Standort-Rangfolge (Standard GLO, RER, RoW); greift sie nicht, ist der Fluss
  **mehrdeutig** - es wird nicht geraten.
* Treffer in mehreren Datenbanken sind immer mehrdeutig.
* Mehrere Datensaetze am selben Standort: ein eindeutiger Markt
  (`market for ...`) wird genommen, sonst mehrdeutig.
* biosphere: `categories` (z. B. `air::urban air close to ground`) bzw. ein
  Suffix wie `(air)`; ohne Unterkategorie die unspezifische Kategorie.
* Die Einheit muss passen (`kg` = `kilogram` usw.).

Fuer alles, was nicht eindeutig ist, werden bis zu 5 Kandidaten
vorgeschlagen: Wortaehnlichkeit gegen `reference product` und `name`
(tippfehlertolerant), Einheit, Standort, Bonus fuer Maerkte (`core/search.py`,
Gewichte in `WEIGHTS`).

## Was in die Excel-Datei geschrieben wird

| Blatt | Inhalt |
|---|---|
| `_ex2bw_Pruefung` | eine Zeile je Meldung, mit Link zur Zelle, Auswahlliste, Eigener Suche, Aktion (`ersetzen`/`Proxy`), Kommentar |
| `_ex2bw_Kandidaten` | ausgeblendet: Kandidaten hinter den Auswahllisten, Pruefsumme der Daten |
| `_ex2bw_Aenderungen` | Protokoll jeder Uebernahme (alt -> neu) |
| `_ex2bw_Migration` | Protokoll einer Umstellung (nur in umgestellten Dateien) |

Alle diese Blaetter haben `skip` in A1 und werden von Brightway nie gelesen.
Betroffene Zellen werden rot markiert und kommentiert; die Markierungen
verschwinden beim naechsten Pruefen bzw. Uebernehmen.

Beim Uebernehmen gilt:

* Die Einheit wird nie geaendert. Passt sie nicht zum Kandidaten, wird die
  Zeile nicht uebernommen (Menge bitte von Hand umrechnen).
* Fehlende Spalten (`database`, `location`, `reference product`, `comment`)
  werden hinten an die Exchanges-Tabelle angefuegt.
* Haben sich die Daten seit dem Pruefen geaendert, wird abgebrochen.
* Vorher wird eine Sicherung in `ex2bw_sicherung/` abgelegt.
* Ist die Datei in Excel geoeffnet, kommt die Meldung "Bitte Datei in Excel
  schliessen".

Alles, was nicht in die Excel gehoert (gesperrte Dateien, Ausnahmen, ...)
steht zusammen mit allen Meldungen in **`ex2bw.log`** im Ordner.

## Schreiben nach Brightway

* Eine Datenbank wird aus **allen** Dateien des Ordners gebildet, die sie
  definieren - auch aus nicht ausgewaehlten (sie werden automatisch
  mitgenommen und gemeldet).
* **Fassungen derselben Datei** (gleicher Name, Endung `_v0.1`, `_v0.3`, ...;
  Varianten wie `..._final.xlsx` oder `... - Kopie.xlsx`) zaehlen nicht als
  weitere Dateien - es zaehlt genau eine:
  * Einzelauswahl: die ausgewaehlte Datei. Andere Versionen -> HINWEIS,
    Varianten ohne Versionsnummer -> WARNUNG; beide werden ignoriert.
  * Ordner oder Mehrfachauswahl: die neueste Version. Ist das nicht
    eindeutig (Variante ohne Versionsnummer, gleiche Nummer), blockiert ein
    FEHLER das Schreiben - dann nur eine Fassung auswaehlen oder die anderen
    verschieben. Weitere Meldungen zu diesen Dateien (z. B. doppelte
    `code`s) werden bis dahin ausgeblendet.
* Die Datenbank in Brightway wird durch diesen Stand ersetzt. Aktivitaeten,
  die in Brightway stehen, aber in keiner Datei mehr, werden nur nach
  Bestaetigung geloescht; zeigen andere Datenbanken darauf, ist eine weitere
  Bestaetigung noetig.
* Der `code` einer Aktivitaet kommt aus der Excel (ohne `code`: der Name).
  Wer eine Aktivitaet ohne `code` umbenennt, aendert damit ihren `code` -
  der Verweis-Test zeigt, ob andere Datenbanken betroffen sind.

## Datenbank/Version umstellen (Grundform)

Voraussetzung: alle Tests gegen die Quelle sind gruen. Je Fluss der
Quell-Datenbank wird im Ziel exakt derselbe Datensatz gesucht (reference
product, name, location bzw. name und categories). Eindeutige Treffer werden
umgestellt, alles andere landet im Pruefblatt. Die neuen Dateien liegen in
einem Unterordner mit dem Namen der Ziel-Datenbank; die Originale bleiben
unveraendert.

## Konsole

Fuer Tests und Automatisierung (Rueckgabe: 0 = in Ordnung, 1 = Meldungen,
2 = Fehler):

```
ex2bw-check   PFAD... [--project P] [--databases a,b] [--tests format,fluss,verweis] [--format-ueberbruecken]
ex2bw-apply   DATEI...
ex2bw-write   PFAD... --project P [--loeschen-erlauben] [--verweise-brechen-erlauben] [--ja]
ex2bw-migrate PFAD... --project P --source DB --target DB [--target-project P2] [--source-biosphere B --target-biosphere B2]
```

## Aufbau

```
src/excel2brightway/
  core/
    workbook.py       Excel-Leser mit Zellpositionen (fuer Pruefen UND Schreiben)
    format_check.py   Format-Test, FORMAT_SCHEMA
    backend.py        Zugriff auf Brightway (BrightwayBackend) bzw. MemoryBackend fuer Tests
    search.py         Kandidatensuche und Re-Ranking
    resolve.py        Fluss-Test
    references.py     Verweis-Test
    checks.py         Tests ausfuehren, Pruefblatt + ex2bw.log schreiben
    review_sheet.py   Pruefblatt lesen/schreiben
    apply.py          Auswahl uebernehmen
    write.py          Schreiben planen und ausfuehren
    migrate.py        Datenbank/Version umstellen
    versions.py       Fassungen derselben Datei (_v0.1, _v0.3, _final, ...)
    importer.py       ecoinvent (ecoSpold02, .7z/.zip/Ordner) in ein Projekt einlesen
    settings.py       zuletzt verwendete Einstellungen des Fensters
  brightway_writer.py Schreiben ueber bw2io/bw2data
  gui.py              Fenster (tkinter)
  cli.py              Konsolenbefehle
```

Tests: `uv run pytest` (laufen ohne Brightway, mit `MemoryBackend`).

## Noch zu testen (TODO)

Am 2026-11-24 mit echtem Brightway (bw2data 4.7) geprueft: Fassungen,
Schreiben, Verweis-Test, Pruefblatt-Ablauf ueber die Konsole, ecoinvent-Import
aus entpacktem Ordner. Offen:

* [ ] **Parameter und Formeln** nach Brightway schreiben - braucht eine
  LCI-Datei mit Parametern. Dabei klaeren, ob `ActivityBlock.parameter_group`
  als `group` geschrieben werden soll.
* [ ] **Fenster (GUI) von Hand**: Excel oeffnet sich nach dem Pruefen,
  Auswahlliste im Pruefblatt, Meldung "Bitte Datei in Excel schliessen",
  Import-Dialog (Archiv/Ordner waehlen, Namensvorschlag, Fenster bleibt
  waehrend des Imports bedienbar).
* [ ] **Activity Browser**: geschriebene Datenbanken in den Testprojekten
  `ecoinvent-3.11-cutoff_ex2bw_test` und `ex2bw_import_test` ansehen.
* [ ] **Import aus `.7z`** direkt (bisher nur Mini-Archiv und entpackter Ordner).
* [ ] **Import mit unpassender Biosphere** (ungelinkte Exchanges -> nichts
  geschrieben, eben angelegte Biosphere wieder entfernt).

## Geplant

* Standarddatei, die Pflicht/optional und Metadaten beschreibt und
  `FORMAT_SCHEMA` ersetzt.
* Migrationspakete (randonneur/randonneur_data) beim Umstellen anwenden;
  eigenes Mapping fuer GaBi ueber das randonneur-Excel-Template.
* Testsatz aus echten Fehlfaellen, um die Gewichte der Kandidatensuche zu
  kalibrieren.
