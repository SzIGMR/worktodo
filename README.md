# skillrunner

`skillrunner` ist ein plugin-basiertes CLI-Tool zum Ausführen von Skills (Fachaufgaben), die per `pip` installiert werden können. Dieses Repository liefert den Kern und den Skill **budget_matcher**.

## Installation

```bash
pip install -e .
```

Optionaler Solver-Support:

```bash
pip install -e .[opt]
```

## Updates & Reinstall

Wenn du den Code aktualisieren willst (z. B. nach einem `git pull`), installiere die Pakete erneut im Editable-Modus:

```bash
git pull
pip install -e .
```

Falls du die optionalen Solver-Abhängigkeiten nutzt, installiere sie ebenfalls neu:

```bash
pip install -e .[opt]
```

## CLI (Deutsch)

```bash
skillrunner list-skills
```

```bash
skillrunner run budget_matcher \
  --bookings bookings.xlsx \
  --budgets budgets.xlsx \
  --out result.xlsx \
  --mode optimize \
  --fy-start-year 2025
```

### Template-Dateien erzeugen

```bash
skillrunner run budget_matcher --generate-templates ./templates
```

### Schnellstart-Skripte (Explorer-klickbar)

Im Ordner `scripts/` liegen drei Varianten als Shell- und Windows-Batch-Dateien, die sich per Doppelklick starten lassen:

- `run_templates.(sh|bat)` erzeugt Template-Dateien in `scripts/templates/`.
- `run_forecast.(sh|bat)` erstellt eine Forecast-Auswertung.
- `run_optimize.(sh|bat)` führt die Optimierung aus.

Passe in den Skripten bei Bedarf die Pfade zu `bookings.xlsx`, `budgets.xlsx` und das Geschäftsjahr (`FY_START_YEAR`) an.

### Eingaben konvertieren

```bash
skillrunner run budget_matcher \
  --convert-inputs ./converted \
  --bookings rohdaten_bookings.csv \
  --budgets rohdaten_budgets.csv
```

## Skill: budget_matcher

### Eingaben

**Bookings-Datei** (CSV/Excel):

- `Vorname`
- `Bezugsmonat` (Format `YYYY/MM`)
- `EUR` (deutsche Zahlenformate werden erkannt, z. B. `1.640,47` oder ` 7.452,00 € `)

**Budgets-Datei** (CSV/Excel):

- `Projekt`
- `Bewilligt`

CSV-Delimiter werden automatisch erkannt, Spaltennamen werden normalisiert (Trimmen, Kleinschreibung, Sonderzeichen -> `_`).

### Modi

**optimize**
- Bucht jede Zeile entweder auf ein Projekt oder lässt sie als `UNALLOCATED`.
- Keine Zeilen-Splittung.
- Budgets dürfen den Mindestrest `min_project_rest_eur` nicht unterschreiten.
- Ziel: maximale Allokation (dominiert alle Strafen), danach weiche Nebenbedingungen:
  - Begrenzung der Projekte pro Person
  - Glättung monatlicher Schwankungen
  - Minimale Anzahl unzugeordneter Personen (sehr niedrige Priorität)

**forecast**
- Aggregiert Ist-Werte im Geschäftsjahr (April bis März).
- Forecast mit Trendlinie (ab 3 Monaten) oder Run-Rate.
- Confidence-Band: niedrige Bandbreite = Run-Rate, hohe Bandbreite = Trend + kleiner Aufschlag.

### Ausgabe (Excel)

- `Inputs_Budgets` (normalisiert)
- `Inputs_Bookings` (normalisiert, inkl. `out_of_period`)
- `Allocation`
- `Project_Summary`
- `Unallocated`
- `Monthly_Project_Summary`
- `Diagnostics`
- `Out_of_Period` (falls vorhanden und nicht inkludiert)
- `Forecast` (nur im Forecast-Modus)

## Tests

```bash
pytest
```

## Hinweise

- Standardmäßig wird PuLP (CBC) verwendet, sofern installiert. Andernfalls greift eine heuristische Zuordnung.
- `.xls`-Dateien werden über `xlrd` gelesen (bereits als Abhängigkeit enthalten).
- Keine Abhängigkeit von `pandas`, der Import ist nicht notwendig.
