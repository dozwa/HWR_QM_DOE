# Statapult – Virtueller Six Sigma Katapult-Simulator

Ein CLI-Tool, das ein realistisches Tisch-Katapult (Statapult) simuliert.
Erzeugt Messdaten fuer DMAIC-Uebungen: DOE, MSA, Regelkarten und Prozessfaehigkeit.

## Installation

```bash
pip install -e ./catapult
```

Fuer Batch-Modus mit Excel/CSV:

```bash
pip install -e "./catapult[batch]"
```

## Schnellstart

```bash
# Einzelschuss mit Default-Einstellungen
statapult shoot --seed 42

# Detaillierte Ausgabe
statapult shoot --abzugswinkel 160 --stoppwinkel 100 --seed 42 --verbose

# 10 Wiederholungen
statapult shoot --repeat 10 --seed 42

# Faktor-Uebersicht
statapult info

# Excel-Vorlage des Notebooks befuellen
statapult fill DoE_Versuchsergebnisse.xlsx --seed 42
```

## Faktoren

| Faktor | Bereich | Einheit | Beschreibung |
|--------|---------|---------|-------------|
| `--abzugswinkel` | 130–170 | Grad | Rueckzugswinkel des Arms |
| `--stoppwinkel` | 70–110 | Grad | Winkel bei dem der Arm stoppt |
| `--gummiband-position` | 8–18 | cm | Position des Gummibands auf dem Arm |
| `--becherposition` | 8–22 | cm | Position des Bechers auf dem Arm |
| `--pin-hoehe` | 8–18 | cm | Hoehe des Drehpunkts |
| `--ballgewicht` | 5–30 | g | Masse des Balls |
| `--wind` | -2 bis 2 | m/s | Windgeschwindigkeit (positiv = Rueckenwind) |

Erreichbarer Wurfweiten-Bereich: **100–600 cm** (abhaengig von Faktorkombination).

## Physik-Modell

Additives Modell mit physikalisch motivierten Termen:

```
d = D_BASE + Σ(aᵢ·xᵢ) + Σ(qᵢ·xᵢ²) + Σ(bᵢⱼ·xᵢ·xⱼ) + Rauschen
```

- **Haupteffekte** (linear): Jeder Faktor hat einen kalibrierten Effekt auf die Wurfweite
- **Kruemmung** (quadratisch): Nur bei Stoppwinkel (optimaler Abwurfwinkel) und Becherposition (Hebelarm vs. Traegheit)
- **Wechselwirkungen**: Physikalische Kopplungen (z.B. Abzugswinkel x Gummiband = Energie)
- **Rauschen**: Mehrstufig (Messung, Setup, Gummiband, Release, Wind-Turbulenz)

## Befehle

### `shoot` – Einzelschuss / Wiederholungen

```bash
# Mit bestimmten Einstellungen
statapult shoot --abzugswinkel 170 --stoppwinkel 110 \
  --gummiband-position 18 --becherposition 15 --pin-hoehe 13 \
  --ballgewicht 10 --seed 42

# JSON-Ausgabe (z.B. fuer Weiterverarbeitung)
statapult shoot --seed 42 --format json

# CSV-Ausgabe
statapult shoot --repeat 20 --seed 42 --format csv
```

### `batch` – Versuchsplan ausfuehren

Fuehrt einen kompletten Versuchsplan aus einer CSV-Datei aus:

```bash
statapult batch -i versuchsplan.csv -o ergebnisse.csv --seed 42
```

Die CSV-Datei muss Spalten mit den Faktornamen enthalten.
Fehlende Faktoren werden mit Defaults aufgefuellt.

Beispiel `versuchsplan.csv`:

```csv
abzugswinkel,stoppwinkel,gummiband_position,becherposition,pin_hoehe
130,70,8,8,8
170,110,18,22,18
150,90,13,15,13
```

### `msa` – MSA-Daten generieren

Erzeugt Daten fuer Measurement System Analysis (Type-1 / Gage R&R):

```bash
statapult msa --operators 3 --measurements-per-operator 10 \
  --seed 42 -o msa_daten.csv
```

Jeder Operator bekommt einen eigenen systematischen Mess-Bias.

### `fill` – Excel-Vorlagen des Notebooks befuellen

Befuellt die Excel-Dateien, die das DMAIC-Notebook erzeugt, mit simulierten Daten.
Erkennt automatisch den Vorlagentyp (MSA, DoE, Konfirmation) und unterstuetzt
sowohl kodierte (-1/+1) als auch natuerliche Faktorwerte.

```bash
# MSA-Vorlage befuellen (Type-1 + Reproduzierbarkeit)
statapult fill MSA_Messung_Template.xlsx --seed 42

# DoE-Versuchsplan ausfuehren (liest Faktoren, schreibt Ergebnisse)
statapult fill DoE_Versuchsergebnisse.xlsx --seed 42

# Konfirmationswuerfe simulieren
statapult fill Konfirmation.xlsx --seed 42

# Ausgabe in neue Datei (Original bleibt erhalten)
statapult fill DoE_Versuchsergebnisse.xlsx -o DoE_filled.xlsx --seed 42
```

### `control` – Regelkarten-Daten generieren

Erzeugt Daten fuer I-MR-Regelkarten:

```bash
# Stabiler Prozess
statapult control --shots 25 --seed 42 -o regelkarte.csv

# Mit Drift (Gummiband-Ermuedung)
statapult control --shots 25 --drift 0.1 --seed 42 -o regelkarte_drift.csv
```

### `info` – Faktoruebersicht

```bash
statapult info
```

## Verwendung als Python-Modul

```python
from statapult import Statapult

katapult = Statapult(seed=42)

# Einzelschuss
result = katapult.shoot({
    "abzugswinkel": 160,
    "stoppwinkel": 100,
    "gummiband_position": 15,
    "becherposition": 18,
    "pin_hoehe": 12,
})
print(f"Wurfweite: {result.wurfweite_cm:.1f} cm")

# Deterministisch (ohne Rauschen)
result = katapult.shoot({"abzugswinkel": 160}, noise_level=0)

# Batch mit pandas DataFrame
import pandas as pd
plan = pd.read_csv("versuchsplan.csv")
ergebnisse = katapult.batch(plan)
```

### Integration mit helper.py

```python
from statapult import Statapult, STANDARD_FACTORS

# Faktoren im helper.py-Format
projekt.faktoren = [f.to_helper_format() for f in STANDARD_FACTORS.values()]

# Versuchsplan ausfuehren
katapult = Statapult(seed=projekt.seed)
projekt.doe_ergebnisse = katapult.batch(projekt.versuchsplan)
```

## Optionen

| Option | Beschreibung |
|--------|-------------|
| `--seed INT` | Zufalls-Seed fuer Reproduzierbarkeit |
| `--repeat N` | Anzahl Wiederholungen (nur `shoot`) |
| `--noise-level FLOAT` | Rausch-Multiplikator, 0 = deterministisch |
| `--format text\|csv\|json` | Ausgabeformat |
| `--verbose` / `-v` | Zeigt physikalische Zwischenwerte |
| `--config FILE` | Eigene YAML-Konfiguration laden |

## Konfiguration

Das Katapult kann ueber eine YAML-Datei angepasst werden (siehe `src/statapult/defaults.yaml`):

```bash
statapult shoot --config mein_katapult.yaml --seed 42
```

Anpassbar sind u.a. Federkonstante, Arm-Masse, Rauschparameter und Balltypen.

## Tests

```bash
cd catapult
pip install -e ".[dev]"
python -m pytest tests/ -v
```

115 Tests in 8 Modulen:

| Modul | Tests | Prueft |
|-------|-------|--------|
| `test_physics.py` | 14 | Physik-Modell, Effektrichtungen, Distanzbereich |
| `test_noise.py` | 10 | Rauschmodell, Operator-Bias, Drift |
| `test_factors.py` | 10 | Faktor-Definitionen, Coding, helper.py-Format |
| `test_simulator.py` | 12 | Orchestrator, Batch, Seed-Reproduzierbarkeit |
| `test_cli.py` | 8 | CLI-Subcommands, Ausgabeformate |
| `test_statistical_properties.py` | 8 | DOE/MSA/Cpk/Regelkarten-Eignung |
| `test_integration.py` | 32 | DOE -> Modell -> Optimierung -> Konfirmation |
| `test_excel_filler.py` | 18 | MSA/DoE/Konfirmation Excel-Roundtrip |
