# DMAIC Katapult-Versuch

Interaktives Jupyter Notebook zur Durchfuehrung eines vollstaendigen **DMAIC-Zyklus** (Define, Measure, Analyze, Improve, Control) am Beispiel eines Katapult-Experiments. Entwickelt fuer den Qualitaetsmanagement-Kurs an der HWR Berlin.

Studierende optimieren die Wurfweite eines selbstgebauten Katapults mithilfe statistischer Methoden -- von der Messsystemanalyse ueber die Versuchsplanung (Design of Experiments) bis zur Prozessfaehigkeitsanalyse.

## Aufbau

```
DMAIC_Katapult_Versuch.ipynb   # Interaktives Notebook (Google Colab)
helper.py                      # Statistische Logik (~3.300 Zeilen)
```

- **Notebook** (`DMAIC_Katapult_Versuch.ipynb`): Fuehrt Studierende Schritt fuer Schritt durch die 5 DMAIC-Phasen. Alle komplexen Berechnungen sind in `helper.py` gekapselt, sodass die Interaktion ueber einfache Funktionsaufrufe und Formularfelder erfolgt.
- **Helper-Modul** (`helper.py`): Enthaelt die gesamte statistische Logik, Visualisierungen und Export-Funktionen. Wird beim Start des Notebooks automatisch von GitHub geladen.

## DMAIC-Phasen im Notebook

### 1. DEFINE
- Zielweite wird per Gruppen-Seed deterministisch zugewiesen (200-450 cm)
- 5 Testwuerfe zur Bestandsaufnahme mit CV-Bewertung
- Projektcharter als strukturiertes Formular

### 2. MEASURE
- **MSA Type-1:** Bias und Repeatability je Messende Person
- **Gage R&R (ANOVA, AIAG-konform):** Varianzzerlegung in Repeatability, Reproducibility und Interaktion; %GRR-Bewertung
- **Baseline:** 10-15 Wuerfe mit Histogramm, Normalverteilungskurve und Shapiro-Wilk-Test
- Excel-Template-Generierung fuer die Datenerfassung

### 3. ANALYZE
- **Design of Experiments:** Voll-, halb- und viertelfaktoriell (2^k) mit Blocking, Centerpoints und Randomisierung
- **OLS-Regression:** Haupteffekte + Zweifach-Interaktionen
- **Hierarchisches Pruning:** Backward Elimination mit Hierarchie-Schutz (signifikante Interaktion schuetzt zugehoerige Haupteffekte)
- Pareto-Diagramm, Koeffiziententabelle, R²-Bewertung
- VIF-Pruefung, Lack-of-Fit-Test, Residuenanalyse

### 4. IMPROVE
- **Konturplots:** Wurfweite, Dispersionsmodell (Taguchi), Fehlerfortpflanzung
- **Analytische Optimierung:** Drei Strategien -- Mittelwert (Accuracy), Varianz (Precision), Dual (gewichteter Kompromiss)
- Strategievergleich in Uebersichtstabelle
- Konfirmationsexperiment mit Vorhersageintervall-Pruefung

### 5. CONTROL
- **I-MR-Kontrollkarte:** Stabilitaetspruefung mit UCL/LCL
- **Normalverteilungspruefung:** Shapiro-Wilk + Q-Q-Plot
- **Prozessfaehigkeit:** Cp/Cpk mit Dual-Skala (Industrie vs. Katapult-Kontext)
- Vorher/Nachher-Vergleich (Baseline vs. Konfirmation)

## Voraussetzungen

### Umgebung
- **Google Colab** (empfohlen): Notebook ist fuer Colab optimiert, inkl. Google Drive Auto-Save
- **Lokaler Jupyter Server:** Funktioniert ebenfalls, ohne Auto-Save auf Drive

### Python-Pakete
```
numpy
pandas
scipy
statsmodels
matplotlib
openpyxl
```

In Colab werden fehlende Pakete automatisch installiert (Zelle 1).

## Schnellstart

1. Notebook in Google Colab oeffnen
2. Zellen 1-4 ausfuehren (Bibliotheken, Drive, Initialisierung, Projektsetup)
3. Gruppenname und -nummer eingeben
4. Den Anweisungen im Notebook folgen -- Phase fuer Phase

### Fortschritt wiederherstellen

Bei Colab-Session-Abbruch:
1. Zellen 1-4 erneut ausfuehren
2. Bei "Projekt einrichten" den Modus **Fortschritt laden** waehlen
3. Gleichen Gruppennamen und -nummer eingeben

Der Fortschritt wird als JSON in Google Drive gespeichert (`MyDrive/DMAIC_Katapult/`) und bei Bedarf inkl. Regressionsmodell automatisch rekonstruiert.

## Export

Am Ende des Tages koennen alle Ergebnisse als ZIP exportiert werden:
- Alle Plots als PNG (150 dpi)
- Datentabellen als CSV
- Textzusammenfassung aller Phasen

## Lizenz

Fuer den internen Gebrauch im QM-Kurs an der HWR Berlin.
