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

## Nutzung in Google Colab

Das Notebook ist fuer Google Colab optimiert. Alle Eingaben erfolgen ueber Colab-Formularfelder (`#@param`), sodass kein Python-Wissen noetig ist.

### Notebook oeffnen

Das Notebook direkt in Colab oeffnen:

```
https://colab.research.google.com/github/dozwa/HWR_QM_DOE/blob/main/DMAIC_Katapult_Versuch.ipynb
```

Alternativ: In Colab ueber **Datei > Notebook oeffnen > GitHub** die URL `dozwa/HWR_QM_DOE` eingeben.

### Setup (Zellen 1-4)

Die ersten vier Code-Zellen muessen bei **jedem Sitzungsstart** ausgefuehrt werden:

| Zelle | Funktion |
|-------|----------|
| 1 | Installiert `statsmodels` und `openpyxl` (in Colab nicht vorinstalliert) |
| 2 | Verbindet Google Drive fuer Auto-Speicherung -- beim ersten Mal Zugriff erlauben |
| 3 | Laedt `helper.py` automatisch von GitHub und initialisiert das Theme |
| 4 | Projekt einrichten: neues Projekt anlegen oder Fortschritt laden |

### Projekt einrichten (Zelle 4)

Beim Einrichten gibt es zwei Modi:

**Neues Projekt:**
1. Modus auf `Neues Projekt` belassen
2. Euren **Gruppennamen** und **Gruppennummer** eintragen
3. Zelle ausfuehren -- die Zielweite wird automatisch aus dem Gruppennamen abgeleitet

**Fortschritt laden (nach Session-Abbruch):**
1. Modus auf `Fortschritt laden` aendern
2. Zelle ausfuehren -- verfuegbare Speicherstaende werden angezeigt
3. Gewuenschte Nr. bei `speicherstand_nr` eintragen
4. Zelle erneut ausfuehren

Das Notebook erkennt auch automatisch vorhandene Speicherstaende anhand des Gruppennamens. Wenn ihr einfach denselben Gruppennamen eingebt, wird der letzte Stand geladen.

### Arbeitsablauf waehrend des Tages

1. **Zellen der Reihe nach ausfuehren** -- keine Zellen ueberspringen
2. **Formularfelder ausfuellen** und die jeweilige Zelle ausfuehren (Shift+Enter)
3. **Excel-Dateien herunterladen**, ausfuellen und wieder hochladen (MSA-Template, DoE-Ergebnisse, Konfirmation)
4. **Fortschritt wird automatisch gespeichert** nach jeder Eingabe in `MyDrive/DMAIC_Katapult/`

### Datei-Upload und -Download

An mehreren Stellen generiert das Notebook Excel-Dateien und erwartet ausgefuellte Versionen zurueck:

| Phase | Download | Ausfuellen | Upload |
|-------|----------|------------|--------|
| MEASURE | `MSA_Messung_Template.xlsx` | Type-1 + Reproduzierbarkeit Messungen eintragen | Ausgefuellte Datei hochladen |
| ANALYZE | `DoE_Versuchsergebnisse.xlsx` | Wurfweiten der Versuche eintragen | Ausgefuellte Datei hochladen |
| IMPROVE | `Konfirmation.xlsx` | Konfirmationswuerfe eintragen | Werte direkt im Notebook eingeben |

In Colab erscheint bei Downloads ein Dialog automatisch. Bei Uploads wird ein Datei-Auswahl-Dialog angezeigt.

### Auto-Speicherung und Wiederherstellung

Der Fortschritt wird als `fortschritt.json` in Google Drive gespeichert:

```
MyDrive/
  DMAIC_Katapult/
    GruppeA_1/
      fortschritt.json      # Projektdaten (alle Phasen)
      DEFINE/               # Plots und Daten der Phase
      MEASURE/
      ...
```

**Was gespeichert wird:** Alle eingegebenen Werte, Messdaten, DoE-Ergebnisse und Optimierungseinstellungen. Das Regressionsmodell wird beim Laden automatisch aus den gespeicherten Rohdaten neu berechnet.

**Was nicht gespeichert wird:** Plots und Figuren (werden neu erzeugt), temporaere Variablen.

### Haeufige Probleme

| Problem | Loesung |
|---------|---------|
| Session abgelaufen / getrennt | Zellen 1-4 erneut ausfuehren, Fortschritt wird automatisch geladen |
| Google Drive nicht verbunden | Zelle 2 ausfuehren und Zugriff erlauben |
| `helper.py` Fehler beim Laden | Netzwerkverbindung pruefen; bei Offline-Nutzung `helper.py` manuell hochladen |
| Excel-Upload funktioniert nicht | Dateiformat pruefen (`.xlsx`), keine Formeln in Datenzellen |
| Fehlermeldung "Mindestens 3 Faktoren" | In der Analyze-Phase mindestens 3 Faktoren definieren |
| Plots werden nicht angezeigt | Zelle erneut ausfuehren; ggf. `plt.show()` pruefen |

### Tipps

- **Nicht parallel arbeiten:** Nur ein Tab pro Gruppe, da Colab Sessions nicht synchronisiert werden
- **Regelmaessig speichern:** Die Phase-Speicherung (`exportiere_phase_auf_drive`) sichert zusaetzlich Plots und CSVs auf Google Drive
- **Ergebnisse exportieren:** Am Ende des Tages die ZIP-Export-Zelle ausfuehren -- sie enthaelt alle Plots und Daten fuer den Bericht

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
