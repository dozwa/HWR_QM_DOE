# DMAIC Katapult-Versuch

Interaktives Jupyter Notebook zur Durchführung eines vollständigen **DMAIC-Zyklus** (Define, Measure, Analyze, Improve, Control) am Beispiel eines Katapult-Experiments. Entwickelt für den Qualitätsmanagement-Kurs an der HWR Berlin.

Studierende optimieren die Wurfweite eines selbstgebauten Katapults mithilfe statistischer Methoden – von der Messsystemanalyse über die Versuchsplanung (Design of Experiments) bis zur Prozessfähigkeitsanalyse.

## Aufbau

```
DMAIC_Katapult_Versuch.ipynb   # Interaktives Notebook (Google Colab)
helper.py                      # Statistische Logik (~3.300 Zeilen)
```

- **Notebook** (`DMAIC_Katapult_Versuch.ipynb`): Führt Studierende Schritt für Schritt durch die 5 DMAIC-Phasen. Alle komplexen Berechnungen sind in `helper.py` gekapselt, sodass die Interaktion über einfache Funktionsaufrufe und Formularfelder erfolgt.
- **Helper-Modul** (`helper.py`): Enthält die gesamte statistische Logik, Visualisierungen und Export-Funktionen. Wird beim Start des Notebooks automatisch von GitHub geladen.

## DMAIC-Phasen im Notebook

### 1. DEFINE
- Zielweite wird per Gruppen-Seed deterministisch zugewiesen (200–450 cm)
- 5 Testwürfe zur Bestandsaufnahme mit CV-Bewertung
- Projektcharter als strukturiertes Formular

### 2. MEASURE
- **MSA Type-1:** Bias und Repeatability je messende Person
- **Gage R&R (ANOVA, AIAG-konform):** Varianzzerlegung in Repeatability, Reproducibility und Interaktion; %GRR-Bewertung
- **Baseline:** 10–15 Würfe mit Histogramm, Normalverteilungskurve und Shapiro-Wilk-Test
- Excel-Template-Generierung für die Datenerfassung

### 3. ANALYZE
- **Design of Experiments:** Voll-, halb- und viertelfaktoriell (2^k) mit Blocking, Centerpoints und Randomisierung
- **OLS-Regression:** Haupteffekte + Zweifach-Interaktionen
- **Hierarchisches Pruning:** Backward Elimination mit Hierarchie-Schutz (signifikante Interaktion schützt zugehörige Haupteffekte)
- Pareto-Diagramm, Koeffiziententabelle, R²-Bewertung
- VIF-Prüfung, Lack-of-Fit-Test, Residuenanalyse

### 4. IMPROVE
- **Konturplots:** Wurfweite, Dispersionsmodell (Taguchi), Fehlerfortpflanzung
- **Analytische Optimierung:** Drei Strategien – Mittelwert (Accuracy), Varianz (Precision), Dual (gewichteter Kompromiss)
- Strategievergleich in Übersichtstabelle
- Konfirmationsexperiment mit Vorhersageintervall-Prüfung

### 5. CONTROL
- **I-MR-Kontrollkarte:** Stabilitätsprüfung mit UCL/LCL
- **Normalverteilungsprüfung:** Shapiro-Wilk + Q-Q-Plot
- **Prozessfähigkeit:** Cp/Cpk mit Dual-Skala (Industrie vs. Katapult-Kontext)
- Vorher/Nachher-Vergleich (Baseline vs. Konfirmation)

## Voraussetzungen

### Umgebung
- **Google Colab** (empfohlen): Notebook ist für Colab optimiert, inkl. Google Drive Auto-Save
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

1. Notebook in Google Colab öffnen
2. Zellen 1–4 ausführen (Bibliotheken, Drive, Initialisierung, Projektsetup)
3. Gruppenname und -nummer eingeben
4. Den Anweisungen im Notebook folgen – Phase für Phase

## Nutzung in Google Colab

Das Notebook ist für Google Colab optimiert. Alle Eingaben erfolgen über Colab-Formularfelder (`#@param`), sodass kein Python-Wissen nötig ist.

### Notebook öffnen

Das Notebook direkt in Colab öffnen:

```
https://colab.research.google.com/github/dozwa/HWR_QM_DOE/blob/main/DMAIC_Katapult_Versuch.ipynb
```

Alternativ: In Colab über **Datei > Notebook öffnen > GitHub** die URL `dozwa/HWR_QM_DOE` eingeben.

### Setup (Zellen 1–4)

Die ersten vier Code-Zellen müssen bei **jedem Sitzungsstart** ausgeführt werden:

| Zelle | Funktion |
|-------|----------|
| 1 | Installiert `statsmodels` und `openpyxl` (in Colab nicht vorinstalliert) |
| 2 | Verbindet Google Drive für Auto-Speicherung – beim ersten Mal Zugriff erlauben |
| 3 | Lädt `helper.py` automatisch von GitHub und initialisiert das Theme |
| 4 | Projekt einrichten: neues Projekt anlegen oder Fortschritt laden |

### Projekt einrichten (Zelle 4)

Beim Einrichten gibt es zwei Modi:

**Neues Projekt:**
1. Modus auf `Neues Projekt` belassen
2. Euren **Gruppennamen** und **Gruppennummer** eintragen
3. Zelle ausführen – die Zielweite wird automatisch aus dem Gruppennamen abgeleitet

**Fortschritt laden (nach Session-Abbruch):**
1. Modus auf `Fortschritt laden` ändern
2. Zelle ausführen – verfügbare Speicherstände werden angezeigt
3. Gewünschte Nr. bei `speicherstand_nr` eintragen
4. Zelle erneut ausführen

Das Notebook erkennt auch automatisch vorhandene Speicherstände anhand des Gruppennamens. Wenn ihr einfach denselben Gruppennamen eingebt, wird der letzte Stand geladen.

### Arbeitsablauf während des Tages

1. **Zellen der Reihe nach ausführen** – keine Zellen überspringen
2. **Formularfelder ausfüllen** und die jeweilige Zelle ausführen (Shift+Enter)
3. **Excel-Dateien herunterladen**, ausfüllen und wieder hochladen (MSA-Template, DoE-Ergebnisse, Konfirmation)
4. **Fortschritt wird automatisch gespeichert** nach jeder Eingabe in `MyDrive/DMAIC_Katapult/`

### Datei-Upload und -Download

An mehreren Stellen generiert das Notebook Excel-Dateien und erwartet ausgefüllte Versionen zurück:

| Phase | Download | Ausfüllen | Upload |
|-------|----------|-----------|--------|
| MEASURE | `MSA_Messung_Template.xlsx` | Type-1 + Reproduzierbarkeit Messungen eintragen | Ausgefüllte Datei hochladen |
| ANALYZE | `DoE_Versuchsergebnisse.xlsx` | Wurfweiten der Versuche eintragen | Ausgefüllte Datei hochladen |
| IMPROVE | `Konfirmation.xlsx` | Konfirmationswürfe eintragen | Werte direkt im Notebook eingeben |

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

**Was nicht gespeichert wird:** Plots und Figuren (werden neu erzeugt), temporäre Variablen.

### Häufige Probleme

| Problem | Lösung |
|---------|--------|
| Session abgelaufen / getrennt | Zellen 1–4 erneut ausführen, Fortschritt wird automatisch geladen |
| Google Drive nicht verbunden | Zelle 2 ausführen und Zugriff erlauben |
| `helper.py` Fehler beim Laden | Netzwerkverbindung prüfen; bei Offline-Nutzung `helper.py` manuell hochladen |
| Excel-Upload funktioniert nicht | Dateiformat prüfen (`.xlsx`), keine Formeln in Datenzellen |
| Fehlermeldung „Mindestens 3 Faktoren" | In der Analyze-Phase mindestens 3 Faktoren definieren |
| Plots werden nicht angezeigt | Zelle erneut ausführen; ggf. `plt.show()` prüfen |

### Tipps

- **Nicht parallel arbeiten:** Nur ein Tab pro Gruppe, da Colab-Sessions nicht synchronisiert werden
- **Regelmäßig speichern:** Die Phase-Speicherung (`exportiere_phase_auf_drive`) sichert zusätzlich Plots und CSVs auf Google Drive
- **Ergebnisse exportieren:** Am Ende des Tages die ZIP-Export-Zelle ausführen – sie enthält alle Plots und Daten für den Bericht

### Fortschritt wiederherstellen

Bei Colab-Session-Abbruch:
1. Zellen 1–4 erneut ausführen
2. Bei „Projekt einrichten" den Modus **Fortschritt laden** wählen
3. Gleichen Gruppennamen und -nummer eingeben

Der Fortschritt wird als JSON in Google Drive gespeichert (`MyDrive/DMAIC_Katapult/`) und bei Bedarf inkl. Regressionsmodell automatisch rekonstruiert.

## Export

Am Ende des Tages können alle Ergebnisse als ZIP exportiert werden:
- Alle Plots als PNG (150 dpi)
- Datentabellen als CSV
- Textzusammenfassung aller Phasen

## Lizenz

Für den internen Gebrauch im QM-Kurs an der HWR Berlin.
