"""DMAIC phase INTRO: cells 0..4 of the notebook."""
from __future__ import annotations

from .cells import code, colab_code, md, phase_export_cell

_INTRO_00_DMAIC_KATAPULT_VERSUCH = r"""# 🎯 DMAIC Katapult-Versuch

## Qualitätsmanagement – Praxistag

Willkommen zum zweiten QM-Tag! Heute durchlauft ihr den **kompletten DMAIC-Zyklus** an eurem selbstgebauten Katapult.

**Was ihr braucht:** Katapult, Tischtennisbälle, Maßband, Klebeband, Laptop

**Ablauf:**
| Phase | Inhalt | Dauer |
|-------|--------|-------|
| Define | Zielweite, Testwürfe, Charter | 45 min |
| Measure | MSA + Baseline | 75 min |
| Analyze | DoE-Planung + Durchführung + Modell | ~4h |
| Improve | Optimierung + Konfirmation | 60 min |
| Control | Kontrollkarte, Cpk, Vorher/Nachher | 30 min |

⚠️ **Wichtig:** Zellen der Reihe nach ausführen! Keine Zellen überspringen."""

_INTRO_01_TITLE_BIBLIOTHEKEN_INSTALLIERE = r"""!pip install -q statsmodels openpyxl
print("✅ Alle Bibliotheken installiert!")
"""

_INTRO_02_TITLE_GOOGLE_DRIVE_VERBINDEN_F = r"""try:
    from google.colab import drive
    drive.mount('/content/drive')
    print("✅ Google Drive verbunden – Auto-Speicherung aktiv!")
except ImportError:
    print("ℹ️ Nicht in Colab – Fortschritt wird lokal gespeichert.")
"""

_INTRO_03_AUTO_SPEICHERUNG_AKTIV_EUER_FO = r"""> **💾 Auto-Speicherung aktiv.** Euer Fortschritt wird nach jeder Eingabe automatisch
> in Google Drive gespeichert (`MyDrive/DMAIC_Katapult/`). Falls die Colab-Session
> abstürzt, führt die Zellen 1–4 erneut aus — eure Daten werden automatisch geladen.
>
> **Neu starten?** Gruppennamen ändern oder den Ordner in Google Drive löschen.
"""

_INTRO_04_TITLE_INITIALISIERUNG = r"""import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy import stats
from IPython.display import HTML, display
import warnings, os, importlib, sys
warnings.filterwarnings('ignore')

# helper.py von GitHub laden (immer aktuelle Version)
import urllib.request
_helper_url = "https://raw.githubusercontent.com/dozwa/HWR_QM_DOE/main/helper.py"
try:
    urllib.request.urlretrieve(_helper_url, "helper.py")
    print("✅ helper.py von GitHub geladen")
except Exception as _e:
    if os.path.exists("helper.py"):
        print(f"⚠️ GitHub nicht erreichbar – nutze lokale Version")
    else:
        print(f"❌ helper.py konnte nicht geladen werden: {_e}")
        raise

if 'helper' in sys.modules:
    importlib.reload(sys.modules['helper'])
import helper
helper.setup_theme()
print("✅ Initialisierung abgeschlossen!")
"""


def cells():
    return [
        md(_INTRO_00_DMAIC_KATAPULT_VERSUCH),
        colab_code("🔧 Bibliotheken installieren (einmalig ausführen)", _INTRO_01_TITLE_BIBLIOTHEKEN_INSTALLIERE),
        colab_code("🔧 Google Drive verbinden (für Auto-Speicherung)", _INTRO_02_TITLE_GOOGLE_DRIVE_VERBINDEN_F),
        md(_INTRO_03_AUTO_SPEICHERUNG_AKTIV_EUER_FO),
        colab_code("🔧 Initialisierung", _INTRO_04_TITLE_INITIALISIERUNG),
    ]
