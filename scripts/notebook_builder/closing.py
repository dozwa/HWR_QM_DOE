"""DMAIC phase CLOSING: cells 81..85 of the notebook."""
from __future__ import annotations

from .cells import code, colab_code, md, phase_export_cell

_CLOSING_81_X = r"""---
# Abschluss: Lessons Learned

Nehmt euch 5 Minuten und reflektiert euren DMAIC-Prozess."""

_CLOSING_82_REFLEXION_LESSONS_LEARNED = r"""## 📝 Reflexion: Lessons Learned

Nehmt euch 10 Minuten Zeit, um die folgenden Fragen als Gruppe zu beantworten.
Die Antworten werden in eurem Projektexport gespeichert.
"""

_CLOSING_83_TITLE_REFLEXION_AUSF_LLEN = r"""import ipywidgets as _widgets

_reflexion_fragen = [
    "Was hat in eurem DMAIC-Projekt gut funktioniert?",
    "Welche Fehlerquellen habt ihr identifiziert (Messung, Durchführung, Modell)?",
    "Was würdet ihr beim nächsten Mal anders machen?",
    "Welche Six-Sigma / QM-Konzepte waren am lehrreichsten und warum?",
    "Offene Fragen oder Verständnisprobleme?",
]

_antwort_widgets = []
for _i, _frage in enumerate(_reflexion_fragen, 1):
    display(HTML(f"<h4 style='color:#2563EB; margin:18px 0 4px 0;'>{_i}. {_frage}</h4>"))
    _w = _widgets.Textarea(
        value='',
        placeholder='Hier eure Antwort eingeben...',
        layout=_widgets.Layout(width='100%', height='120px')
    )
    display(_w)
    _antwort_widgets.append(_w)

def _speichere_reflexion(_btn):
    antworten = [w.value.strip() for w in _antwort_widgets]
    if not any(antworten):
        print("ℹ️ Bitte mindestens eine Frage beantworten.")
        return
    projekt.charter["Lessons Learned"] = "\n\n".join(
        f"{i}. {frage}\n   {antwort}"
        for i, (frage, antwort) in enumerate(zip(_reflexion_fragen, antworten), 1)
    )
    helper.speichere_fortschritt(projekt)
    print("✅ Reflexion gespeichert!")

_btn = _widgets.Button(
    description='💾 Reflexion speichern',
    button_style='primary',
    layout=_widgets.Layout(width='250px', height='38px', margin='20px 0 0 0')
)
_btn.on_click(_speichere_reflexion)
display(_btn)"""

_CLOSING_84_TITLE_ALLE_ERGEBNISSE_EXPORTIE = r'''filepath = helper.exportiere_zip(projekt)
print(f"✅ ZIP erstellt: {filepath}")

# Kopie im Drive-Ordner ablegen
import shutil
_save_dir = helper._fortschritt_verzeichnis(projekt)
if _save_dir:
    os.makedirs(_save_dir, exist_ok=True)
    shutil.copy2(filepath, os.path.join(_save_dir, os.path.basename(filepath)))
    print(f"💾 ZIP auch in Google Drive gespeichert.")

try:
    from google.colab import files
    files.download(filepath)
    print("📥 Download gestartet!")
except ImportError:
    print(f"📁 Datei: {filepath}")

print("""
📋 Die ZIP-Datei enthält:
   📂 plots/     – Alle Grafiken als PNG
   📂 daten/     – Alle Datentabellen als CSV
   📄 zusammenfassung.txt – Ergebnisübersicht
""")
'''

_CLOSING_85_X = r"""---
# 🏁 Geschafft!

Ihr habt heute den **kompletten DMAIC-Zyklus** durchlaufen:

| Phase | ✓ |
|-------|----|
| **Define** | Problem quantifiziert, Charter erstellt |
| **Measure** | Messsystem geprüft, Baseline erhoben |
| **Analyze** | Einflussfaktoren identifiziert, Modell bewertet |
| **Improve** | Optimale Einstellungen gefunden und bestätigt |
| **Control** | Prozessfähigkeit bewertet |

## Nächste Schritte
- **Bericht einreichen bis 30. April 2026**
- Nutzt die exportierten PNGs und CSVs als Grundlage
- Nicht nur Ergebnisse zeigen – **erklären, was sie bedeuten!**

> *Das Ziel dieser Session ist nicht, ein perfektes Katapult zu bauen, sondern den DMAIC-Prozess als Werkzeug für systematische Verbesserung zu erlernen und anzuwenden.*"""


def cells():
    return [
        md(_CLOSING_81_X),
        md(_CLOSING_82_REFLEXION_LESSONS_LEARNED),
        colab_code("📝 Reflexion ausfüllen", _CLOSING_83_TITLE_REFLEXION_AUSF_LLEN),
        colab_code("📦 Alle Ergebnisse exportieren (ZIP)", _CLOSING_84_TITLE_ALLE_ERGEBNISSE_EXPORTIE),
        md(_CLOSING_85_X),
    ]
