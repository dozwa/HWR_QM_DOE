"""DMAIC phase DEFINE: cells 5..16 of the notebook."""
from __future__ import annotations

from .cells import code, colab_code, md, phase_export_cell

_DEFINE_05_TITLE_PROJEKT_EINRICHTEN = r"""modus = "Neues Projekt" #@param ["Neues Projekt", "Fortschritt laden"]
gruppenname = "Gruppe A" #@param {type:"string"}
gruppennummer = 1 #@param {type:"integer"}
toleranz = 15 #@param {type:"number"}
speicherstand_nr = 0 #@param {type:"integer"}

if modus == "Fortschritt laden":
    _staende = helper.finde_speicherstaende()
    if not _staende:
        print("❌ Keine Speicherstände im Google Drive gefunden.")
        print("   Wechsle zu 'Neues Projekt' um ein neues Projekt zu starten.")
    elif speicherstand_nr == 0:
        helper.zeige_speicherstand_auswahl(_staende)
        print("\n👆 Trage die gewünschte Nr. bei 'speicherstand_nr' ein und führe diese Zelle erneut aus.")
    else:
        if speicherstand_nr > len(_staende):
            print(f"❌ Nr. {speicherstand_nr} existiert nicht (max. {len(_staende)}).")
        else:
            _s = _staende[speicherstand_nr - 1]
            projekt = helper.lade_fortschritt(_s["gruppenname"], _s["gruppennummer"])
            if projekt is not None:
                helper.zeige_restore_zusammenfassung(projekt)
            else:
                print("❌ Laden fehlgeschlagen.")
else:
    # Neues Projekt oder auto-detect existierenden Speicherstand
    _restored = helper.lade_fortschritt(gruppenname, gruppennummer)
    if _restored is not None:
        projekt = _restored
        helper.zeige_restore_zusammenfassung(projekt)
    else:
        projekt = helper.init_projekt(gruppenname, gruppennummer, toleranz)
        helper.speichere_fortschritt(projekt)
        print(f"✅ Neues Projekt initialisiert!")
        print(f"   Gruppe: {projekt.gruppenname} (Nr. {projekt.gruppennummer})")
        print(f"   Zielweite: {projekt.zielweite:.0f} cm ± {projekt.toleranz:.0f} cm")
"""

_DEFINE_06_X = r"""---
# Phase 1: DEFINE

## Was ist das Problem? Wie schlecht ist es?

In der Define-Phase definiert ihr das Problem, das ihr lösen wollt. Euer Katapult hat eine **Zielweite** erhalten – könnt ihr diese reproduzierbar treffen?"""

_DEFINE_07_TITLE_EURE_ZIELWEITE = r"""print(f"{'='*50}")
print(f"  EURE AUFGABE")
print(f"  Zielweite:  {projekt.zielweite:.0f} cm")
print(f"  Toleranz:   ±{projekt.toleranz:.0f} cm")
print(f"  Zielband:   [{projekt.zielweite - projekt.toleranz:.0f}, {projekt.zielweite + projekt.toleranz:.0f}] cm")
print(f"{'='*50}")
print(f"\nStellt euer Katapult auf eine Einstellung und macht 5 Testwürfe!")"""

_DEFINE_08_TESTW_RFE_DURCHF_HREN = r"""### Testwürfe durchführen

1. Stellt euer Katapult auf eine **feste Einstellung**
2. Macht **5 Würfe** und messt die Wurfweite
3. Tragt die Werte unten ein

> ⚠️ Noch keine Optimierung! Einfach eure aktuelle Standardeinstellung testen."""

_DEFINE_09_TITLE_TESTW_RFE_EINGEBEN = r"""wurf_1 = 0.0 #@param {type:"number"}
wurf_2 = 0.0 #@param {type:"number"}
wurf_3 = 0.0 #@param {type:"number"}
wurf_4 = 0.0 #@param {type:"number"}
wurf_5 = 0.0 #@param {type:"number"}

wuerfe = np.array([wurf_1, wurf_2, wurf_3, wurf_4, wurf_5])
wuerfe = wuerfe[wuerfe > 0]  # Nullen entfernen

if len(wuerfe) > 0:
    projekt.testwuerfe = wuerfe
    print(f"✅ {len(wuerfe)} Testwürfe eingetragen: {wuerfe}")
    helper.speichere_fortschritt(projekt)
elif len(projekt.testwuerfe) > 0:
    print(f"ℹ️ {len(projekt.testwuerfe)} Testwürfe aus gespeichertem Fortschritt geladen.")
else:
    print("⚠️ Bitte mindestens einen Testwurf eingeben!")
"""

_DEFINE_10_TITLE_TESTWURF_AUSWERTUNG = r"""helper.zeige_testwurf_ergebnis(projekt)"""

_DEFINE_11_TITLE_MESSMODUS_W_HLEN = r'''messmodus = "1D (nur Weite)" #@param ["1D (nur Weite)", "2D (Weite + Querversatz)"]
projekt.messmodus = "2D" if "2D" in messmodus else "1D"
print(f"✅ Messmodus: {projekt.messmodus}")
if projekt.messmodus == "2D":
    display(HTML("""
    <div style="padding:10px; border-left:4px solid #EA580C; background:#FEF9C3; border-radius:4px;">
        ⚠️ <strong>Hinweis:</strong> Bei 2D-Messung verdoppelt sich der Messaufwand pro Wurf.
        Empfohlen nur für Gruppen, die seitliche Präzision einbeziehen wollen.
    </div>"""))
helper.speichere_fortschritt(projekt)
'''

_DEFINE_12_PROJEKTCHARTER = r"""### Projektcharter

Die Projektcharter dokumentiert euer Qualitätsverbesserungsprojekt. Sie ist das zentrale Dokument der Define-Phase.

> 📋 **Für den Bericht:** Die Projektcharter gehört vollständig in euren Bericht!"""

_DEFINE_13_TITLE_PROJEKTCHARTER_AUSF_LLEN = r"""problemstellung = "z.B. Katapult trifft die Zielweite nicht reproduzierbar" #@param {type:"string"}
projektziel = "z.B. Wurfweite 450 cm ± 15 cm mit Cpk > 1.0" #@param {type:"string"}
scope = "z.B. Optimierung der Faktoreinstellungen, nicht der Konstruktion" #@param {type:"string"}
projektleiter = "" #@param {type:"string"}
protokollant = "" #@param {type:"string"}
versuchsdurchfuehrende = "" #@param {type:"string"}

projekt.charter = {
    "Gruppenname": projekt.gruppenname,
    "Zielweite": f"{projekt.zielweite:.0f} cm ± {projekt.toleranz:.0f} cm",
    "Problemstellung": problemstellung,
    "Projektziel": projektziel,
    "Scope": scope,
    "Projektleiter/in": projektleiter,
    "Protokollant/in": protokollant,
    "Versuchsdurchführende": versuchsdurchfuehrende,
    "Datum": "23.04.2026",
}

display(HTML(helper.formatiere_charter(projekt)))
helper.speichere_fortschritt(projekt)
"""

_DEFINE_14_DIV_STYLE_PADDING_10PX_BORDER = r"""<div style="padding:10px; border-left:4px solid #2563EB; background:#DBEAFE; border-radius:4px;">
📋 <strong>Für den Bericht:</strong> Die Projektcharter und die Zielscheibe der Testwürfe gehören in die Define-Phase eures Berichts. Exportiert die Grafiken am Ende des Tages.
</div>"""

_DEFINE_15_DETAILS_STYLE_MARGIN_10PX_0_PA = r"""<details style="margin:10px 0; padding:8px; background:#F9FAFB; border:1px solid #E5E7EB; border-radius:6px;">
<summary style="cursor:pointer; font-weight:bold; color:#2563EB;">
🔍 Für Neugierige: Was ist der Variationskoeffizient (CV)?
</summary>
<div style="margin-top:8px; padding:8px; font-size:0.95em;">

Der **Variationskoeffizient** (CV) setzt die Standardabweichung ins Verhältnis zum Mittelwert:

$$CV = \frac{\sigma}{\bar{x}} \times 100\%$$

Er ist ein **dimensionsloser** Streuungsmaß – damit kann man die Streuung verschiedener Prozesse vergleichen, unabhängig von der absoluten Größe.

**Interpretation für Katapulte:**
- CV < 15%: Katapult ist reproduzierbar
- CV 15-30%: Grenzwertig
- CV > 30%: Grundlegendes Problem mit der Reproduzierbarkeit
</div>
</details>"""


def cells():
    return [
        colab_code("📝 Projekt einrichten", _DEFINE_05_TITLE_PROJEKT_EINRICHTEN),
        md(_DEFINE_06_X),
        colab_code("🎯 Eure Zielweite", _DEFINE_07_TITLE_EURE_ZIELWEITE),
        md(_DEFINE_08_TESTW_RFE_DURCHF_HREN),
        colab_code("📝 Testwürfe eingeben", _DEFINE_09_TITLE_TESTW_RFE_EINGEBEN),
        colab_code("📊 Testwurf-Auswertung", _DEFINE_10_TITLE_TESTWURF_AUSWERTUNG),
        colab_code("📏 Messmodus wählen", _DEFINE_11_TITLE_MESSMODUS_W_HLEN),
        md(_DEFINE_12_PROJEKTCHARTER),
        colab_code("📝 Projektcharter ausfüllen", _DEFINE_13_TITLE_PROJEKTCHARTER_AUSF_LLEN),
        md(_DEFINE_14_DIV_STYLE_PADDING_10PX_BORDER),
        md(_DEFINE_15_DETAILS_STYLE_MARGIN_10PX_0_PA),
        phase_export_cell("DEFINE"),
    ]
