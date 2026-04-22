"""DMAIC phase DEFINE: cells 5..16 of the notebook."""
from __future__ import annotations

from .cells import code, colab_code, factor_def_cell, md, phase_export_cell

# Faktor-Definitionen: einfach editieren. Werden als 5 gleichartige Zellen
# emittiert (3 Pflicht + 2 optional). Diese Definition ist bewusst früh in
# DEFINE, damit Min/Max-Vermessung und die Projektbeschreibung schon auf die
# benannten Faktoren zugreifen können.
FACTOR_DEFAULTS = [
    dict(n=1, name="Winkel",       unit="Grad", low=30.0, high=45.0),
    dict(n=2, name="Spannung",     unit="cm",   low=5.0,  high=15.0),
    dict(n=3, name="Ballposition", unit="cm",   low=0.0,  high=5.0),
    dict(n=4, optional=True,
         optional_hint="optional – leer lassen wenn nur 3 Faktoren",
         no_factor_msg="Kein 4. Faktor definiert (3-Faktor-Design)"),
    dict(n=5, optional=True,
         optional_hint="optional – leer lassen wenn nur 3–4 Faktoren",
         no_factor_msg="Kein 5. Faktor definiert"),
]

_DEFINE_05_TITLE_PROJEKT_EINRICHTEN = r"""modus = "Neues Projekt" #@param ["Neues Projekt", "Fortschritt laden"]
gruppenname = "Gruppe A" #@param {type:"string"}
gruppennummer = 1 #@param {type:"integer"}
zielweite = 300.0 #@param {type:"number"}
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
        projekt = helper.init_projekt(gruppenname, gruppennummer,
                                      zielweite=zielweite, toleranz=toleranz)
        helper.speichere_fortschritt(projekt)
        print(f"✅ Neues Projekt initialisiert!")
        print(f"   Gruppe: {projekt.gruppenname} (Nr. {projekt.gruppennummer})")
        print(f"   Zielweite: {projekt.zielweite:.0f} cm ± {projekt.toleranz:.0f} cm")
        print(f"   (Zielweite könnt ihr nach der Katapult-Vermessung unten noch anpassen.)")
"""

_DEFINE_06_X = r"""---
# Phase 1: DEFINE

## Was ist das Problem? Wie schlecht ist es?

In der Define-Phase beschreibt ihr zunächst euer Katapult, vermesst dessen Reichweite und legt dann die **Zielweite** fest, die ihr reproduzierbar treffen wollt."""

_DEFINE_06A_FAKTOREN_INTRO = r"""## Schritt 1 – Faktoren definieren

Welche **Einstellungen** lassen sich an eurem Katapult verändern? Tragt Name, Einheit und die beiden geplanten Stufen (Low/High) ein. Diese Faktoren werden gleich für die Min/Max-Vermessung und später in ANALYZE für den Versuchsplan verwendet.

> **Stetig vs. zweistufig:** Wenn der Faktor frei einstellbar ist (z.B. Winkel, Spannung), `centerpoint` auf `True` lassen. Bei nur zwei festen Positionen auf `False`."""

_DEFINE_06B_FAKTOREN_UEBERNEHMEN = r"""_faktoren_neu = [faktor1, faktor2, faktor3]
try:
    if faktor4 and faktor4.get("name"):
        _faktoren_neu.append(faktor4)
except NameError:
    pass
try:
    if faktor5 and faktor5.get("name"):
        _faktoren_neu.append(faktor5)
except NameError:
    pass

for f in _faktoren_neu:
    if f["low"] == f["high"]:
        print(f"❌ Faktor '{f['name']}': Low und High sind identisch ({f['low']}). Bitte unterschiedliche Stufen wählen!")
    elif f["low"] > f["high"]:
        f["low"], f["high"] = f["high"], f["low"]
        print(f"⚠️ Faktor '{f['name']}': Low und High vertauscht – automatisch korrigiert.")

projekt.faktoren = _faktoren_neu
print(f"✅ {len(projekt.faktoren)} Faktoren übernommen:")
for f in projekt.faktoren:
    print(f"   • {f['name']} [{f['low']} – {f['high']} {f['einheit']}]")
helper.speichere_fortschritt(projekt)
"""

_DEFINE_06C_VERMESSUNG_INTRO = r"""## Schritt 2 – Katapult vermessen

Ermittelt die **Reichweiten-Spanne** eures Katapults:

1. Stellt alle Faktoren so ein, dass ihr die **kürzeste** Wurfweite erwartet. Werft 3× und tragt die Werte unten ein.
2. Stellt alle Faktoren so ein, dass ihr die **längste** Wurfweite erwartet. Werft 3× und tragt ein.
3. Die tatsächlich eingestellten Werte pro Faktor werden ebenfalls dokumentiert – vorbelegt sind Low (Min) bzw. High (Max) eurer Faktordefinition.

> So erhaltet ihr einen realistischen Bereich, in dem ihr gleich die Zielweite festlegt, und eine reproduzierbare Katapult-Beschreibung."""

_DEFINE_06D_VERMESSUNG_MIN_EINSTELLUNG = r'''helper.zeige_faktoren_legende(projekt)

# Min-Einstellung: tatsächlich eingestellte Werte pro Faktor (Reihenfolge 1..5).
# Lass 0 stehen, wenn du den Low-Wert aus der Faktor-Definition übernehmen willst.
min_val_1 = 0.0 #@param {type:"number"}
min_val_2 = 0.0 #@param {type:"number"}
min_val_3 = 0.0 #@param {type:"number"}
min_val_4 = 0.0 #@param {type:"number"}
min_val_5 = 0.0 #@param {type:"number"}

_min_vals_raw = [min_val_1, min_val_2, min_val_3, min_val_4, min_val_5]
_min_einst = {}
for i, f in enumerate(projekt.faktoren):
    _min_einst[f["name"]] = _min_vals_raw[i] if _min_vals_raw[i] != 0.0 else f["low"]

helper.speichere_vermessung(
    projekt,
    min_wuerfe=[],
    max_wuerfe=[],
    min_einstellung=_min_einst,
    max_einstellung={},
)
print("✅ Min-Einstellung gespeichert:")
for name, val in _min_einst.items():
    einheit = next((f["einheit"] for f in projekt.faktoren if f["name"] == name), "")
    print(f"   • {name}: {val} {einheit}")
'''

_DEFINE_06D_VERMESSUNG_MIN_WUERFE = r'''if projekt.vermessung_min_einstellung:
    print("Aktuelle Min-Einstellung:")
    for name, val in projekt.vermessung_min_einstellung.items():
        einheit = next((f["einheit"] for f in projekt.faktoren if f["name"] == name), "")
        print(f"   • {name}: {val} {einheit}")

# Drei Würfe mit der oben eingestellten Min-Konfiguration.
min_wurf_1 = 0.0 #@param {type:"number"}
min_wurf_2 = 0.0 #@param {type:"number"}
min_wurf_3 = 0.0 #@param {type:"number"}

_min_wuerfe = [min_wurf_1, min_wurf_2, min_wurf_3]
if any(w > 0 for w in _min_wuerfe):
    helper.speichere_vermessung(
        projekt,
        min_wuerfe=_min_wuerfe,
        max_wuerfe=[],
        min_einstellung={},
        max_einstellung={},
    )
    print(f"✅ Min-Würfe gespeichert (μ={float(np.mean(projekt.vermessung_min_wuerfe)):.1f} cm).")
elif len(projekt.vermessung_min_wuerfe) > 0:
    print(f"ℹ️ Min-Würfe aus gespeichertem Fortschritt: μ={float(np.mean(projekt.vermessung_min_wuerfe)):.1f} cm.")
else:
    print("⚠️ Bitte mindestens einen Min-Wurf eingeben (Wert > 0).")
'''

_DEFINE_06E_VERMESSUNG_MAX_EINSTELLUNG = r'''helper.zeige_faktoren_legende(projekt)

# Max-Einstellung: tatsächlich eingestellte Werte pro Faktor (Reihenfolge 1..5).
# Lass 0 stehen, wenn du den High-Wert aus der Faktor-Definition übernehmen willst.
max_val_1 = 0.0 #@param {type:"number"}
max_val_2 = 0.0 #@param {type:"number"}
max_val_3 = 0.0 #@param {type:"number"}
max_val_4 = 0.0 #@param {type:"number"}
max_val_5 = 0.0 #@param {type:"number"}

beschreibung = "" #@param {type:"string"}

_max_vals_raw = [max_val_1, max_val_2, max_val_3, max_val_4, max_val_5]
_max_einst = {}
for i, f in enumerate(projekt.faktoren):
    _max_einst[f["name"]] = _max_vals_raw[i] if _max_vals_raw[i] != 0.0 else f["high"]

helper.speichere_vermessung(
    projekt,
    min_wuerfe=[],
    max_wuerfe=[],
    min_einstellung={},
    max_einstellung=_max_einst,
    beschreibung=beschreibung,
)
print("✅ Max-Einstellung gespeichert:")
for name, val in _max_einst.items():
    einheit = next((f["einheit"] for f in projekt.faktoren if f["name"] == name), "")
    print(f"   • {name}: {val} {einheit}")
'''

_DEFINE_06E_VERMESSUNG_MAX_WUERFE = r'''if projekt.vermessung_max_einstellung:
    print("Aktuelle Max-Einstellung:")
    for name, val in projekt.vermessung_max_einstellung.items():
        einheit = next((f["einheit"] for f in projekt.faktoren if f["name"] == name), "")
        print(f"   • {name}: {val} {einheit}")

# Drei Würfe mit der oben eingestellten Max-Konfiguration.
max_wurf_1 = 0.0 #@param {type:"number"}
max_wurf_2 = 0.0 #@param {type:"number"}
max_wurf_3 = 0.0 #@param {type:"number"}

_max_wuerfe = [max_wurf_1, max_wurf_2, max_wurf_3]
if any(w > 0 for w in _max_wuerfe):
    helper.speichere_vermessung(
        projekt,
        min_wuerfe=[],
        max_wuerfe=_max_wuerfe,
        min_einstellung={},
        max_einstellung={},
    )
    print(f"✅ Max-Würfe gespeichert (μ={float(np.mean(projekt.vermessung_max_wuerfe)):.1f} cm).")
elif len(projekt.vermessung_max_wuerfe) > 0:
    print(f"ℹ️ Max-Würfe aus gespeichertem Fortschritt: μ={float(np.mean(projekt.vermessung_max_wuerfe)):.1f} cm.")
else:
    print("⚠️ Bitte mindestens einen Max-Wurf eingeben (Wert > 0).")
'''

_DEFINE_06F_VERMESSUNG_ANZEIGE = r"""helper.zeige_vermessung(projekt)"""

_DEFINE_06G_ZIELWEITE_ANPASSEN = r"""## Schritt 3 – Zielweite festlegen

Eure Zielweite ist aktuell **300 cm** (Default aus der Projekt-Einrichtung). Passt den Wert an, wenn euer Katapult die 300 cm nicht stabil erreicht oder ihr eine andere Herausforderung wählen wollt. Die Toleranz (siehe oben) definiert zusammen mit der Zielweite die **Testgrenzen** (LSL/USL), auf die wir in der CONTROL-Phase (Cpk) zurückkommen."""

_DEFINE_06H_ZIELWEITE_SET = r"""neue_zielweite = 300.0 #@param {type:"number"}
helper.setze_zielweite(projekt, neue_zielweite)
"""

_DEFINE_06I_ANNAEHERUNG_INTRO = r"""## Schritt 4 – Zielweite manuell anpeilen (OFAT)

Bevor ihr die Messsystemanalyse macht, **sucht manuell** eine Einstellung, mit der ihr eure Zielweite erreicht.

**Spielregeln:**
1. Ändert pro Iteration **immer nur einen einzigen Faktor** ("one factor at a time" / OFAT).
2. Werft 3× und tragt die Werte ein.
3. Führt die Zelle erneut aus — das Notebook protokolliert die Iterationen und warnt, falls ihr mehr als einen Faktor verändert habt.
4. Wenn ihr zufrieden seid, übernehmt die Einstellung mit der Zelle „✅ Typische Einstellung übernehmen" — sie ist ab dann die Referenzeinstellung für Testwürfe und MEASURE-Baseline."""

_DEFINE_06J_ANNAEHERUNG_ITERATION = r'''helper.zeige_faktoren_legende(projekt)

# Aktuelle Einstellung für diese Iteration (ein Wert pro Faktor, Reihenfolge 1..5).
# Ändert pro Durchlauf bitte nur *einen* Faktor gegenüber der letzten Iteration.
cur_val_1 = 0.0 #@param {type:"number"}
cur_val_2 = 0.0 #@param {type:"number"}
cur_val_3 = 0.0 #@param {type:"number"}
cur_val_4 = 0.0 #@param {type:"number"}
cur_val_5 = 0.0 #@param {type:"number"}

# Drei Würfe mit dieser Einstellung.
iter_wurf_1 = 0.0 #@param {type:"number"}
iter_wurf_2 = 0.0 #@param {type:"number"}
iter_wurf_3 = 0.0 #@param {type:"number"}

_cur_vals = [cur_val_1, cur_val_2, cur_val_3, cur_val_4, cur_val_5]
_einstellung = {f["name"]: _cur_vals[i]
                for i, f in enumerate(projekt.faktoren)
                if _cur_vals[i] != 0.0}
_iter_wuerfe = [iter_wurf_1, iter_wurf_2, iter_wurf_3]

if _einstellung and any(w > 0 for w in _iter_wuerfe):
    helper.protokolliere_annaeherung(projekt, _einstellung, _iter_wuerfe)
elif projekt.annaeherung_log:
    print(f"ℹ️ Bisher {len(projekt.annaeherung_log)} Iteration(en) protokolliert.")
    for eintrag in projekt.annaeherung_log[-3:]:
        abw = eintrag["abweichung_vom_ziel"]
        print(f"   Iter {eintrag['iteration']}: μ={eintrag['mean']:.1f} cm "
              f"(Abweichung: {abw:+.1f} cm)")
else:
    print("⚠️ Bitte alle Faktor-Einstellungen und mindestens einen Wurf (>0) eingeben.")
'''

_DEFINE_06K_TYPISCHE_EINSTELLUNG = r"""# Übernimmt die letzte Annäherungs-Iteration als 'typische Einstellung'.
# Diese Einstellung wird später für die Testwürfe und die MEASURE-Baseline
# verwendet, damit die Daten konsistent auf denselben Parametern basieren.
helper.setze_typische_einstellung(projekt)
"""

_DEFINE_07_TITLE_EURE_ZIELWEITE = r"""print(f"{'='*50}")
print(f"  EURE AUFGABE")
print(f"  Zielweite:  {projekt.zielweite:.0f} cm")
print(f"  Toleranz:   ±{projekt.toleranz:.0f} cm")
print(f"  Zielband:   [{projekt.zielweite - projekt.toleranz:.0f}, {projekt.zielweite + projekt.toleranz:.0f}] cm")
print(f"{'='*50}")
if projekt.typische_einstellung:
    print(f"\nTypische Einstellung aus der Annäherung:")
    for name, val in projekt.typische_einstellung.items():
        einheit = next((f["einheit"] for f in projekt.faktoren if f["name"] == name), "")
        print(f"   • {name}: {val} {einheit}")
    print(f"\nMacht 5 Testwürfe mit genau dieser Einstellung.")
else:
    print(f"\nStellt euer Katapult auf eure typische Einstellung und macht 5 Testwürfe.")"""

_DEFINE_08_TESTW_RFE_DURCHF_HREN = r"""## Schritt 5 – Testwürfe (CV bei typischer Einstellung)

1. Katapult auf die oben angezeigte **typische Einstellung** bringen (aus der Annäherung).
2. **5 Würfe** – Wurfweite messen (1D, nur die Weite).
3. Werte unten eintragen.

> ⚠️ Noch keine Optimierung – es geht hier um die Reproduzierbarkeit eurer typischen Einstellung."""

_DEFINE_09_TITLE_TESTW_RFE_EINGEBEN = r"""if projekt.typische_einstellung:
    print("Werft mit dieser Einstellung:")
    for name, val in projekt.typische_einstellung.items():
        einheit = next((f["einheit"] for f in projekt.faktoren if f["name"] == name), "")
        print(f"   • {name}: {val} {einheit}")
else:
    print("⚠️ Hinweis: Noch keine typische Einstellung festgelegt — führt Schritt 4 zuerst aus.")

wurf_1 = 0.0 #@param {type:"number"}
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

_DEFINE_12_PROJEKTCHARTER = r"""### Projektcharter

Die Projektcharter dokumentiert euer Qualitätsverbesserungsprojekt. Sie ist das zentrale Dokument der Define-Phase.

> 📋 **Für den Bericht:** Die Projektcharter gehört vollständig in euren Bericht!"""

_DEFINE_13_TITLE_PROJEKTCHARTER_AUSF_LLEN = r"""problemstellung = "z.B. Katapult trifft die Zielweite nicht reproduzierbar" #@param {type:"string"}
projektziel = "z.B. Wurfweite 450 cm ± 15 cm mit Cpk > 1.0" #@param {type:"string"}
scope = "z.B. Optimierung der Faktoreinstellungen, nicht der Konstruktion" #@param {type:"string"}
projektleiter = "" #@param {type:"string"}
protokollant = "" #@param {type:"string"}
versuchsdurchfuehrende = "" #@param {type:"string"}

_einst_str = ", ".join(f"{n}={v}" for n, v in projekt.typische_einstellung.items()) \
             if projekt.typische_einstellung else "(noch nicht festgelegt)"
projekt.charter = {
    "Gruppenname": projekt.gruppenname,
    "Zielweite": f"{projekt.zielweite:.0f} cm ± {projekt.toleranz:.0f} cm",
    "Typische Einstellung": _einst_str,
    "Problemstellung": problemstellung,
    "Projektziel": projektziel,
    "Scope": scope,
    "Projektleiter/in": projektleiter,
    "Protokollant/in": protokollant,
    "Versuchsdurchführende": versuchsdurchfuehrende,
    "Datum": "22.04.2026",
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
        md(_DEFINE_06A_FAKTOREN_INTRO),
        *[factor_def_cell(**kw) for kw in FACTOR_DEFAULTS],
        colab_code("📋 Faktoren übernehmen", _DEFINE_06B_FAKTOREN_UEBERNEHMEN),
        md(_DEFINE_06C_VERMESSUNG_INTRO),
        colab_code("🎯 Min-Einstellung eintragen", _DEFINE_06D_VERMESSUNG_MIN_EINSTELLUNG),
        colab_code("🎯 Min-Würfe eintragen", _DEFINE_06D_VERMESSUNG_MIN_WUERFE),
        colab_code("🎯 Max-Einstellung eintragen", _DEFINE_06E_VERMESSUNG_MAX_EINSTELLUNG),
        colab_code("🎯 Max-Würfe eintragen", _DEFINE_06E_VERMESSUNG_MAX_WUERFE),
        colab_code("📊 Vermessung anzeigen", _DEFINE_06F_VERMESSUNG_ANZEIGE),
        md(_DEFINE_06G_ZIELWEITE_ANPASSEN),
        colab_code("🎯 Zielweite anpassen (optional)", _DEFINE_06H_ZIELWEITE_SET),
        md(_DEFINE_06I_ANNAEHERUNG_INTRO),
        colab_code("🎯 Annäherung: Einstellung + 3 Würfe", _DEFINE_06J_ANNAEHERUNG_ITERATION),
        colab_code("✅ Typische Einstellung übernehmen", _DEFINE_06K_TYPISCHE_EINSTELLUNG),
        colab_code("🎯 Eure Zielweite", _DEFINE_07_TITLE_EURE_ZIELWEITE),
        md(_DEFINE_08_TESTW_RFE_DURCHF_HREN),
        colab_code("📝 Testwürfe eingeben", _DEFINE_09_TITLE_TESTW_RFE_EINGEBEN),
        colab_code("📊 Testwurf-Auswertung", _DEFINE_10_TITLE_TESTWURF_AUSWERTUNG),
        md(_DEFINE_12_PROJEKTCHARTER),
        colab_code("📝 Projektcharter ausfüllen", _DEFINE_13_TITLE_PROJEKTCHARTER_AUSF_LLEN),
        md(_DEFINE_14_DIV_STYLE_PADDING_10PX_BORDER),
        md(_DEFINE_15_DETAILS_STYLE_MARGIN_10PX_0_PA),
        phase_export_cell("DEFINE"),
    ]
