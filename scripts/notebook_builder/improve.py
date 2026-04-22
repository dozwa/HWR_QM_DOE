"""DMAIC phase IMPROVE: cells 55..71 of the notebook."""
from __future__ import annotations

from .cells import code, colab_code, md, phase_export_cell

_IMPROVE_55_X = r"""---
# Phase 4: IMPROVE

## Optimale Einstellung finden und bestätigen

Das Modell sagt euch, welche Faktoren wichtig sind. Jetzt nutzt ihr es, um die **optimale Einstellung** für eure Zielweite zu finden.

**Schritt 1:** Antwortfläche visuell erkunden (Konturplots)
**Schritt 2:** Optimale Einstellungen analytisch berechnen
**Schritt 3:** Konfirmation — Würfe an der neuen Einstellung"""

_IMPROVE_55_SCHRITT1 = r"""## Schritt 1 – Antwortfläche visuell erkunden"""

_IMPROVE_55_SCHRITT2 = r"""## Schritt 2 – Optimale Einstellungen berechnen"""

_IMPROVE_55_SCHRITT3 = r"""## Schritt 3 – Konfirmation"""

_IMPROVE_56_TITLE_KONTURPLOT_WURFWEITE = r"""faktor_x = 1 #@param {type:"integer"}
faktor_y = 2 #@param {type:"integer"}

# Faktornamen anzeigen zur Orientierung
_fn = projekt.modell._faktor_namen
print("Verfügbare Faktoren:")
for i, name in enumerate(_fn, 1):
    print(f"  {i}: {name}")
print(f"\nGewählt: X-Achse = {_fn[faktor_x - 1]}, Y-Achse = {_fn[faktor_y - 1]}")

if faktor_x == faktor_y:
    print("❌ Bitte zwei verschiedene Faktoren wählen!")
else:
    _fak = helper._effektive_faktoren(projekt)
    fig = helper.plot_kontur(
        projekt.modell, _fak, projekt.zielweite,
        faktor_idx=(faktor_x - 1, faktor_y - 1)
    )
    helper._save_fig(projekt, fig, f"improve_kontur_{_fn[faktor_x-1]}_{_fn[faktor_y-1]}")
    plt.show()
"""

_IMPROVE_57_VARIANZ_KONTURPLOTS = r"""### 📊 Varianz-Konturplots

Der Konturplot oben zeigt, **wo** ihr eure Zielweite trefft (Mittelwert). Aber wo ist der Prozess auch **konsistent**?

Zwei Ansätze:
- **Dispersionsmodell:** Modelliert die tatsächliche Streuung aus euren Wiederholungen
- **Fehlerfortpflanzung:** Berechnet analytisch, wo Einstellfehler die Wurfweite am wenigsten beeinflussen"""

_IMPROVE_58_TITLE_VARIANZ_KONTURPLOT_DISPE = r"""faktor_x = 1 #@param {type:"integer"}
faktor_y = 2 #@param {type:"integer"}

_fn = projekt.modell._faktor_namen
print(f"Faktoren: {', '.join(f'{i+1}={n}' for i, n in enumerate(_fn))}")
print(f"Gewählt: X={_fn[faktor_x-1]}, Y={_fn[faktor_y-1]}")

if faktor_x != faktor_y:
    _df = projekt.modell._daten
    _doe_X = _df[_fn].values
    _doe_response = _df['Y'].values

    fig = helper.plot_kontur_varianz_dispersion(
        projekt.modell, _doe_X, _doe_response,
        helper._effektive_faktoren(projekt),
        faktor_idx=(faktor_x - 1, faktor_y - 1)
    )
    helper._save_fig(projekt, fig, f"improve_varianz_disp_{_fn[faktor_x-1]}_{_fn[faktor_y-1]}")
    plt.show()
else:
    print("❌ Bitte zwei verschiedene Faktoren wählen!")
"""

_IMPROVE_59_TITLE_VARIANZ_KONTURPLOT_FEHLE = r"""faktor_x = 1 #@param {type:"integer"}
faktor_y = 2 #@param {type:"integer"}

_fn = projekt.modell._faktor_namen
print(f"Faktoren: {', '.join(f'{i+1}={n}' for i, n in enumerate(_fn))}")
print(f"Gewählt: X={_fn[faktor_x-1]}, Y={_fn[faktor_y-1]}")

if faktor_x != faktor_y:
    fig = helper.plot_kontur_varianz_transmitted(
        projekt.modell, helper._effektive_faktoren(projekt),
        faktor_idx=(faktor_x - 1, faktor_y - 1)
    )
    helper._save_fig(projekt, fig, f"improve_varianz_trans_{_fn[faktor_x-1]}_{_fn[faktor_y-1]}")
    plt.show()
else:
    print("❌ Bitte zwei verschiedene Faktoren wählen!")
"""

_IMPROVE_60_TITLE_OPTIMALE_EINSTELLUNGEN_B = r"""# Strategie wählen: "mittelwert", "varianz", oder "dual"
strategie = "dual" #@param ["mittelwert", "varianz", "dual"] {type:"string"}
lambda_gewicht = 0.01 #@param {type:"number"}

projekt.optimale_einstellung = helper.optimiere_einstellungen(
    projekt.modell, projekt.zielweite, helper._effektive_faktoren(projekt),
    strategie=strategie, lambda_gewicht=lambda_gewicht
)
helper.zeige_optimierung(projekt.optimale_einstellung)

projekt.optimierung_config = {
    "strategie": strategie,
    "lambda_gewicht": lambda_gewicht,
}
helper.speichere_fortschritt(projekt)
"""

_IMPROVE_61_TITLE_REGRESSIONSFORMEL_ANZEIG = r"""helper.zeige_regressionsformel(projekt.modell, helper._effektive_faktoren(projekt))"""

_IMPROVE_62_PROGNOSETOOL = r"""### 🔮 Prognosetool

Nutzt das Regressionsmodell, um für **beliebige Faktoreinstellungen** die Wurfweite vorherzusagen.

- Tragt eure gewünschten Werte als **kodierte Werte** ein (−1 = niedrig, +1 = hoch, 0 = Mitte)
- Oder als **Originalwerte** (z.B. Winkel in Grad) — die Kodierung erfolgt automatisch
- Die Kodierungstabelle findet ihr in der Formel-Anzeige oben"""

_IMPROVE_63_TITLE_PROGNOSETOOL_WURFWEITE_V = r"""import ipywidgets as _widgets

# Slider pro Faktor erzeugen
_faktor_sliders = []
_slider_box = []
_fak_prog = helper._effektive_faktoren(projekt)
for _f in _fak_prog:
    _center = (_f["low"] + _f["high"]) / 2
    _half = (_f["high"] - _f["low"]) / 2
    _slider = _widgets.FloatSlider(
        value=_center,
        min=_f["low"] - 0.2 * _half,
        max=_f["high"] + 0.2 * _half,
        step=round(_half / 10, 2) or 0.1,
        description=f'{_f["name"]} ({_f["einheit"]})',
        style={'description_width': '180px'},
        layout=_widgets.Layout(width='500px'),
        readout_format='.1f',
    )
    # Kodierter Wert als Label daneben
    _code_label = _widgets.Label(value="(kodiert: 0.00)")
    def _update_label(change, lbl=_code_label, f=_f):
        c = (f["low"] + f["high"]) / 2
        h = (f["high"] - f["low"]) / 2
        coded = (change["new"] - c) / h if h > 0 else 0
        lbl.value = f"(kodiert: {coded:+.2f})"
    _slider.observe(_update_label, names='value')
    _faktor_sliders.append(_slider)
    _slider_box.append(_widgets.HBox([_slider, _code_label]))

_output = _widgets.Output()

def _berechne_prognose(_btn):
    _output.clear_output()
    with _output:
        werte = {}
        for _f, _s in zip(_fak_prog, _faktor_sliders):
            werte[_f["name"]] = _s.value
        ergebnis = helper.prognostiziere(
            projekt.modell, _fak_prog, werte
        )
        helper.zeige_prognose(
            ergebnis, _fak_prog, werte,
            zielweite=projekt.zielweite
        )

_btn = _widgets.Button(
    description='🔮 Prognose berechnen',
    button_style='primary',
    layout=_widgets.Layout(width='250px', height='38px', margin='12px 0')
)
_btn.on_click(_berechne_prognose)

# Anzeige
display(HTML(f"<p><strong>Zielweite:</strong> {projekt.zielweite:.0f} cm ± {projekt.toleranz:.0f} cm</p>"))
for _box in _slider_box:
    display(_box)
display(_btn)
display(_output)

# Initial-Prognose auslösen
_berechne_prognose(None)"""

_IMPROVE_64_VISUELLE_VS_ANALYTISCHE_OPTIMI = r"""### 🔍 Visuelle vs. analytische Optimierung

- **Visuell (Konturplot):** Ihr seht die Antwortfläche und könnt das Optimum abschätzen. Gut für Verständnis, aber ungenau.
- **Analytisch (Optimizer):** Der Computer findet die exakte Einstellung. Genauer, aber als Black Box weniger lehrreich.

➡️ Der Konturplot hilft euch zu **verstehen**, der Optimizer liefert die **exakten Werte**."""

_IMPROVE_65_TITLE_VERGLEICH_ALLE_DREI_OPTI = r"""ergebnisse = helper.vergleiche_optimierungen(
    projekt.modell, projekt.zielweite, projekt.faktoren
)
# Das Ergebnis der gewählten Strategie wird für die Konfirmation verwendet.
print(f"\n→ Für die Konfirmation wird die Strategie '{strategie}' verwendet.")"""

_IMPROVE_66_ROBUSTHEITSHINWEIS_SUCHT_EINST = r"""> 💡 **Robustheitshinweis:** Sucht Einstellungen, bei denen nicht nur der Mittelwert stimmt, sondern auch die **Streuung zwischen euren Wiederholungen klein war**. Prüft die Standardabweichungen in eurer Versuchstabelle – eine Kombination mit kleiner Streuung ist robuster als eine mit hohem Mittelwert aber großer Schwankung."""

_IMPROVE_67_DETAILS_STYLE_MARGIN_10PX_0_PA = r"""<details style="margin:10px 0; padding:8px; background:#F9FAFB; border:1px solid #E5E7EB; border-radius:6px;">
<summary style="cursor:pointer; font-weight:bold; color:#2563EB;">
🔍 Für Neugierige: Konfidenzintervall vs. Vorhersageintervall
</summary>
<div style="margin-top:8px; padding:8px; font-size:0.95em;">

Das Notebook zeigt ein **Vorhersageintervall** (prediction interval), nicht ein Konfidenzintervall:

- **Konfidenzintervall:** Wo liegt der wahre Mittelwert? (schmaler)
- **Vorhersageintervall:** Wo wird der nächste einzelne Wurf landen? (breiter)

Das Vorhersageintervall ist breiter, weil es nicht nur die Unsicherheit über den Mittelwert, sondern auch die **natürliche Streuung** des Prozesses berücksichtigt.

Für die Konfirmation ist das Vorhersageintervall das richtige Maß: Ihr wollt wissen, ob eure **einzelnen Würfe** in den erwarteten Bereich fallen.
</div>
</details>"""

_IMPROVE_68_TITLE_KONFIRMATIONS_TEMPLATE_H = r"""filepath = helper.erstelle_konfirmation_excel(
    projekt.optimale_einstellung["einstellungen"],
    projekt.zielweite,
)
print(f"✅ Template erstellt!")

# Kopie im Drive-Ordner ablegen
import shutil
_save_dir = helper._fortschritt_verzeichnis(projekt)
if _save_dir:
    os.makedirs(_save_dir, exist_ok=True)
    shutil.copy2(filepath, os.path.join(_save_dir, os.path.basename(filepath)))

try:
    from google.colab import files
    files.download(filepath)
except ImportError:
    print(f"📁 Datei: {filepath}")

print(f"\n🎯 Stellt euer Katapult auf die empfohlenen Einstellungen und macht mindestens 10 Würfe!")
"""

_IMPROVE_69_TITLE_KONFIRMATIONSW_RFE_EINGE = r"""wurf_01 = 0.0 #@param {type:"number"}
wurf_02 = 0.0 #@param {type:"number"}
wurf_03 = 0.0 #@param {type:"number"}
wurf_04 = 0.0 #@param {type:"number"}
wurf_05 = 0.0 #@param {type:"number"}
wurf_06 = 0.0 #@param {type:"number"}
wurf_07 = 0.0 #@param {type:"number"}
wurf_08 = 0.0 #@param {type:"number"}
wurf_09 = 0.0 #@param {type:"number"}
wurf_10 = 0.0 #@param {type:"number"}
wurf_11 = 0.0 #@param {type:"number"}
wurf_12 = 0.0 #@param {type:"number"}
wurf_13 = 0.0 #@param {type:"number"}
wurf_14 = 0.0 #@param {type:"number"}
wurf_15 = 0.0 #@param {type:"number"}
wurf_16 = 0.0 #@param {type:"number"}
wurf_17 = 0.0 #@param {type:"number"}
wurf_18 = 0.0 #@param {type:"number"}
wurf_19 = 0.0 #@param {type:"number"}
wurf_20 = 0.0 #@param {type:"number"}

_alle = [
         wurf_01, wurf_02, wurf_03, wurf_04, wurf_05,
         wurf_06, wurf_07, wurf_08, wurf_09, wurf_10,
         wurf_11, wurf_12, wurf_13, wurf_14, wurf_15,
         wurf_16, wurf_17, wurf_18, wurf_19, wurf_20]
werte = [round(w, 1) for w in _alle if w > 0]

if werte:
    projekt.konfirmation_wuerfe = np.array(werte)
    print(f"✅ {len(werte)} Konfirmationswürfe eingetragen")
    print(f"   Werte: {werte}")
    helper.speichere_fortschritt(projekt)
elif len(projekt.konfirmation_wuerfe) > 0:
    print(f"ℹ️ {len(projekt.konfirmation_wuerfe)} Konfirmationswürfe aus gespeichertem Fortschritt geladen.")
else:
    print("⚠️ Bitte Konfirmationswürfe eingeben (mindestens 10 Würfe empfohlen).")"""

_IMPROVE_70_TITLE_KONFIRMATION_AUSWERTEN = r"""if len(projekt.konfirmation_wuerfe) > 0:
    opt = projekt.optimale_einstellung
    ergebnis = helper.analysiere_konfirmation(
        projekt.konfirmation_wuerfe,
        opt["vorhersage"], opt["pi_low"], opt["pi_high"],
        projekt.zielweite, projekt.toleranz
    )
    projekt.konfirmation_ergebnis = ergebnis
    helper.zeige_konfirmation(ergebnis)

    # Zielscheibe (1D: nur Weite)
    fig = helper.plot_zielscheibe(
        projekt.konfirmation_wuerfe, projekt.zielweite, projekt.toleranz,
        modus="1D", titel="Konfirmation – Optimierte Einstellung",
        farbe=helper.GREEN
    )
    helper._save_fig(projekt, fig, "improve_konfirmation_zielscheibe")
    plt.show()

    helper.hinweis_bericht("Die Konfirmation ist der zentrale Nachweis, dass der DMAIC-Zyklus funktioniert hat. Zielscheiben-Plot und Bewertung gehören in den Bericht.")"""


def cells():
    return [
        md(_IMPROVE_55_X),
        md(_IMPROVE_55_SCHRITT1),
        colab_code("📊 Konturplot (Wurfweite)", _IMPROVE_56_TITLE_KONTURPLOT_WURFWEITE),
        md(_IMPROVE_57_VARIANZ_KONTURPLOTS),
        colab_code("📊 Varianz-Konturplot: Dispersionsmodell", _IMPROVE_58_TITLE_VARIANZ_KONTURPLOT_DISPE),
        colab_code("📊 Varianz-Konturplot: Fehlerfortpflanzung", _IMPROVE_59_TITLE_VARIANZ_KONTURPLOT_FEHLE),
        md(_IMPROVE_55_SCHRITT2),
        colab_code("🎯 Optimale Einstellungen berechnen", _IMPROVE_60_TITLE_OPTIMALE_EINSTELLUNGEN_B),
        colab_code("📐 Regressionsformel anzeigen", _IMPROVE_61_TITLE_REGRESSIONSFORMEL_ANZEIG),
        md(_IMPROVE_62_PROGNOSETOOL),
        colab_code("🔮 Prognosetool: Wurfweite vorhersagen", _IMPROVE_63_TITLE_PROGNOSETOOL_WURFWEITE_V),
        md(_IMPROVE_64_VISUELLE_VS_ANALYTISCHE_OPTIMI),
        colab_code("📊 Vergleich: Alle drei Optimierungsstrategien", _IMPROVE_65_TITLE_VERGLEICH_ALLE_DREI_OPTI),
        md(_IMPROVE_66_ROBUSTHEITSHINWEIS_SUCHT_EINST),
        md(_IMPROVE_67_DETAILS_STYLE_MARGIN_10PX_0_PA),
        md(_IMPROVE_55_SCHRITT3),
        colab_code("📥 Konfirmations-Template herunterladen", _IMPROVE_68_TITLE_KONFIRMATIONS_TEMPLATE_H),
        colab_code("📝 Konfirmationswürfe eingeben", _IMPROVE_69_TITLE_KONFIRMATIONSW_RFE_EINGE),
        colab_code("📊 Konfirmation auswerten", _IMPROVE_70_TITLE_KONFIRMATION_AUSWERTEN),
        phase_export_cell("IMPROVE"),
    ]
