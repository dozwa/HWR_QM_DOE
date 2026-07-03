"""DMAIC phase CONTROL: cells 72..80 of the notebook."""
from __future__ import annotations

from .cells import code, colab_code, md, phase_export_cell

_CONTROL_72_X = r"""---
# Phase 5: CONTROL

## Ist der verbesserte Prozess stabil und fähig?

In der Control-Phase prüft ihr:
1. **Stabilität:** Läuft der Prozess gleichmäßig? (I-MR-Kontrollkarte)
2. **Normalverteilung:** Voraussetzung für Cpk (Shapiro-Wilk + Q-Q-Plot)
3. **Prozessfähigkeit:** Passt der Prozess in die Spezifikation? (Cpk)
4. **Vorher/Nachher:** Visueller Verbesserungsvergleich"""

_CONTROL_73_TITLE_I_MR_KONTROLLKARTE_STABI = r"""if len(projekt.konfirmation_wuerfe) > 0:
    projekt.imr_ergebnis = helper.berechne_imr(projekt.konfirmation_wuerfe)

    fig = helper.plot_imr(projekt.konfirmation_wuerfe)
    helper._save_fig(projekt, fig, "control_imr")
    plt.show()

    helper.zeige_stabilitaet(projekt.imr_ergebnis)
else:
    print("⚠️ Keine Konfirmationswürfe vorhanden — führt in IMPROVE die")
    print("   Konfirmation durch (Excel-Upload oder manuelle Eingabe) und")
    print("   führt diese Zelle danach erneut aus.")"""

_CONTROL_74_DETAILS_STYLE_MARGIN_10PX_0_PA = r"""<details style="margin:10px 0; padding:8px; background:#F9FAFB; border:1px solid #E5E7EB; border-radius:6px;">
<summary style="cursor:pointer; font-weight:bold; color:#2563EB;">
🔍 Für Neugierige: Wie werden die Kontrollgrenzen berechnet?
</summary>
<div style="margin-top:8px; padding:8px; font-size:0.95em;">

Die I-MR-Kontrollkarte berechnet die Grenzen aus den **Moving Ranges** (gleitende Spannweiten):

$$MR_i = |x_i - x_{i-1}|$$
$$\overline{MR} = \frac{1}{n-1}\sum MR_i$$

**I-Chart (Einzelwerte):**
$$UCL = \bar{x} + 2{,}66 \cdot \overline{MR}$$
$$LCL = \bar{x} - 2{,}66 \cdot \overline{MR}$$

**MR-Chart:**
$$UCL_{MR} = 3{,}267 \cdot \overline{MR}$$

Der Faktor 2,66 kommt aus d₂ = 1,128 für n=2 (Moving-Range-Subgruppe): 3/d₂ ≈ 2,66.
</div>
</details>"""

_CONTROL_75_TITLE_NORMALVERTEILUNGSPR_FUNG = r"""if len(projekt.konfirmation_wuerfe) > 0:
    norm_test = helper.pruefe_normalverteilung(projekt.konfirmation_wuerfe)

    fig = helper.plot_qq(projekt.konfirmation_wuerfe)
    helper._save_fig(projekt, fig, "control_qq")
    plt.show()

    if not np.isnan(norm_test['shapiro_p']):
        shapiro_schwellen = [
            (0.05, "⚠️", "Cpk mit Vorsicht interpretieren – Daten möglicherweise nicht normalverteilt"),
            (float('inf'), "✅", "Normalverteilungsannahme beibehalten – Cpk ist aussagekräftig"),
        ]
        helper.zeige_ampel(norm_test['shapiro_p'], shapiro_schwellen,
                          titel="Shapiro-Wilk p-Wert:")
else:
    print("⚠️ Keine Konfirmationswürfe vorhanden — erst IMPROVE-Konfirmation, dann hier weiter.")"""

_CONTROL_76_PROZESSF_HIGKEIT_CPK_WAS_BEDEU = r"""### Prozessfähigkeit (Cpk) – Was bedeutet das?

Der **Cpk** misst, ob euer Prozess dauerhaft in die Spezifikation passt:

$$C_{pk} = \min\left(\frac{USL - \bar{x}}{3\sigma},\; \frac{\bar{x} - LSL}{3\sigma}\right)$$

- **USL** (Upper Specification Limit) = Zielweite + Toleranz
- **LSL** (Lower Specification Limit) = Zielweite − Toleranz

| Cpk | Industrie | Euer Katapult |
|-----|-----------|---------------|
| < 0,67 | ❌ Nicht fähig | ❌ Verbesserungsbedürftig |
| 0,67–1,0 | ❌ Nicht fähig | ⚠️ Verbesserung ggü. Baseline |
| 1,0–1,33 | ⚠️ Bedingt fähig | ✅ Gut für ein Experiment |
| ≥ 1,33 | ✅ Prozess fähig | ✅ Hervorragend |

> *In der Industrie wird Cpk ≥ 1,33 gefordert. Für euer selbstgebautes Katapult ist ein Cpk > 0,67 bereits ein Erfolg.*"""

_CONTROL_77_TITLE_PROZESSF_HIGKEIT_CPK = r"""if len(projekt.konfirmation_wuerfe) > 0:
    usl = projekt.zielweite + projekt.toleranz
    lsl = projekt.zielweite - projekt.toleranz

    cpk = helper.berechne_cpk(projekt.konfirmation_wuerfe, usl, lsl)
    projekt.cpk_ergebnis = cpk
    helper.zeige_cpk(cpk)

    fig = helper.plot_cpk_verteilung(cpk)
    helper._save_fig(projekt, fig, "control_cpk_verteilung")
    plt.show()
else:
    print("⚠️ Keine Konfirmationswürfe vorhanden — erst IMPROVE-Konfirmation, dann hier weiter.")"""

_CONTROL_78_TITLE_VORHER_NACHHER_ZIELSCHEI = r"""if len(projekt.baseline_wuerfe) > 0 and len(projekt.konfirmation_wuerfe) > 0:
    fig = helper.plot_vorher_nachher(
        projekt.baseline_wuerfe, projekt.konfirmation_wuerfe,
        projekt.zielweite, projekt.toleranz, projekt.messmodus
    )
    helper._save_fig(projekt, fig, "control_vorher_nachher")
    plt.show()

    helper.hinweis_bericht("Cpk-Wert, Kontrollkarte und Vorher/Nachher-Zielscheibe sind die drei zentralen Control-Outputs.")
else:
    _fehlt = []
    if len(projekt.baseline_wuerfe) == 0:
        _fehlt.append("Baseline-Würfe (MEASURE)")
    if len(projekt.konfirmation_wuerfe) == 0:
        _fehlt.append("Konfirmationswürfe (IMPROVE)")
    print(f"⚠️ Für den Vorher/Nachher-Vergleich fehlen: {', '.join(_fehlt)}.")"""

_CONTROL_79_DETAILS_STYLE_MARGIN_10PX_0_PA = r"""<details style="margin:10px 0; padding:8px; background:#F9FAFB; border:1px solid #E5E7EB; border-radius:6px;">
<summary style="cursor:pointer; font-weight:bold; color:#2563EB;">
🔍 Für Neugierige: Was sagt der Cpk-Wert aus?
</summary>
<div style="margin-top:8px; padding:8px; font-size:0.95em;">

Was steckt hinter den Schwellen aus der Tabelle oben (< 0,67 / 1,0 / 1,33)?
Der Cpk übersetzt sich direkt in einen **Ausschuss-Anteil** (bei Normalverteilung):

| Cpk | Werte innerhalb der Spezifikation | Ausschuss |
|-----|-----------------------------------|-----------|
| 0,67 | ≈ 95,45 % | ≈ 45.500 ppm |
| 1,0 | ≈ 99,73 % | ≈ 2.700 ppm |
| 1,33 | ≈ 99,994 % | ≈ 63 ppm |
| 2,0 („Six Sigma") | ≈ 99,9999998 % | ≈ 0,002 ppm* |

Zwei Details:
- Cpk < 0 bedeutet: Der **Mittelwert** liegt bereits außerhalb der Spezifikation.
- Der Cpk nimmt das *schlechtere* der beiden Abstände (zur oberen und unteren Grenze) — ein dezentrierter Prozess wird also bestraft, selbst wenn er wenig streut.

*\*Das berühmte „3,4 ppm" von Six Sigma rechnet zusätzlich einen 1,5σ-Langzeit-Shift ein.*
</div>
</details>"""


def cells():
    return [
        md(_CONTROL_72_X),
        colab_code("📊 I-MR-Kontrollkarte (Stabilitätsprüfung)", _CONTROL_73_TITLE_I_MR_KONTROLLKARTE_STABI),
        md(_CONTROL_74_DETAILS_STYLE_MARGIN_10PX_0_PA),
        colab_code("📊 Normalverteilungsprüfung (Voraussetzung für Cpk)", _CONTROL_75_TITLE_NORMALVERTEILUNGSPR_FUNG),
        md(_CONTROL_76_PROZESSF_HIGKEIT_CPK_WAS_BEDEU),
        colab_code("📊 Prozessfähigkeit (Cpk)", _CONTROL_77_TITLE_PROZESSF_HIGKEIT_CPK),
        colab_code("📊 Vorher / Nachher – Zielscheiben-Vergleich", _CONTROL_78_TITLE_VORHER_NACHHER_ZIELSCHEI),
        md(_CONTROL_79_DETAILS_STYLE_MARGIN_10PX_0_PA),
        phase_export_cell("CONTROL"),
    ]
