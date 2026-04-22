"""DMAIC phase CONTROL: cells 72..80 of the notebook."""
from __future__ import annotations

from .cells import code, colab_code, md, phase_export_cell

_CONTROL_72_X = r"""---
# Phase 5: CONTROL

## Ist der verbesserte Prozess stabil und fähig?

In der Control-Phase prüft ihr die Konfirmationsdaten aus IMPROVE auf vier Kriterien:

**Schritt 1:** Stabilität — I-MR-Kontrollkarte
**Schritt 2:** Normalverteilung — Shapiro-Wilk + Q-Q-Plot (Voraussetzung für Cpk)
**Schritt 3:** Prozessfähigkeit — Cpk
**Schritt 4:** Vorher/Nachher — Baseline vs. Konfirmation"""

_CONTROL_72_SCHRITT1 = r"""## Schritt 1 – Stabilität prüfen"""

_CONTROL_72_SCHRITT2 = r"""## Schritt 2 – Normalverteilung prüfen"""

_CONTROL_72_SCHRITT3 = r"""## Schritt 3 – Prozessfähigkeit (Cpk) berechnen"""

_CONTROL_72_SCHRITT4 = r"""## Schritt 4 – Vorher/Nachher-Vergleich"""

_CONTROL_73_TITLE_I_MR_KONTROLLKARTE_STABI = r"""if len(projekt.konfirmation_wuerfe) > 0:
    projekt.imr_ergebnis = helper.berechne_imr(projekt.konfirmation_wuerfe)

    fig = helper.plot_imr(projekt.konfirmation_wuerfe)
    helper._save_fig(projekt, fig, "control_imr")
    plt.show()

    helper.zeige_stabilitaet(projekt.imr_ergebnis)"""

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
                          titel="Shapiro-Wilk p-Wert:")"""

_CONTROL_76_PROZESSF_HIGKEIT_CPK_WAS_BEDEU = r"""### Cpk – Was bedeutet das?

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
    plt.show()"""

_CONTROL_78_TITLE_VORHER_NACHHER_ZIELSCHEI = r"""if len(projekt.baseline_wuerfe) > 0 and len(projekt.konfirmation_wuerfe) > 0:
    fig = helper.plot_vorher_nachher(
        projekt.baseline_wuerfe, projekt.konfirmation_wuerfe,
        projekt.zielweite, projekt.toleranz, "1D"
    )
    helper._save_fig(projekt, fig, "control_vorher_nachher")
    plt.show()

    helper.hinweis_bericht("Cpk-Wert, Kontrollkarte und Vorher/Nachher-Zielscheibe sind die drei zentralen Control-Outputs.")"""

_CONTROL_79_DETAILS_STYLE_MARGIN_10PX_0_PA = r"""<details style="margin:10px 0; padding:8px; background:#F9FAFB; border:1px solid #E5E7EB; border-radius:6px;">
<summary style="cursor:pointer; font-weight:bold; color:#2563EB;">
🔍 Für Neugierige: Was sagt der Cpk-Wert aus?
</summary>
<div style="margin-top:8px; padding:8px; font-size:0.95em;">

Der **Cpk** (Process Capability Index) misst, wie gut euer Prozess in die Spezifikationsgrenzen passt:

$$Cpk = \min\left(\frac{USL - \bar{x}}{3\sigma}, \frac{\bar{x} - LSL}{3\sigma}\right)$$

**Interpretation:**
- Cpk < 0: Mittelwert außerhalb der Spezifikation
- Cpk 0–1.0: Prozess passt nicht sicher in die Toleranz
- Cpk 1.0–1.33: Akzeptabel
- Cpk > 1.33: Gut (industrieller Standard)
- Cpk > 1.67: Exzellent

Ein Cpk von 1.0 bedeutet: 99,73% der Werte liegen innerhalb der Spezifikation (bei Normalverteilung). Bei Cpk = 1.33 sind es 99,994%.
</div>
</details>"""


def cells():
    return [
        md(_CONTROL_72_X),
        md(_CONTROL_72_SCHRITT1),
        colab_code("📊 I-MR-Kontrollkarte", _CONTROL_73_TITLE_I_MR_KONTROLLKARTE_STABI),
        md(_CONTROL_74_DETAILS_STYLE_MARGIN_10PX_0_PA),
        md(_CONTROL_72_SCHRITT2),
        colab_code("📊 Shapiro-Wilk + Q-Q-Plot", _CONTROL_75_TITLE_NORMALVERTEILUNGSPR_FUNG),
        md(_CONTROL_72_SCHRITT3),
        md(_CONTROL_76_PROZESSF_HIGKEIT_CPK_WAS_BEDEU),
        colab_code("📊 Cpk berechnen", _CONTROL_77_TITLE_PROZESSF_HIGKEIT_CPK),
        md(_CONTROL_72_SCHRITT4),
        colab_code("📊 Vorher/Nachher-Zielscheibe", _CONTROL_78_TITLE_VORHER_NACHHER_ZIELSCHEI),
        md(_CONTROL_79_DETAILS_STYLE_MARGIN_10PX_0_PA),
        phase_export_cell("CONTROL"),
    ]
