"""DMAIC phase ANALYZE: cells 32..54 of the notebook."""
from __future__ import annotations

from .cells import code, colab_code, md, phase_export_cell

_ANALYZE_32_X = r"""---
# Phase 3: ANALYZE

## Welche Faktoren beeinflussen die Wurfweite?

Die Faktoren habt ihr in DEFINE bereits festgelegt. Hier **verfeinert** ihr sie für das DoE und führt die Versuche durch.

**Schritt 1:** Faktoren für das DoE verfeinern (Teilmenge + Centerpoint-Entscheidung)
**Schritt 2:** Versuchsplan generieren und als Excel herunterladen
**Schritt 3:** Versuche durchführen und Ergebnisse hochladen
**Schritt 4:** Regressionsmodell berechnen und bewerten"""

_ANALYZE_32_SCHRITT1 = r"""## Schritt 1 – Faktoren fürs DoE verfeinern"""

_ANALYZE_32_SCHRITT2 = r"""## Schritt 2 – Versuchsplan generieren"""

_ANALYZE_32_SCHRITT4 = r"""## Schritt 4 – Regressionsmodell auswerten"""

_ANALYZE_33_FAKTOREN_VERFEINERN = r'''# Aktive Faktoren und Centerpoint-Entscheidung (vorbelegt aus DEFINE)
faktor1_aktiv = True if len(projekt.faktoren) > 0 else False #@param {type:"boolean"}
faktor1_centerpoint = True #@param {type:"boolean"}
faktor2_aktiv = True if len(projekt.faktoren) > 1 else False #@param {type:"boolean"}
faktor2_centerpoint = True #@param {type:"boolean"}
faktor3_aktiv = True if len(projekt.faktoren) > 2 else False #@param {type:"boolean"}
faktor3_centerpoint = True #@param {type:"boolean"}
faktor4_aktiv = True if len(projekt.faktoren) > 3 else False #@param {type:"boolean"}
faktor4_centerpoint = True #@param {type:"boolean"}
faktor5_aktiv = True if len(projekt.faktoren) > 4 else False #@param {type:"boolean"}
faktor5_centerpoint = True #@param {type:"boolean"}

if not projekt.faktoren:
    print("❌ In DEFINE wurden noch keine Faktoren definiert. Bitte zurück zu DEFINE → 'Faktoren übernehmen'.")
else:
    _aktiv_flags = [faktor1_aktiv, faktor2_aktiv, faktor3_aktiv, faktor4_aktiv, faktor5_aktiv]
    _cp_flags = [faktor1_centerpoint, faktor2_centerpoint, faktor3_centerpoint,
                 faktor4_centerpoint, faktor5_centerpoint]

    _faktoren_doe = []
    for i, f in enumerate(projekt.faktoren):
        if not _aktiv_flags[i]:
            continue
        f_copy = dict(f)
        f_copy["centerpoint_moeglich"] = bool(_cp_flags[i] and f.get("centerpoint_moeglich", True))
        _faktoren_doe.append(f_copy)

    if len(_faktoren_doe) < 3:
        print(f"❌ Mindestens 3 aktive Faktoren nötig (derzeit {len(_faktoren_doe)}).")
    else:
        projekt.faktoren_doe = _faktoren_doe
        print(f"✅ {len(_faktoren_doe)} Faktoren für das DoE gewählt:")
        for f in _faktoren_doe:
            cp = "mit CP" if f["centerpoint_moeglich"] else "ohne CP"
            print(f"   • {f['name']} [{f['low']} – {f['high']} {f['einheit']}] ({cp})")
        helper.speichere_fortschritt(projekt)
'''

_ANALYZE_38_TITLE_VERSUCHSPLAN_GENERIEREN = r"""design_typ = "Vollfaktoriell (2^k)" #@param ["Vollfaktoriell (2^k)", "Halbfraktionell (2^(k-1))", "Viertelfraktionell (2^(k-2))"]
wiederholungen = 3 #@param {type:"integer"}
blocking = False #@param {type:"boolean"}
centerpoints = 3 #@param {type:"integer"}

# Design-Typ mappen
_design_map = {"Vollfaktoriell": "voll", "Halbfraktionell": "halb", "Viertelfraktionell": "viertel"}
_design = next(v for k, v in _design_map.items() if k in design_typ)

# Wiederholungen validieren
wiederholungen = max(1, min(10, int(wiederholungen)))

# DoE-Faktoren aus ANALYZE-Verfeinerung (Fallback: Master-Liste aus DEFINE)
_fak_doe = helper._effektive_faktoren(projekt)
if not _fak_doe:
    print("❌ Keine Faktoren verfügbar. Bitte zuerst DEFINE → 'Faktoren übernehmen' ausführen.")
    raise RuntimeError("keine Faktoren")

# Validierung
for f in _fak_doe:
    if f["low"] == f["high"]:
        print(f"❌ Faktor '{f['name']}': Low und High sind identisch ({f['low']}). Bitte Faktor-Definition in DEFINE anpassen!")
    if f["low"] > f["high"]:
        f["low"], f["high"] = f["high"], f["low"]
        print(f"⚠️ Faktor '{f['name']}': Low und High vertauscht – automatisch korrigiert.")

# Centerpoints: nur möglich wenn mindestens ein Faktor stetig ist
_cp_faktoren = [f for f in _fak_doe if f.get("centerpoint_moeglich", True)]
if not _cp_faktoren:
    if centerpoints > 0:
        print("ℹ️ Keine stetigen Faktoren → Centerpoints nicht möglich, auf 0 gesetzt.")
    centerpoints = 0
elif centerpoints < 0:
    centerpoints = 0

_bin_namen = [f["name"] for f in _fak_doe if not f.get("centerpoint_moeglich", True)]
if _bin_namen:
    print(f"ℹ️ Zweistufige Faktoren (kein Centerpoint): {', '.join(_bin_namen)}")

projekt.versuchsplan_config = {
    "wiederholungen": wiederholungen,
    "blocking": blocking,
    "centerpoints": centerpoints,
    "design": _design,
}

projekt.versuchsplan = helper.generiere_versuchsplan(
    _fak_doe,
    wiederholungen=wiederholungen,
    blocking=blocking,
    centerpoints=centerpoints,
    seed=projekt.seed,
    design=_design,
)

helper.zeige_versuchsplan_info(projekt.versuchsplan, _fak_doe)
print(f"\nDie ersten 10 Versuche:")
display(projekt.versuchsplan.head(10))
helper.speichere_fortschritt(projekt)
"""

_ANALYZE_39_DETAILS_STYLE_MARGIN_10PX_0_PA = r"""<details style="margin:10px 0; padding:8px; background:#F9FAFB; border:1px solid #E5E7EB; border-radius:6px;">
<summary style="cursor:pointer; font-weight:bold; color:#2563EB;">
🔍 Für Neugierige: Was bedeutet die Kodierung (−1 / +1)?
</summary>
<div style="margin-top:8px; padding:8px; font-size:0.95em;">

In der Versuchsmatrix werden die Faktorstufen **kodiert**:
- **−1** = Low-Stufe (euer unterer Wert)
- **+1** = High-Stufe (euer oberer Wert)
- **0** = Centerpoint (Mittelwert zwischen Low und High)

Die Kodierung hat einen wichtigen Vorteil: Alle Faktoren sind auf die gleiche Skala normiert. Dadurch kann man die **Effektgrößen direkt vergleichen** – egal ob ein Faktor in Grad und der andere in Zentimetern gemessen wird.

Ein Koeffizient von +25 bei kodierter Eingabe bedeutet: Wenn der Faktor von −1 auf +1 wechselt (also von Low auf High), steigt die Wurfweite um 25 cm.
</div>
</details>"""

_ANALYZE_40_TITLE_VERSUCHSPLAN_ALS_EXCEL_H = r"""filepath = helper.erstelle_doe_excel(
    projekt.versuchsplan, helper._effektive_faktoren(projekt)
)
print(f"✅ Excel erstellt!")

# Kopie im Drive-Ordner ablegen
import shutil
_save_dir = helper._fortschritt_verzeichnis(projekt)
if _save_dir:
    os.makedirs(_save_dir, exist_ok=True)
    shutil.copy2(filepath, os.path.join(_save_dir, os.path.basename(filepath)))

try:
    from google.colab import files
    files.download(filepath)
    print("📥 Download gestartet!")
except ImportError:
    print(f"📁 Datei: {filepath}")
"""

_ANALYZE_41_X = r"""---
## Schritt 3 – Versuche durchführen

Jetzt geht es los! Führt die Versuche in der **randomisierten Reihenfolge** aus dem Excel durch.

**Pro Versuch (~2 Min):**
1. Faktoren laut Plan einstellen
2. Wurf durchführen
3. Auftreffpunkt markieren
4. Wurfweite messen
5. Ergebnis ins Excel eintragen
6. → Nächster Versuch

> ⚠️ **Wichtig:** Reihenfolge einhalten! Die Randomisierung schützt vor systematischen Verfälschungen.

> 📋 **Für den Bericht:** Dokumentiert besondere Vorkommnisse (Katapult-Probleme, Ausreißer-Würfe, etc.)"""

_ANALYZE_42_X = r"""---
## Schritt 4 – Regressionsmodell berechnen und bewerten

Bevor ihr die Ergebnisse seht, hier die drei wichtigsten Werkzeuge:

### Pareto-Diagramm (Hauptwerkzeug)
Das Pareto-Diagramm sortiert alle Effekte nach Stärke. **Je länger der Balken, desto größer der Einfluss** des Faktors auf die Wurfweite. Balken über der roten Linie sind **statistisch signifikant**.

### p-Wert (Signifikanzfilter)
Der p-Wert beantwortet die Frage: *Könnte dieser Effekt auch Zufall sein?*
- **p < 0,05** → Der Effekt ist real (statistisch signifikant)
- **p > 0,05** → Nicht sicher genug – könnte Zufall sein

### R² (Modellgüte)
R² sagt euch: **Wie viel Prozent der Streuung erklärt euer Modell?**
- R² = 0,80 bedeutet: 80% der Variation wird durch eure Faktoren erklärt
- Die restlichen 20% sind Rauschen (Messungenauigkeit, unbekannte Einflüsse)

Ladet jetzt eure ausgefüllte Versuchsergebnis-Excel hoch."""

_ANALYZE_43_TITLE_VERSUCHSERGEBNISSE_HOCHL = r"""try:
    from google.colab import files
    print("⬆️ Bitte DoE_Versuchsergebnisse.xlsx (ausgefüllt) hochladen:")
    uploaded = files.upload()
    doe_file = list(uploaded.keys())[0]
except ImportError:
    doe_file = "DoE_Versuchsergebnisse.xlsx"

doe_df = pd.read_excel(doe_file)
print(f"✅ {len(doe_df)} Versuchsergebnisse geladen")
display(doe_df.head())

# In Projekt speichern
projekt.doe_ergebnisse = doe_df
projekt.csv_daten["doe_ergebnisse"] = doe_df

# Excel-Backup auf Drive speichern
import shutil
_save_dir = helper._fortschritt_verzeichnis(projekt)
if _save_dir:
    os.makedirs(_save_dir, exist_ok=True)
    shutil.copy2(doe_file, os.path.join(_save_dir, "DoE_Versuchsergebnisse.xlsx"))

helper.speichere_fortschritt(projekt)
"""

_ANALYZE_44_TITLE_REGRESSIONSMODELL_BERECH = r"""try:
    # Konsistenzprüfung: Excel-Faktoren ggü. den in DEFINE/ANALYZE definierten.
    # Maßgeblich für die Regression sind die Excel-Namen — das sind die
    # tatsächlichen Bedingungen, unter denen gemessen wurde.
    _fak_excel = []
    if projekt.doe_ergebnisse is not None:
        _fak_excel = helper._parse_faktoren_aus_excel(projekt.doe_ergebnisse)

    _vorher_cp = {f["name"]: f.get("centerpoint_moeglich", True)
                  for f in helper._effektive_faktoren(projekt)}

    if _fak_excel and _vorher_cp:
        helper.pruefe_excel_faktoren_konsistenz(projekt, _fak_excel)

    if _fak_excel:
        # Excel-Parser trägt centerpoint_moeglich nicht mit — aus DEFINE/ANALYZE
        # übernehmen, damit stetige/zweistufige Unterscheidung nicht verloren geht.
        for f in _fak_excel:
            f["centerpoint_moeglich"] = _vorher_cp.get(f["name"], True)
        projekt.faktoren_doe = _fak_excel
        print(f"ℹ️ {len(_fak_excel)} Faktoren aus Excel übernommen:")
        for f in _fak_excel:
            print(f"   • {f['name']} [{f['low']} – {f['high']} {f['einheit']}]")

    _fak = helper._effektive_faktoren(projekt)
    projekt.modell = helper.fitte_modell(projekt.doe_ergebnisse, _fak)
    print(f"✅ Modell berechnet!")
    print(f"   R² = {projekt.modell.rsquared:.4f}")
    print(f"   Adj. R² = {projekt.modell.rsquared_adj:.4f}")
    print(f"   Faktoren: {', '.join(projekt.modell._faktor_namen)}")
    helper.speichere_fortschritt(projekt)
except Exception as e:
    print(f"❌ Fehler bei der Modellberechnung: {e}")
    print("\nPrüft eure Daten:")
    print("  - Enthält die Excel-Datei eine Spalte 'Ergebnis: Weite (cm)'?")
    print("  - Sind alle Werte Zahlen (keine leeren Zellen, kein Text)?")
    print("  - Habt ihr Dezimalpunkte statt Kommas verwendet?")
    if projekt.doe_ergebnisse is not None:
        print(f"\n  Spalten in der Excel: {projekt.doe_ergebnisse.columns.tolist()}")
    _fak = helper._effektive_faktoren(projekt)
    if _fak:
        print(f"  Verwendete Faktoren: {[f['name'] for f in _fak]}")
    else:
        print("  ⚠️ Keine Faktoren definiert! Bitte DEFINE → 'Faktoren übernehmen' ausführen.")
"""

_ANALYZE_45_TITLE_AUTOMATISCHES_MODELL_PRU = r"""projekt.modell_gepruned, projekt.pruning_log = helper.hierarchisches_pruning(projekt.modell)

print("Pruning-Protokoll:")
for msg in projekt.pruning_log:
    print(f"  {msg}")

# Modell aktualisieren
if projekt.modell_gepruned is not None:
    projekt.modell = projekt.modell_gepruned
    print(f"\n✅ Finales Modell: R² = {projekt.modell.rsquared:.4f}")"""

_ANALYZE_46_DETAILS_STYLE_MARGIN_10PX_0_PA = r"""<details style="margin:10px 0; padding:8px; background:#F9FAFB; border:1px solid #E5E7EB; border-radius:6px;">
<summary style="cursor:pointer; font-weight:bold; color:#2563EB;">
🔍 Für Neugierige: Was ist das Hierarchieprinzip beim Pruning?
</summary>
<div style="margin-top:8px; padding:8px; font-size:0.95em;">
Ein Interaktionseffekt bedeutet, dass die Wirkung von Faktor A davon abhängt, auf welcher Stufe Faktor B steht. Ohne den Haupteffekt A im Modell wäre die Interaktion nicht korrekt interpretierbar. Deshalb bleibt Faktor A im Modell, auch wenn er allein nicht signifikant ist.

**Beispiel:** Wenn die Interaktion "Winkel × Spannung" signifikant ist (p < 0.05), bleiben sowohl "Winkel" als auch "Spannung" im Modell – auch wenn einer der beiden Haupteffekte allein p > 0.05 hat.
</div>
</details>"""

_ANALYZE_47_TITLE_PARETO_DIAGRAMM_DER_STAN = r"""fig = helper.plot_pareto_effekte(projekt.modell)
helper._save_fig(projekt, fig, "analyze_pareto")
plt.show()"""

_ANALYZE_48_TITLE_KOEFFIZIENTENTABELLE = r'''helper.zeige_koeffizienten(projekt.modell)

# Warnung wenn kein Faktor signifikant
p_vals = projekt.modell.pvalues.drop("Intercept", errors="ignore")
if (p_vals > 0.05).all():
    display(HTML("""
    <div style="padding:10px; border-left:4px solid #DC2626; background:#FEE2E2; border-radius:4px; margin:8px 0;">
        ❌ <strong>Kein Faktor ist signifikant (alle p > 0,05).</strong><br>
        Mögliche Ursachen:<br>
        • Stufenabstände zu klein – war der Unterschied zwischen Low und High spürbar?<br>
        • Zu wenig Wiederholungen – Effekte gehen im Rauschen unter<br>
        • Falsche Faktoren gewählt – relevanter Parameter wurde nicht variiert<br>
        <strong>→ Trotzdem weitermachen!</strong> Interpretiert die Richtung der Effekte und diskutiert die Ursachen im Bericht.
    </div>"""))'''

_ANALYZE_49_TITLE_MODELLG_TE = r"""helper.zeige_modellguete(projekt.modell)"""

_ANALYZE_50_TITLE_HINTERGRUNDPR_FUNGEN_VIF = r"""# VIF
vifs = helper.pruefe_vif(projekt.modell)

# Lack-of-Fit (nur wenn Centerpoints vorhanden)
if projekt.doe_ergebnisse is not None:
    lof = helper.pruefe_lack_of_fit(projekt.modell, projekt.modell._daten)"""

_ANALYZE_51_TITLE_RESIDUENPLOTS_ANOVA_TABE = r"""fig = helper.plot_residuen(projekt.modell)
helper._save_fig(projekt, fig, "analyze_residuen")
plt.show()

# ANOVA-Tabelle (Prio 2)
helper.zeige_anova_tabelle(projekt.modell)
"""

_ANALYZE_52_DETAILS_STYLE_MARGIN_10PX_0_PA = r"""<details style="margin:10px 0; padding:8px; background:#F9FAFB; border:1px solid #E5E7EB; border-radius:6px;">
<summary style="cursor:pointer; font-weight:bold; color:#2563EB;">
🔍 Für Neugierige: Wie liest man Residuenplots?
</summary>
<div style="margin-top:8px; padding:8px; font-size:0.95em;">

**Residuen** = Gemessener Wert – Vorhergesagter Wert. Sie zeigen, was das Modell **nicht** erklären kann.

**Gute Zeichen:**
- Residuen zufällig um Null verteilt (keine Muster)
- Keine Trichterform (konstante Varianz)
- Punkte nahe der Linie im Q-Q-Plot (Normalverteilung)

**Schlechte Zeichen:**
- Systematische Muster (z.B. U-Form) → nichtlinearer Effekt fehlt
- Trichterform → Varianz hängt von der Größe ab
- Einzelne Ausreißer weit weg → eventuell Messfehler
</div>
</details>"""

_ANALYZE_53_DIV_STYLE_PADDING_10PX_BORDER = r"""<div style="padding:10px; border-left:4px solid #2563EB; background:#DBEAFE; border-radius:4px;">
📋 <strong>Für den Bericht:</strong> Pareto-Diagramm, Koeffiziententabelle und R² sind die drei zentralen Outputs der Analyze-Phase. Exportiert sie am Ende des Tages.
</div>"""


def cells():
    return [
        md(_ANALYZE_32_X),
        md(_ANALYZE_32_SCHRITT1),
        colab_code("🧩 Faktoren für das DoE verfeinern", _ANALYZE_33_FAKTOREN_VERFEINERN),
        md(_ANALYZE_32_SCHRITT2),
        colab_code("⚙️ Versuchsplan generieren", _ANALYZE_38_TITLE_VERSUCHSPLAN_GENERIEREN),
        md(_ANALYZE_39_DETAILS_STYLE_MARGIN_10PX_0_PA),
        colab_code("📥 Versuchsplan als Excel herunterladen", _ANALYZE_40_TITLE_VERSUCHSPLAN_ALS_EXCEL_H),
        md(_ANALYZE_41_X),
        md(_ANALYZE_42_X),
        colab_code("⬆️ Versuchsergebnisse hochladen", _ANALYZE_43_TITLE_VERSUCHSERGEBNISSE_HOCHL),
        colab_code("📊 Regressionsmodell berechnen", _ANALYZE_44_TITLE_REGRESSIONSMODELL_BERECH),
        colab_code("⚙️ Automatisches Modell-Pruning (Hierarchieprinzip)", _ANALYZE_45_TITLE_AUTOMATISCHES_MODELL_PRU),
        md(_ANALYZE_46_DETAILS_STYLE_MARGIN_10PX_0_PA),
        colab_code("📊 Pareto-Diagramm der standardisierten Effekte", _ANALYZE_47_TITLE_PARETO_DIAGRAMM_DER_STAN),
        colab_code("📊 Koeffiziententabelle", _ANALYZE_48_TITLE_KOEFFIZIENTENTABELLE),
        colab_code("📊 Modellgüte", _ANALYZE_49_TITLE_MODELLG_TE),
        colab_code("⚙️ Hintergrundprüfungen (VIF, Lack-of-Fit)", _ANALYZE_50_TITLE_HINTERGRUNDPR_FUNGEN_VIF),
        colab_code("🔍 Residuenplots + ANOVA-Tabelle (Prio 2)", _ANALYZE_51_TITLE_RESIDUENPLOTS_ANOVA_TABE),
        md(_ANALYZE_52_DETAILS_STYLE_MARGIN_10PX_0_PA),
        md(_ANALYZE_53_DIV_STYLE_PADDING_10PX_BORDER),
        phase_export_cell("ANALYZE"),
    ]
