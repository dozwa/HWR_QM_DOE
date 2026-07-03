"""Regressionstests für die Qualitäts-Fixes (Juli 2026).

Abgedeckte Bugs:
- T-1:  hierarchisches_pruning terminiert auch bei geschützten Haupteffekten
- T-2:  keine kollinearen per-Faktor-x²-Terme bei 2^k+CP-Designs
- T-3:  prognostiziere ist konsistent mit dem Modell (inkl. x²-Terme)
- T-4:  saturiertes Rausch-Modell wird nicht als gut bewertet (R²_pred)
- T-10: Blocking konfundiert nicht mit einem Haupteffekt
- T-11: Centerpoints werden nicht mit Wiederholungen multipliziert
- T-12: Centerpoint-Erkennung funktioniert mit binären Faktoren
"""
from __future__ import annotations

import itertools
import signal

import numpy as np
import pandas as pd
import pytest

import helper


def _fak(n=3, binaer_letzter=False):
    fak = [{"name": name, "einheit": "u", "low": -1.0, "high": 1.0,
            "centerpoint_moeglich": True} for name in ["A", "B", "C", "D"][:n]]
    if binaer_letzter:
        fak[-1]["centerpoint_moeglich"] = False
    return fak


# ─────────────────────────────────────────────────────────────
# T-1: Pruning terminiert
# ─────────────────────────────────────────────────────────────

def test_pruning_terminiert_bei_geschuetzten_haupteffekten():
    """Starke Interaktion + insignifikante Eltern-Haupteffekte + weiterer
    insignifikanter Haupteffekt — führte früher zur Endlosschleife."""
    fak = _fak(3)
    rng = np.random.default_rng(11)
    plan = helper.generiere_versuchsplan(fak, wiederholungen=3, centerpoints=0,
                                          design="voll", seed=3)
    plan["Ergebnis (cm)"] = (300 + 2 * plan["A_coded"] + 1 * plan["B_coded"]
                              + 40 * plan["A_coded"] * plan["B_coded"]
                              + rng.normal(0, 8, len(plan)))
    modell = helper.fitte_modell(plan, fak)

    def _timeout(sig, frm):
        raise TimeoutError("hierarchisches_pruning hängt")

    signal.signal(signal.SIGALRM, _timeout)
    signal.alarm(30)
    try:
        gepruned, log = helper.hierarchisches_pruning(modell)
    finally:
        signal.alarm(0)

    terme = [t for t in gepruned.params.index if t != "Intercept"]
    # Hierarchie: A:B signifikant → A und B bleiben trotz hoher p-Werte
    assert "A:B" in terme
    assert "A" in terme and "B" in terme


# ─────────────────────────────────────────────────────────────
# T-2: Krümmung bei 2^k+CP → globaler Test, keine kollinearen x²-Terme
# ─────────────────────────────────────────────────────────────

def test_keine_kollinearen_quadratterme_bei_2level_cp():
    fak = _fak(3)
    rng = np.random.default_rng(1)
    plan = helper.generiere_versuchsplan(fak, wiederholungen=2, centerpoints=4,
                                          design="voll", seed=1)
    # Echte Krümmung einbauen
    plan["Ergebnis (cm)"] = (300 + 30 * plan["A_coded"]
                              - 25 * plan["A_coded"] ** 2
                              + rng.normal(0, 3, len(plan)))
    modell = helper.fitte_modell(plan, fak)
    # Keine per-Faktor-x²-Terme (nicht identifizierbar), aber Krümmung erkannt
    assert modell._quad_namen == []
    assert modell._kruemmung is not None
    assert modell._kruemmung["signifikant"]
    # Kein numerisch degeneriertes Modell
    assert modell.condition_number < 1e6


def test_perfaktor_quadratterme_bei_ccd_weiter_moeglich():
    """Mit axialen Punkten (CCD) bleibt die per-Faktor-Zuordnung erhalten."""
    rng = np.random.default_rng(5)
    rows = [(a, b) for a, b in itertools.product([-1, 1], [-1, 1])]
    rows += [(-1.5, 0), (1.5, 0), (0, -1.5), (0, 1.5), (0, 0), (0, 0), (0, 0)]
    rows *= 2
    df = pd.DataFrame(rows, columns=["A_coded", "B_coded"])
    df["C_coded"] = np.tile([-1, 1], len(df) // 2)
    df["Ergebnis (cm)"] = (300 + 30 * df["A_coded"] - 20 * df["A_coded"] ** 2
                            + 10 * df["B_coded"] + rng.normal(0, 3, len(df)))
    modell = helper.fitte_modell(df, _fak(3))
    assert "A_sq" in modell._quad_namen


# ─────────────────────────────────────────────────────────────
# T-3: prognostiziere == Modellvorhersage (auch mit x²-Termen)
# ─────────────────────────────────────────────────────────────

def test_prognose_konsistent_mit_modell():
    rng = np.random.default_rng(5)
    rows = [(a, b) for a, b in itertools.product([-1, 1], [-1, 1])]
    rows += [(-1.5, 0), (1.5, 0), (0, -1.5), (0, 1.5), (0, 0), (0, 0), (0, 0)]
    rows *= 2
    df = pd.DataFrame(rows, columns=["A_coded", "B_coded"])
    df["C_coded"] = np.tile([-1, 1], len(df) // 2)
    df["Ergebnis (cm)"] = (300 + 30 * df["A_coded"] - 20 * df["A_coded"] ** 2
                            + 10 * df["B_coded"] + rng.normal(0, 3, len(df)))
    fak = _fak(3)
    modell = helper.fitte_modell(df, fak)
    assert modell._quad_namen, "Testaufbau: Modell sollte x²-Terme enthalten"

    for punkt in [(1, 1, 1), (-1, 1, -1), (0, 0, -1), (0.5, 0.5, 1)]:
        werte = {"A": punkt[0], "B": punkt[1], "C": punkt[2]}
        prog = helper.prognostiziere(modell, fak, werte)
        ref = pd.DataFrame([{
            "A": punkt[0], "B": punkt[1], "C": punkt[2],
            "A_sq": punkt[0] ** 2, "B_sq": punkt[1] ** 2, "C_sq": punkt[2] ** 2,
        }])
        erwartet = float(modell.predict(ref).iloc[0])
        assert prog["vorhersage"] == pytest.approx(erwartet, abs=1e-8), (
            f"Prognose weicht an {punkt} vom Modell ab")


# ─────────────────────────────────────────────────────────────
# T-4: Saturiertes Rausch-Modell → R²_pred entlarvt es
# ─────────────────────────────────────────────────────────────

def test_rauschmodell_hat_schlechtes_r2_pred():
    fak = _fak(3)
    rng = np.random.default_rng(3)
    plan = helper.generiere_versuchsplan(fak, wiederholungen=1, centerpoints=0,
                                          design="voll", seed=1)
    plan["Ergebnis (cm)"] = 300 + rng.normal(0, 30, len(plan))
    modell = helper.fitte_modell(plan, fak)
    assert modell.rsquared > 0.8, "Testaufbau: Trainings-R² sollte täuschend hoch sein"
    press = helper.berechne_press_r2(modell)
    if press.get("berechenbar"):
        assert press["r2_pred"] < 0.5
    # sonst: nicht berechenbar → zeige_modellguete zeigt Warnbox statt Ampel


# ─────────────────────────────────────────────────────────────
# T-10 / T-11: Versuchsplan-Struktur
# ─────────────────────────────────────────────────────────────

def test_centerpoints_nicht_mit_wiederholungen_multipliziert():
    fak = _fak(3)
    plan = helper.generiere_versuchsplan(fak, wiederholungen=3, centerpoints=3,
                                          design="voll", seed=1)
    assert len(plan) == 8 * 3 + 3
    assert (plan["Typ"] == "Centerpoint").sum() == 3


def test_blocking_nicht_mit_haupteffekt_konfundiert():
    fak = _fak(3)
    plan = helper.generiere_versuchsplan(fak, wiederholungen=1, centerpoints=0,
                                          design="voll", seed=1, blocking=True)
    for col in ["A_coded", "B_coded", "C_coded"]:
        corr = np.corrcoef(plan[col], plan["Block"])[0, 1]
        assert abs(corr) < 0.01, f"Block ist mit {col} konfundiert (r={corr:.2f})"
    # Klassische Konfundierung mit der höchsten Interaktion
    prod = plan["A_coded"] * plan["B_coded"] * plan["C_coded"]
    assert ((prod > 0) == (plan["Block"] == 1)).all()


def test_blocking_fraktioniert_ohne_wiederholung_deaktiviert():
    fak = _fak(4)
    plan = helper.generiere_versuchsplan(fak, wiederholungen=1, centerpoints=0,
                                          design="halb", seed=1, blocking=True)
    assert set(plan["Block"].unique()) == {1}


# ─────────────────────────────────────────────────────────────
# T-12: CP-Erkennung mit binärem Faktor
# ─────────────────────────────────────────────────────────────

def test_kruemmung_erkannt_trotz_binaerem_faktor():
    fak = _fak(3, binaer_letzter=True)
    rng = np.random.default_rng(1)
    plan = helper.generiere_versuchsplan(fak, wiederholungen=2, centerpoints=3,
                                          design="voll", seed=1)
    plan["Ergebnis (cm)"] = (300 + 40 * plan["A_coded"]
                              - 25 * plan["A_coded"] ** 2
                              + 20 * plan["B_coded"]
                              + rng.normal(0, 5, len(plan)))
    modell = helper.fitte_modell(plan, fak)
    assert modell._kruemmung is not None, (
        "Centerpoints wurden trotz binärem Faktor nicht erkannt")
    assert modell._kruemmung["signifikant"]
