"""ANALYZE-Phase Tests: Versuchsplan, Modell-Fit, Pruning, _effektive_faktoren."""
from __future__ import annotations

import numpy as np
import pytest

import helper
from tests import virtuelle_gruppe as vg


# ─────────────────────────────────────────────────────────────
# generiere_versuchsplan
# ─────────────────────────────────────────────────────────────

def test_versuchsplan_groesse_vollfaktoriell():
    """Voll-faktoriell: 2^k × wiederholungen + centerpoints_total."""
    faktoren = [dict(f) for f in vg.FAKTOREN_KATALOG]
    for f in faktoren:
        f.pop("_statapult_key", None)
    plan = helper.generiere_versuchsplan(faktoren, wiederholungen=2, centerpoints=3,
                                          seed=1, design="voll")
    # 2^3 = 8 Ecken × 2 Wiederholungen = 16 Ecken + 3 CPs × 2 Wiederholungen = 22
    assert len(plan) == 8 * 2 + 3 * 2


def test_versuchsplan_kodierte_spalten_sind_pm_1_oder_0():
    faktoren = [dict(f) for f in vg.FAKTOREN_KATALOG]
    for f in faktoren:
        f.pop("_statapult_key", None)
    plan = helper.generiere_versuchsplan(faktoren, wiederholungen=1, centerpoints=2,
                                          seed=1, design="voll")
    coded_cols = [c for c in plan.columns if c.endswith("_coded")]
    assert len(coded_cols) == len(faktoren)
    for col in coded_cols:
        uniq = set(plan[col].unique())
        assert uniq.issubset({-1.0, 0.0, 1.0}), f"unerwartete kodierte Werte in {col}: {uniq}"


def test_versuchsplan_ist_randomisiert():
    """Zwei Pläne mit unterschiedlichem Seed sollen unterschiedliche Reihenfolge haben."""
    faktoren = [dict(f) for f in vg.FAKTOREN_KATALOG]
    for f in faktoren:
        f.pop("_statapult_key", None)
    p1 = helper.generiere_versuchsplan(faktoren, wiederholungen=1, centerpoints=0,
                                        seed=1, design="voll")
    p2 = helper.generiere_versuchsplan(faktoren, wiederholungen=1, centerpoints=0,
                                        seed=99, design="voll")
    assert not (p1["Versuch_Nr"].tolist() == p2["Versuch_Nr"].tolist() and
                p1.iloc[:, 2].tolist() == p2.iloc[:, 2].tolist())


def test_versuchsplan_halb_weniger_als_voll():
    faktoren = [dict(f) for f in vg.FAKTOREN_KATALOG] + [
        {"name": "Extra1", "einheit": "cm", "low": 0, "high": 10, "centerpoint_moeglich": True},
    ]
    for f in faktoren:
        f.pop("_statapult_key", None)
    voll = helper.generiere_versuchsplan(faktoren, wiederholungen=1, centerpoints=0, seed=1,
                                          design="voll")
    halb = helper.generiere_versuchsplan(faktoren, wiederholungen=1, centerpoints=0, seed=1,
                                          design="halb")
    assert len(halb) < len(voll)
    assert len(halb) == len(voll) / 2


# ─────────────────────────────────────────────────────────────
# fitte_modell
# ─────────────────────────────────────────────────────────────

def test_modell_erreicht_hohes_r2_auf_praezisions_daten():
    p = vg.baue_analyze(profile="praezision")
    assert p.modell.rsquared > 0.9, f"R²={p.modell.rsquared:.3f}"


def test_modell_erkennt_statapult_haupteffekte_als_signifikant():
    """Abzugswinkel und Gummiband-Position haben in statapult echte Effekte → p<0.05."""
    p = vg.baue_analyze(profile="typisch")
    sig_namen = []
    for name in p.modell._faktor_namen:
        if name in p.modell.pvalues.index and p.modell.pvalues[name] < 0.05:
            sig_namen.append(name)
    # Mindestens 2 der 3 Haupteffekte signifikant bei realistischen Statapult-Daten.
    assert len(sig_namen) >= 2, f"nur {sig_namen} signifikant von {p.modell._faktor_namen}"


def test_modell_koeffizienten_zeigen_richtige_richtung():
    """Ein höherer Abzugswinkel/höhere Gummiband-Position erhöhen die Weite.
    Die kodierten Koeffizienten sollten also positiv sein."""
    p = vg.baue_analyze(profile="praezision")
    params = p.modell.params
    # Nimm das Feld 'Abzugswinkel' und prüfe positiven Koeffizient.
    abz_key = next((n for n in p.modell._faktor_namen if "Abzugswinkel" in n), None)
    assert abz_key is not None
    assert params[abz_key] > 0


# ─────────────────────────────────────────────────────────────
# hierarchisches_pruning
# ─────────────────────────────────────────────────────────────

def test_pruning_bewahrt_haupteffekte_signifikanter_interaktionen():
    """Ein Haupteffekt darf nicht entfernt werden, wenn eine Interaktion mit
    ihm signifikant ist."""
    p = vg.baue_analyze(profile="streu")  # mehr Rauschen → Pruning aktiver
    terme = list(p.modell.params.index)
    interaktionen = [t for t in terme if ":" in t]
    for inter in interaktionen:
        komponenten = inter.split(":")
        for k in komponenten:
            assert k in terme, (
                f"Haupteffekt {k} fehlt trotz vorhandener Interaktion {inter}"
            )


def test_pruning_reduziert_oder_behaelt_r2_adj():
    """Pruning sollte R²_adj nicht drastisch verschlechtern (schlechte Terme entfernen)."""
    p = vg.baue_analyze(profile="streu")
    # p.modell ist bereits das gepruned (see baue_analyze).
    assert p.modell.rsquared_adj > 0.5, f"R²_adj zu niedrig: {p.modell.rsquared_adj}"


# ─────────────────────────────────────────────────────────────
# _effektive_faktoren
# ─────────────────────────────────────────────────────────────

def test_effektive_faktoren_fallback_auf_master():
    p = helper.init_projekt("Eff", 1)
    p.faktoren = [{"name": "A", "einheit": "cm", "low": 0, "high": 1,
                   "centerpoint_moeglich": True}]
    p.faktoren_doe = []
    assert helper._effektive_faktoren(p) == p.faktoren


def test_effektive_faktoren_doe_ueberschreibt():
    p = helper.init_projekt("Eff2", 1)
    p.faktoren = [{"name": "A", "einheit": "cm", "low": 0, "high": 1,
                   "centerpoint_moeglich": True},
                  {"name": "B", "einheit": "cm", "low": 0, "high": 1,
                   "centerpoint_moeglich": True}]
    p.faktoren_doe = [dict(p.faktoren[0], centerpoint_moeglich=False)]
    eff = helper._effektive_faktoren(p)
    assert len(eff) == 1
    assert eff[0]["name"] == "A"
    assert eff[0]["centerpoint_moeglich"] is False
