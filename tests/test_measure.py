"""MEASURE-Phase Tests: Type-1, Gage R&R, Baseline."""
from __future__ import annotations

import numpy as np
import pandas as pd
import pytest

import helper
from tests import virtuelle_gruppe as vg


# ─────────────────────────────────────────────────────────────
# Type-1 Gage Study
# ─────────────────────────────────────────────────────────────

def test_analysiere_type1_bias_und_repeatability():
    """Systematischer Bias muss sichtbar werden, Repeatability proportional zum Noise."""
    rng = np.random.default_rng(42)
    ref = 300.0
    # Person A misst mit Bias +5, σ=1
    messungen_a = rng.normal(ref + 5.0, 1.0, size=20)
    # Person B misst ohne Bias, σ=3
    messungen_b = rng.normal(ref, 3.0, size=20)
    df = pd.DataFrame({"A": messungen_a, "B": messungen_b})

    erg = helper.analysiere_type1(df, referenzwert=ref)

    # Bias-Vorzeichen und Größenordnung
    assert erg["A"]["bias"] > 3.0 and erg["A"]["bias"] < 7.0
    assert abs(erg["B"]["bias"]) < 2.0

    # Repeatability sollte A < B zeigen (A hatte σ=1, B hatte σ=3)
    assert erg["A"]["repeatability"] < erg["B"]["repeatability"]

    # n stimmt
    assert erg["A"]["n"] == 20


# ─────────────────────────────────────────────────────────────
# Gage R&R
# ─────────────────────────────────────────────────────────────

def test_gage_rr_sauber_ist_akzeptabel():
    """Ohne Operator-Bias → %GRR klein (grün oder gelb), Bewertung hat ✅ oder ⚠️."""
    p = vg.baue_measure(profile="typisch", operator_bias_on=False)
    grr = p.msa_grr
    assert "pct_grr" in grr
    assert grr["pct_grr"] < 15.0, f"erwartet niedriges %GRR, bekommen {grr['pct_grr']:.1f}"
    assert grr["var_repeatability"] >= 0
    assert grr["var_teil"] > 0


def test_gage_rr_mit_bias_ist_schlechter():
    """Mit Operator-Bias → %GRR deutlich höher als ohne."""
    p_clean = vg.baue_measure(profile="typisch", operator_bias_on=False)
    p_biased = vg.baue_measure(profile="typisch", operator_bias_on=True)
    assert p_biased.msa_grr["pct_grr"] > p_clean.msa_grr["pct_grr"]
    # Reproducibility-Varianz sollte beim Bias-Fall größer sein.
    assert p_biased.msa_grr["var_reproducibility"] > p_clean.msa_grr["var_reproducibility"]


def test_gage_rr_varianzkomponenten_summieren_sich():
    p = vg.baue_measure(profile="typisch", operator_bias_on=True)
    grr = p.msa_grr
    summe = grr["var_teil"] + grr["var_grr"]
    assert abs(summe - grr["var_total"]) < 1e-6


# ─────────────────────────────────────────────────────────────
# Baseline
# ─────────────────────────────────────────────────────────────

def test_baseline_akzeptiert_gausssche_daten():
    rng = np.random.default_rng(123)
    wuerfe = rng.normal(300.0, 5.0, size=30)
    stats = helper.analysiere_baseline(wuerfe)
    assert abs(stats["mean"] - 300.0) < 3.0
    assert abs(stats["std"] - 5.0) < 2.0
    # Shapiro-p sollte ≥ 0.05 sein bei sauberer Normalverteilung
    assert stats["shapiro_p"] > 0.05


def test_baseline_lehnt_bimodal_ab():
    """Klar bimodale Daten → Shapiro-p << 0.05."""
    rng = np.random.default_rng(321)
    a = rng.normal(250.0, 3.0, size=15)
    b = rng.normal(350.0, 3.0, size=15)
    bimodal = np.concatenate([a, b])
    stats = helper.analysiere_baseline(bimodal)
    assert stats["shapiro_p"] < 0.05


def test_baseline_from_virtuelle_gruppe_hat_mehr_streuung_bei_streu():
    """Das Streu-Profil muss höhere Baseline-Streuung haben als Präzision."""
    p_praez = vg.baue_measure(profile="praezision")
    p_streu = vg.baue_measure(profile="streu")
    assert p_streu.baseline_stats["std"] > p_praez.baseline_stats["std"]
