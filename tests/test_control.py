"""CONTROL-Phase Tests: I-MR, Cpk, Normalverteilungsprüfung."""
from __future__ import annotations

import numpy as np
import pytest

import helper
from tests import virtuelle_gruppe as vg


# ─────────────────────────────────────────────────────────────
# I-MR-Kontrollkarte
# ─────────────────────────────────────────────────────────────

def test_imr_stabiler_prozess_ist_stabil():
    """Gauss-verteilte Daten ohne Drift → keine Regelverletzung."""
    rng = np.random.default_rng(7)
    daten = rng.normal(300.0, 2.0, size=25)
    erg = helper.berechne_imr(daten)
    assert bool(erg["stabil"]) is True
    assert erg["ausserhalb_i"] == 0


def test_imr_drift_wird_erkannt():
    """Linearer Drift über 25 Würfe → mindestens ein Punkt außerhalb der I-Grenzen."""
    rng = np.random.default_rng(8)
    base = rng.normal(300.0, 1.5, size=25)
    drift = np.linspace(0, 20, 25)  # Drift bis +20 cm
    daten = base + drift
    erg = helper.berechne_imr(daten)
    assert bool(erg["stabil"]) is False
    assert erg["ausserhalb_i"] >= 1


def test_imr_ucl_lcl_symmetrisch_um_xbar():
    daten = np.array([10, 12, 11, 13, 10, 14, 12, 11], dtype=float)
    erg = helper.berechne_imr(daten)
    assert abs((erg["ucl_i"] - erg["x_bar"]) - (erg["x_bar"] - erg["lcl_i"])) < 1e-9


def test_imr_auf_konfirmation_mit_drift_profil():
    """Drift-Profil der virtuellen Gruppe produziert erkennbare Instabilität."""
    p = vg.baue_control(profile="drift")
    assert bool(p.imr_ergebnis["stabil"]) is False


def test_imr_auf_konfirmation_mit_praezisionsprofil():
    p = vg.baue_control(profile="praezision")
    assert bool(p.imr_ergebnis["stabil"]) is True


# ─────────────────────────────────────────────────────────────
# Normalverteilungsprüfung
# ─────────────────────────────────────────────────────────────

def test_pruefe_normalverteilung_akzeptiert_gauss():
    rng = np.random.default_rng(11)
    daten = rng.normal(0, 1, size=30)
    erg = helper.pruefe_normalverteilung(daten)
    assert erg["shapiro_p"] > 0.05


def test_pruefe_normalverteilung_bei_schwer_bimodal():
    rng = np.random.default_rng(12)
    daten = np.concatenate([rng.normal(0, 1, 20), rng.normal(10, 1, 20)])
    erg = helper.pruefe_normalverteilung(daten)
    assert erg["shapiro_p"] < 0.05


# ─────────────────────────────────────────────────────────────
# Cpk
# ─────────────────────────────────────────────────────────────

def test_cpk_hoch_bei_enger_streuung():
    """σ sehr klein, Mittelwert = Ziel → Cpk > 1.33 (industriell fähig)."""
    rng = np.random.default_rng(13)
    daten = rng.normal(300.0, 1.0, size=30)
    erg = helper.berechne_cpk(daten, usl=330, lsl=270)
    assert erg["cpk"] > 1.33
    assert "✅" in erg["bewertung_industrie"]


def test_cpk_niedrig_bei_breiter_streuung():
    """σ groß, Ziel mittig → Cpk klein (< 0.67)."""
    rng = np.random.default_rng(14)
    daten = rng.normal(300.0, 25.0, size=30)
    erg = helper.berechne_cpk(daten, usl=315, lsl=285)
    assert erg["cpk"] < 0.67
    assert "❌" in erg["bewertung_industrie"]


def test_cpk_niedriger_bei_offset_mittelwert():
    """Verschobener Mittelwert → Cpk kleiner als Cp (Asymmetrie)."""
    # Gleiche σ, aber Mittelwert 5 cm off-center.
    rng = np.random.default_rng(15)
    daten_zentriert = rng.normal(300, 3, size=40)
    daten_offset = rng.normal(310, 3, size=40)
    usl, lsl = 315, 285
    erg_z = helper.berechne_cpk(daten_zentriert, usl=usl, lsl=lsl)
    erg_o = helper.berechne_cpk(daten_offset, usl=usl, lsl=lsl)
    assert erg_o["cpk"] < erg_z["cpk"]


def test_cpk_profile_reihenfolge_passt():
    """praezision > typisch > streu > drift bei Cpk."""
    cpks = {}
    for prof in ("praezision", "typisch", "streu", "drift"):
        cpks[prof] = vg.baue_control(profile=prof).cpk_ergebnis["cpk"]
    assert cpks["praezision"] > cpks["typisch"]
    assert cpks["typisch"] > cpks["streu"]
    assert cpks["streu"] > cpks["drift"]


def test_cpk_sigma_null_fallback():
    """Bei σ=0 liefert Cpk Infinity ohne Crash."""
    daten = np.array([300.0, 300.0, 300.0, 300.0])
    erg = helper.berechne_cpk(daten, usl=310, lsl=290)
    assert np.isinf(erg["cpk"])
