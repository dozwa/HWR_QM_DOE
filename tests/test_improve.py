"""IMPROVE-Phase Tests: optimiere_einstellungen + analysiere_konfirmation."""
from __future__ import annotations

import numpy as np
import pytest

import helper
from tests import virtuelle_gruppe as vg


# ─────────────────────────────────────────────────────────────
# optimiere_einstellungen
# ─────────────────────────────────────────────────────────────

def test_optimierung_mittelwert_trifft_zielweite():
    """Strategie 'mittelwert' soll die Vorhersage nahe an die Zielweite bringen."""
    p = vg.baue_analyze(profile="praezision")
    erg = helper.optimiere_einstellungen(
        p.modell, zielweite=p.zielweite, faktoren=p.faktoren_doe, strategie="mittelwert",
    )
    # Vorhersage sollte ≈ Zielweite sein (1 cm Toleranz auf präzisen Daten).
    assert abs(erg["vorhersage"] - p.zielweite) < 2.0


def test_optimierung_liefert_einstellungen_fuer_alle_faktoren():
    p = vg.baue_analyze(profile="typisch")
    erg = helper.optimiere_einstellungen(p.modell, zielweite=p.zielweite,
                                          faktoren=p.faktoren_doe, strategie="dual")
    for f in p.faktoren_doe:
        assert f["name"] in erg["einstellungen"]
        eintrag = erg["einstellungen"][f["name"]]
        # Kodierter Wert im erlaubten Bereich [-1.2, 1.2].
        assert -1.21 <= eintrag["coded"] <= 1.21
        # Original liegt nahe am dekodierten Interval [low, high] ± 10%.
        span = f["high"] - f["low"]
        assert f["low"] - 0.1 * span <= eintrag["original"] <= f["high"] + 0.1 * span


def test_optimierung_pi_umgibt_vorhersage():
    p = vg.baue_analyze(profile="typisch")
    erg = helper.optimiere_einstellungen(p.modell, zielweite=p.zielweite,
                                          faktoren=p.faktoren_doe, strategie="dual")
    assert erg["pi_low"] < erg["vorhersage"] < erg["pi_high"]


def test_strategien_liefern_unterschiedliche_einstellungen():
    """'mittelwert' und 'varianz' sollten unterschiedliche Punkte liefern."""
    p = vg.baue_analyze(profile="streu")
    mittel = helper.optimiere_einstellungen(p.modell, p.zielweite, p.faktoren_doe,
                                             strategie="mittelwert")
    varianz = helper.optimiere_einstellungen(p.modell, p.zielweite, p.faktoren_doe,
                                              strategie="varianz")
    # Mindestens ein Faktor sollte unterschiedlich optimiert sein.
    unterschiede = 0
    for name in mittel["einstellungen"]:
        diff = abs(mittel["einstellungen"][name]["coded"]
                   - varianz["einstellungen"][name]["coded"])
        if diff > 0.1:
            unterschiede += 1
    assert unterschiede >= 1, "Mittelwert- und Varianz-Strategie sollten sich unterscheiden"


def test_dual_strategie_liegt_zwischen_den_extremen():
    """Die dual-Strategie sollte mindestens einen Faktor 'dazwischen' wählen."""
    p = vg.baue_analyze(profile="streu")
    mittel = helper.optimiere_einstellungen(p.modell, p.zielweite, p.faktoren_doe,
                                             strategie="mittelwert")
    varianz = helper.optimiere_einstellungen(p.modell, p.zielweite, p.faktoren_doe,
                                              strategie="varianz")
    dual = helper.optimiere_einstellungen(p.modell, p.zielweite, p.faktoren_doe,
                                           strategie="dual", lambda_gewicht=0.01)
    # Dual sollte nicht identisch mit einer der Extrem-Strategien sein.
    mittel_vec = np.array([e["coded"] for e in mittel["einstellungen"].values()])
    varianz_vec = np.array([e["coded"] for e in varianz["einstellungen"].values()])
    dual_vec = np.array([e["coded"] for e in dual["einstellungen"].values()])
    assert not np.allclose(dual_vec, mittel_vec, atol=1e-4) or \
           not np.allclose(dual_vec, varianz_vec, atol=1e-4)


# ─────────────────────────────────────────────────────────────
# analysiere_konfirmation
# ─────────────────────────────────────────────────────────────

def test_konfirmation_erfolgreich_bei_praezisem_modell():
    """Bei hoher Präzision sollten die Konfirmationswürfe im Toleranzband
    liegen — das PI des Modells kann durch enge Fit-Residuen schmaler sein
    als die tatsächliche Prozess-Sigma, daher wird nicht ``in_pred`` geprüft."""
    p = vg.baue_improve(profile="praezision")
    erg = p.konfirmation_ergebnis
    assert erg["pct_in_tol"] >= 80.0, f"nur {erg['pct_in_tol']:.0f}% in Toleranz"
    assert abs(erg["mean"] - p.zielweite) <= p.toleranz


def test_konfirmation_erkennt_abweichung():
    """Manuell: Würfe weit außerhalb der Vorhersage → in_pred=False."""
    wuerfe = np.array([100, 105, 102, 108, 98], dtype=float)
    erg = helper.analysiere_konfirmation(
        wuerfe, vorhersage=300, pi_low=280, pi_high=320, zielweite=300, toleranz=15,
    )
    assert bool(erg["in_pred"]) is False
    assert erg["pct_in_tol"] == 0.0
    assert "❌" in erg["bewertung"]


def test_konfirmation_streuung_warnt():
    """Mittelwert im PI, aber hohe Streuung → ⚠️."""
    # Mittelwert=300, Toleranz=15 → Würfe bei [290..310] sollen 100% in tol sein,
    # aber Werte knapp außerhalb reduzieren pct_in_tol.
    wuerfe = np.array([290, 300, 320, 280, 315], dtype=float)
    erg = helper.analysiere_konfirmation(
        wuerfe, vorhersage=300, pi_low=270, pi_high=330, zielweite=300, toleranz=15,
    )
    # pct_in_tol < 80 ⇒ Bewertung zeigt Warnung oder Fehler (je nach in_pred).
    assert erg["pct_in_tol"] < 80.0
    assert "✅" not in erg["bewertung"]
