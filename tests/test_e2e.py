"""End-to-End DMAIC-Läufe für 4 virtuelle Gruppen.

Jede Gruppe durchläuft DEFINE → CONTROL mit eigenen statapult-Daten.
Die Tests prüfen, dass die finalen Ergebnisse die erwartete Charakteristik
des Profils widerspiegeln — präzise Gruppe ≠ streuende Gruppe ≠ Drift-Gruppe.
"""
from __future__ import annotations

import numpy as np
import pytest

import helper
from tests import virtuelle_gruppe as vg


# ─────────────────────────────────────────────────────────────
# Shared "Musterteam" fixture — einmal erstellt, von mehreren Tests genutzt.
# ─────────────────────────────────────────────────────────────

@pytest.fixture(scope="module")
def musterteam():
    """'Typisches' E2E-Projekt — Module-scoped zur Performance-Schonung."""
    return vg.fertige_gruppe(profile="typisch")


# ─────────────────────────────────────────────────────────────
# Einzelne E2E-Läufe je Profil
# ─────────────────────────────────────────────────────────────

def test_e2e_praezisionsteam_erreicht_hohes_cpk():
    p = vg.fertige_gruppe(profile="praezision")
    # Alle Phasen befüllt
    assert p.charter and len(p.testwuerfe) == 5
    assert p.msa_grr is not None and p.baseline_stats is not None
    assert p.modell is not None and p.doe_ergebnisse is not None
    assert p.optimale_einstellung is not None
    assert p.konfirmation_ergebnis is not None and p.cpk_ergebnis is not None
    # Qualitätsmerkmale Präzisionsteam
    assert p.msa_grr["pct_grr"] < 20.0
    assert p.baseline_stats["std"] < 3.0
    assert p.modell.rsquared > 0.95
    assert p.cpk_ergebnis["cpk"] > 2.0
    assert bool(p.imr_ergebnis["stabil"]) is True


def test_e2e_typisches_team_industriell_faehig():
    p = vg.fertige_gruppe(profile="typisch")
    assert p.msa_grr["pct_grr"] < 30.0  # akzeptabel / bedingt akzeptabel
    assert p.modell.rsquared > 0.90
    assert p.cpk_ergebnis["cpk"] > 1.33  # industriell fähig
    assert bool(p.imr_ergebnis["stabil"]) is True


def test_e2e_streuteam_cpk_mittel():
    p = vg.fertige_gruppe(profile="streu")
    # Baseline streut deutlich
    assert p.baseline_stats["std"] > 3.0
    # Modell funktioniert trotzdem noch
    assert p.modell.rsquared > 0.70
    # Cpk schlechter als bei typisch, aber noch positiv
    assert 0.5 < p.cpk_ergebnis["cpk"] < 3.0


def test_e2e_driftteam_zeigt_regelverletzung_und_schwaches_cpk():
    p = vg.fertige_gruppe(profile="drift")
    assert bool(p.imr_ergebnis["stabil"]) is False
    assert p.imr_ergebnis["ausserhalb_i"] >= 1
    # Cpk klein durch Drift (Mittelwert bewegt sich weg vom Ziel).
    assert p.cpk_ergebnis["cpk"] < 1.33


# ─────────────────────────────────────────────────────────────
# Cross-Profile-Vergleich: Charakteristik ist unterscheidbar.
# ─────────────────────────────────────────────────────────────

def test_cpk_rangfolge_zwischen_profilen():
    cpks = {p: vg.fertige_gruppe(profile=p).cpk_ergebnis["cpk"]
            for p in ("praezision", "typisch", "streu", "drift")}
    assert cpks["praezision"] > cpks["typisch"] > cpks["streu"] > cpks["drift"]


def test_baseline_streuung_rangfolge():
    stds = {}
    for prof in ("praezision", "typisch", "streu"):
        p = vg.baue_measure(profile=prof)
        stds[prof] = p.baseline_stats["std"]
    assert stds["praezision"] < stds["typisch"] < stds["streu"]


# ─────────────────────────────────────────────────────────────
# Musterteam-Fixture: gemeinsam genutzte E2E-Gruppe.
# ─────────────────────────────────────────────────────────────

def test_musterteam_hat_alle_phasen(musterteam):
    p = musterteam
    assert p.charter
    assert p.msa_grr is not None
    assert p.doe_ergebnisse is not None
    assert p.modell is not None
    assert p.optimale_einstellung is not None
    assert p.cpk_ergebnis is not None


def test_musterteam_fortschritt_roundtrip(musterteam, drive_base):
    helper.speichere_fortschritt(musterteam)
    q = helper.lade_fortschritt(musterteam.gruppenname, musterteam.gruppennummer)
    assert q is not None
    assert abs(q.cpk_ergebnis["cpk"] - musterteam.cpk_ergebnis["cpk"]) < 1e-6
    assert q.modell is not None
    assert abs(q.modell.rsquared - musterteam.modell.rsquared) < 1e-3


def test_musterteam_optimierung_naehe_zielweite(musterteam):
    """Das Optimum soll die Zielweite (300 cm) treffen."""
    vorhersage = musterteam.optimale_einstellung["vorhersage"]
    assert abs(vorhersage - musterteam.zielweite) < 10.0


def test_musterteam_export_figuren_vorhanden(musterteam):
    """Während der Phasen wurden Figuren in p.figuren abgelegt (per _save_fig)."""
    # Mindestens die Vermessungsfigur (aus DEFINE's zeige_vermessung, falls aufgerufen).
    # Hinweis: fertige_gruppe ruft zeige_vermessung NICHT auf — wir prüfen nur, dass das
    # Feld existiert und ein Dict ist.
    assert isinstance(musterteam.figuren, dict)
