"""End-to-End Integrationstest: DOE -> Modell -> Optimierung -> Konfirmation.

Prueft ob der komplette DMAIC-Workflow mit dem virtuellen Katapult
zuverlaessig funktioniert: Versuchsplan ausfuehren, Modell fitten,
Einstellungen optimieren, Konfirmation bestehen.
"""

import os
import sys

import numpy as np
import pytest

# helper.py aus dem uebergeordneten Verzeichnis importieren
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))
os.environ.setdefault("MPLBACKEND", "Agg")

from statapult import Statapult, STANDARD_FACTORS


def _run_scenario(faktoren, zielweite, seed_doe=42, seed_konf=None):
    """Fuehrt ein komplettes DOE-Szenario durch.

    Returns
    -------
    dict mit: vorhersage, konfirmation_mean, abweichung, cpk, in_toleranz_pct, ok
    """
    from helper import (
        generiere_versuchsplan, fitte_modell, hierarchisches_pruning,
        optimiere_einstellungen, analysiere_konfirmation, berechne_cpk,
    )

    if seed_konf is None:
        seed_konf = zielweite * 13 + 7

    plan = generiere_versuchsplan(
        faktoren, wiederholungen=3, centerpoints=3, seed=seed_doe
    )
    katapult = Statapult(seed=seed_doe)
    ergebnisse = katapult.batch(plan)

    modell = fitte_modell(ergebnisse, faktoren)
    modell_p, _ = hierarchisches_pruning(modell)

    opt = optimiere_einstellungen(
        modell_p, zielweite, faktoren, strategie="mittelwert"
    )

    # Codes pruefen (Extrapolation?)
    codes = [d["coded"] for d in opt["einstellungen"].values()]
    in_design = all(abs(c) <= 1.05 for c in codes)

    # Optimale Settings in Simulator-Format
    opt_settings = {}
    for name, data in opt["einstellungen"].items():
        for key, factor in STANDARD_FACTORS.items():
            if factor.name.lower() in name.lower():
                opt_settings[key] = data["original"]
                break

    # Konfirmation: 15 Schuesse
    k2 = Statapult(seed=seed_konf)
    konf = np.array([k2.shoot(opt_settings).wurfweite_cm for _ in range(15)])

    toleranz = 15.0
    result = analysiere_konfirmation(
        konf, opt["vorhersage"],
        opt["pi_low"], opt["pi_high"],
        zielweite, toleranz,
    )
    cpk = berechne_cpk(konf, zielweite + toleranz, zielweite - toleranz)

    return {
        "vorhersage": opt["vorhersage"],
        "konfirmation_mean": np.mean(konf),
        "abweichung": abs(np.mean(konf) - zielweite),
        "cpk": cpk["cpk"],
        "in_toleranz_pct": result["pct_in_tol"],
        "ok": "\u2705" in result["bewertung"],
        "in_design": in_design,
        "r2_adj": modell_p.rsquared_adj,
    }


# ======================================================================
# Faktor-Definitionen
# ======================================================================

FAKTOREN_3F_ABZ_GUM_BECH = [
    {"name": "Abzugswinkel", "einheit": "Grad", "low": 130, "high": 170},
    {"name": "Gummiband-Position", "einheit": "cm", "low": 8, "high": 18},
    {"name": "Becherposition", "einheit": "cm", "low": 8, "high": 22},
]

FAKTOREN_3F_ABZ_STO_GUM = [
    {"name": "Abzugswinkel", "einheit": "Grad", "low": 130, "high": 170},
    {"name": "Stoppwinkel", "einheit": "Grad", "low": 70, "high": 110},
    {"name": "Gummiband-Position", "einheit": "cm", "low": 8, "high": 18},
]

FAKTOREN_4F = [
    {"name": "Abzugswinkel", "einheit": "Grad", "low": 130, "high": 170},
    {"name": "Stoppwinkel", "einheit": "Grad", "low": 70, "high": 110},
    {"name": "Gummiband-Position", "einheit": "cm", "low": 8, "high": 18},
    {"name": "Becherposition", "einheit": "cm", "low": 8, "high": 22},
]

FAKTOREN_5F = [
    {"name": "Abzugswinkel", "einheit": "Grad", "low": 130, "high": 170},
    {"name": "Stoppwinkel", "einheit": "Grad", "low": 70, "high": 110},
    {"name": "Gummiband-Position", "einheit": "cm", "low": 8, "high": 18},
    {"name": "Becherposition", "einheit": "cm", "low": 8, "high": 22},
    {"name": "Pin-Hoehe", "einheit": "cm", "low": 8, "high": 18},
]


# ======================================================================
# Tests: 5 Faktoren (voller Bereich)
# ======================================================================

class TestOptimierung5F:
    """5 Faktoren: sollte alle Zielweiten 200-500 cm treffen."""

    @pytest.mark.parametrize("zielweite", [200, 280, 350, 420, 500])
    def test_zielweite_erreichbar(self, zielweite):
        r = _run_scenario(FAKTOREN_5F, zielweite)
        assert r["ok"], (
            f"5F, Ziel {zielweite}: Abw={r['abweichung']:.1f} cm, "
            f"Cpk={r['cpk']:.2f}, Konf={r['konfirmation_mean']:.1f}"
        )

    @pytest.mark.parametrize("zielweite", [200, 280, 350, 420, 500])
    def test_abweichung_unter_5cm(self, zielweite):
        r = _run_scenario(FAKTOREN_5F, zielweite)
        assert r["abweichung"] < 5.0, (
            f"5F, Ziel {zielweite}: Abw={r['abweichung']:.1f} cm zu gross"
        )

    @pytest.mark.parametrize("zielweite", [200, 280, 350, 420, 500])
    def test_cpk_ueber_1(self, zielweite):
        r = _run_scenario(FAKTOREN_5F, zielweite)
        assert r["cpk"] > 1.0, (
            f"5F, Ziel {zielweite}: Cpk={r['cpk']:.2f} unter 1.0"
        )


# ======================================================================
# Tests: 4 Faktoren
# ======================================================================

class TestOptimierung4F:
    """4 Faktoren: sollte alle Zielweiten 200-500 cm treffen."""

    @pytest.mark.parametrize("zielweite", [200, 280, 350, 420, 500])
    def test_zielweite_erreichbar(self, zielweite):
        r = _run_scenario(FAKTOREN_4F, zielweite)
        assert r["ok"], (
            f"4F, Ziel {zielweite}: Abw={r['abweichung']:.1f} cm, "
            f"Cpk={r['cpk']:.2f}"
        )


# ======================================================================
# Tests: 3 Faktoren (innerhalb erreichbarem Bereich)
# ======================================================================

class TestOptimierung3F:
    """3 Faktoren: sollte Zielweiten innerhalb des DOE-Bereichs treffen."""

    @pytest.mark.parametrize("zielweite", [280, 350, 420, 500])
    def test_abz_gum_bech(self, zielweite):
        """3F Abzugswinkel/Gummiband/Becher: Range ~226-508 cm."""
        r = _run_scenario(FAKTOREN_3F_ABZ_GUM_BECH, zielweite)
        assert r["ok"], (
            f"3F Abz/Gum/Bech, Ziel {zielweite}: Abw={r['abweichung']:.1f} cm"
        )

    @pytest.mark.parametrize("zielweite", [280, 350, 420, 500])
    def test_abz_sto_gum(self, zielweite):
        """3F Abzugswinkel/Stoppwinkel/Gummiband: Range ~216-518 cm."""
        r = _run_scenario(FAKTOREN_3F_ABZ_STO_GUM, zielweite)
        assert r["ok"], (
            f"3F Abz/Sto/Gum, Ziel {zielweite}: Abw={r['abweichung']:.1f} cm"
        )


# ======================================================================
# Tests: Modellqualitaet
# ======================================================================

class TestModellqualitaet:
    """Prueft ob die DOE-Modelle hohe Guete haben."""

    @pytest.mark.parametrize("faktoren,label", [
        (FAKTOREN_3F_ABZ_GUM_BECH, "3F"),
        (FAKTOREN_4F, "4F"),
        (FAKTOREN_5F, "5F"),
    ])
    def test_r2_adj_hoch(self, faktoren, label):
        r = _run_scenario(faktoren, 350)
        assert r["r2_adj"] > 0.98, (
            f"{label}: R2_adj={r['r2_adj']:.4f} zu niedrig"
        )

    def test_5f_praezision_konsistent(self):
        """5F-Modell sollte bei verschiedenen Seeds konsistent treffen."""
        abweichungen = []
        for seed in [42, 123, 456, 789, 1000]:
            r = _run_scenario(FAKTOREN_5F, 350, seed_doe=seed, seed_konf=seed + 1)
            abweichungen.append(r["abweichung"])
        assert np.mean(abweichungen) < 5.0, (
            f"Mittlere Abweichung {np.mean(abweichungen):.1f} cm ueber Seeds"
        )
        assert max(abweichungen) < 10.0, (
            f"Max Abweichung {max(abweichungen):.1f} cm bei einem Seed"
        )
