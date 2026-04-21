"""Factories for building `helper.Projekt` instances populated by statapult.

Each factory is independently callable so phase-specific tests can request
only what they need. The profile preset controls the catapult character
(precise, typical, noisy, drifting).
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional

import numpy as np
import pandas as pd

import helper
from statapult import Statapult


# ─────────────────────────────────────────────────────────────────────
# Profile presets — used by the four E2E groups.
# Values are the noise multipliers / drift values passed to statapult.
# ─────────────────────────────────────────────────────────────────────
@dataclass(frozen=True)
class Profile:
    name: str
    noise_level: float
    drift: float = 0.0
    seed: int = 0


PROFILES: Dict[str, Profile] = {
    "praezision":   Profile(name="praezision",   noise_level=0.3,  seed=1001),
    "typisch":      Profile(name="typisch",      noise_level=1.0,  seed=1002),
    "streu":        Profile(name="streu",        noise_level=2.0,  seed=1003),
    "drift":        Profile(name="drift",        noise_level=1.0,  drift=0.15, seed=1004),
}


# Three factors with helper.py-compatible dicts. These are the factors the
# notebook exposes (Winkel/Spannung/Ballposition) mapped onto statapult keys.
FAKTOREN_KATALOG: List[Dict] = [
    {"name": "Abzugswinkel",       "einheit": "Grad", "low": 140, "high": 165,
     "centerpoint_moeglich": True, "_statapult_key": "abzugswinkel"},
    {"name": "Gummiband-Position", "einheit": "cm",   "low": 10,  "high": 17,
     "centerpoint_moeglich": True, "_statapult_key": "gummiband_position"},
    {"name": "Becherposition",     "einheit": "cm",   "low": 10,  "high": 20,
     "centerpoint_moeglich": True, "_statapult_key": "becherposition"},
]


def _clean_factor(f: Dict) -> Dict:
    """Return a helper-compatible copy (strip the _statapult_key)."""
    return {k: v for k, v in f.items() if not k.startswith("_")}


def _settings_from_coded(coded_row: pd.Series, faktoren: List[Dict]) -> Dict[str, float]:
    """Map a coded DoE row onto statapult natural settings."""
    settings = {}
    for f in faktoren:
        key = f["_statapult_key"]
        # Coded columns in the plan are either `{name}_coded` or `{name} (kodiert)`.
        coded_col = next((c for c in coded_row.index
                          if c == f"{f['name']}_coded" or c == f"{f['name']} (kodiert)"),
                         None)
        if coded_col is None:
            raise KeyError(f"Keine kodierte Spalte für {f['name']} in {list(coded_row.index)}")
        coded = float(coded_row[coded_col])
        center = (f["low"] + f["high"]) / 2
        half = (f["high"] - f["low"]) / 2
        settings[key] = center + coded * half
    return settings


# ─────────────────────────────────────────────────────────────────────
# Phase factories
# ─────────────────────────────────────────────────────────────────────

def neue_gruppe(profile: str = "typisch", zielweite: float = 300.0,
                toleranz: float = 30.0) -> helper.Projekt:
    """Create an empty Projekt with the three Standard-Faktoren already defined."""
    p = helper.init_projekt(f"Test_{profile}", 1, zielweite=zielweite, toleranz=toleranz)
    p.faktoren = [_clean_factor(f) for f in FAKTOREN_KATALOG]
    return p


def baue_define(profile: str = "typisch", **kwargs) -> helper.Projekt:
    """DEFINE: init + Faktoren + Min/Max-Vermessung + Testwürfe + Charter."""
    prof = PROFILES[profile]
    p = neue_gruppe(profile=profile, **kwargs)

    # Min/Max-Vermessung: Katapult auf Low/High-Stufen aller Faktoren.
    kat = Statapult(seed=prof.seed)
    min_settings = {f["_statapult_key"]: f["low"] for f in FAKTOREN_KATALOG}
    max_settings = {f["_statapult_key"]: f["high"] for f in FAKTOREN_KATALOG}
    min_wuerfe = [kat.shoot(min_settings, noise_level=prof.noise_level).wurfweite_cm
                  for _ in range(3)]
    max_wuerfe = [kat.shoot(max_settings, noise_level=prof.noise_level).wurfweite_cm
                  for _ in range(3)]
    helper.speichere_vermessung(
        p,
        min_wuerfe=min_wuerfe,
        max_wuerfe=max_wuerfe,
        min_einstellung={f["name"]: f["low"] for f in FAKTOREN_KATALOG},
        max_einstellung={f["name"]: f["high"] for f in FAKTOREN_KATALOG},
        beschreibung=f"Statapult Profil {profile}",
    )

    # Testwürfe bei typischer (Mittel-)Einstellung, 5 Würfe
    mid_settings = {f["_statapult_key"]: (f["low"] + f["high"]) / 2 for f in FAKTOREN_KATALOG}
    p.testwuerfe = np.array([kat.shoot(mid_settings, noise_level=prof.noise_level).wurfweite_cm
                              for _ in range(5)])

    # Projektcharter
    p.charter = {
        "Gruppenname": p.gruppenname,
        "Zielweite": f"{p.zielweite:.0f} cm ± {p.toleranz:.0f} cm",
        "Problemstellung": "Katapult reproduzierbar auf Zielweite bringen",
        "Projektziel": f"Cpk ≥ 1.0 bei Zielweite {p.zielweite:.0f} cm",
        "Scope": "Optimierung der Faktoreinstellungen",
    }
    return p


def baue_measure(projekt: Optional[helper.Projekt] = None,
                 profile: str = "typisch",
                 operator_bias_on: bool = False) -> helper.Projekt:
    """MEASURE: MSA (Type-1 + Gage R&R) und Baseline."""
    prof = PROFILES[profile]
    p = projekt if projekt is not None else baue_define(profile=profile)

    # Type-1: 25 Messungen auf derselben Referenz-Position durch eine Person.
    kat = Statapult(seed=prof.seed + 100)
    ref_settings = {f["_statapult_key"]: (f["low"] + f["high"]) / 2 for f in FAKTOREN_KATALOG}
    ref_wuerfe = [kat.shoot(ref_settings, noise_level=prof.noise_level).wurfweite_cm
                  for _ in range(25)]
    referenzwert = float(np.mean(ref_wuerfe))
    type1_df = pd.DataFrame({"Person_A": ref_wuerfe})
    p.msa_type1 = helper.analysiere_type1(type1_df, referenzwert=referenzwert)

    # Gage R&R: 5 Wurf-IDs × 3 Personen × 2 Wiederholungen (30 Messwerte).
    # Jedes Teil ist ein echter, unterschiedlich weit gelandeter Auftreffpunkt,
    # damit die Teil-zu-Teil-Variation deutlich größer als die Messunsicherheit
    # ist. Mit ``operator_bias_on`` wird systematische Reproduzierbarkeitslücke
    # addiert — das sollte %GRR nach oben treiben.
    kat_msa = Statapult(seed=prof.seed + 200)
    # 5 unterschiedliche Auftreffpunkte entlang des Kennfelds erzeugen.
    teile_settings = [
        {f["_statapult_key"]: f["low"] for f in FAKTOREN_KATALOG},
        {f["_statapult_key"]: f["low"] + 0.25 * (f["high"] - f["low"]) for f in FAKTOREN_KATALOG},
        {f["_statapult_key"]: (f["low"] + f["high"]) / 2 for f in FAKTOREN_KATALOG},
        {f["_statapult_key"]: f["low"] + 0.75 * (f["high"] - f["low"]) for f in FAKTOREN_KATALOG},
        {f["_statapult_key"]: f["high"] for f in FAKTOREN_KATALOG},
    ]
    teil_true_values = [kat_msa.shoot(s, noise_level=0.0).wurfweite_cm for s in teile_settings]

    rng = np.random.default_rng(prof.seed + 201)
    rows = []
    for teil_nr, true_val in enumerate(teil_true_values, start=1):
        for person in ("Anna", "Ben", "Cara"):
            for _ in range(2):
                if operator_bias_on:
                    bias = {"Anna": -8.0, "Ben": 0.0, "Cara": 8.0}[person]
                else:
                    bias = 0.0
                # Repeatability-Noise 0.7 cm — winzig gegen Teile-Spanne (~300 cm).
                messwert = true_val + rng.normal(0, 0.7) + bias
                rows.append({"Wurf_ID": teil_nr, "Person": person, "Messwert": messwert})
    grr_df = pd.DataFrame(rows)
    p.msa_grr = helper.analysiere_gage_rr(grr_df)
    p.msa_rohdaten = grr_df

    # Baseline: 15 Würfe bei typischer Einstellung.
    kat_base = Statapult(seed=prof.seed + 300)
    p.baseline_wuerfe = np.array([kat_base.shoot(ref_settings, noise_level=prof.noise_level)
                                   .wurfweite_cm for _ in range(15)])
    p.baseline_stats = helper.analysiere_baseline(p.baseline_wuerfe)
    return p


def baue_analyze(projekt: Optional[helper.Projekt] = None,
                 profile: str = "typisch",
                 design: str = "voll",
                 wiederholungen: int = 2,
                 centerpoints: int = 3) -> helper.Projekt:
    """ANALYZE: Versuchsplan + Statapult-Batch + Regression + Pruning."""
    prof = PROFILES[profile]
    p = projekt if projekt is not None else baue_measure(profile=profile)

    # Alle 3 Faktoren aktiv mit CP für DoE.
    p.faktoren_doe = [_clean_factor(f) for f in FAKTOREN_KATALOG]

    p.versuchsplan = helper.generiere_versuchsplan(
        p.faktoren_doe, wiederholungen=wiederholungen, centerpoints=centerpoints,
        seed=prof.seed + 400, design=design,
    )
    p.versuchsplan_config = {
        "wiederholungen": wiederholungen,
        "blocking": False,
        "centerpoints": centerpoints,
        "design": design,
    }

    # Statapult-Batch: pro Zeile aus kodierten Stufen die Settings bauen.
    kat = Statapult(seed=prof.seed + 500)
    ergebnisse = []
    for _, row in p.versuchsplan.iterrows():
        settings = _settings_from_coded(row, FAKTOREN_KATALOG)
        ergebnisse.append(kat.shoot(settings, noise_level=prof.noise_level).wurfweite_cm)
    p.doe_ergebnisse = p.versuchsplan.copy()
    p.doe_ergebnisse["Ergebnis: Weite (cm)"] = ergebnisse

    # Modell + Pruning
    p.modell = helper.fitte_modell(p.doe_ergebnisse, p.faktoren_doe)
    p.modell_gepruned, p.pruning_log = helper.hierarchisches_pruning(p.modell)
    p.modell = p.modell_gepruned
    return p


def baue_improve(projekt: Optional[helper.Projekt] = None,
                 profile: str = "typisch",
                 strategie: str = "dual") -> helper.Projekt:
    """IMPROVE: Optimierung + Konfirmation (10 Würfe am Optimum)."""
    prof = PROFILES[profile]
    p = projekt if projekt is not None else baue_analyze(profile=profile)

    p.optimale_einstellung = helper.optimiere_einstellungen(
        p.modell, zielweite=p.zielweite, faktoren=p.faktoren_doe, strategie=strategie,
    )
    p.optimierung_config = {"strategie": strategie, "lambda_gewicht": 0.01}

    # Konfirmation: 10 Würfe am optimalen Punkt.
    kat = Statapult(seed=prof.seed + 600)
    opt_settings = {}
    for f in FAKTOREN_KATALOG:
        opt_settings[f["_statapult_key"]] = p.optimale_einstellung["einstellungen"][f["name"]]["original"]
    p.konfirmation_wuerfe = np.array([kat.shoot(opt_settings, noise_level=prof.noise_level)
                                       .wurfweite_cm for _ in range(10)])
    p.konfirmation_ergebnis = helper.analysiere_konfirmation(
        p.konfirmation_wuerfe,
        vorhersage=p.optimale_einstellung["vorhersage"],
        pi_low=p.optimale_einstellung["pi_low"],
        pi_high=p.optimale_einstellung["pi_high"],
        zielweite=p.zielweite, toleranz=p.toleranz,
    )
    return p


def baue_control(projekt: Optional[helper.Projekt] = None,
                 profile: str = "typisch") -> helper.Projekt:
    """CONTROL: I-MR + Cpk auf den Konfirmationswürfen.

    Für das `drift`-Profil werden die Konfirmationswürfe um einen linearen
    Drift ergänzt, damit die Regelkarte eine Verletzung zeigen kann.
    """
    prof = PROFILES[profile]
    p = projekt if projekt is not None else baue_improve(profile=profile)

    if prof.drift > 0:
        # Overlay eines linearen Drifts auf die Konfirmationswürfe.
        n = len(p.konfirmation_wuerfe)
        drift_offsets = np.arange(n) * (prof.drift * 20.0)  # bis +30 cm am Ende
        p.konfirmation_wuerfe = p.konfirmation_wuerfe + drift_offsets

    p.imr_ergebnis = helper.berechne_imr(p.konfirmation_wuerfe)
    usl = p.zielweite + p.toleranz
    lsl = p.zielweite - p.toleranz
    p.cpk_ergebnis = helper.berechne_cpk(p.konfirmation_wuerfe, usl, lsl)
    return p


def fertige_gruppe(profile: str = "typisch") -> helper.Projekt:
    """Convenience: full DEFINE → CONTROL chain in one call."""
    p = baue_define(profile=profile)
    p = baue_measure(p, profile=profile)
    p = baue_analyze(p, profile=profile)
    p = baue_improve(p, profile=profile)
    p = baue_control(p, profile=profile)
    return p
