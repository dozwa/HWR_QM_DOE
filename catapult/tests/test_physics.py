"""Tests fuer die Physik-Engine."""

import math

import pytest

from statapult.factors import ALL_FACTORS
from statapult.physics import (
    CatapultPhysics,
    compute_distance,
    compute_launch_info,
    simulate_shot,
)


class TestComputeDistance:
    """Tests fuer das kalibrierte Distanzmodell."""

    def test_center_point_range(self):
        """Center-point sollte im Bereich 300-400 cm liegen."""
        defaults = {k: f.default for k, f in ALL_FACTORS.items()}
        d = compute_distance(defaults)
        assert 300 <= d <= 400, f"Center distance {d:.1f} cm out of range"

    def test_min_distance_reachable(self):
        """Minimale Distanz (alle auf 'kurz') sollte nahe 100 cm sein."""
        settings = {k: f.low for k, f in ALL_FACTORS.items()}
        settings["ballgewicht"] = ALL_FACTORS["ballgewicht"].high  # schwer = kurz
        d = compute_distance(settings)
        assert 50 < d < 150, f"Min distance {d:.1f} cm out of range"

    def test_max_distance_reachable(self):
        """Maximale Distanz (alle auf 'weit') sollte nahe 600 cm sein."""
        settings = {k: f.high for k, f in ALL_FACTORS.items()}
        settings["ballgewicht"] = ALL_FACTORS["ballgewicht"].low  # leicht = weit
        d = compute_distance(settings)
        assert 500 < d < 700, f"Max distance {d:.1f} cm out of range"

    def test_abzugswinkel_positive_effect(self):
        """Hoeherer Abzugswinkel = laengere Distanz."""
        defaults = {k: f.default for k, f in ALL_FACTORS.items()}
        low = dict(defaults, abzugswinkel=ALL_FACTORS["abzugswinkel"].low)
        high = dict(defaults, abzugswinkel=ALL_FACTORS["abzugswinkel"].high)
        assert compute_distance(high) > compute_distance(low)

    def test_ballgewicht_negative_effect(self):
        """Schwererer Ball = kuerzere Distanz."""
        defaults = {k: f.default for k, f in ALL_FACTORS.items()}
        light = dict(defaults, ballgewicht=ALL_FACTORS["ballgewicht"].low)
        heavy = dict(defaults, ballgewicht=ALL_FACTORS["ballgewicht"].high)
        assert compute_distance(light) > compute_distance(heavy)

    def test_gummiband_positive_effect(self):
        """Hoehere Gummiband-Position = laengere Distanz."""
        defaults = {k: f.default for k, f in ALL_FACTORS.items()}
        low = dict(defaults, gummiband_position=ALL_FACTORS["gummiband_position"].low)
        high = dict(defaults, gummiband_position=ALL_FACTORS["gummiband_position"].high)
        assert compute_distance(high) > compute_distance(low)

    def test_pin_hoehe_positive_effect(self):
        """Hoehere Pin-Hoehe = laengere Distanz."""
        defaults = {k: f.default for k, f in ALL_FACTORS.items()}
        low = dict(defaults, pin_hoehe=ALL_FACTORS["pin_hoehe"].low)
        high = dict(defaults, pin_hoehe=ALL_FACTORS["pin_hoehe"].high)
        assert compute_distance(high) > compute_distance(low)

    def test_wind_positive_effect(self):
        """Rueckenwind = laengere Distanz."""
        defaults = {k: f.default for k, f in ALL_FACTORS.items()}
        headwind = dict(defaults, wind=ALL_FACTORS["wind"].low)
        tailwind = dict(defaults, wind=ALL_FACTORS["wind"].high)
        assert compute_distance(tailwind) > compute_distance(headwind)

    def test_never_negative(self):
        """Distanz darf nie negativ sein."""
        settings = {k: f.low for k, f in ALL_FACTORS.items()}
        d = compute_distance(settings)
        assert d >= 0

    def test_interactions_exist(self):
        """Wechselwirkungen sollten vorhanden sein (multiplikatives Modell)."""
        defaults = {k: f.default for k, f in ALL_FACTORS.items()}
        f1, f2 = "abzugswinkel", "gummiband_position"

        ll = dict(defaults, **{f1: ALL_FACTORS[f1].low, f2: ALL_FACTORS[f2].low})
        lh = dict(defaults, **{f1: ALL_FACTORS[f1].low, f2: ALL_FACTORS[f2].high})
        hl = dict(defaults, **{f1: ALL_FACTORS[f1].high, f2: ALL_FACTORS[f2].low})
        hh = dict(defaults, **{f1: ALL_FACTORS[f1].high, f2: ALL_FACTORS[f2].high})

        interaction = (
            compute_distance(hh) - compute_distance(hl)
            - compute_distance(lh) + compute_distance(ll)
        ) / 4.0
        assert abs(interaction) > 1.0, "Interaction too small"


class TestSimulateShot:
    """Tests fuer die Gesamtsimulation."""

    def test_returns_three_values(self):
        """simulate_shot gibt (distance, launch, trajectory) zurueck."""
        d, l, t = simulate_shot(150, 90, 13, 15, 13, 10, 0)
        assert isinstance(d, float)
        assert d > 0

    def test_launch_info_plausible(self):
        """Launch-Infos sollten physikalisch plausibel sein."""
        d, l, t = simulate_shot(150, 90, 13, 15, 13, 10, 0)
        assert l.spring_energy_j >= 0
        assert l.ball_speed_m_s >= 0
        assert 0 < l.release_angle_deg < 90
        assert l.release_height_m >= 0


class TestComputeLaunchInfo:
    """Tests fuer die Verbose-Infos."""

    def test_release_angle_range(self):
        """Abwurfwinkel sollte im Bereich 20-50 Grad liegen."""
        for stopp in [70, 90, 110]:
            settings = {
                "abzugswinkel": 150, "stoppwinkel": stopp,
                "gummiband_position": 13, "becherposition": 15,
                "pin_hoehe": 13, "ballgewicht": 10,
            }
            launch = compute_launch_info(settings)
            assert 15 <= launch.release_angle_deg <= 55

    def test_energy_increases_with_pullback(self):
        """Mehr Rueckzug = mehr Federenergie."""
        base = {
            "stoppwinkel": 90, "gummiband_position": 13,
            "becherposition": 15, "pin_hoehe": 13, "ballgewicht": 10,
        }
        e_low = compute_launch_info({**base, "abzugswinkel": 130}).spring_energy_j
        e_high = compute_launch_info({**base, "abzugswinkel": 170}).spring_energy_j
        assert e_high > e_low
