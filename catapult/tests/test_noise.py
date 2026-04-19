"""Tests fuer das Rauschmodell."""

import math

import numpy as np
import pytest

from statapult.noise import NoiseModel


class TestNoiseModel:
    def test_sigma_total(self):
        """sigma_total sollte die Quadratwurzel der Varianz-Summe sein."""
        nm = NoiseModel()
        expected = math.sqrt(
            nm.sigma_measurement**2 + nm.sigma_setup**2
            + nm.sigma_rubber**2 + nm.sigma_release**2
            + nm.sigma_wind_turbulence**2
        )
        assert abs(nm.sigma_total - expected) < 1e-10

    def test_zero_mean(self):
        """Rauschen sollte im Mittel null sein."""
        nm = NoiseModel()
        rng = np.random.default_rng(42)
        samples = [nm.total_noise(rng) for _ in range(10000)]
        mean = np.mean(samples)
        assert abs(mean) < 0.2, f"Noise mean {mean:.4f} not close to zero"

    def test_std_matches_sigma(self):
        """Empirische Std.Abw. sollte sigma_total entsprechen."""
        nm = NoiseModel()
        rng = np.random.default_rng(42)
        samples = [nm.total_noise(rng) for _ in range(10000)]
        std = np.std(samples)
        assert abs(std - nm.sigma_total) < 0.2, (
            f"Empirical std {std:.2f} vs sigma_total {nm.sigma_total:.2f}"
        )

    def test_reproducibility_with_seed(self):
        """Gleicher Seed = gleiche Ergebnisse."""
        nm = NoiseModel()
        rng1 = np.random.default_rng(123)
        rng2 = np.random.default_rng(123)
        for _ in range(10):
            assert nm.total_noise(rng1) == nm.total_noise(rng2)

    def test_different_seeds_differ(self):
        """Verschiedene Seeds = verschiedene Ergebnisse."""
        nm = NoiseModel()
        rng1 = np.random.default_rng(1)
        rng2 = np.random.default_rng(2)
        v1 = nm.total_noise(rng1)
        v2 = nm.total_noise(rng2)
        assert v1 != v2


class TestOperatorBias:
    def test_consistent_bias(self):
        """Gleicher Operator bekommt gleichen Bias."""
        nm = NoiseModel()
        rng = np.random.default_rng(42)
        b1 = nm.get_operator_bias("Op_A", rng)
        b2 = nm.get_operator_bias("Op_A", rng)
        assert b1 == b2

    def test_different_operators_differ(self):
        """Verschiedene Operatoren bekommen verschiedene Biases."""
        nm = NoiseModel()
        rng = np.random.default_rng(42)
        b1 = nm.get_operator_bias("Op_A", rng)
        b2 = nm.get_operator_bias("Op_B", rng)
        assert b1 != b2

    def test_reset_clears_biases(self):
        """reset_operators loescht die gespeicherten Biases."""
        nm = NoiseModel()
        rng = np.random.default_rng(42)
        nm.get_operator_bias("Op_A", rng)
        nm.reset_operators()
        assert len(nm._operator_biases) == 0


class TestDrift:
    def test_zero_drift_by_default(self):
        """Standard-Drift-Rate ist 0."""
        nm = NoiseModel()
        assert nm.apply_drift(100) == 0.0

    def test_positive_drift(self):
        """Drift steigt mit Schussnummer."""
        nm = NoiseModel(drift_rate=0.1)
        assert nm.apply_drift(0) == 0.0
        assert nm.apply_drift(10) == pytest.approx(1.0)
        assert nm.apply_drift(50) == pytest.approx(5.0)
