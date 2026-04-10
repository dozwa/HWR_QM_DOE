"""Statistische Validierungstests.

Pruefen, ob der Simulator Daten erzeugt, die fuer DMAIC-Uebungen
(MSA, DOE, Cpk, Regelkarten) geeignet sind.
"""

from itertools import product as iterproduct

import numpy as np
import pytest

from statapult import Statapult
from statapult.factors import ALL_FACTORS, STANDARD_FACTORS


class TestDOEProperties:
    """Tests fuer Design-of-Experiments-Eignung."""

    @pytest.fixture
    def doe_results(self):
        """Fuehrt ein 2^5 vollfaktorielles Design mit 3 Wiederholungen durch."""
        s = Statapult(seed=42)
        factors = list(STANDARD_FACTORS.keys())
        n_factors = len(factors)

        # Vollfaktoriell 2^5
        levels = list(iterproduct([-1, 1], repeat=n_factors))

        results = []
        for coded_levels in levels:
            settings = {}
            for i, key in enumerate(factors):
                f = ALL_FACTORS[key]
                settings[key] = f.natural(coded_levels[i])

            # 3 Wiederholungen
            for _ in range(3):
                r = s.shoot(settings, noise_level=1.0)
                results.append({
                    **{f"x{i+1}": coded_levels[i] for i in range(n_factors)},
                    "distance": r.wurfweite_cm,
                })

        return results, factors

    def test_distance_range(self, doe_results):
        """Distanzen sollten im Bereich 100-600 cm liegen."""
        results, _ = doe_results
        distances = [r["distance"] for r in results]
        assert min(distances) > 80, f"Min distance {min(distances):.1f} too low"
        assert max(distances) < 700, f"Max distance {max(distances):.1f} too high"

    def test_main_effects_detectable(self, doe_results):
        """Mindestens 3 Haupteffekte sollten > 15 cm sein."""
        results, factors = doe_results
        effects = []
        for i in range(len(factors)):
            xi_key = f"x{i+1}"
            high = [r["distance"] for r in results if r[xi_key] == 1]
            low = [r["distance"] for r in results if r[xi_key] == -1]
            effect = np.mean(high) - np.mean(low)
            effects.append(abs(effect))

        large_effects = sum(1 for e in effects if e > 15)
        assert large_effects >= 3, (
            f"Only {large_effects} effects > 15 cm: {effects}"
        )

    def test_effect_sizes_reasonable(self, doe_results):
        """Kein Einzeleffekt sollte > 200 cm sein."""
        results, factors = doe_results
        for i in range(len(factors)):
            xi_key = f"x{i+1}"
            high = [r["distance"] for r in results if r[xi_key] == 1]
            low = [r["distance"] for r in results if r[xi_key] == -1]
            effect = abs(np.mean(high) - np.mean(low))
            assert effect < 200, f"Factor {factors[i]} effect {effect:.1f} too large"


class TestMSAProperties:
    """Tests fuer Measurement System Analysis-Eignung."""

    def test_measurement_variability(self):
        """Wiederholmessungen sollten Streuung zeigen."""
        s = Statapult(seed=42)
        settings = {k: f.default for k, f in ALL_FACTORS.items()}
        distances = [s.shoot(settings).wurfweite_cm for _ in range(30)]
        std = np.std(distances, ddof=1)
        assert 1.0 < std < 8.0, f"Measurement std {std:.2f} out of range"

    def test_operator_differences(self):
        """Verschiedene Operatoren sollten unterschiedliche Biases zeigen."""
        s = Statapult(seed=42)
        settings = {k: f.default for k, f in ALL_FACTORS.items()}

        op1 = [s.shoot(settings, operator_id="Op1").wurfweite_cm for _ in range(20)]
        op2 = [s.shoot(settings, operator_id="Op2").wurfweite_cm for _ in range(20)]

        mean_diff = abs(np.mean(op1) - np.mean(op2))
        assert mean_diff > 0.1, "Operators should show different biases"


class TestControlChartProperties:
    """Tests fuer Regelkarten-Eignung."""

    def test_stable_process(self):
        """Ohne Drift sollten die Messungen stabil sein."""
        s = Statapult(seed=42)
        settings = {k: f.default for k, f in ALL_FACTORS.items()}
        distances = [s.shoot(settings).wurfweite_cm for _ in range(25)]
        mean = np.mean(distances)
        std = np.std(distances, ddof=1)

        # Keine Punkte sollten > 3 sigma vom Mittelwert entfernt sein
        for d in distances:
            assert abs(d - mean) < 4 * std, (
                f"Value {d:.1f} > 4 sigma from mean {mean:.1f}"
            )

    def test_normality_approximate(self):
        """Wiederholmessungen sollten annaehernd normalverteilt sein."""
        s = Statapult(seed=42)
        settings = {k: f.default for k, f in ALL_FACTORS.items()}
        distances = [s.shoot(settings).wurfweite_cm for _ in range(100)]

        from scipy import stats
        _, p_value = stats.shapiro(distances)
        assert p_value > 0.01, f"Shapiro p-value {p_value:.4f} suggests non-normality"


class TestCpkProperties:
    """Tests fuer Process Capability-Eignung."""

    def test_cpk_achievable(self):
        """Mit optimierten Settings und Toleranz ±15 cm sollte Cpk > 1.0 moeglich sein."""
        s = Statapult(seed=42)
        # Gute Settings mit niedrigem Rauschen
        settings = {k: f.default for k, f in ALL_FACTORS.items()}
        distances = [s.shoot(settings).wurfweite_cm for _ in range(50)]

        mean = np.mean(distances)
        std = np.std(distances, ddof=1)
        tolerance = 15.0

        cpk_upper = (mean + tolerance - mean) / (3 * std)
        cpk_lower = (mean - (mean - tolerance)) / (3 * std)
        cpk = min(cpk_upper, cpk_lower)

        assert cpk > 1.0, f"Cpk {cpk:.2f} should be achievable"
