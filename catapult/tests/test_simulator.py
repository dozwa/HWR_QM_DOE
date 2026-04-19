"""Tests fuer den Simulator-Orchestrator."""

import pytest

from statapult import Statapult, ShotResult


class TestStatapult:
    def test_default_construction(self):
        s = Statapult(seed=42)
        assert s._shot_count == 0

    def test_shoot_returns_result(self):
        s = Statapult(seed=42)
        result = s.shoot({"abzugswinkel": 150})
        assert isinstance(result, ShotResult)
        assert result.wurfweite_cm > 0

    def test_reproducible_with_seed(self):
        s1 = Statapult(seed=42)
        s2 = Statapult(seed=42)
        r1 = s1.shoot({"abzugswinkel": 160})
        r2 = s2.shoot({"abzugswinkel": 160})
        assert r1.wurfweite_cm == r2.wurfweite_cm

    def test_different_seeds_differ(self):
        s1 = Statapult(seed=1)
        s2 = Statapult(seed=2)
        r1 = s1.shoot({"abzugswinkel": 160})
        r2 = s2.shoot({"abzugswinkel": 160})
        assert r1.wurfweite_cm != r2.wurfweite_cm

    def test_deterministic_mode(self):
        """noise_level=0 gibt deterministische Ergebnisse."""
        s = Statapult(seed=42)
        r1 = s.shoot({"abzugswinkel": 160}, noise_level=0)
        s.reset(seed=42)
        r2 = s.shoot({"abzugswinkel": 160}, noise_level=0)
        assert r1.wurfweite_cm == r2.wurfweite_cm
        assert r1.noise_cm == 0.0

    def test_shot_counter(self):
        s = Statapult(seed=42)
        s.shoot({})
        s.shoot({})
        assert s._shot_count == 2

    def test_reset(self):
        s = Statapult(seed=42)
        s.shoot({})
        s.shoot({})
        s.reset(seed=42)
        assert s._shot_count == 0


class TestShootMultiple:
    def test_returns_list(self):
        s = Statapult(seed=42)
        results = s.shoot_multiple({"abzugswinkel": 150}, n=5)
        assert len(results) == 5
        assert all(isinstance(r, ShotResult) for r in results)

    def test_variability(self):
        """Mehrfachschuesse sollten unterschiedliche Ergebnisse liefern."""
        s = Statapult(seed=42)
        results = s.shoot_multiple({"abzugswinkel": 150}, n=10)
        distances = [r.wurfweite_cm for r in results]
        assert len(set(distances)) > 1, "All shots identical"


class TestBatchMode:
    def test_batch_adds_result_column(self):
        import pandas as pd
        s = Statapult(seed=42)
        plan = pd.DataFrame({
            "abzugswinkel": [130, 170],
            "stoppwinkel": [90, 90],
            "gummiband_position": [13, 13],
            "becherposition": [15, 15],
            "pin_hoehe": [13, 13],
        })
        result = s.batch(plan)
        assert "Ergebnis" in result.columns
        assert len(result) == 2

    def test_batch_higher_pullback_longer(self):
        """Im Batch-Modus: hoeherer Abzugswinkel = laengere Distanz."""
        import pandas as pd
        s = Statapult(seed=42)
        plan = pd.DataFrame({
            "abzugswinkel": [130, 170],
            "stoppwinkel": [90, 90],
            "gummiband_position": [13, 13],
            "becherposition": [15, 15],
            "pin_hoehe": [13, 13],
        })
        result = s.batch(plan, noise_level=0)
        assert result["Ergebnis"].iloc[1] > result["Ergebnis"].iloc[0]


class TestToDict:
    def test_contains_key_fields(self):
        s = Statapult(seed=42)
        result = s.shoot({})
        d = result.to_dict()
        assert "wurfweite_cm" in d
        assert "settings" in d
        assert "physics" in d
