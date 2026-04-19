"""Tests fuer Faktor-Definitionen."""

import pytest

from statapult.factors import (
    ALL_FACTORS,
    ENVIRONMENTAL_FACTORS,
    STANDARD_FACTORS,
    Factor,
    get_defaults,
    validate_settings,
)


class TestFactorDefinitions:
    def test_five_standard_factors(self):
        assert len(STANDARD_FACTORS) == 5

    def test_two_environmental_factors(self):
        assert len(ENVIRONMENTAL_FACTORS) == 2

    def test_all_factors_combined(self):
        assert len(ALL_FACTORS) == 7

    def test_low_less_than_high(self):
        for key, f in ALL_FACTORS.items():
            assert f.low < f.high, f"Factor {key}: low >= high"

    def test_default_in_range(self):
        for key, f in ALL_FACTORS.items():
            assert f.low <= f.default <= f.high, (
                f"Factor {key}: default {f.default} out of [{f.low}, {f.high}]"
            )


class TestFactorCoding:
    def test_low_codes_to_minus_one(self):
        for key, f in ALL_FACTORS.items():
            assert f.coded(f.low) == pytest.approx(-1.0)

    def test_high_codes_to_plus_one(self):
        for key, f in ALL_FACTORS.items():
            assert f.coded(f.high) == pytest.approx(1.0)

    def test_roundtrip(self):
        for key, f in ALL_FACTORS.items():
            for val in [f.low, f.default, f.high]:
                assert f.natural(f.coded(val)) == pytest.approx(val)


class TestHelperFormat:
    def test_format_has_required_keys(self):
        for f in ALL_FACTORS.values():
            d = f.to_helper_format()
            assert "name" in d
            assert "einheit" in d
            assert "low" in d
            assert "high" in d

    def test_format_values_match(self):
        f = STANDARD_FACTORS["abzugswinkel"]
        d = f.to_helper_format()
        assert d["low"] == f.low
        assert d["high"] == f.high


class TestValidateSettings:
    def test_fills_missing_with_defaults(self):
        result = validate_settings({})
        defaults = get_defaults()
        for key in ALL_FACTORS:
            assert result[key] == defaults[key]

    def test_clamps_out_of_range(self):
        result = validate_settings({"abzugswinkel": 999})
        assert result["abzugswinkel"] == ALL_FACTORS["abzugswinkel"].high

    def test_keeps_valid_values(self):
        result = validate_settings({"abzugswinkel": 160})
        assert result["abzugswinkel"] == 160
