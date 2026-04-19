"""Tests fuer das CLI-Interface."""

import json
import subprocess
import sys

import pytest


def run_cli(*args):
    """Fuehrt statapult CLI als Subprocess aus."""
    result = subprocess.run(
        [sys.executable, "-m", "statapult", *args],
        capture_output=True, text=True,
    )
    return result


class TestShootCommand:
    def test_basic_shoot(self):
        r = run_cli("shoot", "--seed", "42")
        assert r.returncode == 0
        assert "Wurfweite:" in r.stdout

    def test_verbose_output(self):
        r = run_cli("shoot", "--seed", "42", "--verbose")
        assert r.returncode == 0
        assert "Statapult Simulator" in r.stdout
        assert "Federenergie:" in r.stdout

    def test_json_output(self):
        r = run_cli("shoot", "--seed", "42", "--format", "json")
        assert r.returncode == 0
        data = json.loads(r.stdout)
        assert "wurfweite_cm" in data

    def test_repeat(self):
        r = run_cli("shoot", "--seed", "42", "--repeat", "5")
        assert r.returncode == 0
        assert "Mittelwert:" in r.stdout

    def test_custom_factors(self):
        r = run_cli(
            "shoot", "--abzugswinkel", "170", "--stoppwinkel", "110",
            "--seed", "42",
        )
        assert r.returncode == 0


class TestInfoCommand:
    def test_shows_factors(self):
        r = run_cli("info")
        assert r.returncode == 0
        assert "Abzugswinkel" in r.stdout
        assert "Stoppwinkel" in r.stdout


class TestMSACommand:
    def test_msa_output(self):
        r = run_cli(
            "msa", "--seed", "42", "--operators", "2",
            "--measurements-per-operator", "3",
        )
        assert r.returncode == 0
        assert "Operator" in r.stdout
        lines = r.stdout.strip().split("\n")
        assert len(lines) == 7  # header + 2*3 data rows


class TestControlCommand:
    def test_control_output(self):
        r = run_cli("control", "--seed", "42", "--shots", "5")
        assert r.returncode == 0
        assert "Schuss" in r.stdout
        lines = r.stdout.strip().split("\n")
        assert len(lines) == 6  # header + 5 data rows
