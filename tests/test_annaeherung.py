"""Tests für die neuen DEFINE-Helfer: Annäherung und initiale Einstellung."""
from __future__ import annotations

import numpy as np
import pytest

import helper
from tests import virtuelle_gruppe as vg


# ─────────────────────────────────────────────────────────────
# protokolliere_annaeherung
# ─────────────────────────────────────────────────────────────

def test_annaeherung_protokolliert_iteration(drive_base):
    p = vg.baue_define(profile="typisch")
    assert p.annaeherung_log == []

    helper.protokolliere_annaeherung(
        p,
        einstellung={"Abzugswinkel": 150, "Gummiband-Position": 13, "Becherposition": 15},
        wuerfe=[290, 295, 298],
    )
    assert len(p.annaeherung_log) == 1
    eintrag = p.annaeherung_log[0]
    assert eintrag["iteration"] == 1
    assert eintrag["wuerfe"] == [290.0, 295.0, 298.0]
    assert abs(eintrag["mean"] - 294.333) < 0.01
    assert abs(eintrag["abweichung_vom_ziel"] - (294.333 - p.zielweite)) < 0.01


def test_annaeherung_nur_ofat_ohne_warnung(drive_base, capsys):
    p = vg.baue_define(profile="typisch")
    helper.protokolliere_annaeherung(p, {"A": 10, "B": 5}, [290, 295, 300])
    capsys.readouterr()
    helper.protokolliere_annaeherung(p, {"A": 12, "B": 5}, [310, 315, 320])  # nur A geändert
    out = capsys.readouterr().out
    assert "OFAT-Hinweis" not in out


def test_annaeherung_mehrere_aenderungen_warnt(drive_base, capsys):
    p = vg.baue_define(profile="typisch")
    helper.protokolliere_annaeherung(p, {"A": 10, "B": 5, "C": 2}, [290, 295, 300])
    capsys.readouterr()
    helper.protokolliere_annaeherung(p, {"A": 12, "B": 7, "C": 2}, [310, 315, 320])  # A + B
    out = capsys.readouterr().out
    assert "OFAT-Hinweis" in out
    assert "A" in out and "B" in out


def test_annaeherung_log_roundtrip(drive_base):
    p = vg.baue_define(profile="typisch")
    helper.protokolliere_annaeherung(p, {"Abzugswinkel": 150}, [290, 295])
    helper.protokolliere_annaeherung(p, {"Abzugswinkel": 155}, [310, 315])

    q = helper.lade_fortschritt(p.gruppenname, p.gruppennummer)
    assert len(q.annaeherung_log) == 2
    assert q.annaeherung_log[0]["iteration"] == 1
    assert q.annaeherung_log[1]["iteration"] == 2
    assert q.annaeherung_log[1]["einstellung"] == {"Abzugswinkel": 155}


# ─────────────────────────────────────────────────────────────
# setze_initiale_einstellung
# ─────────────────────────────────────────────────────────────

def test_initiale_einstellung_aus_letzter_iteration(drive_base):
    p = vg.baue_define(profile="typisch")
    helper.protokolliere_annaeherung(p, {"A": 10}, [290])
    helper.protokolliere_annaeherung(p, {"A": 12}, [300])
    helper.setze_initiale_einstellung(p)
    assert p.initiale_einstellung == {"A": 12}


def test_initiale_einstellung_explizit_uebergeben(drive_base):
    p = vg.baue_define(profile="typisch")
    helper.setze_initiale_einstellung(p, {"Abzugswinkel": 155, "Becherposition": 16})
    assert p.initiale_einstellung == {"Abzugswinkel": 155, "Becherposition": 16}


def test_initiale_einstellung_ohne_log_warnt(drive_base, capsys):
    p = vg.baue_define(profile="typisch")
    helper.setze_initiale_einstellung(p)  # kein annaeherung_log, kein Argument
    out = capsys.readouterr().out
    assert p.initiale_einstellung == {}
    assert "Keine Annäherungs-Iteration" in out or "keine" in out.lower()


def test_initiale_einstellung_roundtrip(drive_base):
    p = vg.baue_define(profile="typisch")
    helper.setze_initiale_einstellung(p, {"Abzugswinkel": 160})
    q = helper.lade_fortschritt(p.gruppenname, p.gruppennummer)
    assert q.initiale_einstellung == {"Abzugswinkel": 160}


# ─────────────────────────────────────────────────────────────
# zeige_faktoren_legende
# ─────────────────────────────────────────────────────────────

def test_faktoren_legende_druckt_namen(capsys):
    p = helper.init_projekt("L", 1)
    p.faktoren = [
        {"name": "MeinFaktor", "einheit": "mm", "low": 1, "high": 9,
         "centerpoint_moeglich": True},
    ]
    helper.zeige_faktoren_legende(p)
    out = capsys.readouterr().out
    assert "MeinFaktor" in out
    assert "mm" in out


def test_faktoren_legende_warnt_ohne_faktoren(capsys):
    p = helper.init_projekt("Leer", 1)
    helper.zeige_faktoren_legende(p)
    out = capsys.readouterr().out
    assert "⚠️" in out or "Noch keine" in out
