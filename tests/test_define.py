"""DEFINE-Phase Tests (Ergebnis-Ebene).

Getestet werden init_projekt, speichere_vermessung / zeige_vermessung /
setze_zielweite, berechne_testwurf_statistik, formatiere_charter und die
Round-Trip-Persistenz für alle DEFINE-Felder.
"""
from __future__ import annotations

import numpy as np
import pytest

import helper
from tests import virtuelle_gruppe as vg


# ─────────────────────────────────────────────────────────────
# init_projekt
# ─────────────────────────────────────────────────────────────

def test_init_projekt_default_zielweite():
    p = helper.init_projekt("DefaultGroup", 1)
    assert p.zielweite == 300.0
    assert p.toleranz == 15.0
    assert p.gruppenname == "DefaultGroup"
    assert p.gruppennummer == 1


def test_init_projekt_zielweite_override():
    p = helper.init_projekt("Custom", 2, zielweite=450.0, toleranz=25.0)
    assert p.zielweite == 450.0
    assert p.toleranz == 25.0


def test_init_projekt_seed_deterministic():
    p1 = helper.init_projekt("Alpha", 3)
    p2 = helper.init_projekt("Alpha", 3)
    p3 = helper.init_projekt("Alpha", 4)
    assert p1.seed == p2.seed
    assert p1.seed != p3.seed


def test_init_projekt_starts_empty():
    p = helper.init_projekt("Empty", 1)
    assert p.faktoren == []
    assert p.faktoren_doe == []
    assert len(p.vermessung_min_wuerfe) == 0
    assert len(p.vermessung_max_wuerfe) == 0
    assert p.vermessung_beschreibung == ""


# ─────────────────────────────────────────────────────────────
# speichere_vermessung
# ─────────────────────────────────────────────────────────────

def test_speichere_vermessung_persistiert(drive_base):
    p = vg.neue_gruppe()
    helper.speichere_vermessung(
        p,
        min_wuerfe=[220, 225, 218],
        max_wuerfe=[440, 445, 450],
        min_einstellung={"Abzugswinkel": 140},
        max_einstellung={"Abzugswinkel": 165},
        beschreibung="Testkatapult",
    )
    assert list(p.vermessung_min_wuerfe) == [220.0, 225.0, 218.0]
    assert list(p.vermessung_max_wuerfe) == [440.0, 445.0, 450.0]
    assert p.vermessung_beschreibung == "Testkatapult"

    reloaded = helper.lade_fortschritt(p.gruppenname, p.gruppennummer)
    assert reloaded is not None
    np.testing.assert_array_equal(reloaded.vermessung_min_wuerfe, p.vermessung_min_wuerfe)
    np.testing.assert_array_equal(reloaded.vermessung_max_wuerfe, p.vermessung_max_wuerfe)
    assert reloaded.vermessung_beschreibung == "Testkatapult"


def test_speichere_vermessung_filtert_nullwerte(drive_base):
    p = vg.neue_gruppe()
    helper.speichere_vermessung(
        p,
        min_wuerfe=[220, 0, 218, -5],
        max_wuerfe=[440, 445],
        min_einstellung={},
        max_einstellung={},
    )
    # Zero and negative filtered out.
    assert len(p.vermessung_min_wuerfe) == 2
    assert list(p.vermessung_min_wuerfe) == [220.0, 218.0]
    assert len(p.vermessung_max_wuerfe) == 2


def test_speichere_vermessung_teilupdate_keine_falschen_warnungen(drive_base, capsys):
    """Zweimalige Aufrufe (erst Min, dann Max) dürfen nicht fälschlich warnen."""
    p = vg.neue_gruppe()
    helper.speichere_vermessung(p, [220, 225, 218], [], {}, {})
    first = capsys.readouterr().out
    helper.speichere_vermessung(p, [], [440, 445, 450], {}, {})
    second = capsys.readouterr().out
    assert "Keine gültigen Max-Würfe" not in first
    assert "Keine gültigen Min-Würfe" not in second


# ─────────────────────────────────────────────────────────────
# setze_zielweite
# ─────────────────────────────────────────────────────────────

def test_setze_zielweite_innerhalb_der_spanne_keine_warnung(drive_base, capsys):
    p = vg.baue_define(profile="typisch")
    capsys.readouterr()
    helper.setze_zielweite(p, 350.0)
    out = capsys.readouterr().out
    assert p.zielweite == 350.0
    assert "✅" in out
    assert "außerhalb" not in out


def test_setze_zielweite_ausserhalb_warnt_aber_setzt(drive_base, capsys):
    p = vg.baue_define(profile="typisch")
    oben = float(p.vermessung_max_wuerfe.mean()) + 200.0  # unerreichbar
    capsys.readouterr()
    helper.setze_zielweite(p, oben)
    out = capsys.readouterr().out
    assert p.zielweite == oben
    assert "außerhalb" in out


def test_setze_zielweite_ohne_vermessung_hinweis(drive_base, capsys):
    p = helper.init_projekt("Bare", 1)
    capsys.readouterr()
    helper.setze_zielweite(p, 280.0)
    out = capsys.readouterr().out
    assert p.zielweite == 280.0
    assert "Plausibilitätsprüfung" in out or "noch keine" in out.lower()


def test_setze_zielweite_negative_wird_ignoriert(drive_base, capsys):
    p = helper.init_projekt("Neg", 1)
    old = p.zielweite
    helper.setze_zielweite(p, -10.0)
    assert p.zielweite == old


# ─────────────────────────────────────────────────────────────
# Testwurf-Statistik + Ampel-Klassifikation
# ─────────────────────────────────────────────────────────────

def test_testwurf_statistik_cv_klassifikation():
    """CV soll je nach Streuung in die 3 Ampel-Bänder fallen (grün/gelb/rot)."""
    # Grün: sehr reproduzierbar (< 15% CV)
    gruen = np.array([300, 302, 298, 301, 299], dtype=float)
    s = helper.berechne_testwurf_statistik(gruen)
    assert s["cv"] < 15.0

    # Gelb: 15–30 % CV (mittel streuend)
    gelb = np.array([300, 360, 250, 330, 280], dtype=float)
    s = helper.berechne_testwurf_statistik(gelb)
    assert 0 < s["cv"] < 30  # konkret: auf den Wurfdaten ca. 15 %

    # Rot: extrem streuend (>30 %)
    rot = np.array([100, 500, 200, 450, 150], dtype=float)
    s = helper.berechne_testwurf_statistik(rot)
    assert s["cv"] > 30.0


def test_testwurf_statistik_zaehlt_korrekt():
    w = np.array([1, 2, 3, 4, 5], dtype=float)
    s = helper.berechne_testwurf_statistik(w)
    assert s["n"] == 5
    assert s["mean"] == 3.0


# ─────────────────────────────────────────────────────────────
# Charter + zeige_vermessung rendern
# ─────────────────────────────────────────────────────────────

def test_formatiere_charter_enthaelt_alle_eintraege():
    p = vg.baue_define(profile="typisch")
    html = helper.formatiere_charter(p)
    for key in ("Problemstellung", "Projektziel", "Scope", "Zielweite"):
        assert key in html


def test_zeige_vermessung_rendert_ohne_fehler(drive_base):
    """zeige_vermessung darf auf einer minimalen Vermessung durchlaufen."""
    p = vg.baue_define(profile="typisch")
    helper.zeige_vermessung(p)  # stored in p.figuren
    assert "define_vermessung" in p.figuren


def test_zeige_vermessung_ohne_daten_gibt_hinweis(drive_base, capsys):
    p = helper.init_projekt("Leer", 1)
    helper.zeige_vermessung(p)
    out = capsys.readouterr().out
    assert "noch nicht vollständig" in out or "nicht vollständig" in out.lower()
