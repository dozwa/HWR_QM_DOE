"""Tests für das Fortschritt-Save/Load zwischen Phasen (fortschritt.json)."""
from __future__ import annotations

import numpy as np
import pandas as pd
import pytest

import helper
from tests import virtuelle_gruppe as vg


def test_roundtrip_define(drive_base):
    p = vg.baue_define(profile="typisch")
    helper.speichere_fortschritt(p)
    q = helper.lade_fortschritt(p.gruppenname, p.gruppennummer)
    assert q is not None
    assert q.zielweite == p.zielweite
    np.testing.assert_array_almost_equal(q.vermessung_min_wuerfe, p.vermessung_min_wuerfe)
    np.testing.assert_array_almost_equal(q.vermessung_max_wuerfe, p.vermessung_max_wuerfe)
    assert q.vermessung_beschreibung == p.vermessung_beschreibung
    assert [f["name"] for f in q.faktoren] == [f["name"] for f in p.faktoren]
    assert q.charter == p.charter


def test_roundtrip_analyze_regeneriert_modell(drive_base):
    """Nach Save/Load muss `_recompute_derived` das Modell neu bauen."""
    p = vg.baue_analyze(profile="typisch")
    r2_original = p.modell.rsquared
    helper.speichere_fortschritt(p)
    q = helper.lade_fortschritt(p.gruppenname, p.gruppennummer)
    assert q.modell is not None
    assert abs(q.modell.rsquared - r2_original) < 1e-6
    # faktoren_doe muss round-trippen
    assert [f["name"] for f in q.faktoren_doe] == [f["name"] for f in p.faktoren_doe]


def test_roundtrip_improve_restauriert_optimum(drive_base):
    p = vg.baue_improve(profile="typisch")
    helper.speichere_fortschritt(p)
    q = helper.lade_fortschritt(p.gruppenname, p.gruppennummer)
    # Konfirmationswürfe und optimale Einstellung müssen zurückkommen.
    np.testing.assert_array_almost_equal(q.konfirmation_wuerfe, p.konfirmation_wuerfe)
    assert q.optimale_einstellung is not None
    assert abs(q.optimale_einstellung["vorhersage"]
               - p.optimale_einstellung["vorhersage"]) < 1.0  # numerische Toleranz


def test_roundtrip_control_cpk_identisch(drive_base):
    p = vg.baue_control(profile="streu")
    cpk_before = p.cpk_ergebnis["cpk"]
    helper.speichere_fortschritt(p)
    q = helper.lade_fortschritt(p.gruppenname, p.gruppennummer)
    assert q.cpk_ergebnis is not None
    assert abs(q.cpk_ergebnis["cpk"] - cpk_before) < 1e-6
    assert q.imr_ergebnis is not None
    assert bool(q.imr_ergebnis["stabil"]) == bool(p.imr_ergebnis["stabil"])


def test_multiple_save_load_zyklen(drive_base):
    """Mehrmals speichern+laden darf die Daten nicht drift-behaftet machen."""
    p = vg.baue_analyze(profile="typisch")
    original_faktoren = [f["name"] for f in p.faktoren]
    original_doe_len = len(p.doe_ergebnisse)
    for _ in range(3):
        helper.speichere_fortschritt(p)
        p = helper.lade_fortschritt(p.gruppenname, p.gruppennummer)
    assert [f["name"] for f in p.faktoren] == original_faktoren
    assert len(p.doe_ergebnisse) == original_doe_len


def test_aktuelle_phase_wird_richtig_erkannt(drive_base):
    """_aktuelle_phase klassifiziert den Fortschritt anhand befüllter Felder.

    Hinweis: Sobald ``konfirmation_wuerfe`` gesetzt sind, zählt das als CONTROL —
    auch wenn die Konfirmation laut Notebook-Reihenfolge in IMPROVE eingegeben
    wird. Die Fixture ``baue_improve`` befüllt beide Felder, deshalb fällt sie
    bereits in CONTROL.
    """
    start = helper.init_projekt("PhaseTest", 1)
    assert helper._aktuelle_phase(start) == "START"

    define_p = vg.baue_define(profile="typisch")
    assert helper._aktuelle_phase(define_p) == "DEFINE"

    measure_p = vg.baue_measure(profile="typisch")
    assert helper._aktuelle_phase(measure_p) == "MEASURE"

    analyze_p = vg.baue_analyze(profile="typisch")
    assert helper._aktuelle_phase(analyze_p) == "ANALYZE"

    improve_p = vg.baue_improve(profile="typisch")
    # Fixture füllt bereits konfirmation_wuerfe → CONTROL laut helper._aktuelle_phase
    assert helper._aktuelle_phase(improve_p) == "CONTROL"

    control_p = vg.baue_control(profile="typisch")
    assert helper._aktuelle_phase(control_p) == "CONTROL"

    # Reine IMPROVE-Phase: optimale Einstellung, aber noch keine Konfirmation.
    improve_only = vg.baue_analyze(profile="typisch")
    improve_only.optimale_einstellung = {"vorhersage": 300.0, "pi_low": 290.0,
                                          "pi_high": 310.0, "einstellungen": {}}
    assert helper._aktuelle_phase(improve_only) == "IMPROVE"


def test_finde_speicherstaende_entdeckt_gruppen(drive_base):
    """Nach zwei speichere_fortschritt-Aufrufen müssen beide Gruppen gefunden werden."""
    p1 = helper.init_projekt("GruppeEins", 1)
    helper.speichere_fortschritt(p1)
    p2 = helper.init_projekt("GruppeZwei", 2)
    helper.speichere_fortschritt(p2)

    staende = helper.finde_speicherstaende()
    namen = {(s["gruppenname"], s["gruppennummer"]) for s in staende}
    assert ("GruppeEins", 1) in namen
    assert ("GruppeZwei", 2) in namen
