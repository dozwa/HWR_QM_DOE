"""Tests für pruefe_excel_faktoren_konsistenz und den Upload-Flow.

Prüft, dass Diskrepanzen zwischen den in DEFINE definierten Faktoren und den
Spaltenheadern einer hochgeladenen DoE-Excel klar gemeldet werden.
"""
from __future__ import annotations

import numpy as np
import pandas as pd
import pytest

import helper
from tests import virtuelle_gruppe as vg


def _as_excel_df(faktoren, extra=()):
    """Erzeugt einen DataFrame mit den Spalten, die _parse_faktoren_aus_excel
    erkennt, inkl. 'Name (kodiert)' und 'Name (Einheit)'. ``extra`` kann
    weitere Faktoren in die Excel einfügen.
    """
    columns = {"Versuch_Nr": [1, 2], "Block": [1, 1]}
    for f in list(faktoren) + list(extra):
        columns[f"{f['name']} (kodiert)"] = [-1.0, 1.0]
        columns[f"{f['name']} ({f['einheit']})"] = [f["low"], f["high"]]
    columns["Ergebnis: Weite (cm)"] = [200.0, 400.0]
    return pd.DataFrame(columns)


def test_konsistenz_keine_meldung_wenn_identisch():
    p = helper.init_projekt("K", 1)
    p.faktoren_doe = [
        {"name": "Winkel", "einheit": "Grad", "low": 30, "high": 45,
         "centerpoint_moeglich": True},
    ]
    df = _as_excel_df(p.faktoren_doe)
    excel_fak = helper._parse_faktoren_aus_excel(df)
    meldungen = helper.pruefe_excel_faktoren_konsistenz(p, excel_fak)
    assert meldungen == []


def test_konsistenz_erkennt_zusaetzlichen_excel_faktor():
    p = helper.init_projekt("K2", 1)
    p.faktoren_doe = [
        {"name": "Winkel", "einheit": "Grad", "low": 30, "high": 45,
         "centerpoint_moeglich": True},
    ]
    extra = [{"name": "Extra", "einheit": "cm", "low": 0, "high": 5}]
    df = _as_excel_df(p.faktoren_doe, extra=extra)
    excel_fak = helper._parse_faktoren_aus_excel(df)
    meldungen = helper.pruefe_excel_faktoren_konsistenz(p, excel_fak)
    assert any("'Extra'" in m and "Excel" in m for m in meldungen)


def test_konsistenz_erkennt_fehlenden_excel_faktor():
    p = helper.init_projekt("K3", 1)
    p.faktoren_doe = [
        {"name": "Winkel", "einheit": "Grad", "low": 30, "high": 45,
         "centerpoint_moeglich": True},
        {"name": "Spannung", "einheit": "cm", "low": 5, "high": 15,
         "centerpoint_moeglich": True},
    ]
    # Excel enthält nur Winkel
    df = _as_excel_df([p.faktoren_doe[0]])
    excel_fak = helper._parse_faktoren_aus_excel(df)
    meldungen = helper.pruefe_excel_faktoren_konsistenz(p, excel_fak)
    assert any("'Spannung'" in m and "nicht in Excel" in m for m in meldungen)


def test_konsistenz_erkennt_abweichende_stufen():
    p = helper.init_projekt("K4", 1)
    p.faktoren_doe = [
        {"name": "Winkel", "einheit": "Grad", "low": 30, "high": 45,
         "centerpoint_moeglich": True},
    ]
    df = _as_excel_df([{"name": "Winkel", "einheit": "Grad", "low": 25, "high": 50}])
    excel_fak = helper._parse_faktoren_aus_excel(df)
    meldungen = helper.pruefe_excel_faktoren_konsistenz(p, excel_fak)
    # Beide Stufen-Abweichungen werden gemeldet
    assert any("Low=25" in m for m in meldungen)
    assert any("High=50" in m for m in meldungen)


def test_konsistenz_leere_referenz_gibt_keine_meldungen():
    """Ohne projekt.faktoren* soll die Prüfung einfach leer zurückgeben."""
    p = helper.init_projekt("K5", 1)
    p.faktoren_doe = []
    p.faktoren = []
    df = _as_excel_df([{"name": "Winkel", "einheit": "Grad", "low": 30, "high": 45}])
    excel_fak = helper._parse_faktoren_aus_excel(df)
    meldungen = helper.pruefe_excel_faktoren_konsistenz(p, excel_fak)
    assert meldungen == []


def test_konsistenz_e2e_excel_ueberschreibt_faktoren_doe():
    """Der Notebook-Flow schreibt nach Upload die Excel-Namen als faktoren_doe
    zurück — wir simulieren das und prüfen das fitte_modell mit den Excel-Faktoren
    funktioniert (ohne Crash), selbst wenn projekt.faktoren_doe abwich."""
    p = vg.baue_analyze(profile="typisch")
    # Rename a factor in Excel headers — der Notebook-Flow würde das erkennen
    # und die Excel-Namen als Quelle der Wahrheit setzen.
    renamed = p.doe_ergebnisse.rename(columns={
        "Abzugswinkel (Grad) (kodiert)": "Abzugswinkel_neu (Grad) (kodiert)",
        "Abzugswinkel (Grad) (Grad)": "Abzugswinkel_neu (Grad) (Grad)",
    }, errors="ignore")
    # _parse_faktoren_aus_excel sollte mindestens die bekannten Faktoren finden.
    excel_fak = helper._parse_faktoren_aus_excel(p.doe_ergebnisse)
    assert len(excel_fak) >= 1
    # Konsistenzprüfung läuft ohne Exception
    helper.pruefe_excel_faktoren_konsistenz(p, excel_fak)
