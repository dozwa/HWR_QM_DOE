"""Ein-/Ausgabe-Utilities fuer CSV und JSON."""

from __future__ import annotations

import csv
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional


def read_csv(path: str | Path) -> List[Dict[str, float]]:
    """Liest einen Versuchsplan aus einer CSV-Datei.

    Jede Zeile wird als Dict {spaltenname: wert} zurueckgegeben.
    """
    path = Path(path)
    rows = []
    with open(path, "r", encoding="utf-8-sig") as f:
        # Auto-detect delimiter
        sample = f.read(4096)
        f.seek(0)
        sniffer = csv.Sniffer()
        try:
            dialect = sniffer.sniff(sample, delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        reader = csv.DictReader(f, dialect=dialect)
        for row in reader:
            parsed = {}
            for key, value in row.items():
                if key is None:
                    continue
                try:
                    parsed[key.strip()] = float(value.replace(",", "."))
                except (ValueError, AttributeError):
                    parsed[key.strip()] = value
            rows.append(parsed)
    return rows


def write_csv(
    rows: List[Dict[str, Any]],
    path: Optional[str | Path] = None,
    delimiter: str = ",",
) -> Optional[str]:
    """Schreibt Ergebnisse als CSV.

    Wenn path None ist, wird nach stdout geschrieben und der String zurueckgegeben.
    """
    if not rows:
        return ""

    fieldnames = list(rows[0].keys())
    if path is not None:
        with open(path, "w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, delimiter=delimiter)
            writer.writeheader()
            writer.writerows(rows)
        return None

    import io
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=fieldnames, delimiter=delimiter)
    writer.writeheader()
    writer.writerows(rows)
    return buf.getvalue()


def format_result_text(result, verbose: bool = False) -> str:
    """Formatiert ein ShotResult als lesbaren Text."""
    lines = []
    if verbose:
        lines.append("=== Statapult Simulator ===")
        lines.append("")
        lines.append("Einstellungen:")
        for key, value in result.settings.items():
            from .factors import ALL_FACTORS
            f = ALL_FACTORS.get(key)
            name = f.name if f else key
            einheit = f.einheit if f else ""
            lines.append(f"  {name:<22s} {value:>8.1f} {einheit}")

        lines.append("")
        lines.append("Physik:")
        lines.append(f"  Federenergie:        {result.launch.spring_energy_j:>8.4f} J")
        lines.append(f"  Abwurfgeschw.:       {result.launch.ball_speed_m_s:>8.2f} m/s")
        lines.append(f"  Abwurfwinkel:        {result.launch.release_angle_deg:>8.1f} Grad")
        lines.append(f"  Abwurfhoehe:         {result.launch.release_height_m * 100:>8.1f} cm")
        lines.append("")
        lines.append("Ergebnis:")

    lines.append(f"  Wurfweite:           {result.wurfweite_cm:>8.1f} cm")
    if verbose:
        lines.append(f"  (davon Rauschen:     {result.noise_cm:>8.1f} cm)")
        lines.append(f"  Flugzeit:            {result.trajectory.flight_time_s:>8.3f} s")
        lines.append(f"  Max. Hoehe:          {result.trajectory.max_height_cm:>8.1f} cm")

    return "\n".join(lines)


def format_result_json(result) -> str:
    """Formatiert ein ShotResult als JSON."""
    return json.dumps(result.to_dict(), indent=2, ensure_ascii=False)


def format_multiple_results_text(results, show_stats: bool = True) -> str:
    """Formatiert mehrere ShotResults als Tabelle."""
    lines = []
    for i, r in enumerate(results, 1):
        lines.append(f"  Wurf {i:>3d}: {r.wurfweite_cm:>8.1f} cm")

    if show_stats and len(results) > 1:
        distances = [r.wurfweite_cm for r in results]
        import numpy as np
        mean = np.mean(distances)
        std = np.std(distances, ddof=1)
        lines.append(f"\n  Mittelwert: {mean:.1f} cm | Std.Abw.: {std:.1f} cm")

    return "\n".join(lines)
