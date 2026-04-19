"""Faktor-Definitionen fuer den Statapult-Simulator.

Jeder Faktor hat Name, Einheit, Low/High-Level und einen Default-Wert.
Das Format ist kompatibel mit helper.py ``generiere_versuchsplan()``.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict


@dataclass(frozen=True)
class Factor:
    """Ein einzelner experimenteller Faktor."""

    name: str
    key: str
    einheit: str
    low: float
    high: float
    default: float
    description: str = ""

    # ------------------------------------------------------------------
    def to_helper_format(self) -> Dict:
        """Erzeugt das Dict-Format fuer ``helper.generiere_versuchsplan()``."""
        return {
            "name": f"{self.name} ({self.einheit})",
            "einheit": self.einheit,
            "low": self.low,
            "high": self.high,
        }

    def coded(self, natural_value: float) -> float:
        """Natuerlichen Wert in kodierten Wert (-1 / +1) umrechnen."""
        center = (self.high + self.low) / 2
        half_range = (self.high - self.low) / 2
        if half_range == 0:
            return 0.0
        return (natural_value - center) / half_range

    def natural(self, coded_value: float) -> float:
        """Kodierten Wert (-1 / +1) in natuerlichen Wert umrechnen."""
        center = (self.high + self.low) / 2
        half_range = (self.high - self.low) / 2
        return center + coded_value * half_range

    def clamp(self, value: float) -> float:
        """Wert auf den gueltigen Bereich begrenzen."""
        return max(self.low, min(self.high, value))


# ======================================================================
# Standard-Faktoren (entsprechen dem realen Statapult)
# ======================================================================

STANDARD_FACTORS: Dict[str, Factor] = {
    "abzugswinkel": Factor(
        name="Abzugswinkel",
        key="abzugswinkel",
        einheit="Grad",
        low=130,
        high=170,
        default=150,
        description="Rueckzugswinkel des Arms (wie weit zurueckgezogen)",
    ),
    "stoppwinkel": Factor(
        name="Stoppwinkel",
        key="stoppwinkel",
        einheit="Grad",
        low=70,
        high=110,
        default=90,
        description="Winkel bei dem der Arm gestoppt wird (Release-Punkt)",
    ),
    "gummiband_position": Factor(
        name="Gummiband-Position",
        key="gummiband_position",
        einheit="cm",
        low=8,
        high=18,
        default=13,
        description="Position des Gummibands auf dem Arm (Abstand vom Pivot)",
    ),
    "becherposition": Factor(
        name="Becherposition",
        key="becherposition",
        einheit="cm",
        low=8,
        high=22,
        default=15,
        description="Position des Bechers auf dem Arm (Abstand vom Pivot)",
    ),
    "pin_hoehe": Factor(
        name="Pin-Hoehe",
        key="pin_hoehe",
        einheit="cm",
        low=8,
        high=18,
        default=13,
        description="Hoehe des Drehpunkts (Pivot)",
    ),
}

ENVIRONMENTAL_FACTORS: Dict[str, Factor] = {
    "ballgewicht": Factor(
        name="Ballgewicht",
        key="ballgewicht",
        einheit="g",
        low=5,
        high=30,
        default=10,
        description="Masse des Balls",
    ),
    "wind": Factor(
        name="Wind",
        key="wind",
        einheit="m/s",
        low=-2,
        high=2,
        default=0,
        description="Windgeschwindigkeit (positiv = Rueckenwind)",
    ),
}

ALL_FACTORS: Dict[str, Factor] = {**STANDARD_FACTORS, **ENVIRONMENTAL_FACTORS}


def get_defaults() -> Dict[str, float]:
    """Gibt die Default-Werte aller Faktoren zurueck."""
    return {key: f.default for key, f in ALL_FACTORS.items()}


def validate_settings(settings: Dict[str, float]) -> Dict[str, float]:
    """Prueft und begrenzt Settings auf gueltige Bereiche.

    Fehlende Faktoren werden mit Defaults aufgefuellt.
    """
    result = get_defaults()
    for key, value in settings.items():
        if key in ALL_FACTORS:
            result[key] = ALL_FACTORS[key].clamp(value)
    return result


def factors_info() -> str:
    """Gibt eine formatierte Uebersicht aller Faktoren zurueck."""
    lines = ["=== Statapult Faktoren ===", ""]
    lines.append("Standard-Faktoren:")
    for f in STANDARD_FACTORS.values():
        lines.append(
            f"  {f.name:<22s} {f.low:>6.0f} - {f.high:<6.0f} {f.einheit:<6s}"
            f"  (Default: {f.default})"
        )
    lines.append("")
    lines.append("Umwelt-Faktoren:")
    for f in ENVIRONMENTAL_FACTORS.values():
        lines.append(
            f"  {f.name:<22s} {f.low:>6.1f} - {f.high:<6.1f} {f.einheit:<6s}"
            f"  (Default: {f.default})"
        )
    return "\n".join(lines)
