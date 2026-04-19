"""Statapult -- Virtueller Six Sigma Katapult-Simulator.

Beispiel::

    from statapult import Statapult

    katapult = Statapult(seed=42)
    result = katapult.shoot({
        "abzugswinkel": 160,
        "stoppwinkel": 90,
        "gummiband_position": 15,
        "becherposition": 18,
        "pin_hoehe": 12,
    })
    print(f"Wurfweite: {result.wurfweite_cm:.1f} cm")
"""

__version__ = "0.1.0"

from .simulator import Statapult, ShotResult
from .config import CatapultConfig
from .factors import ALL_FACTORS, STANDARD_FACTORS, ENVIRONMENTAL_FACTORS, Factor

__all__ = [
    "Statapult",
    "ShotResult",
    "CatapultConfig",
    "ALL_FACTORS",
    "STANDARD_FACTORS",
    "ENVIRONMENTAL_FACTORS",
    "Factor",
]
