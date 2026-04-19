"""Mehrstufiges Rauschmodell fuer realistische statistische Eigenschaften.

Die einzelnen Rauschquellen sind auf die DMAIC-Uebungen abgestimmt:
- sigma_measurement: Messungenauigkeit (fuer MSA Type-1 und Gage R&R)
- sigma_setup: Einstellungsreproduzierbarkeit (fuer DOE-Replikation)
- sigma_rubber: Gummiband-Variabilitaet (Materialschwankung)
- sigma_release: Release-Mechanismus-Impraezision
- sigma_wind_turbulence: Umwelt-Turbulenz
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Dict, Optional

import numpy as np


@dataclass
class NoiseModel:
    """Konfigurierbare Rauschquellen."""

    sigma_measurement: float = 1.5
    sigma_setup: float = 2.0
    sigma_rubber: float = 0.5
    sigma_release: float = 1.0
    sigma_wind_turbulence: float = 0.3
    operator_sigma: float = 1.0
    drift_rate: float = 0.0   # cm pro Schuss

    # Interne Speicherung der Operator-Biases
    _operator_biases: Dict[str, float] = field(default_factory=dict, repr=False)

    @property
    def sigma_total(self) -> float:
        """Gesamte Standardabweichung (Quadratwurzel der Varianz-Summe)."""
        return math.sqrt(
            self.sigma_measurement ** 2
            + self.sigma_setup ** 2
            + self.sigma_rubber ** 2
            + self.sigma_release ** 2
            + self.sigma_wind_turbulence ** 2
        )

    def total_noise(self, rng: np.random.Generator) -> float:
        """Erzeugt Gesamtrauschen aus allen unabhaengigen Quellen."""
        return (
            rng.normal(0, self.sigma_measurement)
            + rng.normal(0, self.sigma_setup)
            + rng.normal(0, self.sigma_rubber)
            + rng.normal(0, self.sigma_release)
            + rng.normal(0, self.sigma_wind_turbulence)
        )

    def get_operator_bias(self, operator_id: str, rng: np.random.Generator) -> float:
        """Gibt den systematischen Bias eines Operators zurueck.

        Der Bias wird beim ersten Zugriff erzeugt und dann beibehalten,
        damit derselbe Operator konsistente Messfehler hat.
        """
        if operator_id not in self._operator_biases:
            self._operator_biases[operator_id] = rng.normal(0, self.operator_sigma)
        return self._operator_biases[operator_id]

    def measurement_noise(self, rng: np.random.Generator) -> float:
        """Nur Mess-Rauschen (fuer MSA-Simulation)."""
        return rng.normal(0, self.sigma_measurement)

    def apply_drift(self, shot_number: int) -> float:
        """Berechnet den Drift-Beitrag basierend auf der Schussnummer."""
        return self.drift_rate * shot_number

    def reset_operators(self) -> None:
        """Setzt die Operator-Biases zurueck."""
        self._operator_biases.clear()
