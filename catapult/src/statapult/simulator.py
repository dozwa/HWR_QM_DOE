"""Statapult-Simulator -- Orchestrator-Klasse.

Verbindet Physik-Engine, Rauschmodell und Faktoren zu einer
einheitlichen Schnittstelle.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import numpy as np

from .config import CatapultConfig
from .factors import ALL_FACTORS, validate_settings
from .noise import NoiseModel
from .physics import CatapultPhysics, LaunchResult, TrajectoryResult, simulate_shot


@dataclass
class ShotResult:
    """Ergebnis eines Katapult-Schusses."""

    wurfweite_cm: float
    true_distance_cm: float
    noise_cm: float
    drift_cm: float
    operator_bias_cm: float
    settings: Dict[str, float]
    launch: LaunchResult
    trajectory: TrajectoryResult

    def to_dict(self) -> Dict[str, Any]:
        """Konvertiert das Ergebnis in ein Dictionary."""
        return {
            "wurfweite_cm": round(self.wurfweite_cm, 1),
            "true_distance_cm": round(self.true_distance_cm, 1),
            "noise_cm": round(self.noise_cm, 1),
            "drift_cm": round(self.drift_cm, 1),
            "settings": self.settings,
            "physics": {
                "spring_energy_j": round(self.launch.spring_energy_j, 4),
                "ball_speed_m_s": round(self.launch.ball_speed_m_s, 2),
                "release_angle_deg": round(self.launch.release_angle_deg, 1),
                "release_height_m": round(self.launch.release_height_m, 3),
                "max_height_cm": round(self.trajectory.max_height_cm, 1),
                "flight_time_s": round(self.trajectory.flight_time_s, 3),
            },
        }


class Statapult:
    """Virtueller Statapult-Simulator.

    Beispiel::

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

    def __init__(
        self,
        config: Optional[CatapultConfig] = None,
        seed: Optional[int] = None,
    ):
        self.config = config or CatapultConfig.default()
        self.rng = np.random.default_rng(seed)
        self._shot_count = 0

    def shoot(
        self,
        settings: Dict[str, float],
        noise_level: float = 1.0,
        operator_id: Optional[str] = None,
    ) -> ShotResult:
        """Fuehrt einen Schuss mit den gegebenen Einstellungen durch.

        Parameters
        ----------
        settings : dict
            Faktor-Einstellungen (z.B. {"abzugswinkel": 160, ...}).
            Fehlende Faktoren werden mit Defaults aufgefuellt.
        noise_level : float
            Multiplikator fuer das Rauschen (0 = deterministisch).
        operator_id : str, optional
            Operator-ID fuer MSA-Simulation (erzeugt systematischen Bias).

        Returns
        -------
        ShotResult
        """
        validated = validate_settings(settings)

        # Deterministische Physik
        distance_cm, launch, trajectory = simulate_shot(
            abzugswinkel=validated["abzugswinkel"],
            stoppwinkel=validated["stoppwinkel"],
            gummiband_position=validated["gummiband_position"],
            becherposition=validated["becherposition"],
            pin_hoehe=validated["pin_hoehe"],
            ballgewicht=validated["ballgewicht"],
            wind=validated["wind"],
            phys=self.config.physics,
        )

        # Rauschen
        noise = self.config.noise.total_noise(self.rng) * noise_level

        # Drift
        drift = self.config.noise.apply_drift(self._shot_count)

        # Operator-Bias (fuer MSA)
        operator_bias = 0.0
        if operator_id is not None:
            operator_bias = self.config.noise.get_operator_bias(
                operator_id, self.rng
            )

        final_distance = max(0.0, distance_cm + noise + drift + operator_bias)
        self._shot_count += 1

        return ShotResult(
            wurfweite_cm=final_distance,
            true_distance_cm=distance_cm,
            noise_cm=noise,
            drift_cm=drift,
            operator_bias_cm=operator_bias,
            settings=validated,
            launch=launch,
            trajectory=trajectory,
        )

    def shoot_multiple(
        self,
        settings: Dict[str, float],
        n: int = 1,
        noise_level: float = 1.0,
        operator_id: Optional[str] = None,
    ) -> List[ShotResult]:
        """Fuehrt mehrere Schuesse mit gleichen Einstellungen durch."""
        return [
            self.shoot(settings, noise_level=noise_level, operator_id=operator_id)
            for _ in range(n)
        ]

    def batch(
        self,
        plan: Any,  # pd.DataFrame -- lazy import
        noise_level: float = 1.0,
        result_column: str = "Ergebnis",
    ) -> Any:
        """Fuehrt eine Reihe von Schuessen aus einem Versuchsplan (DataFrame) durch.

        Parameters
        ----------
        plan : pd.DataFrame
            Versuchsplan mit Spalten die den Faktor-Keys entsprechen.
        noise_level : float
            Rausch-Multiplikator.
        result_column : str
            Name der Ergebnis-Spalte.

        Returns
        -------
        pd.DataFrame
            Kopie des Plans mit einer zusaetzlichen Ergebnis-Spalte.
        """
        import pandas as pd

        # Spalten-Mapping: Unterstuetzt sowohl Keys als auch formatierte Namen
        col_map = _build_column_map(plan.columns)

        results = []
        for _, row in plan.iterrows():
            settings = {}
            for col in plan.columns:
                mapped_key = col_map.get(col)
                if mapped_key:
                    settings[mapped_key] = float(row[col])
            result = self.shoot(settings, noise_level=noise_level)
            results.append(result.wurfweite_cm)

        out = plan.copy()
        out[result_column] = results
        return out

    def reset(self, seed: Optional[int] = None) -> None:
        """Setzt den Simulator zurueck (Schusszaehler, Operator-Biases)."""
        if seed is not None:
            self.rng = np.random.default_rng(seed)
        self._shot_count = 0
        self.config.noise.reset_operators()


def _build_column_map(columns) -> Dict[str, str]:
    """Baut ein Mapping von DataFrame-Spaltennamen auf Faktor-Keys."""
    col_map = {}
    for col in columns:
        col_lower = str(col).lower().strip()
        # Direkter Match
        if col_lower in ALL_FACTORS:
            col_map[col] = col_lower
            continue
        # Normalisierter Match (Bindestriche, Leerzeichen -> Unterstriche)
        normalized = col_lower.replace("-", "_").replace(" ", "_")
        if normalized in ALL_FACTORS:
            col_map[col] = normalized
            continue
        # Match ueber Faktor-Name (z.B. "Abzugswinkel (Grad)")
        for key, factor in ALL_FACTORS.items():
            if factor.name.lower() in col_lower or key in col_lower:
                col_map[col] = key
                break
    return col_map
