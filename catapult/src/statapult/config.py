"""Konfigurationsmanagement fuer den Statapult-Simulator.

Laedt physikalische Konstanten und Rauschparameter aus YAML-Dateien
mit Fallback auf eingebaute Defaults.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Optional

import yaml

from .noise import NoiseModel
from .physics import CatapultPhysics

_DEFAULTS_PATH = Path(__file__).parent / "defaults.yaml"


@dataclass
class BallType:
    """Konfiguration fuer einen Balltyp."""

    name: str
    mass_g: float
    radius_cm: float
    drag_coefficient: float


@dataclass
class CatapultConfig:
    """Gesamtkonfiguration des Simulators."""

    physics: CatapultPhysics = field(default_factory=CatapultPhysics)
    noise: NoiseModel = field(default_factory=NoiseModel)
    enable_drag: bool = False
    ball_types: Dict[str, BallType] = field(default_factory=dict)

    @classmethod
    def default(cls) -> CatapultConfig:
        """Erzeugt die Standardkonfiguration aus defaults.yaml."""
        return cls.from_yaml(_DEFAULTS_PATH)

    @classmethod
    def from_yaml(cls, path: str | Path) -> CatapultConfig:
        """Laedt Konfiguration aus einer YAML-Datei."""
        path = Path(path)
        if not path.exists():
            raise FileNotFoundError(f"Konfigurationsdatei nicht gefunden: {path}")

        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)

        return cls._from_dict(data)

    @classmethod
    def _from_dict(cls, data: Dict[str, Any]) -> CatapultConfig:
        cat = data.get("catapult", {})
        phys_data = data.get("physics", {})
        noise_data = data.get("noise", {})
        ball_data = data.get("ball_types", {})

        # Kalibrierungskoeffizienten aus YAML laden (falls vorhanden)
        calib_data = data.get("calibration", {})
        factor_coefficients = None
        if calib_data.get("factor_coefficients"):
            factor_coefficients = {}
            for key, vals in calib_data["factor_coefficients"].items():
                factor_coefficients[key] = (vals["linear"], vals.get("quadratic", 0.0))

        physics = CatapultPhysics(
            arm_length_m=cat.get("arm_length_cm", 30.0) / 100.0,
            arm_mass_kg=cat.get("arm_mass_g", 50.0) / 1000.0,
            k_base=cat.get("k_base", 80.0),
            efficiency=cat.get("efficiency", 0.78),
            rest_angle_deg=cat.get("rest_angle_deg", 110.0),
            gravity=phys_data.get("gravity", 9.81),
            air_density=phys_data.get("air_density", 1.225),
            drag_coefficient=phys_data.get("drag_coefficient", 0.47),
            d_base=calib_data.get("d_base", 280.0),
            factor_coefficients=factor_coefficients,
        )

        noise = NoiseModel(
            sigma_measurement=noise_data.get("sigma_measurement", 1.5),
            sigma_setup=noise_data.get("sigma_setup", 2.0),
            sigma_rubber=noise_data.get("sigma_rubber", 0.5),
            sigma_release=noise_data.get("sigma_release", 1.0),
            sigma_wind_turbulence=noise_data.get("sigma_wind_turbulence", 0.3),
            operator_sigma=noise_data.get("operator_sigma", 1.0),
            drift_rate=noise_data.get("drift_rate", 0.0),
        )

        ball_types = {}
        for name, bt in ball_data.items():
            ball_types[name] = BallType(
                name=name,
                mass_g=bt["mass_g"],
                radius_cm=bt["radius_cm"],
                drag_coefficient=bt["drag_coefficient"],
            )

        return cls(
            physics=physics,
            noise=noise,
            enable_drag=phys_data.get("enable_drag", False),
            ball_types=ball_types,
        )
