"""Physik-Engine: Kalibriertes Energie-Modell fuer den Statapult.

Das Modell verbindet physikalische Mechanismen mit empirischer Kalibrierung,
um realistische Wurfweiten fuer DOE-Uebungen zu erzeugen.

Physikalische Grundlage:
- Jeder Faktor beeinflusst die Wurfweite ueber einen spezifischen Mechanismus
- Abzugswinkel: gespeicherte elastische Energie (mehr Rueckzug = mehr Energie)
- Stoppwinkel: Abwurfwinkel (bestimmt die Flugbahn-Geometrie)
- Gummiband-Position: Federkraft (hoehere Position = staerkere Kraft)
- Becherposition: Hebelarm vs. Traegheitsmoment (Optimum in der Mitte)
- Pin-Hoehe: Abwurfhoehe (hoeher = laengere Flugzeit)
- Ballgewicht: Traegheit (schwerer = langsamer)
- Wind: Ruecken-/Gegenwind (veraendert effektive Reichweite)

Die Wechselwirkungen entstehen natuerlich aus dem multiplikativen Modell:
die Gesamtwirkung ist das Produkt der Einzelwirkungen, nicht ihre Summe.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Dict, Optional

from .factors import ALL_FACTORS


# ======================================================================
# Kalibrierungskoeffizienten
# ======================================================================

# Basis-Wurfweite bei Mittelstellung aller Faktoren (cm)
# Kalibriert fuer ein realistisches Tisch-Statapult (60-300 cm Bereich)
D_BASE: float = 150.0

# Faktor-Koeffizienten: (linearer_Anteil c, quadratischer_Anteil q)
# d = D_BASE * prod(1 + c_i * x_i + q_i * x_i^2)
# x_i: kodierter Faktorwert (-1 bis +1)
# c_i > 0: hoehere Einstellung -> laengere Distanz
# q_i < 0: Optimum liegt in der Mitte (Kruemmung)
FACTOR_COEFFICIENTS: Dict[str, tuple[float, float]] = {
    "abzugswinkel":       (0.18, 0.00),    # Energie ~ Rueckzugswinkel
    "stoppwinkel":        (0.13, -0.04),   # Abwurfwinkel, leichte Kruemmung
    "gummiband_position": (0.16, 0.00),    # Federkraft ~ Position
    "becherposition":     (0.10, -0.05),   # Hebelarm-Optimum (Kruemmung)
    "pin_hoehe":          (0.07, 0.00),    # Abwurfhoehe
    "ballgewicht":        (-0.13, 0.00),   # Schwerer Ball = kuerzer
    "wind":               (0.05, 0.00),    # Rueckenwind hilft
}


@dataclass
class CatapultPhysics:
    """Physikalische Konstanten des Katapults.

    Diese werden fuer die Berechnung der Zwischenwerte (Energie,
    Geschwindigkeit, Abwurfwinkel) im Verbose-Modus verwendet.
    """

    arm_length_m: float = 0.30
    arm_mass_kg: float = 0.050
    k_base: float = 80.0
    efficiency: float = 0.78
    rest_angle_deg: float = 110.0
    gravity: float = 9.81
    air_density: float = 1.225
    drag_coefficient: float = 0.47
    d_base: float = D_BASE
    factor_coefficients: Dict[str, tuple[float, float]] = None

    def __post_init__(self):
        if self.factor_coefficients is None:
            self.factor_coefficients = dict(FACTOR_COEFFICIENTS)


@dataclass
class LaunchResult:
    """Ergebnis der Abwurfberechnung (fuer Anzeige/Verbose-Modus)."""

    spring_energy_j: float
    angular_velocity_rad_s: float
    ball_speed_m_s: float
    release_angle_deg: float
    release_height_m: float
    v_x: float
    v_y: float


@dataclass
class TrajectoryResult:
    """Ergebnis der Flugbahnberechnung."""

    distance_cm: float
    max_height_cm: float
    flight_time_s: float


def _code_factor(key: str, value: float) -> float:
    """Kodiert einen Faktorwert in [-1, +1]."""
    f = ALL_FACTORS[key]
    center = (f.high + f.low) / 2.0
    half_range = (f.high - f.low) / 2.0
    if half_range == 0:
        return 0.0
    return (value - center) / half_range


def compute_distance(
    settings: Dict[str, float],
    phys: Optional[CatapultPhysics] = None,
) -> float:
    """Berechnet die deterministische Wurfweite (ohne Rauschen).

    Kalibriertes multiplikatives Modell:
    d = D_BASE * prod(1 + c_i * x_i + q_i * x_i^2)

    Jeder Faktor wirkt als unabhaengiger Multiplikator.
    Wechselwirkungen entstehen natuerlich aus der Multiplikation.
    """
    if phys is None:
        phys = CatapultPhysics()

    d = phys.d_base
    coeffs = phys.factor_coefficients

    for key, (c, q) in coeffs.items():
        if key not in ALL_FACTORS:
            continue
        val = settings.get(key, ALL_FACTORS[key].default)
        x = _code_factor(key, val)
        d *= (1.0 + c * x + q * x ** 2)

    return max(0.0, d)


def compute_launch_info(
    settings: Dict[str, float],
    phys: Optional[CatapultPhysics] = None,
) -> LaunchResult:
    """Berechnet physikalische Zwischenwerte fuer den Verbose-Modus.

    Verwendet das vereinfachte Energie-Modell um anschauliche Werte
    wie Abwurfgeschwindigkeit und -winkel zu liefern.
    """
    if phys is None:
        phys = CatapultPhysics()

    abzug = settings.get("abzugswinkel", 150.0)
    stopp = settings.get("stoppwinkel", 90.0)
    gummi_cm = settings.get("gummiband_position", 13.0)
    becher_cm = settings.get("becherposition", 15.0)
    pin_cm = settings.get("pin_hoehe", 13.0)
    ball_g = settings.get("ballgewicht", 10.0)

    gummi_m = gummi_cm / 100.0
    becher_m = becher_cm / 100.0
    pin_m = pin_cm / 100.0
    m_ball = ball_g / 1000.0

    # Elastische Energie (illustrativ)
    k_eff = phys.k_base * (gummi_m / phys.arm_length_m) ** 2
    delta_x = 2.0 * gummi_m * math.sin(
        math.radians((abzug - phys.rest_angle_deg) / 2.0)
    )
    delta_x = max(0.0, delta_x)
    U_spring = 0.5 * k_eff * delta_x ** 2

    # Traegheitsmoment und Geschwindigkeit (illustrativ)
    I_arm = (1.0 / 3.0) * phys.arm_mass_kg * phys.arm_length_m ** 2
    I_ball = m_ball * becher_m ** 2
    I_total = I_arm + I_ball
    E_avail = phys.efficiency * U_spring

    if I_total > 0 and E_avail > 0:
        omega = math.sqrt(2.0 * E_avail / I_total)
        v_ball = omega * becher_m
    else:
        omega = 0.0
        v_ball = 0.0

    # Abwurfgeometrie (illustrativ)
    launch_angle_deg = (stopp - 70.0) * 0.75 + 20.0
    launch_angle_rad = math.radians(launch_angle_deg)
    h_release = pin_m + becher_m * math.sin(launch_angle_rad)

    # Aus kalibrierter Distanz die effektive Geschwindigkeit rueckrechnen
    distance_cm = compute_distance(settings, phys)
    distance_m = distance_cm / 100.0

    g = phys.gravity
    if launch_angle_deg > 0 and launch_angle_deg < 90:
        # Effektive Geschwindigkeit aus Wurfweite
        sin2a = math.sin(2.0 * launch_angle_rad)
        if sin2a > 0:
            v_eff = math.sqrt(
                distance_m * g / (sin2a + 0.01)
            )
        else:
            v_eff = v_ball
    else:
        v_eff = v_ball

    vx = v_eff * math.cos(launch_angle_rad)
    vy = v_eff * math.sin(launch_angle_rad)

    return LaunchResult(
        spring_energy_j=U_spring,
        angular_velocity_rad_s=omega,
        ball_speed_m_s=v_eff,
        release_angle_deg=launch_angle_deg,
        release_height_m=max(0.0, h_release),
        v_x=vx,
        v_y=vy,
    )


def compute_trajectory_info(
    launch: LaunchResult,
    distance_cm: float,
    gravity: float = 9.81,
) -> TrajectoryResult:
    """Berechnet Flugbahn-Kennwerte aus der kalibrierten Distanz."""
    vx = launch.v_x
    vy = launch.v_y
    h = launch.release_height_m
    g = gravity

    if vx > 0 and vy >= 0:
        t_apex = vy / g
        max_h = h + vy * t_apex - 0.5 * g * t_apex ** 2
        t_land = distance_cm / 100.0 / vx if vx > 0 else 0
    else:
        max_h = h
        t_land = 0.0

    return TrajectoryResult(
        distance_cm=distance_cm,
        max_height_cm=max(h, max_h) * 100.0,
        flight_time_s=max(0.0, t_land),
    )


def simulate_shot(
    abzugswinkel: float,
    stoppwinkel: float,
    gummiband_position: float,
    becherposition: float,
    pin_hoehe: float,
    ballgewicht: float = 10.0,
    wind: float = 0.0,
    enable_drag: bool = False,
    ball_radius_cm: float = 2.0,
    phys: Optional[CatapultPhysics] = None,
) -> tuple[float, LaunchResult, TrajectoryResult]:
    """Fuehrt einen kompletten Schuss durch.

    Returns
    -------
    (distance_cm, launch_result, trajectory_result)
    """
    if phys is None:
        phys = CatapultPhysics()

    settings = {
        "abzugswinkel": abzugswinkel,
        "stoppwinkel": stoppwinkel,
        "gummiband_position": gummiband_position,
        "becherposition": becherposition,
        "pin_hoehe": pin_hoehe,
        "ballgewicht": ballgewicht,
        "wind": wind,
    }

    distance_cm = compute_distance(settings, phys)
    launch = compute_launch_info(settings, phys)
    traj = compute_trajectory_info(launch, distance_cm, phys.gravity)

    return distance_cm, launch, traj
