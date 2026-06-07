"""Physical constants and default probe geometry values."""

import math

ELEMENTARY_CHARGE_C = 1.602e-19
VACUUM_PERMITTIVITY_F_PER_M = 8.8541878188e-12
BOLTZMANN_J_PER_K = 1.38e-23
ARGON_ION_MASS_KG = 6.62e-26
XENON_ION_MASS_KG = 2.1801714e-25
ELECTRON_MASS_KG = 9.11e-31
JOULE_TO_EV = 6.241509e18

DEFAULT_PROBE_LENGTH_M = 33.4e-3
DEFAULT_PROBE_DIAMETER_M = 1.27e-3

def cylindrical_probe_area(diameter_m: float = DEFAULT_PROBE_DIAMETER_M, length_m: float = DEFAULT_PROBE_LENGTH_M) -> float:
    """Return approximate exposed cylindrical Langmuir probe area in m^2."""
    return math.pi / 4 * diameter_m**2 + math.pi * diameter_m * length_m
