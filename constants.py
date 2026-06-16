from dataclasses import dataclass 



@dataclass
class FillFactor:

    min_percent: float
    max_percent: float
    average_factor: float 


SWING_DEPTH_RATIO_TABLE = {
    (0, 90): {
        (0, 2): 0.80,
        (2, 4): 0.75,
        (4, 6): 0.70,
    },

    (90, 180): {
        (0, 2): 0.65,
        (2, 4): 0.60,
        (4, 6): 0.55,
    }
}

SHOVEL_PRODUCTION_FACTORS = {
    "default": 1.0,
    40: {
        45: 0.93,
        60: 0.89,
        75: 0.85,
        90: 0.80,
        120: 0.72,
        150: 0.65,
        180: 0.59,
    },

    60: {
        45: 1.10,
        60: 1.03,
        75: 0.96,
        90: 0.91,
        120: 0.81,
        150: 0.73,
        180: 0.66,
    },

    80: {
        45: 1.22,
        60: 1.12,
        75: 1.04,
        90: 0.98,
        120: 0.86,
        150: 0.77,
        180: 0.69,
    },

    100: {
        45: 1.26,
        60: 1.16,
        75: 1.07,
        90: 1.00,
        120: 0.88,
        150: 0.79,
        180: 0.71,
    },

    120: {
        45: 1.20,
        60: 1.11,
        75: 1.03,
        90: 0.97,
        120: 0.86,
        150: 0.77,
        180: 0.70,
    },

    140: {
        45: 1.12,
        60: 1.04,
        75: 0.97,
        90: 0.91,
        120: 0.81,
        150: 0.73,
        180: 0.66,
    },

    160: {
        45: 1.03,
        60: 0.96,
        75: 0.90,
        90: 0.85,
        120: 0.75,
        150: 0.67,
        180: 0.62,
    }
}

# D - Fill factor 
FILL_FACTOR_TABLE = {
    "moist_loam_sandy_clay": {
        "range_percent": (100, 110),
        "fill_factor": 1.05
    },

    "sand_gravel": {
        "range_percent": (95, 100),
        "fill_factor": 0.98
    },

    "rock_poorly_blasted": {
        "range_percent": (40, 50),
        "fill_factor": 0.45
    },

    "rock_well_blasted": {
        "range_percent": (60, 75),
        "fill_factor": 0.68
    },

    "hard_tough_clay": {
        "range_percent": (80, 90),
        "fill_factor": 0.85
    }
}

SOIL_VOLUME_CORRECTION = {
    "common": 1.00,
    "rock": 0.85,
    "wet": 0.90,
    "clay": 0.95,
    "sand": 1.10
}