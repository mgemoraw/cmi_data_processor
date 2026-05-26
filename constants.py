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