from constants import FILL_FACTOR_TABLE, SOIL_VOLUME_CORRECTION

SOIL_TYPE_ALIASES = {

    "hard_tough_clay": [
        "clay",
        "hard clay",
        "tough clay",
        "soft clay",
        "clay soil",
        "hard clay soil",
        "tough clay soil",
        "soft clay soil",
    ],

    "sand_gravel": [
        "sand",
        "gravel",
        "sand gravel",
        "clay gravel",
        "sand and gravel",
        "clay and gravel",
        "gravelly sand",
        "gravel and sand",
        "gravel and clay",
    ],

    "rock_poorly_blasted": [
        "poorly blasted rock",
        "hard rock",
    ],

    "rock_well_blasted": [
        "well blasted rock",
        "blasted rock",
        "rock",
        "soft rock",
    ],

    "moist_loam_sandy_clay": [
        "loam",
        "sandy clay",
        "moist loam",
        "clay and sand",
        "sand and clay",
        "earth",
    ]
}


def normalize_text(value):

    if not value:
        return ""

    return (
        str(value)
        .lower()
        .strip()
        .replace("-", " ")
        .replace("_", " ")
    )

def resolve_soil_type(raw_soil_value):

    normalized = normalize_text(raw_soil_value)
    # print("normalized: ", normalized)

    for canonical_name, aliases in SOIL_TYPE_ALIASES.items():
       
        # print("canonical name: ", canonical_name)
        # print("aliases: ", aliases)
    
        # for alias in aliases:
        #     if alias in normalized:
        if normalized in aliases:
            return canonical_name

    return "unknown"

def get_fill_factor(raw_soil):

    soil_key = resolve_soil_type(raw_soil)
    # print("soil key:", soil_key)

    return FILL_FACTOR_TABLE.get(
        soil_key,
        {"fill_factor": 0.98}
    )


def get_volume_correction(raw_soil):
    for key in SOIL_VOLUME_CORRECTION.keys():
        if key in raw_soil.lower():
            return SOIL_VOLUME_CORRECTION[key]
        
    soil_key = resolve_soil_type(raw_soil)
    return SOIL_VOLUME_CORRECTION.get(
        soil_key,
        0.8
    )

if __name__ == "__main__":
    soil = "soft rock  "
    resolved = resolve_soil_type(soil)
    factor = get_fill_factor(soil)
    print(f"{resolved}: {factor}")
    # factor = get_fill_factor(
    #     source_ws[f"G{row}"].value
    # )