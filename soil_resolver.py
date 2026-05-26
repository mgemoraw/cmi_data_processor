from constants import FILL_FACTOR_TABLE

SOIL_TYPE_ALIASES = {

    "clay": [
        "clay",
        "hard clay",
        "tough clay",
        "soft clay",
        "clay soil",
    ],

    "sand_gravel": [
        "sand",
        "gravel",
        "sand gravel",
        "sand and gravel",
        "gravelly sand",
    ],

    "rock_poorly_blasted": [
        "poorly blasted rock",
        "soft rock",
        "broken rock",
    ],

    "rock_well_blasted": [
        "well blasted rock",
        "hard rock",
        "blasted rock",
    ],

    "moist_loam_sandy_clay": [
        "loam",
        "sandy clay",
        "moist loam",
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

    for canonical_name, aliases in SOIL_TYPE_ALIASES.items():

        for alias in aliases:

            if alias in normalized:
                return canonical_name

    return "unknown"

def get_fill_factor(raw_soil):

    soil_key = resolve_soil_type(raw_soil)

    return FILL_FACTOR_TABLE.get(
        soil_key,
        {"fill_factor": 1}
    )


if __name__ == "__main__":
    soil = "clay"
    resolved = resolve_soil_type(soil)
    factor = get_fill_factor(resolved)
    print(f"{resolved}: {factor}")
    # factor = get_fill_factor(
    #     source_ws[f"G{row}"].value
    # )