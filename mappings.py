
COLUMN_MAPPINGS = {
    "truck": {
        "source_sheet": "truck",
        "destination_sheet": "Truck",
        "source_start_row": 7,
        "dest_start_row": 11,

        "header_mappings": {
            "date": "A7",
            "project_code": "B7",
            "data_collector": "C7",
            "number_of_equipment_types": "D7",
        },

        "column_mappings": {
            "E": "B",  # Equipment Tag
            "F": "C",  # Man power
            "G": "D",  # Truck Plate or Tag
            "H": "E",  # Task Type
            "I": "F",  # Description
            "J": "G",  # Soil Type
            "P": "L",  # Unit
            "O": "N",  # Total Cycle Time
            "Q": "M",  # Q Actual Bucket capacity
        }

    },

    "dozer": {
        "source_sheet": "dozer",
        "destination_sheet": "Dozer",

        "source_start_row": 7,
        "dest_start_row": 11,

        "header_mappings": {
            "date": "A7",
            "project_code": "B7",
            "data_collector": "C7",
            "number_of_equipment_types": "D7",
        },


        "column_mappings": {
            'E': 'B',  # Equipment Tag (Dozer Cyle)
            'F': 'C',  # Man power
            'G': 'D',  # Dozer Blade Type
            'H': 'E',  # Task Type
            'I': 'F',  # Description
            'J': 'G',  # Soil Type
            'K': 'H',  # Blade Height (m)
            'L': 'I',  # Blade Width (m)
            'M': 'J',  # Blade Length (m)
            'N': 'K',  # unit (m3, m, etc)
            #'O': 'L',  # Blade Load (m3, m, etc) - This will be calculated, so we can skip copying this column
            'P': 'M',  # Cycle Time (seconds)
            # 'Q': 'N',
        },
        "custom_fields": {
            "blade_load":"L",

            "unit": "K",
        }
    },

    "excavator": {
        "source_sheet": "excavator",
        "destination_sheet": "Excavator",
        "source_start_row": 7,
        "dest_start_row": 11,

        "header_mappings": {
            "date": "A7",
            "project_code": "B7",
            "data_collector": "C7",
            "number_of_equipment_types": "D7",
        },


        "column_mappings": {
            'E': 'B',  # Equipment Tag (Dozer Cyle)
            'F': 'C',  # Man power
            'G': 'D',  # Dozer Blade Type
            'H': 'E',  # Task Type
            'I': 'F',  # Description
            'J': 'G',  # Soil Type
            'K': 'H',  # Bucket Fill factor
            #'L': 'I',  # Angle of swing
            #'M': 'I',  # Depth of Cut
            'N': 'J',   # Volume Correction
            'O': 'K',  # Efficiency (60m/60m)
            'P': 'L',  # unit (m3, m, etc)
            'Q': 'M',  # Q Heaped Bucket capacity(m3, m, etc) - 
            'R': 'N',  # Cycle Time (seconds)
        },
        "custom_fields": {
            "swing_ratio": {
                "source_angle_col": "L",
                "source_depth_col": "M",
                "dest_col": "I"
            },

            "volume_correction": {
                "soil_type_col": "G",
                "dest_col": "J"
            },

            "efficiency": {
                "default": 60,
                "dest_col": "K"
            }
        }
    }
}
