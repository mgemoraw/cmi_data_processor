from template_mapping_engine import TemplateMappingEngine
from processing_engine import DataProcessingEngine

if __name__ == "__main__":
    # Initialize the engine
    engine = DataProcessingEngine(
        folder_path="templates",
        output_folder="output",
        logger=print,
        equipment="truck"
    )

    # Example usage
    template_path = "templates/1.8. Truck _ Template_2026_05_01_BiT.xlsx"
    output_path = "output/truck_report.xlsx"

    engine.copy_template(template_path, output_path)