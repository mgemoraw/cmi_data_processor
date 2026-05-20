import shutil
from pathlib import Path
from openpyxl import load_workbook


class TemplateMappingEngine:    
    def __init__(self, template_dir=None, **kwargs):
        self.template_dir = Path(template_dir)
        self.output_folder = kwargs.get("output_folder", Path("output"))
        self.logger = kwargs.get("logger", print)
        self.progress_callback = kwargs.get("progress_callback", None)
        self.equipment = kwargs.get("equipment", None)


   
    def log(self, message):
        if self.logger:
            self.logger(message)

    def update_progress(self, value):
        if self.progress_callback:
            self.progress_callback(value)
            
    def copy_template(
        self,
        template_path,
        output_path
    ):

        template_path = Path(template_path)
        output_path = Path(output_path)

        # Ensure template exists
        if not template_path.exists():
            raise FileNotFoundError(
                f"Template not found: {template_path}"
            )

        # Create output folder
        output_path.parent.mkdir(
            parents=True,
            exist_ok=True
        )

        # Remove existing destination file
        if output_path.exists():
            output_path.unlink()

        # Copy template
        shutil.copy2(
            template_path,
            output_path
        )

        self.log(
            f"Copied template → {output_path.name}"
        )

    def read_template(self, template_path):
        try:
            workbook = load_workbook(template_path)
            self.logger(f"Successfully read template: {template_path}")
            return workbook
        except Exception as e:
            self.logger(f"Error reading template: {e}")
            return None