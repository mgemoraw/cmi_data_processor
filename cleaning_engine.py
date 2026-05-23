

class DataCleaningEngine:
    def __init__(self, *args, **kwargs):
        self.input_folder = kwargs.get("input_folder", None)
        self.output_folder = kwargs.get("output_folder", None)
        self.logger = kwargs.get("logger", print)
        self.progress_callback = kwargs.get("progress_callback", None)
        self.equipment = kwargs.get("equipment", None)
        self.template_path = kwargs.get("template_path", None)

    def log(self, message):
        if self.logger:
            self.logger(message)

    def update_progress(self, value):
        if self.progress_callback:
            self.progress_callback(value)
