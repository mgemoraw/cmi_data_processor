"""
Data Aggregating and Analysis Engine
"""
from openpyxl import Workbook, load_workbook
from pathlib import Path 
import os 
import win32com.client  as win32

from dataclasses import dataclass

@dataclass
class OutputData:
    day: str
    date: str

    operation_activity: str
    equipment_type: str
    description_work: str

    productivity: float
    ideal_productivity: float
    overall_method_productivity: float
    ideal_cycle_variability: float
    overall_cycle_variability: float

    HEADERS = [
        "Day",
        "Date",
        "Operation/Activity",
        "Equipment Type",
        "Description Work",
        "Productivity (Units/hr)",
        "Ideal Productivity (Unit/hr)",
        "Overall Method Productivity (Unit/hr)",
        "Ideal Cycle Variability (%)",
        "Overall Cycle Variability (%)",
    ]
    COLUMN_MAP = {
        "A": "day",
        "B": "date",
        "C": "operation_activity",
        "D": "equipment_type",
        "E": "description_work",
        "F": "productivity",
        "G": "ideal_productivity",
        "H": "overall_method_productivity",
        "I": "ideal_cycle_variability",
        "J": "overall_cycle_variability",
    }


    @classmethod
    def write_headers(cls, ws, row=1):
        for col, header in enumerate(cls.HEADERS, start=1):
            ws.cell(row=row, column=col, value=header)

    def write_row(self, ws, row):
        values = [
            self.day,
            self.date,
            self.operation_activity,
            self.equipment_type,
            self.description_work,
            self.productivity,
            self.ideal_productivity,
            self.overall_method_productivity,
            self.ideal_cycle_variability,
            self.overall_cycle_variability,
        ]

        for col, value in enumerate(values, start=1):
            ws.cell(row=row, column=col, value=value)
            

    def write_to_row(
        self,
        ws,
        row: int,
        activity: str,
        equipment: str,
        particular: str
    ):
        ws[f"C{row}"] = self.date
        ws[f"D{row}"] = activity
        ws[f"E{row}"] = equipment
        ws[f"F{row}"] = particular
        ws[f"G{row}"] = self.average_productivity
        ws[f"H{row}"] = self.ideal_productivity
        ws[f"I{row}"] = self.overall_method_productivity
        ws[f"J{row}"] = self.ideal_cycle_variability
        ws[f"K{row}"] = self.overall_cycle_variability


class AnalysisEngine:
    def __init__(self, data_folder, template_path=None, equipment=None, particular:str=None, activity:str=None):
        self.source_folder = Path(data_folder)
        self.output_folder = self.source_folder / "output"
        self.template_path = Path(template_path) if template_path else None
        self.equipment = equipment
        self.particular = particular
        self.activity = activity

        self.output_folder.mkdir(exist_ok=True)

    def run(self):
        return self._start_processing()


    def aggregate_data(self, data_folder, output_file):
        """
        This method aggregates data from multiple Excel files in the specified folder and saves the combined data into a new Excel file.
        """

    def refresh_excel_file(self, file_path):
        """
        This method refreshes the data connections in an Excel file using win32com.
        """
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(str(file_path))
        workbook.RefreshAll()
        workbook.Save()
        workbook.Close()
        excel.Quit()

    def _start_processing(self):
        """
        This method will process the excel files and extract required output data
        """
        new_wb = Workbook()
        new_ws = new_wb.create_sheet(title=self.activity)
        td_start_row = 15

        OutputData.write_headers(new_ws, row=14)
        current_row = 15

        for index, file in enumerate(os.listdir(self.source_folder)):
            if file.endswith(".xlsx"):

                # refresh the excel file to get the latest data from the data connections
                self.refresh_excel_file(os.path.join(self.source_folder, file))

                file_path = os.path.join(self.source_folder, file)
                wb = load_workbook(filename=file_path, data_only=True)
                sheet_name = f'{self.equipment.title()}'
                ws = wb[sheet_name]

                data = OutputData(
                    day=index+1,
                    date=ws["L6"].value,
                    operation_activity=self.activity,
                    equipment_type=self.equipment,
                    description_work=self.particular,
                    productivity=ws["N111"].value,
                    ideal_productivity=ws["O11"].value,
                    overall_method_productivity=ws["P11"].value,
                    ideal_cycle_variability=ws["Q11"].value,
                    overall_cycle_variability=ws["R11"].value,
                )
                # new_ws[f'C{td_start_row+index}'] = data.date
                # new_ws[f'D{td_start_row+index}'] = self.activity
                # new_ws[f'E{td_start_row+index}'] = self.equipment
                # new_ws[f'F{td_start_row+index}'] = self.particular
                # new_ws[f'G{td_start_row+index}'] = data.average_productivity
                # new_ws[f'H{td_start_row+index}'] = data.ideal_productivity
                # new_ws[f'I{td_start_row+index}'] = data.overall_method_productivity
                # new_ws[f'J{td_start_row+index}'] = data.ideal_cycle_variability
                # new_ws[f'K{td_start_row+index}'] = data.overall_cycle_variability

                data.write_row(new_ws, current_row)
                current_row += 1
                print(f"Processed file: {file}")
                print(f"Extracted data: {ws['N11'].value}")

        output_file = os.path.join(self.output_folder, f"{self.activity}.xlsx")
        new_wb.save(output_file)

        return True


    def read_excel_files(self, file_path):
        """
        This method reads an Excel file and returns its content as a list of dictionaries, where each dictionary represents a row of data.
        """
        



if __name__ == "__main__":
    source_path = Path("./input_folder/output").resolve()
    template_path = ""
    equipment = "Dozer"
    particular = "Site Clearing using Dozer"
    activity = "Site Clearing using Dozezers"

    engine = AnalysisEngine(
        data_folder=source_path,
        template_path=template_path,
        equipment=equipment,
        particular=particular,
        activity=activity
        )

    engine.run()
