

import os
from pathlib import Path
import re
import shutil
from copy import copy 
from openpyxl import Workbook, load_workbook
from datetime import datetime 


class DataProcessingEngine:
    def __init__(self, input_folder=None, template_path=None, logger=print):
        self.source_path = Path(input_folder) if input_folder else Path.cwd()
        self.template_path = Path(template_path) if template_path else Path.cwd()
        self.template_mapping = {}

        self.logger = logger

        self.logger = logger
        self.progress_callback = None
        
        self.output_folder = os.path.join(self.source_path, "output")

        # Pattern for DD-MM-YYYY
        self.date_pattern = r"(\d{2}-\d{2}-\d{4})"

    def log(self, message):
        self.logger(message)


    def update_progress(self, value):
        if self.progress_callback:
            self.progress_callback(value)


    def sort_files_by_date(self, files):
        def extract_date_object(file_name):
            match = re.search(self.date_pattern, file_name)
            if match:
                date_str = match.group(1)  # Extracts the date substring (e.g., "02-03-2014")
                try:
                    # Convert the string to a real datetime object for accurate chronological sorting
                    # Adjust "%d-%m-%Y" if your pattern captures dates differently (e.g., "%Y-%m-%d")
                    return datetime.strptime(date_str, "%d-%m-%Y")
                except ValueError:
                    # Handle cases where the regex matches a string that isn't a valid calendar date
                    return datetime.max
            
            # If no date pattern is found, return datetime.max to safely push the file to the end of the list
            return datetime.max

        # Sort using the actual datetime object instead of a string
        return sorted(files, key=extract_date_object)
    
    
    def read_excel_contents(self):
        self.logger("🚀 Starting data processing...")
        # Implement the main processing logic here
        self.logger(f"📂 Source folder: {self.source_path}")
        # template = self.read_template_file(self.template_path)
        files = self.sort_files_by_date(os.listdir(self.source_path))
        for index, file in enumerate(files):
            if file.endswith('.xlsx') and file != self.template_path.name:
                self.logger(f"📄 Processing file: {file}")
                # Implement your data processing logic here
                # For example, you can read the file, extract data, etc.
                # You can also update progress using self.update_progress(value)

                date_str = self.get_date_str(os.path.join(self.source_path, file))
                data_count = self.format_data_count(data_count=index+1)
                
                # copy daily variables template
                self.copy_daily_variables(os.path.join(self.source_path, file), date_str, data_count)

                # copy main template
                instance_template = self.copy_template(self.template_path, date_str, data_count)

                # populate productivity and MPDM sheets in the copied template
                # source_file_path = os.path.join(self.source_path, file)
                self.populate_productivity(os.path.join(self.source_path, file), instance_template)
                # self.populate_mpdm(source_file_path, instance_template)
                

        # For example, you can call other methods to read templates, process data, etc.
        self.logger("✅ Data processing completed!")

    def format_data_count(self, data_count=1):
        try:
            return f"{data_count:02d}"
        except Exception as e:
            self.logger(f"❌ Error formatting data count: {e}")
            return "00"

    def populate_productivity(self, source_file, template_path):
        try:
            source_wb = load_workbook(filename=source_file, read_only=False)
            # source_ws = source_wb[self.equipment_sheet_name]
            source_ws = source_wb['dozer']

            new_wb = load_workbook(filename=template_path, read_only=False)
            # new_ws = new_wb[self.equipment_sheet_name.title()]
            new_ws = new_wb['Dozer']
          
            # Copy values and styles from source to new sheet
            project_code = source_ws['B7'].value
            number_of_equipment_types = source_ws['D7'].value
            date = source_ws['A7'].value
            data_collector = source_ws['C7'].value

            DOZER_COLUMN_MAPPING = {
                'B': 'E',  # Equipment Tag (Dozer Cyle)
                'C': 'F',  # Man power
                'D': 'G',  # Dozer Blade Type
                'E': 'H',  # Task Type
                'F': 'I',  # Description
                'G': 'J',  # Soil Type
                'H': 'K',  # Blade Height (m)
                'I': 'L',  # Blade Width (m)
                'J': 'M',  # Blade Length (m)
                'K': 'N',  # unit (m3, m, etc)
                'L': 'O',  # Blade Load (m3, m, etc) - This will be calculated, so we can skip copying this column
                'M': 'P',  # Cycle Time (seconds)
                'N': 'Q',  # Productivity (m3/hr, m/hr, etc) - This will be calculated, so we can skip copying this column
            }
            source_start_row = 7  # Assuming the first row is headers
            dest_start_row = 11  # Assuming the first row is headers
            

            # Copy project code, date, data collector, and number of equipment types to the new sheet
            new_ws['D6'] = project_code
            new_ws['I6'] = number_of_equipment_types
            new_ws['L6'] = date.strftime("%d-%m-%Y") if hasattr(date, "strftime") else date
            new_ws['O6'] = data_collector

            
            # calculate how many rows down the data offset needs to shift
            row_offset = dest_start_row - source_start_row

            for row in source_ws.iter_rows(min_row=source_start_row, values_only=False):
                self.logger(f"Processing row {row[0].row} with values: {[cell.value for cell in row]}")
                for cell in row:
                    
                    if cell.coordinate in DOZER_COLUMN_MAPPING.keys():
                        self.logger(f"Processing cell {cell.coordinate} with value: {cell.value}")
                        # dest_cell = DOZER_COLUMN_MAPPING[cell.coordinate[0]] + str(cell.row[0] + row_offset)
                        dest_cell= DOZER_COLUMN_MAPPING[cell.column_letter] 
                        self.logger(f"Mapping source cell {cell.coordinate} to destination cell {dest_cell}{cell.row}+ {row_offset}")
                        new_cell = new_ws[dest_cell + str(cell.row + row_offset)]
                        new_ws[f"E{row[0].row + row_offset}"] = "Zoned Rock Fill dam shell at the flnaks"  # Copy Equipment Tag to column E in the new sheet
                        new_cell.value = cell.value 
           
            dest_path = os.path.join(self.output_folder, os.path.basename(template_path))
            new_wb.save(template_path)
            self.logger(f"✅ Updated productivity sheet in: {template_path}")
            source_wb.close()
        except Exception as e:
            self.logger(f"❌ Error populating productivity sheet: {e}")

    def populate_mpdm(self, file_path, template_path):
        try:
            wb = load_workbook(filename=file_path, read_only=False)
            ws = wb['mpdm']
            ws['B2'] = "Sample MPDM Data"
            dest_path = os.path.join(self.output_folder, os.path.basename(template_path))
            wb.save(dest_path)
            self.logger(f"✅ Updated MPDM sheet in: {dest_path}")
            wb.close()
        except Exception as e:
            self.logger(f"❌ Error populating MPDM sheet: {e}")

    def copy_daily_variables(self, file_path, date_str, data_count):
        # daily_variables_template = os.path.join(self.template_path.parent, "daily_variables_BiT.xlsx")
        wb = load_workbook(filename=file_path, read_only=False)
        ws = wb['daily_variables']
        new_wb = Workbook()
        new_ws = new_wb.active
        new_ws.title = "daily_variables"
        for row in ws.iter_rows(values_only=False):
            # Grab the matching target cell in the new sheet
            for cell in row:
                new_cell = new_ws.cell(row=cell.row, column=cell.column, value=cell.value)
               
               # deep copy structural attributes
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = cell.number_format
                    new_cell.protection = copy(cell.protection)
                    new_cell.alignment = copy(cell.alignment)

        # 2. Copy Column Dimensions (Widths)
        for col_letter, col_dim in ws.column_dimensions.items():
            new_ws.column_dimensions[col_letter].width = col_dim.width
            new_ws.column_dimensions[col_letter].hidden = col_dim.hidden

        # 3. Copy Row Dimensions (Heights)
        for row_idx, row_dim in ws.row_dimensions.items():
            new_ws.row_dimensions[row_idx].height = row_dim.height
            new_ws.row_dimensions[row_idx].hidden = row_dim.hidden

        # 4. Copy Merged Cell Ranges (if any exist)
        for merged_range in list(ws.merged_cells.ranges):
            new_ws.merge_cells(str(merged_range))    
            # new_ws.append([cell.value for cell in row])

        # Finally update data count and date in the new sheet
        new_ws['E5'] = data_count

        
        # Save logic pipeline
        dest_folder = os.path.join(self.source_path, "daily_variables")
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)
            
        dest_path = os.path.join(dest_folder, f"{data_count}_{date_str}_daily_variables.xlsx")
        new_wb.save(dest_path)
        wb.close()  # Clean up file stream locks
        
        self.logger(f"✅ Copied daily variables with complete styles to: {dest_path}")

    def copy_template(self, template_path, date_str=None, data_count="00"):  

        dest_folder = os.path.join(self.source_path, "output")
        try:
            if not os.path.exists(dest_folder):
                os.makedirs(dest_folder)
                self.logger(f"📁 Created folder: {dest_folder}")
            
            # dest_path = os.path.join(dest_folder, os.path.basename(self.source_path))
            # dest_path = os.path.join(dest_folder, f"{data_count}_{date_str}_{os.path.basename(template_path)}")
            dest_path = os.path.join(dest_folder, f"{data_count}_{date_str}_template.xlsx")

            shutil.copy(template_path, dest_path)
            self.logger(f"✅ Copied template to: {dest_path}")
            return dest_path
          
        except Exception as e:
            self.logger(f"❌ Error copying template: {e}")
            return None

    def get_date_str(self, file_path):
        self.logger(f"reading date: from {file_path}")
        try:
            
            file_name = os.path.basename(file_path)
            mach = re.search(self.date_pattern, file_name)
            if mach:
                return mach.group(1)


            wb = load_workbook(file_path, data_only=True)
            ws = wb['daily_variables']
            date_value = ws['H5'].value
            if hasattr(date_value, "strftime"):
                return date_value.strftime("%d-%m-%Y")
            
            else:
                if "/" in str(date_value):
                    return date_value.replace("/", "-")
                return date_value

        except Exception as e:
            self.logger(f"❌ Error reading date from {file_path}: {e}")
            return None
        
    def read_template_file(self, template_path):
        try:
            wb = load_workbook(filename=template_path, read_only=True)
            self.logger(f"✅ Successfully loaded template: {template_path}")
            return wb
        except Exception as e:
            self.logger(f"❌ Error loading template: {e}")
            return None

    def rename_template_file(self, template_path, new_name):
        try:
            dest_path = os.path.join(self.output_folder, new_name)
            shutil.copy(template_path, dest_path)
            self.logger(f"✅ Renamed template to: {dest_path}")
            return dest_path
        except Exception as e:
            self.logger(f"❌ Error renaming template: {e}")
            return None
        
    
        
    def load_source_data(self, file_path):
        wb = load_workbook(filename=file_path, read_only=True)