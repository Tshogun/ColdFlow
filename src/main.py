import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), './excel')))
from write_excel import ColdFlow  # type: ignore 

if __name__ == "__main__":
    excel_file_path = os.path.join(os.path.dirname(os.path.dirname(__file__)),  'data', 'cold_email_data.xlsx')
    coldflow = ColdFlow(excel_file_path)
    coldflow.process_excel_sheet()
    coldflow.save_workbook()