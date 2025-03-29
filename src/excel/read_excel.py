import openpyxl
from datetime import datetime
import sys
import os

# Add logger directory to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../logger')))
from logger import Logger
log = Logger()

# Add llm directory to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../llm')))
from generate_content import CoverLetterGenerator # type: ignore
generator = CoverLetterGenerator()

class ExcelProcessor:
    def __init__(self):
        pass

    def process_excel_sheet(self, file_path):
        """
        Opens the Excel file, iterates through each row of the first sheet
        (skipping the header row), and captures the data (excluding the first
        column) as lists of cell values.

        Args:
            file_path (str): The path to the Excel file.

        Returns:
            list: A list of lists, where each inner list represents a row
                  (excluding the first column and header) and contains the
                  values of the remaining cells in that row.
                  Returns None if the file is not found or an error occurs.
        """
        try:
            log.info(f"Opening Excel file: {file_path}")
            workbook = openpyxl.load_workbook(file_path)

            # Get the first sheet
            sheet = workbook.active
            log.info(f"Processing sheet: {sheet.title} (dropping first column and header)")

            all_rows_data = []
            row_count = 0
            first_row_skipped = False
            for row in sheet.iter_rows():
                if not first_row_skipped:
                    log.debug("Skipping header row.")
                    first_row_skipped = True
                    continue

                row_data = []
                for cell_index, cell in enumerate(row):
                    if cell_index > 0:  # Skip the first column (index 0)
                        row_data.append(cell.value)
                if row_data:  # Only append if the row had more than just the first column
                    all_rows_data.append(row_data)
                    row_count += 1
                    log.debug(f"Captured data from row {row_count} (excluding first column): {row_data}")

            log.info(f"Successfully processed {row_count} rows (excluding header and first column) from sheet: {sheet.title}")
            return all_rows_data

        except FileNotFoundError:
            log.error(f"Error: Excel file not found at path: {file_path}")
            return None
        except Exception as e:
            log.error(f"An error occurred while processing the Excel file: {e}", exc_info=True)
            return None

if __name__ == "__main__":
    excel_file_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '..', 'data', 'cold_email_data.xlsx')
    processor = ExcelProcessor()
    rows = processor.process_excel_sheet(excel_file_path)

    if rows:
        print("Data from the Excel sheet (without first column and header):")
        for row in rows:
            print(row)
    else:
        print("Failed to process the Excel sheet.")