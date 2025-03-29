import pandas as pd
import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../logger')))
from logger import Logger

log = Logger()

def read_data(file_path):
    """Reads data from the specified Excel file into a pandas DataFrame."""
    try:
        df = pd.read_excel(file_path)
        log.info(f"Successfully read data from: {file_path}")
        return df
    except FileNotFoundError:
        log.error(f"Error: Excel file not found at path: {file_path}")
        return None
    except Exception as e:
        log.error(f"An error occurred while reading the Excel file: {e}")
        return None

if __name__ == "__main__":
    # Example usage if you run this file directly
    excel_file = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'cold_email_data.xlsx')
    df = read_data(excel_file)
    if df is not None:
        log.debug(f"First 5 rows from read_excel:\n{df.head()}")
        log.info(f"Columns from read_excel: {df.columns.tolist()}")