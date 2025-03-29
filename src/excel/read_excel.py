import pandas as pd
from logger.logger import Logger

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
