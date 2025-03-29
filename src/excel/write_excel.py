from read_excel import ExcelProcessor
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

#Add email directory to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../email')))
from send_email import EmailSender, getenv # type: ignore
sender_email, password = getenv()
es = EmailSender(sender_email, password)

excel_file_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '..', 'data', 'cold_email_data.xlsx')
processor = ExcelProcessor()
rows = processor.process_excel_sheet(excel_file_path)