import logging
import os
from datetime import datetime

class Logger:
    def __init__(self, log_directory='logs'):
        self.log_directory = log_directory
        os.makedirs(self.log_directory, exist_ok=True)
        self.log_file_1 = os.path.join(self.log_directory, 'log_file_1.log')
        self.log_file_2 = os.path.join(self.log_directory, 'log_file_2.log')
        self.log_file_3 = os.path.join(self.log_directory, 'log_file_3.log')

        self.logger = logging.getLogger('coldflow_logger')
        self.logger.setLevel(logging.DEBUG)  # Log everything

        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

        # Create and add handlers for each log file
        handler1 = logging.FileHandler(self.log_file_1)
        handler1.setFormatter(formatter)
        self.logger.addHandler(handler1)

        handler2 = logging.FileHandler(self.log_file_2)
        handler2.setFormatter(formatter)
        self.logger.addHandler(handler2)

        handler3 = logging.FileHandler(self.log_file_3)
        handler3.setFormatter(formatter)
        self.logger.addHandler(handler3)

    def info(self, message):
        """Logs an info level message to all three files."""
        self.logger.info(message)

    def error(self, message):
        """Logs an error level message to all three files."""
        self.logger.error(message)
        print(message)

    def debug(self, message):
        """Logs a debug level message to all three files."""
        self.logger.debug(message)
