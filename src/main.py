import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
from cryptography.fernet import Fernet
import configparser
import logging

# Adjust path to be relative to the script's location
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(SCRIPT_DIR, './excel')))
from write_excel import ColdFlow  # type: ignore

CREDENTIALS_DIR = os.path.join(SCRIPT_DIR, 'credentials')
if not os.path.exists(CREDENTIALS_DIR):
    os.makedirs(CREDENTIALS_DIR)

CONFIG_FILE = os.path.join(CREDENTIALS_DIR, 'coldflow_config.ini')
KEY_FILE = os.path.join(CREDENTIALS_DIR, 'coldflow_key.key')

# Setup basic logging to capture output
LOG_FILE = os.path.join(SCRIPT_DIR, 'coldflow.log')
logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def generate_key():
    key = Fernet.generate_key()
    with open(KEY_FILE, 'wb') as key_file:
        key_file.write(key)

def load_key():
    if not os.path.exists(KEY_FILE):
        generate_key()
    with open(KEY_FILE, 'rb') as key_file:
        return key_file.read()

def encrypt_data(data, key):
    f = Fernet(key)
    return f.encrypt(data.encode()).decode()

def decrypt_data(encrypted_data, key):
    f = Fernet(key)
    return f.decrypt(encrypted_data.encode()).decode()

def save_credentials(email, password, groq_key):
    key = load_key()
    config = configparser.ConfigParser()
    config['credentials'] = {
        'email': encrypt_data(email, key),
        'password': encrypt_data(password, key),
        'groq_key': encrypt_data(groq_key, key)
    }
    with open(CONFIG_FILE, 'w') as configfile:
        config.write(configfile)

def load_credentials():
    if os.path.exists(CONFIG_FILE) and os.path.exists(KEY_FILE):
        key = load_key()
        config = configparser.ConfigParser()
        config.read(CONFIG_FILE)
        try:
            email = decrypt_data(config['credentials']['email'], key)
            password = decrypt_data(config['credentials']['password'], key)
            groq_key = decrypt_data(config['credentials']['groq_key'], key)
            return email, password, groq_key
        except Exception as e:
            logging.error(f"Error loading/decrypting credentials: {e}")
            messagebox.showerror("Error", "Error loading saved credentials.")
            return None, None, None
    return None, None, None

def set_environment_variables(email, password, groq_key):
    """Sets the environment variables for the current session."""
    os.environ["COLDFLOW_EMAIL_ADDRESS"] = email
    os.environ["COLDFLOW_EMAIL_PASSWORD"] = password
    os.environ["GROQ_API_KEY"] = groq_key

def run_coldflow_app(log_text_widget, status_label):
    """Runs the ColdFlow application."""
    email, password, groq_key = load_credentials()
    if not all([email, password, groq_key]):
        status_label.config(text="Credentials not set. Go to Settings.", fg="red")
        messagebox.showerror("ColdFlow", "Email and Groq API credentials are not set. Please go to Settings.")
        return

    set_environment_variables(email, password, groq_key)
    status_label.config(text="ColdFlow is running...")
    status_label.configure(foreground="green")
    root.update()

    excel_file_path = os.path.join(os.path.dirname(SCRIPT_DIR), 'data', 'cold_email_data.xlsx')
    coldflow_app = ColdFlow(excel_file_path)

    # Redirect logging to the text widget
    log_handler = TkinterHandler(log_text_widget)
    logger = logging.getLogger()
    logger.addHandler(log_handler)

    try:
        coldflow_app.process_excel_sheet()
        coldflow_app.save_workbook()
        logging.info("ColdFlow finished successfully.")
        status_label.config(text="ColdFlow finished.")
        status_label.configure(foreground="blue")
        messagebox.showinfo("ColdFlow", "Email processing and Excel update complete.")
    except Exception as e:
        logging.error(f"An error occurred during ColdFlow execution: {e}", exc_info=True)
        status_label.config(text="ColdFlow encountered an error.", fg="red")
        messagebox.showerror("ColdFlow Error", f"An error occurred: {e}")
    finally:
        logger.removeHandler(log_handler) # Clean up the handler

def open_settings_window():
    settings_window = tk.Toplevel(root)
    settings_window.title("ColdFlow Settings")

    email_label = ttk.Label(settings_window, text="Email Address:")
    email_label.grid(row=0, column=0, sticky=tk.W, pady=5, padx=10)
    email_entry = ttk.Entry(settings_window, width=40)
    email_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=10)

    password_label = ttk.Label(settings_window, text="Email Password:")
    password_label.grid(row=1, column=0, sticky=tk.W, pady=5, padx=10)
    password_entry = ttk.Entry(settings_window, width=40, show="*")
    password_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=10)

    groq_key_label = ttk.Label(settings_window, text="Groq API Key:")
    groq_key_label.grid(row=2, column=0, sticky=tk.W, pady=5, padx=10)
    groq_key_entry = ttk.Entry(settings_window, width=40, show="*")
    groq_key_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=10)

    def save_settings():
        save_credentials(email_entry.get(), password_entry.get(), groq_key_entry.get())
        messagebox.showinfo("Settings", "Credentials saved.")
        settings_window.destroy()

    save_button = ttk.Button(settings_window, text="Save Settings", command=save_settings)
    save_button.grid(row=3, column=0, columnspan=2, pady=10, padx=10)

    # Load saved credentials into the settings window
    saved_email, saved_password, saved_groq_key = load_credentials()
    if saved_email:
        email_entry.insert(0, saved_email)
    if saved_password:
        password_entry.insert(0, saved_password)
    if saved_groq_key:
        groq_key_entry.insert(0, saved_groq_key)

def main():
    global root
    root = tk.Tk()
    root.title("ColdFlow")

    main_frame = ttk.Frame(root, padding=20)
    main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

    settings_button = ttk.Button(main_frame, text="Settings", command=open_settings_window)
    settings_button.grid(row=0, column=0, sticky=tk.W, pady=10)

    log_label = ttk.Label(main_frame, text="Logs:")
    log_label.grid(row=1, column=0, sticky=tk.W, pady=5)
    log_text_widget = scrolledtext.ScrolledText(main_frame, height=10, width=60)
    log_text_widget.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
    log_text_widget.config(state=tk.DISABLED) # Make it read-only

    status_label = ttk.Label(main_frame, text="", font=("Arial", 10))
    status_label.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=10)

    run_button = ttk.Button(main_frame, text="Run ColdFlow", command=lambda: run_coldflow_app(log_text_widget, status_label))
    run_button.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=10)

    root.mainloop()

class TkinterHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        self.text_widget.config(state=tk.NORMAL)
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.see(tk.END)  # Autoscroll to the bottom
        self.text_widget.config(state=tk.DISABLED)

if __name__ == "__main__":
    main()