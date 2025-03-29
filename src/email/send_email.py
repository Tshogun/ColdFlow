import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import sys
from dotenv import load_dotenv

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../logger')))
from logger import Logger

log = Logger()

def getenv():
    # Load environment variables from .env file
    dotenv_path = os.path.join(os.path.dirname(__file__), '..', 'credentials', 'email.env')
    load_dotenv(dotenv_path=dotenv_path)

    password = os.getenv("EMAIL_PASSWORD")
    if not password:
        log.error("Error: EMAIL_PASSWORD not found in .env file")
    else:
        log.info(f"Successfully loaded EMAIL_PASSWORD: {password[:5]}...") 

    sender_email = os.getenv("EMAIL_ADDRESS")
    if not sender_email:
        log.error("Error: EMAIL_ADDRESS not found in .env file")
    else:
        log.info(f"Successfully loaded address: {sender_email[:5]}")
        
    return sender_email, password

class EmailSender:
    def __init__(self, sender_email, password, smtp_server="smtp.gmail.com", smtp_port=587):
        self.sender_email = sender_email
        self.password = password
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        log.info(f"EmailSender initialized with sender: {self.sender_email}")

    def send_email(self, receiver_email, subject, body, attachment_path=None):
        """Send an email with an optional attachment and body content."""
        log.info(f"Sending email to: {receiver_email}, subject: {subject}")
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = receiver_email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))
            log.debug(f"Email body attached: {body}")

            if attachment_path:
                log.info(f"Attachment path provided: {attachment_path}")
                self._attach_file(msg, attachment_path)

            self._send_email(msg, receiver_email)
            log.info(f"Email sent successfully to: {receiver_email}")
            return True
        except Exception as e:
            log.error(f"Failed to send email to {receiver_email}: {e}", exc_info=True)  # Log exception details
            return False

    def _attach_file(self, msg, attachment_path):
        """Attach a file to the email message."""
        log.info(f"Attaching file: {attachment_path}")
        try:
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
                msg.attach(part)
                log.debug(f"File attached: {attachment_path}")
        except FileNotFoundError:
            log.error(f"Attachment file not found: {attachment_path}")
            raise  # Re-raise the exception to be caught in send_email
        except Exception as e:
            log.error(f"Error opening attachment: {e}", exc_info=True)
            raise  # Re-raise the exception

    def _send_email(self, msg, receiver_email):
        """Send the email using SMTP."""
        log.info(f"Sending email via SMTP to: {receiver_email}")
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.password)
                log.debug("SMTP connection established and authenticated.")
                server.sendmail(self.sender_email, receiver_email, msg.as_string())
                log.debug("Email sent via SMTP.")
        except smtplib.SMTPAuthenticationError as auth_err:
            log.error(f"SMTP Authentication Error: {auth_err}", exc_info=True)
            raise
        except Exception as e:
            log.error(f"Error while sending email via SMTP to {receiver_email}: {e}", exc_info=True)
            raise

if __name__ == "__main__":
    # Load environment variables from .env file
    dotenv_path = os.path.join(os.path.dirname(__file__), '..', 'credentials', 'email.env')
    load_dotenv(dotenv_path=dotenv_path)

    password = os.getenv("EMAIL_PASSWORD")
    if not password:
        log.error("Error: EMAIL_PASSWORD not found in .env file")
    else:
        log.info(f"Successfully loaded EMAIL_PASSWORD: {password[:5]}...") 

    sender_email = os.getenv("EMAIL_ADDRESS")
    if not sender_email:
        log.error("Error: EMAIL_ADDRESS not found in .env file")
    else:
        log.info(f"Successfully loaded address: {sender_email[:5]}")

    if sender_email and password:
        email_sender = EmailSender(sender_email, password)
        receiver_email = "n.b24fnmail@gmail.com"  # Replace with a test email
        subject = "Test Email from ColdFlow"
        body = "This is a test email sent from ColdFlow."

        attachment_path = "src/resume/document.pdf"

        email_sent = email_sender.send_email(receiver_email, subject, body, attachment_path)
        if email_sent:
            log.info("Test email sent successfully!")
        else:
            log.debug("Test email failed to send.")
    else:
        log.error("Email credentials not found in .env file.")