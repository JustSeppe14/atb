import smtplib
import os
from email.message import EmailMessage
from dotenv import load_dotenv
load_dotenv()
import logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Email configuration
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL_ADDRESS = os.getenv("EMAIL_ACCOUNT")      # Replace with your email
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")           # Replace with your app password or email password
EMAIL_RECIPIENTS = os.getenv("EMAIL_RECIPIENTS", "")
RECIPIENT = [email.strip() for email in EMAIL_RECIPIENTS.split(",") if email.strip()]      # Replace with recipient's email

# File to send
ATTACHMENT_PATH = 'wedstrijd_data_2025.xlsx'

def send_email():
    try:
        msg = EmailMessage()
        msg['Subject'] = 'Wedstrijd Data 2025'
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = ', '.join(RECIPIENT)  # Join multiple recipients with a comma
        msg.set_content(
            "Beste,\n\nIn de bijlage vindt u het aangepaste klassement voor deze week ('wedstrijd_data_2025.xlsx').\n\nMet vriendelijke groet,\nUw automatiseringsscript"
        )

        # Attach the file
        with open(ATTACHMENT_PATH, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(ATTACHMENT_PATH)
            msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

        # Send the email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            smtp.send_message(msg)
            logger.info(f"✅ Email sent to {RECIPIENT} with attachment '{file_name}'.")
    except Exception as e:
        logger.error(f"❌ Failed to send email: {e}")
        raise

if __name__ == '__main__':
    send_email()