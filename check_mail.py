import imaplib
import email
import os
import pickle
from dotenv import load_dotenv
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

load_dotenv()

# Config
# Config
IMAP_SERVER = "imap.gmail.com"
EMAIL_ACCOUNT = os.getenv("EMAIL_ACCOUNT")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
SENDER_FILTER = "seppevc@hotmail.be"
DOWNLOAD_FOLDER = "Result"
SAVE_AS_FILENAME = "finish.xlsx"
LAST_ID_FILE = ".last_email_id"


# Validate required environment variables
if not EMAIL_PASSWORD:
    raise ValueError("EMAIL_PASSWORD environment variable is required")

os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

def connect_mailbox():
    """Connect to IMAP server with error handling"""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
        mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
        logger.info("✅ Successfully connected to mailbox")
        return mail
    except imaplib.IMAP4.error as e:
        logger.error(f"❌ Failed to connect to mailbox: {e}")
        raise
    except Exception as e:
        logger.error(f"❌ Unexpected error connecting to mailbox: {e}")
        raise

def get_last_processed_id():
    """Get the last processed email ID"""
    if os.path.isfile(LAST_ID_FILE):
        try:
            with open(LAST_ID_FILE, "rb") as f:
                return pickle.load(f)
        except Exception as e:
            logger.warning(f"⚠️ Could not read last processed ID: {e}")
    return None

def set_last_processed_id(email_id):
    """Save the last processed email ID"""
    try:
        with open(LAST_ID_FILE, "wb") as f:
            pickle.dump(email_id, f)
        logger.info(f"📝 Updated last processed ID: {email_id}")
    except Exception as e:
        logger.error(f"❌ Could not save last processed ID: {e}")

def is_excel_file(filename):
    """Check if file is an Excel file"""
    if not filename:
        return False
    return filename.lower().endswith(('.xlsx', '.xls'))

def fetch_new_mail_and_save_attachment(mail):
    """Fetch only the most recent email and save Excel attachment if present"""
    try:
        mail.select("inbox")
        status, messages = mail.search(None, f'(FROM "{SENDER_FILTER}")')
        
        if status != "OK":
            logger.warning("❌ No emails found or search failed")
            return

        message_ids = messages[0].split()
        if not message_ids:
            logger.info("📭 No emails from sender")
            return

        # Only process the most recent email (last in the list)
        latest_msg_id = message_ids[-1]

        try:
            res, data = mail.fetch(latest_msg_id, "(RFC822)")
            if res != "OK":
                logger.warning(f"⚠️ Could not fetch email {latest_msg_id}")
                return

            msg = email.message_from_bytes(data[0][1])
            subject = msg.get("Subject", "No Subject")
            logger.info(f"📨 Processing email: {subject}")
            
            attachments = []

            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get("Content-Disposition") is None:
                    continue

                filename = part.get_filename()
                if filename and is_excel_file(filename):
                    payload = part.get_payload(decode=True)
                    if payload:
                        attachments.append((filename, payload))

            if attachments:
                # Save largest Excel attachment
                largest = max(attachments, key=lambda x: len(x[1]))
                save_path = os.path.join(DOWNLOAD_FOLDER, SAVE_AS_FILENAME)
                
                with open(save_path, "wb") as f:
                    f.write(largest[1])
                
                logger.info(f"✅ Saved '{largest[0]}' as '{SAVE_AS_FILENAME}' ({len(largest[1])} bytes)")
                
                # Update last processed ID
                set_last_processed_id(latest_msg_id)
                return
            else:
                logger.warning("⚠️ No Excel attachments found in the most recent email")

        except Exception as e:
            logger.error(f"❌ Error processing email {latest_msg_id}: {e}")

    except Exception as e:
        logger.error(f"❌ Error in fetch_new_mail_and_save_attachment: {e}")
        raise
def main():
    """Main function with proper error handling"""
    mail = None
    try:
        mail = connect_mailbox()
        fetch_new_mail_and_save_attachment(mail)
    except Exception as e:
        logger.error(f"❌ Script failed: {e}")
        return 1
    finally:
        if mail:
            try:
                mail.logout()
                logger.info("👋 Disconnected from mailbox")
            except Exception as e:
                logger.warning(f"⚠️ Error during logout: {e}")
    
    return 0

if __name__ == "__main__":
    exit(main())