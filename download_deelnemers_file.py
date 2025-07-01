import pandas as pd
import requests
import logging
import os
from io import BytesIO
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configuration from environment variables
GOOGLE_SHEETS_ID = os.getenv("GOOGLE_SHEETS_ID")
WORKSHEET_GID = os.getenv("GOOGLE_SHEETS_GID")
DEELNEMERS_FOLDER = "Deelnemers"
OUTPUT_FILE = os.path.join(DEELNEMERS_FOLDER, "deelnemerslijst 2025.xlsx")

def download_google_sheets_as_excel():
    """Download Google Sheets data directly as Excel file."""
    try:
        # Validate environment variables
        if not GOOGLE_SHEETS_ID or not WORKSHEET_GID:
            raise ValueError("GOOGLE_SHEETS_ID and GOOGLE_SHEETS_GID must be set in .env file")
        
        # Create Deelnemers folder if it doesn't exist
        os.makedirs(DEELNEMERS_FOLDER, exist_ok=True)
        logger.info(f"📁 Ensured {DEELNEMERS_FOLDER} folder exists")
        
        # Construct the Excel export URL for public sheets
        excel_url = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEETS_ID}/export?format=xlsx&gid={WORKSHEET_GID}"
        
        logger.info("📡 Downloading Excel file from Google Sheets...")
        response = requests.get(excel_url)
        response.raise_for_status()
        
        # Save the Excel file directly
        with open(OUTPUT_FILE, 'wb') as f:
            f.write(response.content)
        
        logger.info(f"💾 Downloaded and saved as {OUTPUT_FILE}")
        
        # Load and show preview of data
        df = pd.read_excel(OUTPUT_FILE)
        logger.info(f"✅ File contains {len(df)} records")
        
        # Show preview of data
        logger.info("📋 Data preview:")
        logger.info(f"   Columns: {list(df.columns)}")
        logger.info(f"   First few rows:")
        for i, row in df.head(3).iterrows():
            logger.info(f"      {dict(row)}")
        
        return df
        
    except Exception as e:
        logger.error(f"❌ Failed to download Google Sheets: {e}")
        raise

if __name__ == '__main__':
    download_google_sheets_as_excel()