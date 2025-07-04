import subprocess
import sys
import logging
import os
from utils import backup_deelnemers_file

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Global variable for second period status
IS_SECOND_PERIOD_STARTED = False

def ask_second_period_status():
    """Ask user if the second period has started and set global variable."""
    global IS_SECOND_PERIOD_STARTED
    
    while True:
        response = input("Is the second period started? (y/n): ").strip().lower()
        if response in ['y', 'yes']:
            IS_SECOND_PERIOD_STARTED = True
            logger.info("✅ Second period is marked as STARTED")
            break
        elif response in ['n', 'no']:
            IS_SECOND_PERIOD_STARTED = False
            logger.info("✅ Second period is marked as NOT STARTED")
            break
        else:
            print("Please enter 'y' or 'n'")

def run_generate_regelmatigheidscriterium():
    try:
        # Pass the global variable as environment variable
        env = os.environ.copy()
        env['IS_SECOND_PERIOD_STARTED'] = str(IS_SECOND_PERIOD_STARTED)
        subprocess.run([sys.executable, 'generate_regelmatigheidscriterium.py'], check=True, env=env)
        logger.info("✅ Regelmatigheidscriterium generated successfully.")
    except Exception as e:
        logger.error(f"❌ Error generating regelmatigheidscriterium: {e}")
        raise

def run_generate_klassement():
    try:
        # Pass the global variable as environment variable
        env = os.environ.copy()
        env['IS_SECOND_PERIOD_STARTED'] = str(IS_SECOND_PERIOD_STARTED)
        subprocess.run([sys.executable, 'generate_klassement.py'], check=True, env=env)
        logger.info("✅ Klassement generated successfully.")
    except Exception as e:
        logger.error(f"❌ Error generating klassement: {e}")
        raise

def run_teams_sta():
    try:
        # Pass the global variable as environment variable
        env = os.environ.copy()
        env['IS_SECOND_PERIOD_STARTED'] = str(IS_SECOND_PERIOD_STARTED)
        subprocess.run([sys.executable, 'team_klassement.py'], check=True, env=env)
        logger.info("✅ Team klassement (STA) generated successfully.")
    except Exception as e:
        logger.error(f"❌ Error generating team klassement (STA): {e}")
        raise

def run_teams_dam():
    try:
        # Pass the global variable as environment variable
        env = os.environ.copy()
        env['IS_SECOND_PERIOD_STARTED'] = str(IS_SECOND_PERIOD_STARTED)
        subprocess.run([sys.executable, 'team_DAM_klassement.py'], check=True, env=env)
        logger.info("✅ Team klassement (DAM) generated successfully.")
    except Exception as e:
        logger.error(f"❌ Error generating team klassement (DAM): {e}")
        raise

def run_combine():
    try:
        # Pass the global variable as environment variable
        env = os.environ.copy()
        env['IS_SECOND_PERIOD_STARTED'] = str(IS_SECOND_PERIOD_STARTED)
        subprocess.run([sys.executable, 'combine_files.py'], check=True, env=env)
        logger.info("✅ Files combined successfully.")
    except Exception as e:
        logger.error(f"❌ Error combining files: {e}")
        raise

def run_search_mail():
    try:
        subprocess.run([sys.executable, 'check_mail.py'], check=True)
        logger.info("✅ Mail search successfully.")
    except Exception as e:
        logger.error(f"❌ Error searching mail: {e}")
        raise
    
def run_send_mail():
    try:
        subprocess.run([sys.executable, 'send_mail.py'], check=True)
        logger.info("✅ Mail sent successfully.")
    except Exception as e:
        logger.error(f"❌ Error sending mail: {e}")
        raise
    
def run_deelnemers_file():
    try:
        subprocess.run([sys.executable, 'download_deelnemers_file.py'], check=True)
        logger.info("✅ Deelnemers file downloaded successfully.")
    except Exception as e:
        logger.error(f"❌ Error downloading deelnemers file: {e}")
        raise

def run_utils():
    try:
        subprocess.run([sys.executable, 'utils.py'], check=True)
        logger.info("✅ Utils executed.")
    except Exception as e:
        logger.error(f"❌ Error executing utils: {e}")
        raise

if __name__ == '__main__':
    logger.info("Starting the generation process...")
    
    # Ask about second period status at the beginning
    ask_second_period_status()
    
    run_deelnemers_file()

    run_search_mail()
    run_generate_klassement()
    run_generate_regelmatigheidscriterium()
    run_teams_sta()
    run_teams_dam()
    run_combine()
    run_send_mail()
    
    try: 
        backup_deelnemers_file()
        logger.info("✅ Deelnemers file backed up.")
    except Exception as e:
        logger.error(f"❌ Error backing up deelnemers file: {e}")

    logger.info("Generation process completed.")