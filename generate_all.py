import subprocess
import sys

def run_generate_regelmatigheidscriterium():
    try:
        # Run the generate_regelmatigheidscriterium.py script
        subprocess.run([sys.executable, 'generate_regelmatigheidscriterium.py'], check=True)
        print("✅ Regelmatigheidscriterium generated successfully.")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to generate Regelmatigheidscriterium: {e}")

def run_generate_klassement():
    try:
        # Run the generate_klassement.py script
        subprocess.run([sys.executable, 'generate_klassement.py'], check=True)
        print("✅ Klassement generated successfully.")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to generate Klassement: {e}")

def run_teams_sta():
    try:
        subprocess.run([sys.executable, 'team_klassement.py'], check=True)
        print("✅ Team klassement (STA) generated successfully.")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to generate Team klassement (STA): {e}")
        
def run_teams_dam():
    try:
        subprocess.run([sys.executable, 'team_DAM_klassement.py'], check=True)
        print("✅ Team klassement (DAM) generated successfully.")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to generate Team klassement (DAM): {e}")
def run_combine():
    try:
        # Run the generate_klassement.py script
        subprocess.run([sys.executable, 'combine_files.py'], check=True)
        print("✅ Files combined successfully.")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to combine files: {e}")
        
def run_utils():
    try:
        # Run the generate_klassement.py script
        subprocess.run([sys.executable, 'utils.py'], check=True)
        print("✅ Utils executed.")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to execute Utils: {e}")

if __name__ == '__main__':
    print("Starting the generation process...")

    run_generate_klassement()
    run_generate_regelmatigheidscriterium()
    run_teams_sta()
    run_teams_dam()
    
    run_combine()

    print("Generation process completed.")
