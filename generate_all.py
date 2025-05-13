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

if __name__ == '__main__':
    print("Starting the generation process...")
    
    run_generate_regelmatigheidscriterium()
    run_generate_klassement()

    print("Generation process completed.")
