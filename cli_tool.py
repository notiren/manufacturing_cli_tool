import subprocess
import sys
import os

def resource_path(relative_path):
    """Get absolute path to resource (for dev and for PyInstaller onefile)."""
    try:
        base_path = sys._MEIPASS  # Temp folder in onefile mode
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# menu list
SCRIPTS = {
    "1": ("Script 1 - Download Images from URL", "download_img_url.py"),
    "2": ("Script 2 - Analyze JSON/ZIP", "analyze_json_zip.py"),
    "3": ("Script 3 - Convert CSV to Excel", "convert_csv_excel_headers.py"),
    "4": ("Script 4 - Split CSV Tests", "split_csv_tests.py"),
    "5": ("Script 5 - File Parser", "FileParser.exe"),
}

def display_menu():
    print("\n=== Manufacturing CLI Tool ===")
    for key, (desc, _) in SCRIPTS.items():
        print(f"{key}. {desc}")
    print("q. Quit")

def get_user_choice():
    while True:
        display_menu()
        choice = input("\nSelect a script to run (1–5): ").strip().lower()
        if choice == 'q':
            print("Exiting.\n")
            sys.exit(0)
        if choice in SCRIPTS:
            return SCRIPTS[choice]
        print("Invalid choice. Please try again.\n")

def main():
    desc, script_rel_path = get_user_choice()
    script_path = resource_path(script_rel_path)

    if not os.path.exists(script_path):
        print(f"Script not found: {script_path}")
        sys.exit(1)

    print(f"\nRunning: {desc}")

    try:
        if script_path.endswith(".py"):
            process = subprocess.Popen([sys.executable, script_path])
        elif script_path.endswith(".exe"):
            process = subprocess.Popen([script_path])
        else:
            print("Unsupported file type.")
            sys.exit(1)
            
        while True:
            try:
                process.wait(timeout=0.5)
                break
            except subprocess.TimeoutExpired:
                continue

    except KeyboardInterrupt:
        print("User interrupted. Terminating the running script...")
        if process.poll() is None:
            try:
                process.terminate()
                try:
                    process.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    process.kill()
            except Exception as e:
                print(f"Error terminating process: {e}")
    sys.exit(0)  # Exit after running one script

if __name__ == "__main__":
    main()
