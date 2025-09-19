import subprocess
import sys
import os

def resource_path(relative_path):
    """Get absolute path to resource (for dev and for PyInstaller onefile)."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# MENU LIST

SCRIPTS = {
    "1": ("Script 1 - Download Images from URL", "download_img_url.py"),
    "2": ("Script 2 - Analyze JSON/ZIP", "analyze_json_zip.py"),
    "3": ("Script 3 - Convert CSV to Excel", "convert_csv_excel_headers.py"),
    "4": ("Script 4 - Split CSV Tests", "split_csv_tests.py"),
    "5": ("Script 5 - File Parser", "FileParser.exe"),
    "6": ("Script 6 - Validate Limits", "validate_limits.py"),
    "7": ("Script 7 - Format Mic Calibration File", "format_mic_calibration_file.py")
}

def display_menu():
    print("\n=== Manufacturing CLI Tool ===\n")
    for key, (desc, _) in SCRIPTS.items():
        print(f"{key}. {desc}")
    print("q. Quit")

def get_user_choice():
    while True:
        display_menu()
        choice = input("\nSelect a script to run (1–7): ").strip().lower()
        if choice == 'q':
            print("Exiting.\n")
            return None
        if choice in SCRIPTS:
            return SCRIPTS[choice]
        print("\nInvalid choice. Please try again.")

def prompt_post_script():
    while True:
        user_input = input("Press r to return to menu, or q to quit: ").strip().lower()
        if user_input == 'r':
            return True  # go back to menu
        elif user_input == 'q':
            print("Exiting.\n")
            return False
        else:
            print("Invalid input. Please enter r or q.")

# ---- Main ----

def main():
    while True:
        result = get_user_choice()
        if result is None:
            break

        desc, script_rel_path = result
        script_path = resource_path(script_rel_path)

        if not os.path.exists(script_path):
            print(f"Script not found: {script_path}")
            continue

        print(f"\nRunning: {desc}")

        try:
            if script_path.endswith(".py"):
                process = subprocess.Popen([sys.executable, script_path])
            elif script_path.endswith(".exe"):
                process = subprocess.Popen([script_path])
            else:
                print("Unsupported file type.")
                continue

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

        # after script ends
        if not prompt_post_script():
            break

if __name__ == "__main__":
    main()
