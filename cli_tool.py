import subprocess
import sys
import os
import argparse

def ensure_package(pkg, imp=None):
    try:
        __import__(imp or pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", pkg])

ensure_package("flask")

from flask import Flask, render_template, request, redirect, url_for, flash

def resource_path(relative_path):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


# ---------- Menu list ----------
SCRIPTS = {
    "analyze_json_zip": ("Analyze JSON/ZIP", "scripts/analyze_json_zip.py"),
    "analyze_mtf_data": ("Analyze MTF Data", "scripts/analyze_mtf_data.py"),
    "camera_qc_analyzer": ("Camera QC Analyzer", "scripts/camera_qc_analyzer.py"),
    "csv_convert_to_excel": ("CSV Convert to Excel", "scripts/csv_convert_to_excel.py"),
    "csv_split_tests": ("CSV Split Tests", "scripts/csv_split_tests.py"),
    "download_img_url": ("Download Images from URL", "scripts/download_img_url.py"),
    "FileParser": ("File Parser", "scripts/FileParser.exe"),
    "format_mic_calibration_file": ("Format Mic Calibration File", "scripts/format_mic_calibration_file.py"),
    "qrcode_gen": ("Generate QR Code", "scripts/qrcode_gen.py"),
    "validate_limits": ("Validate Limits", "scripts/validate_limits.py"),
}

# ---------- Flask App Setup ----------
app = Flask(__name__)
app.secret_key = 'manufacturing_cli_tool_secret_key'

@app.route('/')
def index():
    return render_template('index.html', scripts=SCRIPTS)

@app.route('/run/<script_key>', methods=['GET', 'POST'])
def run_script(script_key):
    if script_key not in SCRIPTS:
        flash('Script not found', 'error')
        return redirect(url_for('index'))

    desc, script_rel_path = SCRIPTS[script_key]
    script_path = resource_path(script_rel_path)

    if not os.path.exists(script_path):
        flash(f'Script not found: {script_path}', 'error')
        return redirect(url_for('index'))

    try:
        if script_path.endswith('.py'):
            process = subprocess.Popen([sys.executable, script_path])
        elif script_path.endswith('.exe'):
            process = subprocess.Popen([script_path])
        else:
            flash('Unsupported file type', 'error')
            return redirect(url_for('index'))

        process.wait()
        flash(f'Successfully ran: {desc}', 'success')
    except Exception as e:
        flash(f'Error running {desc}: {str(e)}', 'error')

    return redirect(url_for('index'))

def run_web_server(port=5000):
    print(f"Starting web server on http://localhost:{port}")
    print("Press Ctrl+C to stop the server")
    app.run(host='localhost', port=port, debug=False)


# ---------- Display Menu ----------
def display_menu():
    print("\n=== Manufacturing CLI Tool ===\n")
    keys = list(SCRIPTS.keys())
    for i, key in enumerate(keys, 1):
        desc, _ = SCRIPTS[key]
        print(f"{i}. {desc}")
    print("q. Quit")


# ---------- Get User Choice ----------
def get_user_choice():
    keys = list(SCRIPTS.keys())
    while True:
        display_menu()
        try:
            choice = input(f"\nSelect a script to run (1–{len(keys)}): ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print("\nExiting.\n")
            return None

        if choice == "q":
            print("Exiting.\n")
            return None

        if choice.isdigit():
            index = int(choice) - 1
            if 0 <= index < len(keys):
                return SCRIPTS[keys[index]]

        print("Invalid choice. Please try again.")


# ---------- Prompt after script runs ----------
def prompt_post_script():
    while True:
        try:
            user_input = input("\nPress r to return to menu, or q to quit: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print("\nExiting.\n")
            return False

        if user_input == "r":
            return True
        elif user_input == "q":
            print("Exiting.\n")
            return False
        else:
            print("Invalid input. Please enter r or q.")


# ---------- Main Loop ----------
def main():
    while True:
        result = get_user_choice()
        if result is None:
            return

        desc, script_rel_path = result
        script_path = resource_path(script_rel_path)

        if not os.path.exists(script_path):
            print(f"Script not found: {script_path}")
            continue

        print(f"\n▶ Running: {desc}\n")

        try:
            if script_path.endswith(".py"):
                process = subprocess.Popen([sys.executable, script_path])
            elif script_path.endswith(".exe"):
                process = subprocess.Popen([script_path])
            else:
                print("Unsupported file type.")
                continue

            try:
                process.wait()
            except KeyboardInterrupt:
                process.wait()

        except KeyboardInterrupt:
            print("\nExiting.\n")
            return

        if not prompt_post_script():
            return


# ---------- Entry Point ----------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Manufacturing CLI Tool')
    parser.add_argument('--web', action='store_true', help='Run as web server')
    parser.add_argument('--port', type=int, default=5000, help='Port for web server (default: 5000)')

    args = parser.parse_args()

    if args.web:
        try:
            run_web_server(args.port)
        except KeyboardInterrupt:
            print("\nWeb server stopped.")
    else:
        try:
            main()
        except KeyboardInterrupt:
            pass
        finally:
            sys.exit(0)
