import subprocess
import sys
import os
import threading
import argparse

def ensure_package(pkg, imp=None):
    try:
        __import__(imp or pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", pkg])

ensure_package("flask")
ensure_package("flask-socketio", "flask_socketio")

from flask import Flask, render_template, request, redirect, url_for, flash, send_from_directory
from flask_socketio import SocketIO, emit

def resource_path(relative_path):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


# ---------- Menu list ----------
SCRIPTS = {
    "analyze_json_zip":            ("Analyze JSON/ZIP",            "scripts/analyze_json_zip.py"),
    "analyze_mtf_data":            ("Analyze MTF Data",            "scripts/analyze_mtf_data.py"),
    "camera_qc_analyzer":          ("Camera QC Analyzer",          "scripts/camera_qc_analyzer.py"),
    "csv_convert_to_excel":        ("CSV Convert to Excel",        "scripts/csv_convert_to_excel.py"),
    "csv_split_tests":             ("CSV Split Tests",             "scripts/csv_split_tests.py"),
    "download_img_url":            ("Download Images from URL",    "scripts/download_img_url.py"),
    "FileParser":                  ("File Parser",                 "scripts/FileParser.exe"),
    "format_mic_calibration_file": ("Format Mic Calibration File", "scripts/format_mic_calibration_file.py"),
    "qrcode_gen":                  ("Generate QR Code",            "scripts/qrcode_gen.py"),
    "validate_limits":             ("Validate Limits",             "scripts/validate_limits.py"),
}

# ---------- Flask + SocketIO ----------
app = Flask(__name__)
app.secret_key = 'manufacturing_cli_tool_secret_key'
socketio = SocketIO(app, async_mode='threading')

# sid -> subprocess
running_processes = {}
# sid -> generation counter (incremented each time a new script starts)
session_generations = {}

@app.route('/favicon.ico')
def favicon():
    return send_from_directory(resource_path('web_images'), 'manufacturing.png', mimetype='image/png')

# ---------- Routes ----------
@app.route('/')
def index():
    return render_template('index.html', scripts=SCRIPTS)

@app.route('/terminal/<script_key>')
def terminal(script_key):
    if script_key not in SCRIPTS:
        flash('Script not found', 'error')
        return redirect(url_for('index'))
    desc, _ = SCRIPTS[script_key]
    return render_template('terminal.html', script_key=script_key, script_desc=desc)

# ---------- SocketIO events ----------
@socketio.on('start_script')
def handle_start_script(data):
    script_key = data.get('script_key')
    if script_key not in SCRIPTS:
        emit('output', {'data': 'Script not found\r\n'})
        return

    desc, script_rel_path = SCRIPTS[script_key]
    script_path = resource_path(script_rel_path)

    if not os.path.exists(script_path):
        emit('output', {'data': f'Script not found: {script_path}\r\n'})
        return

    sid = request.sid

    # Increment generation FIRST so any running stream_output thread stops emitting
    gen = session_generations.get(sid, 0) + 1
    session_generations[sid] = gen

    # Kill any previously running process for this client
    old = running_processes.pop(sid, None)
    if old:
        try:
            old.terminate()
        except Exception:
            pass

    try:
        if script_path.endswith('.py'):
            cmd = [sys.executable, '-u', script_path]
        elif script_path.endswith('.exe'):
            cmd = [script_path]
        else:
            emit('output', {'data': 'Unsupported file type\r\n'})
            return

        env = os.environ.copy()
        env['PYTHONUTF8'] = '1'

        process = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            bufsize=0,
            cwd=resource_path('.'),
            env=env
        )
        running_processes[sid] = process
        emit('clear_terminal', {})

        def stream_output(my_gen=gen):
            try:
                while True:
                    chunk = process.stdout.read(1)
                    if not chunk:
                        break
                    if session_generations.get(sid) != my_gen:
                        return
                    socketio.emit('output', {'data': chunk.decode('utf-8', errors='replace')}, to=sid)
                process.stdout.close()
                if session_generations.get(sid) == my_gen:
                    socketio.emit('output', {'data': '\r\n\r\n[Process finished]\r\n'}, to=sid)
                    socketio.emit('script_done', {}, to=sid)
            except Exception as e:
                if session_generations.get(sid) == my_gen:
                    socketio.emit('output', {'data': f'\r\n[Stream error: {e}]\r\n'}, to=sid)
            finally:
                running_processes.pop(sid, None)

        threading.Thread(target=stream_output, daemon=True).start()

    except Exception as e:
        emit('output', {'data': f'Failed to start: {str(e)}\r\n'})

@socketio.on('stop_script')
def handle_stop_script(data=None):
    process = running_processes.pop(request.sid, None)
    if process:
        try:
            process.terminate()
        except Exception:
            pass

@socketio.on('launch_gui')
def handle_launch_gui(data):
    script_key = data.get('script_key')
    if script_key not in SCRIPTS:
        return
    desc, script_rel_path = SCRIPTS[script_key]
    script_path = resource_path(script_rel_path)
    if not os.path.exists(script_path):
        return
    try:
        subprocess.Popen([sys.executable, script_path], cwd=resource_path('.'))
    except Exception:
        pass

@socketio.on('input')
def handle_input(data):
    process = running_processes.get(request.sid)
    if process and process.stdin:
        try:
            process.stdin.write(data['data'].encode('utf-8'))
            process.stdin.flush()
        except Exception:
            pass

@socketio.on('disconnect')
def handle_disconnect():
    sid = request.sid
    session_generations.pop(sid, None)
    process = running_processes.pop(sid, None)
    if process:
        try:
            process.terminate()
        except Exception:
            pass

def run_web_server(port=5000):
    print(f"Starting web server on http://localhost:{port}")
    print("Press Ctrl+C to stop the server")
    socketio.run(app, host='localhost', port=port, debug=False)


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
