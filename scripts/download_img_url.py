import sys
import subprocess

# Ensure required packages

def ensure_package(pkg, imp=None):
    try:
        __import__(imp or pkg)
    except ImportError:
        print(f"📦 Installing missing package: {pkg}")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", pkg])
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to install {pkg}: {e}")
            sys.exit(1)

ensure_package("pandas")
ensure_package("requests")
ensure_package("tqdm")
ensure_package("openpyxl")

import os
import signal
import pandas as pd
import requests
import threading
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm
from collections import defaultdict

# Config

possible_id_col = ["FactoryId", "SerialNumber"]
max_workers = 32

# Helper for dynamic output folder

def get_output_folder(folder_name="downloaded_images"):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(script_dir).lower() == "scripts":
        base_dir = os.path.dirname(script_dir)
    else:
        base_dir = script_dir
    output_folder = os.path.join(base_dir, folder_name)
    os.makedirs(output_folder, exist_ok=True)
    return output_folder

output_dir = get_output_folder("downloaded_images")

# Other Helpers

def clean_url(val):
    if not isinstance(val, str):
        return None
    val = val.strip()
    if val.lower().startswith("http://") or val.lower().startswith("https://"):
        return val
    return None

def normalize_url(url):
    return url.strip().lower()

# Thread-safe filename tracking (ADDED)

used_filenames = defaultdict(set)
filename_lock = threading.Lock()

# Get input Excel file
try:
    excel_path = input("Drop the path to an Excel (.xlsx) file: ").strip().strip('"').strip("'")

    if not excel_path:
        print("No file path provided. Exiting.\n")
        sys.exit(1)

    if not os.path.isfile(excel_path):
        print(f"File not found: {excel_path}\n")
        sys.exit(1)

    if not excel_path.lower().endswith(".xlsx"):
        print(f"Invalid file type. Please provide a .xlsx file.\n")
        sys.exit(1)
except KeyboardInterrupt:
        sys.exit(1)
except Exception as e:
    print(f"An unexpected error occurred: {e}")
    sys.exit(1)

# Load Excel

print(f"Using Excel file: {excel_path}")
try:
    df = pd.read_excel(excel_path)
except Exception as e:
    print(f"Failed to read Excel file: {e}\n")
    sys.exit(1)

# Detect which ID column to use
    
id_col = None
for col in possible_id_col:
    if col in df.columns:
        id_col = col
        break
    
if not id_col:
    print(f"None of the expected ID columns found: {possible_id_col}")
    print(f"Available columns: {list(df.columns)}")
    id_col = input("Please enter the column name to use as ID: ").strip()
    if id_col not in df.columns:
        print(f"Column '{id_col}' not found in file.")
        sys.exit(1)

print(f"Using ID column: '{id_col}'")

# Detect URL columns

def detect_url_columns(df, sample_size=10):
    url_cols = []
    for col in df.columns:
        if df[col].dropna().astype(str).head(sample_size).str.startswith("http").any():
            url_cols.append(col)
    return url_cols

url_columns = detect_url_columns(df)
if not url_columns:
    print("No columns contain URLs.")
    sys.exit(1)

print(f"URL columns detected: {url_columns}")

# Create folders

for col in url_columns:
    folder_path = os.path.join(output_dir, col.strip())
    os.makedirs(folder_path, exist_ok=True)

# Download function with retry logic

def download_image(session, record_id, folder_name, url, max_retries=3):
    if not url:
        return "Skipped invalid URL"

    folder_path = os.path.join(output_dir, folder_name.strip())

    file_ext = os.path.splitext(url)[1]
    if not file_ext or len(file_ext) > 5:
        file_ext = '.png'

    # Filename-based duplicate handling

    base_name = str(record_id)

    with filename_lock:
        if f"{base_name}{file_ext}" not in used_filenames[folder_name]:
            final_name = f"{base_name}{file_ext}"
            used_filenames[folder_name].add(final_name)
        else:
            index = 2
            while True:
                final_name = f"{base_name}_{index}{file_ext}"
                if final_name not in used_filenames[folder_name]:
                    used_filenames[folder_name].add(final_name)
                    break
                index += 1

    file_path = os.path.join(folder_path, final_name)

    for attempt in range(1, max_retries + 1):
        try:
            response = session.get(url, timeout=10)
            response.raise_for_status()
            with open(file_path, 'wb') as f:
                f.write(response.content)
            return "Downloaded"
        except Exception:
            if attempt == max_retries:
                return "Failed"

# ---- Main ----

def main():
    tasks = []
    success_count = 0
    fail_count = 0
    stop_flag = threading.Event()

    # Signal handler only sets the flag — message is printed after pbar.close()
    # so tqdm finishes its line before we write anything.
    def handle_interrupt(sig, frame):
        stop_flag.set()

    signal.signal(signal.SIGINT, handle_interrupt)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        with requests.Session() as session:
            for row in df.itertuples(index=False):
                if stop_flag.is_set():
                    break
                record_id = str(getattr(row, id_col)).strip()
                if not record_id:
                    continue
                for col in url_columns:
                    if stop_flag.is_set():
                        break
                    raw_url = getattr(row, col)
                    url = clean_url(raw_url)
                    if not url:
                        continue
                    tasks.append(executor.submit(download_image, session, record_id, col, url))

        print(f"Starting download of {len(tasks)} files... Press Ctrl+C to cancel.")
        completed = set()
        
        class _TtyStream:
            """Wraps a stream and forces isatty()=True so tqdm uses \\r-based updates."""
            def __init__(self, s): self._s = s
            def write(self, data): return self._s.write(data)
            def flush(self): self._s.flush()
            def isatty(self): return True

        pbar = tqdm(
            total=len(tasks), desc="Downloading", unit="file",
            file=_TtyStream(sys.stdout), ascii="░█", ncols=80, dynamic_ncols=False,
            bar_format="{desc}: \033[96m{percentage:3.0f}%\033[0m|{bar}| \033[96m{n_fmt}\033[0m/\033[96m{total_fmt}\033[0m [{elapsed}<{remaining}, {rate_fmt}]",
        )
        try:
            while len(completed) < len(tasks) and not stop_flag.is_set():
                for future in tasks:
                    if future in completed:
                        continue
                    if future.done():
                        completed.add(future)
                        try:
                            result = future.result()
                            if result == "Downloaded":
                                success_count += 1
                            elif result == "Failed":
                                fail_count += 1
                        except Exception:
                            fail_count += 1
                        pbar.update(1)
                threading.Event().wait(0.1)
        except KeyboardInterrupt:
            stop_flag.set()
        finally:
            pbar.close()
            if stop_flag.is_set():
                print("\nInterrupt received, stopping downloads...")
                os._exit(0)
                    
    print(f"All downloads attempted. {success_count} success, {fail_count} failed.")
    print(f"Images can be found inside folder: '{output_dir}'\n")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        sys.exit(1)
