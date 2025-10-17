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

# Config

id_col = 'Id'
factory_col = 'FactoryId'
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

downloaded_urls = set()

# Download function with retry logic

def download_image(session, factory_id, record_id, folder_name, url, max_retries=3):
    if not url:
        return "Skipped invalid URL"

    folder_path = os.path.join(output_dir, folder_name.strip())

    file_ext = os.path.splitext(url)[1]
    if not file_ext or len(file_ext) > 5:
        file_ext = '.png'

    file_name = f"{factory_id}-{record_id}"
    file_path = os.path.join(folder_path, f"{file_name}{file_ext}")
    file_index = 1
    while os.path.exists(file_path):
        file_path = os.path.join(folder_path, f"{file_name}_{file_index}{file_ext}")
        file_index += 1

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

    def handle_interrupt(sig, frame):
        print("\nInterrupt received, stopping downloads...")
        stop_flag.set()

    signal.signal(signal.SIGINT, handle_interrupt)
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        with requests.Session() as session:
            for row in df.itertuples(index=False):
                if stop_flag.is_set():
                    break
                factory_id = str(getattr(row, factory_col)).strip()
                record_id = str(getattr(row, id_col)).strip()
                if not factory_id or not record_id:
                    continue
                for col in url_columns:
                    if stop_flag.is_set():
                        break
                    raw_url = getattr(row, col)
                    url = clean_url(raw_url)
                    if not url:
                        continue
                    url_key = normalize_url(url)
                    if url_key in downloaded_urls:
                        continue
                    downloaded_urls.add(url_key)
                    tasks.append(executor.submit(download_image, session, factory_id, record_id, col, url))

        print(f"Starting download of {len(tasks)} files... Press Ctrl+C to cancel.")
        completed = set()
        
        pbar = tqdm(total=len(tasks), desc="Downloading", unit="file")
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
            print("\nInterrupt received, stopping downloads...")
        finally:
            pbar.close()
            if stop_flag.is_set():
                sys.exit(0)
                    
    print(f"All downloads attempted. {success_count} success, {fail_count} failed.")
    print(f"Images can be found inside folder: '{output_dir}'\n")

if __name__ == "__main__":
    main()
