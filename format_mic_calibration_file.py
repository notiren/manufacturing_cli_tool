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

ensure_package("numpy")

import re
import numpy as np
import os

# START SCRIPT

def format_calibration_file(input_file, output_folder):
    # --- Step 0: Validate file extension ---
    if not input_file.lower().endswith(".txt"):
        raise ValueError("Input file must be a .txt file")

    # --- Step 1: Read file ---
    try:
        with open(input_file, "r") as f:
            lines = f.readlines()
    except Exception as e:
        raise IOError(f"Could not read file: {e}")

    if len(lines) < 3:
        raise ValueError("File does not contain enough lines of data")

    # --- Step 2: Extract metadata ---
    sens_match = re.search(r"Sens Factor\s*=\s*([-\d.]+)dB", lines[0])
    serno_match = re.search(r"SERNO:\s*(\d+)", lines[0])

    if not sens_match or not serno_match:
        raise ValueError("Could not extract SensitivityFactor or SerialNumber from header")

    try:
        sensitivity_factor = float(sens_match.group(1))
        serial_number = int(serno_match.group(1))
    except Exception:
        raise ValueError("Invalid SensitivityFactor or SerialNumber format")

    # --- Step 3: Extract frequency/dB data ---
    freqs, db_vals = [], []
    bad_lines = 0
    for line in lines[2:]:
        parts = line.strip().split()
        if len(parts) == 2:
            try:
                freq, db_val = map(float, parts)
                freqs.append(freq)
                db_vals.append(db_val)
            except ValueError:
                bad_lines += 1
        elif line.strip():
            bad_lines += 1

    if not freqs:
        raise ValueError("No valid frequency/dB pairs found in file")

    if bad_lines > 0:
        print(f"Warning: Skipped {bad_lines} malformed data lines")

    freqs = np.array(freqs)
    db_vals = np.array(db_vals)

    # --- Step 4: Reference frequencies ---
    ref_freqs = np.array([
        100.0, 106.0, 112.0, 118.0, 125.0, 132.0, 140.0, 150.0, 160.0, 170.0,
        180.0, 190.0, 200.0, 212.0, 224.0, 236.0, 250.0, 265.0, 280.0, 300.0,
        315.0, 335.0, 355.0, 375.0, 400.0, 425.0, 450.0, 475.0, 500.0, 530.0,
        560.0, 600.0, 630.0, 670.0, 710.0, 750.0, 800.0, 850.0, 900.0, 950.0,
        1000.0, 1060.0, 1120.0, 1180.0, 1250.0, 1320.0, 1400.0, 1500.0, 1600.0,
        1700.0, 1800.0, 1900.0, 2000.0, 2120.0, 2240.0, 2360.0, 2500.0, 2650.0,
        2800.0, 3000.0, 3150.0, 3350.0, 3550.0, 3750.0, 4000.0, 4250.0, 4500.0,
        4750.0, 5000.0, 5300.0, 5600.0, 6000.0, 6300.0, 6700.0, 7100.0, 7500.0,
        8000.0, 8500.0, 9000.0, 9500.0, 10000.0, 10600.0, 11200.0, 11800.0, 
        12500.0
    ])

    # --- Step 5: Interpolate values ---
    ref_db_vals = np.interp(ref_freqs, freqs, db_vals)

    # --- Step 6: Define output folder ---
    script_dir = os.path.dirname(os.path.abspath(__file__))
    extracted_dir = os.path.join(script_dir, output_folder)
    os.makedirs(extracted_dir, exist_ok=True)

    output_file = os.path.join(extracted_dir, f"{serial_number}_filtered.txt")

    # --- Step 7: Write output ---
    with open(output_file, "w") as f:
        f.write("SerialNumber\tSensitivityFactor\tFrequency\tDbValue\tLowerLimit\tUpperLimit\n")
        for freq, db_val in zip(ref_freqs, ref_db_vals):
            f.write(f"{serial_number}\t{sensitivity_factor:.2f}\t{freq:.1f}\t{db_val:.2f}\t-17.00\t-33.00\n")

    print(f"Formatted file saved to: {output_folder}/{serial_number}_filtered.txt\n")

# main

if __name__ == "__main__":
    input_file = input('Drop the path to a mic calibration .txt file: ').strip().strip('"').strip("'")
    output_folder = "extracted"
    
    print(f"Using file: {os.path.basename(input_file)}")

    try:
        format_calibration_file(input_file, output_folder)
    except Exception as e:
        sys.exit(str(e))
