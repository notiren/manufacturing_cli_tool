import zipfile
import json
import csv
import os
from io import TextIOWrapper
import sys

def process_json_data(file_name, data, results):
    serial_number = data.get("serialNumber", os.path.splitext(file_name)[0])
    sequences = data.get("sequences", [])
    vcp_datas = data.get("vcpDatas", [])
    network_tasks = data.get("networkTasks", [])
    all_sequences = []

    # check for Adapter Edac data
    if sequences:
        for sequence in sequences:
            all_sequences.append(sequence.get("sequenceDatas", []))
    # else check for PoE EELoad data
    elif vcp_datas:
        all_sequences.append(vcp_datas)
    # else check for PoE Network Data
    elif network_tasks:
        for task in network_tasks:
            all_sequences.append(task.get("taskSections", []))

    # check for false measurements
    for measures in all_sequences:
        for measure in measures:
            if not measure.get("hasPassed", True):
                name = measure.get("name")
                value = measure.get("value")
                upper = measure.get("upperLimit")
                lower = measure.get("lowerLimit")

                try:
                    value = float(value)
                    upper = float(upper) if upper is not None else None
                    lower = float(lower) if lower is not None else None
                except (ValueError, TypeError):
                    continue

                status = "within limits"
                deviation = ""

                if upper is not None and value > upper:
                    deviation = value - upper
                    status = f"above upper by {deviation:.3f}"
                elif lower is not None and value < lower:
                    deviation = lower - value
                    status = f"below lower by {deviation:.3f}"

                results.append({
                    "file": serial_number,
                    "name": name,
                    "value": value,
                    "lowerLimit": lower,
                    "upperLimit": upper,
                    "status": status,
                    "deviation": deviation
                })

def analyze_failed_measurements(zip_or_json_path, output_dir):
    results = []

    if zipfile.is_zipfile(zip_or_json_path):
        with zipfile.ZipFile(zip_or_json_path, 'r') as zip_ref:
            for file_name in zip_ref.namelist():
                if file_name.endswith('.txt') or file_name.endswith('.json'):
                    with zip_ref.open(file_name) as file:
                        try:
                            data = json.load(TextIOWrapper(file, encoding='utf-8'))
                            process_json_data(file_name, data, results)
                        except Exception as e:
                            print(f"Failed to parse {file_name}: {e}")
    else:
        if zip_or_json_path.endswith('.txt') or zip_or_json_path.endswith('.json'):
            try:
                with open(zip_or_json_path, 'r', encoding='utf-8') as file:
                    data = json.load(file)
                    process_json_data(zip_or_json_path, data, results)
            except Exception as e:
                print(f"Failed to parse {zip_or_json_path}: {e}")
                return
        elif zip_or_json_path == "":
            print("No file path provided. Exiting.\n")
            return
        else:
            print("Provided file is not a supported .json/.txt file or a .zip archive.\n")
            return

    if not results:
        print("No failed entries with complete data were found.\n")
        return

    results.sort(key=lambda x: x['file'])

    # ensure output folder exists
    os.makedirs(output_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(zip_or_json_path))[0]
    output_csv_file = os.path.join(output_dir, f"failed_{base_name}.csv")

    try:
        with open(output_csv_file, 'w', newline='', encoding='utf-8') as f:
            fieldnames = ["file", "name", "value", "lowerLimit", "upperLimit", "status", "deviation"]
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()

            last_file = None
            for row in results:
                if row['file'] != last_file:
                    if last_file is not None:
                        writer.writerow({})
                    last_file = row['file']
                writer.writerow(row)

        print(f"-----------")
        print(f"Analysis complete. Results saved inside folder: '{output_dir}'\n")

    except PermissionError:
        print(f"\nPermission denied: Could not write to '{output_csv_file}'. Is the file open in another program?\n")

if __name__ == "__main__":
    input_path = input("Drop the path to a .zip, .json, or .txt file: ").strip('"').strip("'")
    output_dir = os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), "extracted")

    analyze_failed_measurements(input_path, output_dir)
