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

ensure_package("openpyxl")

import csv
import json
import os
import openpyxl
import re

# Output folder logic

script_dir = os.path.dirname(os.path.abspath(__file__))
if os.path.basename(script_dir).lower() == "scripts":
    output_dir = os.path.join(os.path.dirname(script_dir), "extracted")
else:
    output_dir = os.path.join(script_dir, "extracted")
    
os.makedirs(output_dir, exist_ok=True)
config_path = os.path.join(script_dir, "config.json")

# Parser functions

def parse_csv(path, skip_rows):
    with open(path, "r", encoding="utf-8") as f:
        for _ in range(skip_rows):
            next(f)
        reader = csv.DictReader(f, delimiter=';')
        if not reader.fieldnames:
            raise ValueError("No headers found in CSV file.")
        return [
            {normalize_key(k): v for k, v in row.items()}
            for row in reader
        ]

def parse_xlsx(path, skip_rows=None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = wb.active
    rows = list(sheet.iter_rows(values_only=True))
    headers = [normalize_key(str(cell)) if cell is not None else "" for cell in rows[0]]
    return [
        {headers[i]: str(cell) if cell is not None else "" for i, cell in enumerate(row)}
        for row in rows[1:]
    ]

def parse_txt_json_array(path, skip_rows=None):
    with open(path, "r", encoding="utf-8") as f:
        try:
            data = json.load(f)
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON from TXT file: {e}")
            sys.exit(1)

    rows = []
    for item in data:
        serial = item.get("SerialNumber", "N/A")
        for task in item.get("NetworkTasks", []):
            for section in task.get("TaskSections", []):
                row = {
                    "SerialNumber": serial,
                    "Name": section.get("Name", ""),
                    "Value": section.get("Value", ""),
                    "IsDataSet": section.get("IsDataSet", False),
                    "NetworkChartType": section.get("NetworkChartType", "")
                }
                rows.append(row)
    return rows

# Parser mapping

PARSERS = {
    "parse_csv": parse_csv,
    "parse_xlsx": parse_xlsx,
    "parse_txt_json_array": parse_txt_json_array,
}

# Utilities

def normalize_key(key: str) -> str:
    key = key.replace("\n", " ").replace("\r", "").replace('"', '').strip()
    key = re.sub(r'\s+\(', '(', key)
    return key

def get_field(row, field_names):
    if isinstance(field_names, str):
        return row.get(field_names)
    for name in field_names:
        if name in row:
            return row.get(name)
    return None

def load_config(path="config.json"):
    if not os.path.exists(path):
        print(f"Config file '{path}' not found.")
        sys.exit(1)
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def load_limits(json_path, root_key):
    json_path = os.path.join(script_dir, json_path) if not os.path.isabs(json_path) else json_path
    if not os.path.exists(json_path):
        print(f"Limits file not found: {json_path}")
        sys.exit(1)
        
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    flat_limits = {}
    for section in data.get(root_key, []):
        for limit in section.get("limits", []):
            key = normalize_key(limit["key"])
            flat_limits[key] = (limit["lowerLimit"], limit["upperLimit"])
    return flat_limits

def validate_rows(rows, limits_dict, field_map):
    results = []
    for i, row in enumerate(rows, start=1):
        serial = get_field(row, field_map.get("serial", ["SerialNumber"])) or "N/A"
        serial = serial.strip() if isinstance(serial, str) else "N/A"

        key_field = field_map.get("key", "NetworkChartType")
        value_field = field_map.get("value", "Value")
        name_field = field_map.get("name", "Name")

        if key_field in row and value_field in row:
            key = row.get(key_field)
            name = row.get(name_field, "")
            val_str = row.get(value_field, "")
            if not key:
                results.append(f"[Row {i} | SN: {serial}] Missing key '{key_field}' (Name: '{name}')")
                continue
            if val_str == "":
                results.append(f"[Row {i} | SN: {serial}] Missing value for '{name}'")
                continue
            try:
                val = float(val_str)
            except ValueError:
                continue

            low, high = limits_dict.get(key, (None, None))
            label = name if name else key
            if (low is not None and val < low) or (high is not None and val > high):
                range_str = f"{low if low is not None else '-∞'}–{high if high is not None else '∞'}"
                results.append(f"[Row {i} | SN: {serial}] ❌ '{label}' = {val} (Out of range: {range_str})")
        else:
            for key, (low, high) in limits_dict.items():
                raw = row.get(key, "")
                if not raw:
                    results.append(f"[Row {i} | SN: {serial}] Missing value for '{key}'")
                    continue
                try:
                    val = float(raw.replace(",", "."))
                except ValueError:
                    continue
                if (low is not None and val < low) or (high is not None and val > high):
                    range_str = f"{low if low is not None else '-∞'}–{high if high is not None else '∞'}"
                    results.append(f"[Row {i} | SN: {serial}] ❌ '{key}' = {val} (Out of range: {range_str})")

    return results

def validate_file(path, limits_dict, parser_func, skip_rows, output_log, field_map):
    rows = parser_func(path, skip_rows) if parser_func == parse_csv else parser_func(path)
    results = validate_rows(rows, limits_dict, field_map)

    print("-----------")
    if not results:
        print(f"No failed entries were found. Tests are validated!")
    else:
        os.makedirs(os.path.dirname(output_log), exist_ok=True)
        with open(output_log, "w", encoding="utf-8") as out:
            out.write("\n".join(results))
        print(f"Analysis complete. Results saved in: {output_log}")

# ---- Main ----

def main():
    config = load_config(config_path)
    SKIP_ROWS = config.get("SKIP_ROWS", 3)
    VALIDATORS = config.get("VALIDATORS", {})
    OUTPUT_LOG = os.path.join(output_dir, "validation_results.txt")

    print("Please select one of the options:")
    keys = list(VALIDATORS.keys())
    for i, key in enumerate(keys, start=1):
        print(f"{i}. {VALIDATORS[key]['label']}")

    choice = input(f"Enter 1-{len(keys)}: ").strip()
    if not choice.isdigit() or not (1 <= int(choice) <= len(keys)):
        print("Invalid choice.")
        sys.exit(1)
    selected_key = keys[int(choice) - 1]
    selected = VALIDATORS[selected_key]

    file_path = input(f"Drop the path to the {selected['label']} file: ").strip().strip('"')

    limits = load_limits(selected['limits']['json'], selected['limits']['root'])
    parser_name = selected.get('parser')
    parser_func = PARSERS.get(parser_name)
    if parser_func is None:
        print(f"No parser function found for '{parser_name}'")
        sys.exit(1)

    field_map = selected.get("fields", {})
    validate_file(file_path, limits, parser_func, SKIP_ROWS, OUTPUT_LOG, field_map)

if __name__ == "__main__":
    main()
