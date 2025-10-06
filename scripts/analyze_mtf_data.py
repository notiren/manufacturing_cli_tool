import sys
import subprocess

# -----------------------------------
# Ensure required packages
# -----------------------------------
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
ensure_package("pandas")

import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment

# -----------------------------------
# Utility functions
# -----------------------------------

def get_output_folder(folder_name="extracted"):
    """Determine and create output folder near script."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(script_dir).lower() == "scripts":
        base_dir = os.path.dirname(script_dir)
    else:
        base_dir = script_dir
    output_folder = os.path.join(base_dir, folder_name)
    os.makedirs(output_folder, exist_ok=True)
    return output_folder


def get_limit(index: int) -> float:
    """Return threshold limit based on B index."""
    if index == 0:
        return 64.0
    elif 1 <= index <= 4:
        return 45.1
    elif 5 <= index <= 10:
        return 47.5
    elif 11 <= index <= 16:
        return 59.7
    else:
        return None


# -----------------------------------
# Core logic
# -----------------------------------

def process_sheet(df, max_pairs=17):
    """
    Process a single sheet dataframe and return averaged results with pass/fail logic.
    """
    output_columns = ["NO", "SN.", "Barcode", "Result"]
    for i in range(max_pairs):
        output_columns.append(f"B{i}_H_V")

    out_df = pd.DataFrame(columns=output_columns)

    for idx, row in df.iterrows():
        new_row = {}
        new_row["NO"] = idx + 1
        new_row["SN."] = row.get("SN.", row.get("SN", ""))
        new_row["Barcode"] = row.get("Barcode", "")
        new_row["Result"] = ""

        pass_flag = True  # assume pass until proven otherwise

        for i in range(max_pairs):
            v_col = f"B{i}_V"
            h_col = f"B{i}_H"
            avg_val = None

            if v_col in df.columns and h_col in df.columns:
                try:
                    avg_val = (float(row[v_col]) + float(row[h_col])) / 2
                except (TypeError, ValueError):
                    avg_val = None

            new_row[f"B{i}_H_V"] = avg_val

            # Check threshold for pass/fail
            limit = get_limit(i)
            if limit is not None and avg_val is not None:
                if avg_val < limit:
                    pass_flag = False

        new_row["Result"] = "PASS" if pass_flag else "FAIL"
        out_df.loc[len(out_df)] = new_row

    return out_df


def apply_styles_and_formatting(output_file, sheet_name, max_pairs=17):
    """
    Apply coloring, auto column widths (with regulated widths for processed columns),
    freeze top row, and add a summary block (fail count + percentage).
    """

    wb = load_workbook(output_file)
    ws = wb[sheet_name]

    # Define fill colors
    good_fill  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # green
    bad_fill   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # red
    accent2_20 = PatternFill(start_color="FDECEC", end_color="FDECEC", fill_type="solid") # accent 2 20%
    accent2_40 = PatternFill(start_color="F7CAC9", end_color="F7CAC9", fill_type="solid") # accent 2 40%
    accent3_40 = PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid") # accent 3 40%

    headers = [cell.value for cell in ws[1]]
    col_map = {header: idx + 1 for idx, header in enumerate(headers)}

    # Apply threshold-based coloring
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for i in range(max_pairs):
            col_name = f"B{i}_H_V"
            if col_name in col_map:
                cell = row[col_map[col_name] - 1]
                if cell.value is not None:
                    limit = get_limit(i)
                    if limit is not None:
                        if cell.value >= limit:
                            cell.fill = good_fill
                        else:
                            cell.fill = bad_fill

    # ✅ Regulated column widths
    for column_cells in ws.columns:
        header = column_cells[0].value
        column = get_column_letter(column_cells[0].column)

        if header and any(header.startswith(f"B{i}_H_V") for i in range(max_pairs)):
            ws.column_dimensions[column].width = 9
            continue

        max_length = 0
        for cell in column_cells:
            if cell.value:
                length = len(str(cell.value))
                max_length = max(max_length, length)
        ws.column_dimensions[column].width = min(max_length + 2, 35)

    # Freeze top row
    ws.freeze_panes = "A2"

    # ✅ Add fail count and percentage summary
    if sheet_name in ("AVG_MTF_FOCUS", "AVG_MTF_LCE"):
        result_col = col_map.get("Result")
        if result_col:
            results = [ws.cell(row=i, column=result_col).value for i in range(2, ws.max_row + 1)]
            total = len(results)
            fails = sum(1 for r in results if str(r).strip().lower() == "fail")
            fail_percent = (fails / total * 100) if total else 0

            # Place summary block 3 columns after last processed Bx_H_V column
            last_data_col = max(c for c in col_map.values())
            start_col = last_data_col + 3
            start_col_letter = get_column_letter(start_col)

            ws[f"{start_col_letter}2"] = "Summary"
            ws[f"{start_col_letter}2"].font = Font(bold=True, size=12)

            ws[f"{start_col_letter}3"] = "Fail Count"
            ws[f"{start_col_letter}4"] = "Total Samples"
            ws[f"{start_col_letter}5"] = "Fail %"

            ws[f"{get_column_letter(start_col + 1)}3"] = fails
            ws[f"{get_column_letter(start_col + 1)}3"].fill = accent2_40
            
            ws[f"{get_column_letter(start_col + 1)}4"] = total
            ws[f"{get_column_letter(start_col + 1)}4"].fill = accent3_40
            
            ws[f"{get_column_letter(start_col + 1)}5"] = f"{fail_percent:.2f}%"
            ws[f"{get_column_letter(start_col + 1)}5"].fill = accent2_20

            # Styling summary cells
            for row in range(2, 6):
                for col_offset in range(2):
                    cell = ws[f"{get_column_letter(start_col + col_offset)}{row}"]
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if row == 2:
                        cell.font = Font(bold=True)
                    ws.column_dimensions[get_column_letter(start_col + col_offset)].width = 15

    wb.save(output_file)
    

def process_excel(input_file, output_folder):
    """Main Excel processor — creates output file, copies sheets, applies formatting."""
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_file = os.path.join(output_folder, f"{base_name}_processed.xlsx")

    # Load all sheet names first
    xls = pd.ExcelFile(input_file)
    sheets = xls.sheet_names

    required_sheets = ["Test_MTF_FOCUS", "Test_MTF_LCE"]
    missing = [s for s in required_sheets if s not in sheets]
    if missing:
        print(f"❌ Missing required sheet(s): {', '.join(missing)}")
        sys.exit(1)

    print("📑 Found sheets:", ", ".join(sheets))

    # Process focus and LCE
    focus_df = pd.read_excel(input_file, sheet_name="Test_MTF_FOCUS")
    lce_df = pd.read_excel(input_file, sheet_name="Test_MTF_LCE")

    focus_out = process_sheet(focus_df)
    lce_out = process_sheet(lce_df)

    # Write output sheets
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        focus_out.to_excel(writer, sheet_name="AVG_MTF_FOCUS", index=False)
        lce_out.to_excel(writer, sheet_name="AVG_MTF_LCE", index=False)

        # Copy remaining sheets
        for sheet in sheets:
            if sheet not in required_sheets:
                print(f"📋 Copying sheet: {sheet}")
                df = pd.read_excel(input_file, sheet_name=sheet)
                df.to_excel(writer, sheet_name=sheet, index=False)

    # Apply formatting to all sheets
    wb = load_workbook(output_file)
    for sheet_name in wb.sheetnames:
        apply_styles_and_formatting(output_file, sheet_name)

    print(f"✅ Output written to: {output_file}")
    return output_file


# -----------------------------------
# CLI entry point
# -----------------------------------
if __name__ == "__main__":
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = input("Drop the path to a .xlsx file: ").strip().strip('"').strip("'")

    if not os.path.exists(input_file):
        print(f"❌ File not found: {input_file}")
        sys.exit(1)

    output_folder = get_output_folder("extracted")
    process_excel(input_file, output_folder)
