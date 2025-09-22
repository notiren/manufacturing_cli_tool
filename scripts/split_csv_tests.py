import os
import sys

# Helper for dynamic output

def get_output_folder(folder_name="extracted"):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.basename(script_dir).lower() == "scripts":
        base_dir = os.path.dirname(script_dir)
    else:
        base_dir = script_dir
    output_folder = os.path.join(base_dir, folder_name)
    os.makedirs(output_folder, exist_ok=True)
    return output_folder

# Split CSV by preserving format

def split_csv_preserve_format(input_file, output_dir, rows_per_split):
    with open(input_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    # extract the first 3 lines + header
    empty_lines = lines[:3]
    header_line = lines[3]
    data_lines = lines[4:]

    # create or ensure output folder path exists
    os.makedirs(output_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_prefix = os.path.join(output_dir, base_name)

    for i in range(0, len(data_lines), rows_per_split):
        part = data_lines[i:i+rows_per_split]
        part_num = i // rows_per_split + 1
        with open(f"{output_prefix}_part{part_num}.csv", 'w', encoding='utf-8') as f_out:
            f_out.writelines(empty_lines)
            f_out.write(header_line)
            f_out.writelines(part)
    print(f"-----------")
    print(f"CSV was split into {part_num} parts. Files located inside folder: '{output_dir}'\n")

# ---- Main ----

def main():
    csv_file = input("Drop the path to a .csv file: ").strip().strip('"').strip("'")
    if not csv_file:
        print("No file path provided. Exiting.\n")
        sys.exit(1)
        
    if not os.path.isabs(csv_file):
        csv_file = os.path.abspath(os.path.join(os.getcwd(), csv_file))

    if not os.path.isfile(csv_file):
        print(f"File does not exist: {csv_file}\n")
        sys.exit(1)

    if not csv_file.lower().endswith('.csv'):
        print(f"The file is not a CSV: {csv_file}\n")
        sys.exit(1)
        
    # prompt user for number of items        
    split_number = int(input("Please enter number of items per split: "))
    output_folder = get_output_folder("extracted")
    split_csv_preserve_format(csv_file, output_folder, split_number)

if __name__ == "__main__":
    main()
