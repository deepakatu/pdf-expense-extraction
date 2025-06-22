import os
import csv
import re
from collections import defaultdict

def get_base_filename(filename):
    base_name = re.sub(r"[_ ]\d+(?=\.csv$)", "", filename)
    return base_name

def merge_csv_files(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    file_groups = defaultdict(list)
    for filename in os.listdir(input_folder):
        if filename.endswith(".csv"):
            base_name = get_base_filename(filename)
            file_groups[base_name].append(filename)
    for base_name, file_list in file_groups.items():
        merged_output_file = os.path.join(output_folder, f"{base_name}")
        first_file = True
        with open(merged_output_file, mode='w', newline='', encoding='utf-8') as outfile:
            writer = None
            for file in sorted(file_list):
                input_file_path = os.path.join(input_folder, file)
                with open(input_file_path, mode='r', newline='', encoding='utf-8') as infile:
                    reader = csv.reader(infile)
                    header = next(reader)
                    if first_file:
                        writer = csv.writer(outfile)
                        writer.writerow(header)
                        first_file = False
                    for row in reader:
                        writer.writerow(row)
        print(f"✅ Merged {len(file_list)} files into: {merged_output_file}")

input_folder = "Processed_Invoices/FY2324 csvs"
output_folder = os.path.join(input_folder, "Merged")
merge_csv_files(input_folder, output_folder)
