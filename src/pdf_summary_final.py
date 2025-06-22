import csv
import os
import re

def clean_amount(amount, filename, row):
    if amount:
        amount = amount.replace(",", "").strip()
        try:
            return float(amount)
        except ValueError:
            error_message = f"❌ ValueError: Could not convert '{amount}' in file '{filename}' (Row: {row})\n"
            print(error_message)
            with open("Processed_Invoices/error_log.txt", mode="a", encoding="utf-8") as error_log:
                error_log.write(error_message)
            return 0.0
    return 0.0

def process_supplier_files(input_folder, output_file):
    summary_data = []
    error_log_path = os.path.join(input_folder, "error_log.txt")
    if os.path.exists(error_log_path):
        os.remove(error_log_path)
    for filename in os.listdir(input_folder):
        if filename.endswith(".csv") and filename != "error_log.txt":
            file_path = os.path.join(input_folder, filename)
            supplier_name = os.path.splitext(filename)[0]
            total_expense = 0.0
            with open(file_path, mode='r', newline='', encoding='utf-8') as infile:
                reader = csv.reader(infile)
                header = next(reader)
                for row in reader:
                    try:
                        expense_amount = clean_amount(row[1], filename, row)
                        total_expense += expense_amount
                    except IndexError:
                        error_message = f"⚠️ IndexError: Skipping malformed row in {filename}: {row}\n"
                        print(error_message)
                        with open(error_log_path, mode="a", encoding="utf-8") as error_log:
                            error_log.write(error_message)
            summary_data.append([supplier_name, round(total_expense, 2)])
    with open(output_file, mode='w', newline='', encoding='utf-8') as outfile:
        writer = csv.writer(outfile)
        writer.writerow(["Supplier Name", "Expense Amount"])
        writer.writerows(summary_data)
    print(f"✅ Final summary saved to: {output_file}")

input_folder = "Processed_Invoices/FY2324 csvs/Merged"
output_file = "Processed_Invoices/FY2324 csvs/Merged/FY2324_final_summary.csv"
process_supplier_files(input_folder, output_file)
