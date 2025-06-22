import pdfplumber
import re
import csv
import os
from datetime import datetime

# Ignore and allow lists - generalized for public code
IGNORE_WORDS = {"subtotal", "total", "our fee", "engineering plan applications", "gst", "sub total", "Project Management", "Expenses"}
IGNORED_FILES = {file.lower() for file in {"SupplierX Confidential.pdf"}}
ALLOW_DUPLICATES_SUPPLIERS = {"SupplierA", "SupplierB", "SupplierC"}

START_DATE = datetime(2022, 7, 1)
END_DATE = datetime(2024, 6, 30)

date_patterns = [
    r"\b\d{1,2}-\d{1,2}-\d{4}(?:[^\d]|$)",
    r"\b\d{1,2}/\d{1,2}/\d{4}(?:[^\d]|$)",
    r"\b\d{4}-\d{1,2}-\d{1,2}(?:[^\d]|$)",
    r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s\d{1,2},\s\d{4}(?:[^\d]|$)",
    r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s\d{1,2}\s\d{4}(?:[^\d]|$)",
    r"\b\d{1,2}\s(?:January|February|March|April|May|June|July|August|September|October|November|December)\s\d{4}(?:[^\d]|$)",
    r"\b\d{1,2}[A-Za-z]{3}\d{4}(?:[^\d]|$)",
]

date_formats = [
    "%d %b %Y", "%d %B %Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d",
    "%B %d, %Y", "%B %d %Y", "%d%b%Y"
]

def clean_date_string(date_str):
    date_str = re.sub(r"[^\w\s/-]", "", date_str)
    date_str = re.sub(r"(\d{4})[-A-Za-z]+$", r"\1", date_str)
    return date_str.strip()

def fix_missing_spaces(text):
    text = re.sub(r"([a-z])([A-Z])", r"\1 \2", text)
    text = re.sub(r"([a-zA-Z])(\d)", r"\1 \2", text)
    text = re.sub(r"(\d)([A-Za-z])", r"\1 \2", text)
    return text

def extract_dates(text):
    text = fix_missing_spaces(text)
    text = re.sub(r"[^\x20-\x7E]", "", text)
    text = re.sub(r"\s+", " ", text.strip())
    extracted_dates = []
    for pattern in date_patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            match = clean_date_string(match)
            parsed_date = None
            match = match.strip()
            for fmt in date_formats:
                try:
                    parsed_date = datetime.strptime(match, fmt)
                    break
                except ValueError:
                    continue
            if parsed_date and START_DATE <= parsed_date <= END_DATE:
                extracted_dates.append(parsed_date.strftime("%d/%m/%Y"))
    return extracted_dates

def extract_invoice_data(pdf_path, output_folder):
    extracted_data = []
    unique_records = set()
    supplier_name = os.path.splitext(os.path.basename(pdf_path))[0]
    allow_duplicates = any(supplier_name.startswith(supplier) for supplier in ALLOW_DUPLICATES_SUPPLIERS)

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            possible_dates = []
            text = page.extract_text()
            if text:
                possible_dates.extend(extract_dates(text))
                if not possible_dates:
                    continue
                invoice_date = min(possible_dates)
                lines = text.split("\n")
                prev_line_1 = ""
                prev_line_2 = ""
                i = 0
                while i < len(lines):
                    line = lines[i].strip()
                    # --- Insert general expense extraction logic here ---
                    current_values = re.findall(r"\b\d{1,3}(?:,\d{3})*\.\d{2}\b", line)
                    matched_word = next((kw for kw in IGNORE_WORDS if kw in line.lower()), None)
                    if matched_word and current_values:
                        expense_amount = current_values[-1]
                        unique_key = (supplier_name, expense_amount, matched_word.capitalize())
                        if allow_duplicates or unique_key not in unique_records:
                            unique_records.add(unique_key)
                            extracted_data.append([supplier_name, expense_amount, matched_word.capitalize(), invoice_date, line])
                    prev_line_2, prev_line_1 = prev_line_1, line
                    i += 1
    if not extracted_data:
        print(f"⚠️ No valid data found in {pdf_path}. Skipping file creation.")
        return
    os.makedirs(output_folder, exist_ok=True)
    output_csv = os.path.join(output_folder, f"{supplier_name}.csv")
    with open(output_csv, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["Supplier Name", "Expense Amount", "Matched Word", "Invoice Date", "Matched Line"])
        writer.writerows(extracted_data)
    print(f"✅ Data saved to {output_csv}")

def process_all_pdfs(input_folder, output_folder):
    for filename in os.listdir(input_folder):
        if not filename.endswith(".pdf"):
            continue
        if filename.lower() in IGNORED_FILES:
            print(f"⚠️ Skipping ignored file: {filename}")
            continue
        pdf_path = os.path.join(input_folder, filename)
        print(f"📄 Processing: {filename}")
        extract_invoice_data(pdf_path, output_folder)

# Example usage
input_folder = "Invoices"
output_folder = "Processed_Invoices"
process_all_pdfs(input_folder, output_folder)
