import os
import pdfplumber

def count_pages_and_size(folder_path):
    total_pages = 0
    total_size_bytes = 0
    file_count = 0
    for filename in os.listdir(folder_path):
        file_count += 1
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)
            try:
                with pdfplumber.open(file_path) as pdf:
                    num_pages = len(pdf.pages)
                    print(f"{filename}: {num_pages} pages")
                    total_pages += num_pages
                file_size = os.path.getsize(file_path)
                total_size_bytes += file_size
            except Exception as e:
                print(f"❌ Failed to process {filename}: {e}")

    total_size_gb = total_size_bytes / (1024 ** 3)
    print(f"\n📄 Total number of pages: {total_pages}")
    print(f"💾 Total PDF file size: {total_size_gb:.2f} GB")
    print(f"📁 Total PDF files: {file_count}")
    return total_pages, total_size_gb, file_count

# Example usage
folder_path = "Invoices"  # Replace with your folder path
count_pages_and_size(folder_path)
