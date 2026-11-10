import pdfplumber
import pandas as pd
import re
import os
import shutil
import sys

pdf_path = sys.argv[1] if len(sys.argv) > 1 else "EHC Benchmarking.pdf"
if not os.path.exists(pdf_path):
    raise FileNotFoundError(f"PDF not found: {pdf_path}")
print(f"Extracting from: {pdf_path}")

filename_lower = os.path.basename(pdf_path).lower()

if "ltd" in filename_lower:
    benefit_type = "ltd"
elif "dental" in filename_lower:
    benefit_type = "dental"
elif "std" in filename_lower:
    benefit_type = "std"
else:
    benefit_type = "ehc"

output_folder = f"C:/automation testing/{benefit_type}_output_files_1"

if os.path.exists(output_folder):
    shutil.rmtree(output_folder)
os.makedirs(output_folder, exist_ok=True)

with pdfplumber.open(pdf_path) as pdf:
    first_page_text = pdf.pages[0].extract_text().split("\n")
    company_name = "Unknown"
    lines = first_page_text

    for i, line in enumerate(lines):
        if line.startswith("Active as of:"):
            if i + 1 < len(lines):
                company_name = lines[i + 1].strip()
            break

    for line in first_page_text:
        if "Benefit:" in line:
            b_text = line.split("Benefit:")[-1].lower()
            if "disability" in b_text or "ltd" in b_text:
                benefit_type = "ltd"
            elif "dental" in b_text:
                benefit_type = "dental"
        if "Contract Number:" in line or "Plan InSite:" in line or "Active as of:" in line:
            continue

    if benefit_type not in ["ltd", "dental", "std", "ehc"]:
        for line in first_page_text:
            if "Benefit:" in line:
                b_text = line.split("Benefit:")[-1].lower()
                if "disability" in b_text or "ltd" in b_text:
                    benefit_type = "ltd"
                elif "dental" in b_text:
                    benefit_type = "dental"
                elif "std" in b_text:
                    benefit_type = "std"

    if not company_name.strip() or company_name.strip().lower() == "unknown":
        company_name = "Benchmark Client"

    meta_path = os.path.join(output_folder, "company_meta.txt")

    latest_folder_file = (r"C:\automation testing\latest_output_folder.txt")

    with open(latest_folder_file, "w") as folder_file:
        folder_file.write(output_folder)

    with open(meta_path, "w") as meta_f:
        meta_f.write(company_name)
    
    for pn in range(0, len(pdf.pages)):
        page = pdf.pages[pn]
        raw_text = page.extract_text()

        # Initialize metadata trackers
        table_rows = []
        is_inside_table = False        
        expense_type = ""
        attribute = ""

        # Loop line by line through the document text
        for line in raw_text.split("\n"):
            # Dynamically extract the metadata values from the text lines
            if "Expense Type:" in line:
                expense_type = line.split("Expense Type:")[-1].split("|")[0].replace(" / All", "").strip()
            if "Attribute:" in line:
                attribute = line.split("Attribute:")[-1].split("|")[0].strip()

            # Start capturing data once we hit the column headers
            if "Attribute Range Value" in line:
                is_inside_table = True
                continue
                
            # Stop capturing once we parse past the Totals row
            if "Totals" in line and is_inside_table:
                cleaned_total = re.sub(r'\s{2,}', '  ', line.strip())
                parts = cleaned_total.rsplit(maxsplit=3)
                table_rows.append(parts)
                break
                
            if is_inside_table:
                cleaned_line = line.strip()
                if not cleaned_line:
                    continue
                    
                cleaned_line = re.sub(r'\s{2,}', '  ', cleaned_line)
                parts = cleaned_line.rsplit(maxsplit=3)
                if len(parts) == 4:
                    table_rows.append(parts)

        # Convert array into a structured Pandas DataFrame dataset
        if table_rows:
            columns = ["Attribute Range Value", "Number of Contracts", "Number of Plans", "Percentage of Plans"]
            my_dataset = pd.DataFrame(table_rows, columns=columns)

            filename = f"{benefit_type}_clean_dataset_{pn+1}.csv"
            file_path = os.path.join(output_folder, filename)
            
            print(f"Dataset Successfully Defined for Page {pn+1}!\n")

            with open(file_path, "w") as f:
                # Write the metadata names as commented lines at the top
                f.write(f"# Expense Type: {expense_type}\n")
                f.write(f"# Attribute: {attribute}\n")
                # Append the dataframe seamlessly right underneath
                my_dataset.to_csv(f, index=False)
                
            print(f"Saved dataset to '{filename}' with embedded metadata headers.")
        else:
            print(f"Page {pn+1}: Failed to find or parse the table region.")