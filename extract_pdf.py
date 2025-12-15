import os
import pdfplumber
import pandas as pd
import re

pdf_folder = r"C:\Users\antonlangbruttig\Desktop\pdfs"    # change this to your file page where the file is stored if there is more than one put it all into one folder
output_excel = "C:\\Users\\antonlangbruttig\\Desktop\\combined_tasks_clean.xlsx" # change this to the file path of your system should just have to change "antonlangbruttig" to the name of your system 

all_data = []

# Pattern to match dates
date_pattern = re.compile(r"\d{1,2}/\d{1,2}/\d{2,4}")

for filename in os.listdir(pdf_folder):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, filename)
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Extract text with layout preservation to handle multi-line names
                text = page.extract_text(layout=True)
                if not text:
                    continue
                lines = text.split("\n")
                for i, line in enumerate(lines):
                    # Skip empty lines
                    if not line.strip():
                        continue

                    parts = line.split()
                    if len(parts) < 4:
                        continue

                    # Skip lines where first part is not a number
                    if not parts[0].isdigit():
                        continue

                    line_id = parts[0]

                    # Find the first two dates
                    dates = [p for p in parts if date_pattern.match(p)]
                    if len(dates) < 2:
                        # Check if next line might contain the rest of the data (for "double" lines)
                        if i + 1 < len(lines):
                            next_line = lines[i + 1].strip()
                            combined_line = line + " " + next_line
                            parts = combined_line.split()
                            dates = [p for p in parts if date_pattern.match(p)]
                            if len(dates) < 2:
                                continue
                        else:
                            continue
                    start, finish = dates[0], dates[1]

                    # Name = everything after Line ID up to first date
                    name_parts = parts[1:parts.index(dates[0])]
                    original_name = " ".join(name_parts)  # For debugging and fallback

                    # Extract number from name (assuming it's the first numeric part)
                    number = ""
                    if name_parts and name_parts[0].isdigit():
                        number = name_parts[0]
                        name_parts = name_parts[1:]

                    # Remove code like "123d" and anything after it
                    for j, part in enumerate(name_parts):
                        if re.match(r'^\d+[a-zA-Z]', part):  # Matches "123d" format
                            name_parts = name_parts[:j]
                            break

                    # Join remaining parts as Name, fallback to original if empty
                    name = " ".join(name_parts)
                    
                    all_data.append({
                        "Line": line_id,
                        "Number": number,
                        "Name": name,
                        "Start": start,
                        "Finish": finish
                    })

# Save to Excel
df = pd.DataFrame(all_data)
df.to_excel(output_excel, index=False)
print(f"Done! Clean data saved to {output_excel}")
