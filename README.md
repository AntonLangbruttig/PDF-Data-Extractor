# PDF Data Extractor

An automated batch processing tool that extracts structured data from PDF documents and outputs clean Excel spreadsheets. Built for a client to eliminate manual data entry.

## Features

- **Batch Processing** — Process multiple PDF files from a single folder
- **Layout-Aware Extraction** — Preserves text positioning using pdfplumber
- **Regex Pattern Matching** — Intelligently identifies and parses date formats
- **Multi-Line Handling** — Combines split data when entries span multiple lines
- **Data Cleaning** — Normalizes extracted data, removes codes, formats names
- **Excel Output** — Exports structured data to formatted spreadsheet with pandas
- **99% Time Reduction** — Compared to manual data entry

## Tech Stack

- **Language:** Python
- **PDF Processing:** pdfplumber
- **Data Handling:** pandas
- **Pattern Matching:** Regular Expressions (re)
- **Excel Export:** OpenPyXL (via pandas)

## How It Works

1. Reads all PDF files from a specified folder
2. Extracts text with layout preservation
3. Parses each line to identify structured data (ID, Name, Dates)
4. Handles edge cases like multi-line entries
5. Cleans and normalizes the data
6. Outputs everything to a formatted Excel file

## Output Format

| Line | Number | Name | Start | Finish |
|------|--------|------|-------|--------|
| 1 | 123 | John Smith | 1/15/24 | 2/20/24 |
| 2 | 456 | Jane Doe | 1/18/24 | 3/01/24 |

## Installation

```bash
pip install pdfplumber pandas openpyxl
```

## Usage

1. Update the `pdf_folder` path to your PDF directory
2. Update the `output_excel` path for the output file
3. Run the script:

```bash
python extract_pdfs.py
```

## Configuration

```python
# Set your input folder (containing PDFs)
pdf_folder = r"C:\Users\yourname\Desktop\pdfs"

# Set your output file path
output_excel = "C:\\Users\\yourname\\Desktop\\combined_tasks_clean.xlsx"
```

## Project Structure

```
PDF-Data-Extractor/
├── extract_pdfs.py    # Main extraction script
└── README.md
```

## Use Case

This tool was built for a client who needed to extract task data from hundreds of PDF documents. What previously took hours of manual copy-paste now runs in seconds with consistent, error-free results.

## License

MIT
