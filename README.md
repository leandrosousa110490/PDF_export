# PDF Text Extractor

A Python script that extracts text from all PDF files in a specified folder and saves the extracted text to separate text files. The script handles multi-page PDFs and formats all extracted text as one continuous line.

## Features

- **Batch Processing**: Processes all PDF files in a folder automatically
- **Multi-page Support**: Extracts text from PDFs with multiple pages
- **Single-line Format**: All extracted text is formatted as one continuous line (no line breaks)
- **Configurable Paths**: Easy-to-modify path variables at the top of the script
- **Matching File Names**: Creates `.txt` files with the same name as the original PDFs
- **Error Handling**: Gracefully handles corrupted or encrypted PDFs
- **Progress Feedback**: Shows real-time processing status
- **Flexible Output**: Can save to a custom folder or same folder as PDFs

## Requirements

- Python 3.6 or higher
- PyPDF2 library

## Installation

1. **Clone or download** this repository to your local machine

2. **Install the required dependency**:
   ```bash
   pip install PyPDF2
   ```
   
   Or install from the requirements file:
   ```bash
   pip install -r requirements.txt
   ```

## Configuration

Before running the script, update the path variables at the top of `pdf_text_extractor.py`:

```python
# Path to the folder containing PDF files
PDF_FOLDER_PATH = r"C:\Users\nbaba\Desktop\PDF to Excel\folder"

# Path to the folder where extracted text files will be saved
# Set to None to save in the same folder as PDFs
OUTPUT_FOLDER_PATH = r"C:\Users\nbaba\Desktop\PDF to Excel\extracted_texts"
```

## Usage

1. **Update the path variables** in the script as described above

2. **Run the script**:
   ```bash
   python pdf_text_extractor.py
   ```

3. **The script will**:
   - Use the configured PDF folder path
   - Extract text from each PDF (including all pages)
   - Format all text as one continuous line
   - Save the extracted text to `.txt` files in the configured output folder
   - Show progress and results

## Example

If you have these PDF files in your configured PDF folder:
```
/my-pdfs/
├── invoice_001.pdf
├── contract_2023.pdf
└── receipt_march.pdf
```

After running the script with `OUTPUT_FOLDER_PATH = "/my-texts/"`, you'll get:

**Input folder (`/my-pdfs/`):**
```
/my-pdfs/
├── invoice_001.pdf
├── contract_2023.pdf
└── receipt_march.pdf
```

**Output folder (`/my-texts/`):**
```
/my-texts/
├── invoice_001.txt          ← New text file (single line format)
├── contract_2023.txt        ← New text file (single line format)
└── receipt_march.txt        ← New text file (single line format)
```

**Text Format Example:**
Instead of multi-line text like:
```
INVOICE
Company: ABC Corp
Invoice Number: 12345
Date: 2024-01-01
```

You'll get single-line text like:
```
INVOICE Company: ABC Corp Invoice Number: 12345 Date: 2024-01-01
```

## Output Format

Each text file includes:
- Header with the source PDF filename
- Page separators for multi-page documents
- All extracted text content

Example output:
```
Text extracted from: invoice_001.pdf
==================================================

--- Page 1 ---

[Extracted text from page 1]

--- Page 2 ---

[Extracted text from page 2]
```

## Troubleshooting

**"No module named 'PyPDF2'" error:**
- Install the dependency: `pip install PyPDF2`

**"No PDF files found" message:**
- Check that the folder path is correct
- Ensure PDF files have `.pdf` extension
- Verify you have read permissions for the folder

**Empty or garbled text output:**
- Some PDFs contain images or scanned text that can't be extracted
- Try using OCR tools for image-based PDFs

## Requirements

- Python 3.6+
- PyPDF2 library
- Read/write permissions for the target folder

## License

This script is provided as-is for educational and personal use.