# PDF Value Extractor

A powerful Python system that extracts specific values from PDF files and exports results to Excel. Extract text that appears **after**, **before**, or **between** specified text markers, with comprehensive data type validation and professional Excel reporting.

## 🚀 New Features

- ✅ **Excel Export** - Automatically export all results to formatted Excel files
- ✅ **Enhanced Data Types** - New "both" option for text containing both letters and numbers
- ✅ **Comprehensive Configuration** - Centralized config system with predefined extraction rules
- ✅ **Interactive Menu** - User-friendly interface for different extraction modes
- ✅ **Professional Reporting** - Excel files with summary sheets, formatting, and statistics
- ✅ **Batch Processing** - Process multiple PDFs with multiple configurations simultaneously

## Features

- ✅ Extract text **after** a search string (e.g., "Total: $1,234.56")
- ✅ Extract text **before** a search string (e.g., "$1,234.56 Tax")  
- ✅ Extract text **between** two strings (e.g., "From: 2024-01-01 To: 2024-12-31")
- ✅ Data type validation (numbers, letters, both, or any)
- ✅ **Excel export with professional formatting**
- ✅ **Summary sheets with extraction statistics**
- ✅ Configurable extraction length and formatting
- ✅ Regular expression support for advanced patterns
- ✅ Batch processing of multiple PDF files
- ✅ **Centralized configuration system**

## Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the Interactive System
```bash
python run_extraction.py
```

Choose from the menu:
1. **Comprehensive Extraction** - Run all predefined configurations on all PDFs
2. **Targeted Extraction** - Select specific configurations to run
3. **Interactive Extraction** - Create custom extractions on-the-fly
4. **List Configurations** - View all available predefined configurations

### 3. View Results
- Results are automatically exported to Excel files with timestamps
- Excel files include formatted data, filters, and summary statistics
- Files are saved in the current directory with names like `comprehensive_extraction_20241031_113039.xlsx`

## Git Setup

### Quick Setup — if you've done this kind of thing before

Clone the repository:
```bash
git clone https://github.com/leandrosousa110490/PFD_Export.git
```

Get started by creating a new file or uploading an existing file. We recommend every repository include a README, LICENSE, and .gitignore.

### …or create a new repository on the command line

```bash
echo "# PFD_Export" >> README.md
git init
git add README.md
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/leandrosousa110490/PFD_Export.git
git push -u origin main
```

### …or push an existing repository from the command line

```bash
git remote add origin https://github.com/leandrosousa110490/PFD_Export.git
git branch -M main
git push -u origin main
```

## File Overview

| File | Purpose |
|------|---------|
| `pdf_value_extractor.py` | Main extraction engine with all the logic |
| `config.py` | **Centralized configuration** - all settings and predefined rules |
| `excel_exporter.py` | Excel export functionality with professional formatting |
| `run_extraction.py` | Interactive menu system and batch processing |
| `example_usage.py` | Detailed examples and demonstrations |
| `extraction_config.py` | Legacy configuration (still supported) |
| `requirements.txt` | Python dependencies |

## Configuration System

### Centralized Configuration (`config.py`)

The new centralized configuration system provides:
- **PDF Folder Settings** - Configure where to find PDF files
- **Excel Export Settings** - Control Excel output format and location
- **Predefined Extraction Rules** - 12 ready-to-use extraction configurations

### Excel Export Configuration

```python
# Excel output settings
EXCEL_OUTPUT_FOLDER = ""  # Empty = current directory
EXCEL_FILENAME_PREFIX = "comprehensive_extraction"
EXCEL_INCLUDE_TIMESTAMP = True  # Add timestamp to filename

# Excel formatting options
EXCEL_AUTO_ADJUST_COLUMNS = True  # Auto-adjust column widths
EXCEL_FREEZE_HEADER = True        # Freeze the header row
EXCEL_ADD_FILTERS = True          # Add filters to columns
```

### PDF File Configuration

```python
# Set this to any absolute path where your PDF files are located
# Examples:
#   PDF_FOLDER = r"C:\Users\YourName\Documents\PDFs"
#   PDF_FOLDER = r"D:\Invoice_PDFs"  
#   PDF_FOLDER = "folder"  # Relative to project directory
PDF_FOLDER = "folder"              # Change this to your PDF folder path
FALLBACK_TO_CURRENT_DIR = True     # Use current dir if folder not found
```

The system automatically:
- Accepts both **absolute paths** (e.g., `C:\Users\YourName\Documents\PDFs`) and **relative paths**
- Finds all `.pdf` files in the specified folder
- Falls back to the current directory if the folder doesn't exist
- Processes all found PDF files for extraction

### Excel Output Configuration

```python
# Set this to any absolute path where you want to save Excel files
# Examples:
#   EXCEL_OUTPUT_FOLDER = r"C:\Users\YourName\Documents\Excel_Reports"
#   EXCEL_OUTPUT_FOLDER = r"D:\PDF_Extraction_Results"
#   EXCEL_OUTPUT_FOLDER = ""  # Current directory
EXCEL_OUTPUT_FOLDER = ""           # Change this to your desired output path
EXCEL_FILENAME_PREFIX = "extraction_results"
EXCEL_INCLUDE_TIMESTAMP = True     # Add timestamp to filename
```

## Extraction Modes

The system supports four powerful extraction modes to handle different text extraction scenarios:

### 1. **After Mode** (`extraction_mode="after"`)
Extracts text that appears **after** a search string.

**Example:**
```
PDF Text: "Total Amount: $1,234.56"
Search: "Total Amount:"
Result: "$1,234.56"
```

**Configuration:**
```python
ExtractionConfig(
    name="Invoice Total",
    search_terms=["Total Amount:", "Total:", "Amount Due:"],
    extraction_mode="after",
    expected_type="numbers"
)
```

### 2. **Before Mode** (`extraction_mode="before"`)
Extracts text that appears **before** a search string.

**Example:**
```
PDF Text: "$1,234.56 Total Amount"
Search: "Total Amount"
Result: "$1,234.56"
```

### 3. **Between Mode** (`extraction_mode="between"`) ✨ **NEW!**
Extracts text that appears **between** two specific markers.

**Example:**
```
PDF Text: "Contract Period: From 2024-01-01 To 2024-12-31 Duration"
Search: "From"
End Text: "To"
Result: "2024-01-01"
```

**Configuration:**
```python
ExtractionConfig(
    name="Contract Start Date",
    search_terms=["From", "Start Date:"],
    end_text="To",
    extraction_mode="between",
    expected_type="any"
)
```

**Use Cases for Between Mode:**
- Extract date ranges: "From 2024-01-01 To 2024-12-31"
- Extract amounts: "Subtotal $100.00 Tax $10.00"
- Extract addresses: "Address: 123 Main St City: New York"
- Extract contract periods, warranty terms, etc.

### 4. **Around Mode** (`extraction_mode="around"`)
Extracts text in the vicinity of a search string (within a specified distance).

**Example:**
```
PDF Text: "Invoice #12345 dated 2024-01-15 for customer John Doe"
Search: "Invoice"
Result: "Invoice #12345 dated 2024-01-15"
```

### Data Type Validation

- **"numbers"** - Only accepts numeric values, currency symbols, decimals
- **"letters"** - Only accepts alphabetic characters and spaces  
- **"both"** - **NEW!** Accepts text containing both letters and numbers (e.g., "INV-2024")
- **"any"** - Accepts any text

## Predefined Configurations

The system includes 12 comprehensive extraction configurations ready for immediate use:

### Financial Extractions
- **`invoice_total`** - Extract total amounts from invoices
- **`tax_amount`** - Extract tax amounts  
- **`receipt_total`** - Extract receipt totals

### Document Identifiers
- **`invoice_number`** - Extract invoice numbers (supports formats like "INV-2024")
- **`document_id`** - Extract document IDs and reference numbers

### Personal Information
- **`customer_name`** - Extract customer names
- **`phone_number`** - Extract phone numbers
- **`email_address`** - Extract email addresses

### Business Information  
- **`store_name`** - Extract store or business names
- **`party_names`** - Extract party names from contracts

### Date Extractions
- **`due_date`** - Extract due dates
- **`contract_date`** - Extract contract dates

### View All Configurations
```python
from config import EXTRACTION_CONFIGS
for name, config in EXTRACTION_CONFIGS.items():
    print(f"{name}: {config.search_terms}")
```

## Excel Export Features

### Professional Formatting
- **Auto-adjusted column widths** for optimal readability
- **Frozen header row** for easy navigation
- **Column filters** for data analysis
- **Timestamp columns** for tracking extraction time

### Summary Sheet
Each Excel file includes a comprehensive summary sheet with:
- **Total extractions performed**
- **Success/failure counts**
- **Success rate percentage**
- **Unique PDF files processed**
- **Unique configurations used**
- **Detailed statistics per configuration**

### File Naming
Excel files are automatically named with timestamps:
```
comprehensive_extraction_20241031_113039.xlsx
targeted_extraction_20241031_114522.xlsx
interactive_extraction_20241031_115033.xlsx
```

## Usage Examples

### Interactive Menu System

Run the main script to access the interactive menu:

```bash
python run_extraction.py
```

**Menu Options:**
1. **Comprehensive Extraction** - Run all 12 predefined configurations
2. **Targeted Extraction** - Select specific configurations to run
3. **Interactive Extraction** - Create custom extraction rules on-the-fly
4. **List Configurations** - View all available predefined configurations
5. **Exit** - Close the application

### Example 1: Comprehensive Extraction

Extracts data using all predefined configurations and exports to Excel:

```
Select an option: 1

🔍 Starting comprehensive extraction...
📁 Found 8 PDF files in folder: folder

Processing configurations:
✓ invoice_total: 3/8 successful extractions
✓ tax_amount: 2/8 successful extractions  
✓ customer_name: 4/8 successful extractions
... (and 9 more configurations)

📊 Extraction Summary:
- Total extractions: 96
- Successful: 28 (29.2%)
- Failed: 68 (70.8%)

💾 Excel file created: comprehensive_extraction_20241031_113039.xlsx
```

### Example 2: Interactive Custom Extraction

Create custom extraction rules interactively:

```
Select an option: 3

Enter search text: Purchase Order
Select extraction mode:
1. After the search text
2. Before the search text  
3. Between two texts
Choice: 1

Select data type:
1. Numbers only
2. Letters only
3. Both letters and numbers
4. Any text
Choice: 3

Enter maximum characters to extract (default 50): 20
Case sensitive search? (y/n): n

✓ Custom configuration created successfully!
🔍 Running extraction on 8 PDF files...
```

### Example 3: Programmatic Usage

```python
from pdf_value_extractor import PDFValueExtractor, ExtractionConfig

config = ExtractionConfig(
    search_text="Total:",
    extraction_mode="after", 
    expected_type="numbers",
    max_chars=20
)

extractor = PDFValueExtractor(config)
result = extractor.extract_value("invoice.pdf")
print(result)  # Output: "$1,234.56"
```

### Example 4: Extract Text Between Two Markers

```python
config = ExtractionConfig(
    name="Date Range",
    search_terms=["From:", "Start Date:"],
    end_text="To:",
    extraction_mode="between",
    expected_type="any"
)
```

### Example 5: Extract Multiple Values from Multiple PDFs

```python
# In run_extraction.py, uncomment and modify:
configs_to_run = [
    "invoice_number",
    "invoice_total", 
    "customer_name"
]
extract_multiple_configs(configs_to_run)
```

## Advanced Features

### Batch Processing
- Process multiple PDF files simultaneously
- Automatic file discovery in specified folders
- Fallback to current directory if folder not found

### Data Validation
- **Smart type detection** with "both" option for mixed content
- **Configurable character limits** to prevent over-extraction
- **Whitespace cleaning** for consistent results
- **Case-sensitive/insensitive** search options

### Regular Expression Support
- Use regex patterns for complex extractions
- Combine with standard search modes
- Advanced pattern matching capabilities

### Professional Reporting
- **Detailed extraction logs** with success/failure tracking
- **Statistical summaries** with success rates
- **Per-configuration analytics** in Excel exports
- **Timestamp tracking** for audit trails

## Troubleshooting

### Common Issues

**Issue: "No PDF files found"**
- Check that PDF files exist in the specified folder
- Verify the `PDF_FOLDER` setting in `config.py`
- Ensure PDF files have `.pdf` extension
- **For custom paths**: Use absolute paths like `r"C:\Users\YourName\Documents\PDFs"`
- **For network paths**: Use UNC format like `r"\\server\share\PDFs"`
- Test your path by running `python example_custom_paths.py`

**Issue: "No values extracted"**
- Verify search text exists in the PDF
- Check if text is searchable (not scanned image)
- Try adjusting `MAX_CHARACTERS` setting
- Consider using case-insensitive search

**Issue: "Excel file not created"**
- Check write permissions in output directory
- Verify `EXCEL_OUTPUT_FOLDER` setting
- Ensure sufficient disk space

**Issue: "Wrong data type extracted"**
- Adjust `EXPECTED_TYPE` setting:
  - Use "numbers" for amounts
  - Use "letters" for names
  - Use "both" for mixed content (e.g., "INV-2024")
  - Use "any" for unrestricted extraction

### Performance Tips
- Use specific search terms to improve accuracy
- Limit `MAX_CHARACTERS` to reasonable values
- Enable `CLEAN_WHITESPACE` for cleaner results
- Use targeted extraction instead of comprehensive for faster processing

## Sample Output

### Console Output
```
🔍 Starting comprehensive extraction...
📁 Found 8 PDF files in folder: folder

📄 Processing: sample_invoice.pdf
  ✓ invoice_total: $1,234.56
  ✓ customer_name: John Smith
  ✗ tax_amount: Not found

📄 Processing: receipt.pdf  
  ✓ receipt_total: $45.99
  ✗ invoice_number: Not found

📊 Final Summary:
- Total extractions: 96
- Successful: 28 (29.2%)
- Failed: 68 (70.8%)
- Unique PDFs processed: 8
- Configurations used: 12

💾 Excel file: comprehensive_extraction_20241031_113039.xlsx
✨ Extraction completed successfully!
```

### Excel Output Structure
| PDF File | Configuration | Search Terms | Extracted Value | Data Type | Status | Timestamp |
|----------|---------------|--------------|-----------------|-----------|---------|-----------|
| invoice.pdf | invoice_total | Total:,Amount: | $1,234.56 | numbers | Success | 2024-10-31 11:30:39 |
| receipt.pdf | customer_name | Customer:,Name: | John Smith | letters | Success | 2024-10-31 11:30:40 |

## Requirements

- Python 3.7+
- PyPDF2 (for PDF text extraction)
- openpyxl (for Excel export)
- All dependencies listed in `requirements.txt`

## License

This project is open source and available under the MIT License.