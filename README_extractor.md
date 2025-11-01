# Text to Excel Extractor

A powerful Python tool that extracts specific values from text files using configurable before/after text patterns and exports the results to Excel format.

## Features

- 📄 **Configurable Text Extraction**: Define custom patterns to extract specific values
- 🔧 **JSON Configuration**: Easy-to-modify configuration file for extraction patterns
- 📊 **Excel Export**: Automatically exports results to Excel with proper formatting
- 📁 **Batch Processing**: Processes multiple text files at once
- 🎯 **Flexible Patterns**: Extract text between "before" and "after" patterns
- ⚙️ **Customizable Paths**: Configure input/output folders via variables
- 🛡️ **Error Handling**: Robust error handling and user feedback

## Installation

1. **Install Python dependencies:**
   ```bash
   pip install pandas openpyxl
   ```

2. **Or use the requirements file:**
   ```bash
   pip install -r requirements_extractor.txt
   ```

## Configuration

### Path Variables (in `text_to_excel_extractor.py`)

Edit these variables at the top of the script:

```python
# Path to the folder containing extracted text files
TEXT_FILES_FOLDER_PATH = r"C:\Users\nbaba\Desktop\PDF to Excel\extracted_texts"

# Path to the configuration JSON file
CONFIG_FILE_PATH = r"C:\Users\nbaba\Desktop\PDF to Excel\extraction_config.json"

# Path to the folder where Excel file will be saved
EXPORT_FOLDER_PATH = r"C:\Users\nbaba\Desktop\PDF to Excel\exports"

# Name of the Excel file to be created
EXCEL_FILE_NAME = "extracted_data.xlsx"
```

### Extraction Configuration (`extraction_config.json`)

The configuration file defines what values to extract from your text files:

```json
{
  "configurations": [
    {
      "name": "Invoice_Number",
      "description": "Extract invoice number from invoices",
      "before_text": "Invoice Number:",
      "after_text": "Date:",
      "case_sensitive": false
    },
    {
      "name": "Total_Amount",
      "description": "Extract total amount from invoices",
      "before_text": "Total Amount:",
      "after_text": "$",
      "case_sensitive": false
    }
  ],
  "settings": {
    "max_extraction_length": 100,
    "trim_whitespace": true,
    "remove_special_chars": false,
    "default_value_if_not_found": "NOT_FOUND"
  }
}
```

#### Configuration Parameters:

- **`name`**: Unique identifier for the extraction rule
- **`description`**: Human-readable description
- **`before_text`**: Text that appears before the value you want to extract
- **`after_text`**: Text that appears after the value (optional)
- **`case_sensitive`**: Whether the search should be case-sensitive

### OR Statement Functionality

The extractor now supports **OR statements** for more flexible pattern matching. Instead of a single pattern, you can define multiple alternative patterns for each extraction rule.

#### Single Pattern (Legacy Format):
```json
{
  "name": "Invoice_Number",
  "description": "Extract invoice number",
  "before_text": "Invoice Number:",
  "after_text": "Date:",
  "case_sensitive": false
}
```

#### Multiple Patterns (OR Statement Format):
```json
{
  "name": "Invoice_Number",
  "description": "Extract invoice number with multiple pattern options",
  "patterns": [
    {
      "before_text": "Invoice Number:",
      "after_text": "Date:",
      "case_sensitive": false
    },
    {
      "before_text": "Invoice #:",
      "after_text": "Date:",
      "case_sensitive": false
    },
    {
      "before_text": "INV-",
      "after_text": " ",
      "case_sensitive": false
    }
  ]
}
```

#### How OR Statements Work:

1. **Sequential Pattern Testing**: The extractor tries each pattern in order
2. **First Match Wins**: Returns the value from the first successful pattern
3. **Fallback Support**: If Pattern 1 fails, try Pattern 2, then Pattern 3, etc.
4. **Backward Compatibility**: Single pattern configurations still work perfectly

#### Example Use Cases:

- **Invoice Numbers**: Try "Invoice Number:", then "Invoice #:", then "INV-"
- **Total Amounts**: Try "Total Amount:", then "Total:", then "Amount Due:"
- **Company Names**: Try "Company:", then "From:", then "Bill To:"
- **Dates**: Try "Date:", then "Invoice Date:", then "Issued:"

This flexibility ensures higher extraction success rates across different document formats!

#### Settings:

- **`max_extraction_length`**: Maximum characters to extract
- **`trim_whitespace`**: Remove leading/trailing spaces
- **`remove_special_chars`**: Remove special characters from extracted values
- **`default_value_if_not_found`**: Value to use when pattern is not found

## Usage

1. **Configure your paths** in `text_to_excel_extractor.py`
2. **Set up your extraction patterns** in `extraction_config.json`
3. **Run the script:**
   ```bash
   python text_to_excel_extractor.py
   ```

## Output Format

The Excel file contains three columns:

| Column | Description |
|--------|-------------|
| **filename** | Name of the source text file |
| **config_name** | Name of the extraction configuration used |
| **extracted_value** | The extracted value or "NOT_FOUND" |

### Example Output:

| filename | config_name | extracted_value |
|----------|-------------|-----------------|
| test_invoice_1.txt | Invoice_Number | INV-2024010 |
| test_invoice_1.txt | Total_Amount | $4631.04 |
| test_invoice_2.txt | Invoice_Number | INV-2024011 |
| test_invoice_2.txt | Total_Amount | $2150.00 |

## Example Text Content

If your text file contains:
```
INVOICE Company: Innovation Partners Invoice Number: INV-2024010 Date: 10/11/2025 Total Amount: $4631.04
```

With this configuration:
```json
{
  "name": "Invoice_Number",
  "before_text": "Invoice Number:",
  "after_text": "Date:",
  "case_sensitive": false
}
```

The extracted value would be: `INV-2024010`

## Adding New Extraction Patterns

To add a new extraction pattern, edit `extraction_config.json`:

```json
{
  "name": "Customer_Name",
  "description": "Extract customer name from invoices",
  "before_text": "Customer Name:",
  "after_text": "Description",
  "case_sensitive": false
}
```

## Troubleshooting

### Common Issues:

1. **"No configurations found"**
   - Check that `extraction_config.json` exists and is valid JSON
   - Verify the `CONFIG_FILE_PATH` variable

2. **"No text files found"**
   - Check that the `TEXT_FILES_FOLDER_PATH` contains .txt files
   - Verify the folder path is correct

3. **"NOT_FOUND" values**
   - Check that your `before_text` and `after_text` patterns match the content
   - Try making the search case-insensitive
   - Verify the text content format

4. **Excel export fails**
   - Ensure the export folder exists or can be created
   - Check file permissions
   - Make sure pandas and openpyxl are installed

## Files Created

- `text_to_excel_extractor.py` - Main extraction script
- `extraction_config.json` - Configuration file for extraction patterns
- `requirements_extractor.txt` - Python dependencies
- `README_extractor.md` - This documentation file
- `exports/extracted_data.xlsx` - Generated Excel file with results