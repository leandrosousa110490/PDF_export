"""
Comprehensive Configuration File for PDF Value Extraction System
================================================================

This file contains all configuration settings for the PDF value extraction system,
including folder paths, Excel output settings, and extraction configurations.
"""

import os
from datetime import datetime

# ============================================================================
# FOLDER AND FILE PATHS
# ============================================================================

# PDF Files Configuration
# Set this to any absolute path where your PDF files are located
# Examples:
#   PDF_FOLDER = r"C:\Users\YourName\Documents\PDFs"
#   PDF_FOLDER = r"D:\Invoice_PDFs"
#   PDF_FOLDER = "folder"  # Relative to project directory
PDF_FOLDER = "folder"  # Change this to your PDF folder path
FALLBACK_TO_CURRENT_DIR = True  # If True, falls back to current directory if PDF_FOLDER doesn't exist

# Excel Output Configuration
# Set this to any absolute path where you want to save Excel files
# Examples:
#   EXCEL_OUTPUT_FOLDER = r"C:\Users\YourName\Documents\Excel_Reports"
#   EXCEL_OUTPUT_FOLDER = r"D:\PDF_Extraction_Results"
#   EXCEL_OUTPUT_FOLDER = ""  # Current directory
EXCEL_OUTPUT_FOLDER = ""  # Change this to your desired output path (empty = current directory)
EXCEL_FILENAME_PREFIX = "extraction_results"  # Prefix for Excel filenames
EXCEL_INCLUDE_TIMESTAMP = True  # Include timestamp in Excel filename

# ============================================================================
# EXCEL EXPORT SETTINGS
# ============================================================================

# Excel Column Names
EXCEL_COLUMNS = {
    'pdf_name': 'PDF Name',
    'config_name': 'Config Name',
    'extracted_value': 'Extracted Value',
    'extraction_status': 'Status',
    'timestamp': 'Extraction Time'
}

# Excel Formatting Options
EXCEL_AUTO_ADJUST_COLUMNS = True  # Auto-adjust column widths
EXCEL_FREEZE_HEADER = True  # Freeze the header row
EXCEL_ADD_FILTERS = True  # Add filters to columns

# ============================================================================
# EXTRACTION CONFIGURATIONS
# ============================================================================

class ExtractionConfig:
    """Configuration class for PDF value extraction"""
    
    def __init__(self, name, search_terms, expected_type="any", extraction_mode="after", 
                 max_distance=50, case_sensitive=False, end_text=None):
        self.name = name
        self.search_terms = search_terms
        self.expected_type = expected_type  # "any", "numbers", "letters", "both"
        self.extraction_mode = extraction_mode  # "after", "before", "between", "around"
        self.max_distance = max_distance
        self.case_sensitive = case_sensitive
        self.end_text = end_text  # Used for "between" mode

# ============================================================================
# PREDEFINED EXTRACTION CONFIGURATIONS
# ============================================================================

EXTRACTION_CONFIGS = {
    # Invoice and Financial Document Configurations
    "invoice_total": ExtractionConfig(
        name="Invoice Total",
        search_terms=["Total:", "Total Amount:", "Amount Due:", "Grand Total:"],
        expected_type="numbers",
        extraction_mode="after",
        max_distance=30
    ),
    
    "invoice_number": ExtractionConfig(
        name="Invoice Number",
        search_terms=["Invoice #:", "Invoice No:", "Invoice Number:", "Inv #:"],
        expected_type="both",
        extraction_mode="after",
        max_distance=20
    ),
    
    "customer_name": ExtractionConfig(
        name="Customer Name",
        search_terms=["Bill To:", "Customer:", "Client:", "Sold To:"],
        expected_type="letters",
        extraction_mode="after",
        max_distance=50
    ),
    
    "due_date": ExtractionConfig(
        name="Due Date",
        search_terms=["Due Date:", "Payment Due:", "Due By:"],
        expected_type="both",
        extraction_mode="after",
        max_distance=30
    ),
    
    "tax_amount": ExtractionConfig(
        name="Tax Amount",
        search_terms=["Tax:", "Sales Tax:", "VAT:", "GST:"],
        expected_type="numbers",
        extraction_mode="after",
        max_distance=25
    ),
    
    # Receipt Configurations
    "receipt_total": ExtractionConfig(
        name="Receipt Total",
        search_terms=["Total:", "Amount:", "Total Amount:"],
        expected_type="numbers",
        extraction_mode="after",
        max_distance=30
    ),
    
    "store_name": ExtractionConfig(
        name="Store Name",
        search_terms=["Store:", "Location:", "Branch:"],
        expected_type="letters",
        extraction_mode="after",
        max_distance=40
    ),
    
    # Contract and Legal Document Configurations
    "contract_date": ExtractionConfig(
        name="Contract Date",
        search_terms=["Date:", "Contract Date:", "Effective Date:"],
        expected_type="both",
        extraction_mode="after",
        max_distance=30
    ),
    
    "party_names": ExtractionConfig(
        name="Party Names",
        search_terms=["Party 1:", "Party 2:", "Between:", "Parties:"],
        expected_type="letters",
        extraction_mode="after",
        max_distance=60
    ),
    
    # General Document Configurations
    "document_id": ExtractionConfig(
        name="Document ID",
        search_terms=["Document ID:", "Doc ID:", "Reference:", "Ref:"],
        expected_type="both",
        extraction_mode="after",
        max_distance=25
    ),
    
    "phone_number": ExtractionConfig(
        name="Phone Number",
        search_terms=["Phone:", "Tel:", "Telephone:", "Contact:"],
        expected_type="both",
        extraction_mode="after",
        max_distance=30
    ),
    
    "email_address": ExtractionConfig(
        name="Email Address",
        search_terms=["Email:", "E-mail:", "Contact Email:"],
        expected_type="both",
        extraction_mode="after",
        max_distance=40
    ),
    
    # Between Text Configurations
    "date_range": ExtractionConfig(
        name="Date Range",
        search_terms=["From:", "Period:", "Between:"],
        expected_type="any",
        extraction_mode="between",
        end_text="To:",
        max_distance=50
    ),
    
    "contract_period": ExtractionConfig(
        name="Contract Period",
        search_terms=["Valid from:", "Effective from:", "Start date:"],
        expected_type="any",
        extraction_mode="between",
        end_text="until",
        max_distance=60
    ),
    
    "amount_breakdown": ExtractionConfig(
        name="Amount Breakdown",
        search_terms=["Subtotal:", "Net amount:"],
        expected_type="numbers",
        extraction_mode="between",
        end_text="Tax:",
        max_distance=40
    ),
    
    "address_details": ExtractionConfig(
        name="Address Details",
        search_terms=["Client Address:", "Blvd, Business"],
        expected_type="any",
        extraction_mode="between",
        end_text="Blvd, Business",
        max_distance=100
    )
}

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def get_pdf_folder_path():
    """Get the full path to the PDF folder"""
    return os.path.abspath(PDF_FOLDER)

def get_excel_output_path():
    """Get the full path to the Excel output folder"""
    output_path = os.path.abspath(EXCEL_OUTPUT_FOLDER)
    os.makedirs(output_path, exist_ok=True)
    return output_path

def generate_excel_filename():
    """Generate Excel filename with optional timestamp"""
    if EXCEL_INCLUDE_TIMESTAMP:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{EXCEL_FILENAME_PREFIX}_{timestamp}.xlsx"
    else:
        return f"{EXCEL_FILENAME_PREFIX}.xlsx"

def get_full_excel_path():
    """Get the full path for the Excel output file"""
    return os.path.join(get_excel_output_path(), generate_excel_filename())

def get_pdf_files():
    """Get list of PDF files from the configured folder"""
    pdf_files = []
    
    # Try to get files from the specified PDF folder
    if os.path.exists(PDF_FOLDER) and os.path.isdir(PDF_FOLDER):
        for file in os.listdir(PDF_FOLDER):
            if file.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(PDF_FOLDER, file))
    elif FALLBACK_TO_CURRENT_DIR:
        # Fallback to current directory
        for file in os.listdir('.'):
            if file.lower().endswith('.pdf'):
                pdf_files.append(file)
    
    return pdf_files

# ============================================================================
# CONFIGURATION VALIDATION
# ============================================================================

def validate_config():
    """Validate the configuration settings"""
    errors = []
    
    # Check if PDF folder exists or fallback is enabled
    if not os.path.exists(PDF_FOLDER) and not FALLBACK_TO_CURRENT_DIR:
        errors.append(f"PDF folder '{PDF_FOLDER}' does not exist and fallback is disabled")
    
    # Validate extraction configurations
    for config_name, config in EXTRACTION_CONFIGS.items():
        if not config.search_terms:
            errors.append(f"Configuration '{config_name}' has no search terms")
        
        if config.expected_type not in ["any", "numbers", "letters", "both"]:
            errors.append(f"Configuration '{config_name}' has invalid expected_type: {config.expected_type}")
        
        if config.extraction_mode not in ["after", "before", "between", "around"]:
            errors.append(f"Configuration '{config_name}' has invalid extraction_mode: {config.extraction_mode}")
    
    return errors

# ============================================================================
# CONFIGURATION INFO
# ============================================================================

def print_config_info():
    """Print configuration information"""
    print("PDF Value Extraction System Configuration")
    print("=" * 50)
    print(f"PDF Folder: {get_pdf_folder_path()}")
    print(f"Excel Output Folder: {get_excel_output_path()}")
    print(f"Excel Filename: {generate_excel_filename()}")
    print(f"Available Configurations: {len(EXTRACTION_CONFIGS)}")
    print(f"PDF Files Found: {len(get_pdf_files())}")
    
    # Validate configuration
    errors = validate_config()
    if errors:
        print("\nConfiguration Errors:")
        for error in errors:
            print(f"  - {error}")
    else:
        print("\nConfiguration is valid!")

if __name__ == "__main__":
    print_config_info()
