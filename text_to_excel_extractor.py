#!/usr/bin/env python3
"""
Text to Excel Extractor
Extracts specific values from text files using configurable patterns and exports to Excel.
"""

import os
import json
import re
import pandas as pd
from pathlib import Path
from typing import Dict, List, Any, Optional

# ================================
# CONFIGURATION VARIABLES
# ================================

# Path to the folder containing extracted text files
TEXT_FILES_FOLDER_PATH = r"C:\Users\nbaba\Desktop\PDF to Excel\extracted_texts"

# Path to the configuration JSON file
CONFIG_FILE_PATH = r"C:\Users\nbaba\Desktop\PDF to Excel\extraction_config.json"

# Path to the folder where Excel file will be saved
EXPORT_FOLDER_PATH = r"C:\Users\nbaba\Desktop\PDF to Excel\exports"

# Name of the Excel file to be created
EXCEL_FILE_NAME = "extracted_data.xlsx"

# ================================
# MAIN CLASSES AND FUNCTIONS
# ================================

class TextExtractor:
    """Handles text extraction using before/after patterns."""
    
    def __init__(self, config_path: str):
        """Initialize with configuration file."""
        self.config = self.load_config(config_path)
        self.settings = self.config.get('settings', {})
    
    def load_config(self, config_path: str) -> Dict[str, Any]:
        """Load configuration from JSON file."""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"❌ Configuration file not found: {config_path}")
            return {"configurations": [], "settings": {}}
        except json.JSONDecodeError as e:
            print(f"❌ Error parsing configuration file: {e}")
            return {"configurations": [], "settings": {}}
    
    def extract_value_between_text(self, text: str, before_text: str, after_text: str, 
                                 case_sensitive: bool = False) -> Optional[str]:
        """Extract text between before_text and after_text."""
        if not text or not before_text:
            return None
        
        # Prepare text for searching
        search_text = text if case_sensitive else text.lower()
        before_pattern = before_text if case_sensitive else before_text.lower()
        after_pattern = after_text if case_sensitive else after_text.lower()
        
        # Find the start position after before_text
        start_pos = search_text.find(before_pattern)
        if start_pos == -1:
            return None
        
        start_pos += len(before_pattern)
        
        # If no after_text specified, extract to end or max length
        if not after_text:
            extracted = text[start_pos:start_pos + self.settings.get('max_extraction_length', 100)]
        else:
            # Find the end position before after_text
            end_pos = search_text.find(after_pattern, start_pos)
            if end_pos == -1:
                extracted = text[start_pos:start_pos + self.settings.get('max_extraction_length', 100)]
            else:
                extracted = text[start_pos:end_pos]
        
        # Clean up the extracted text
        if self.settings.get('trim_whitespace', True):
            extracted = extracted.strip()
        
        if self.settings.get('remove_special_chars', False):
            extracted = re.sub(r'[^\w\s.-]', '', extracted)
        
        return extracted if extracted else None
    
    def extract_from_text_file(self, file_path: str) -> Dict[str, str]:
        """Extract all configured values from a single text file."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                text_content = f.read()
        except Exception as e:
            print(f"❌ Error reading file {file_path}: {e}")
            return {}
        
        results = {}
        configurations = self.config.get('configurations', [])
        
        for config in configurations:
            config_name = config.get('name', 'Unknown')
            before_text = config.get('before_text', '')
            after_text = config.get('after_text', '')
            case_sensitive = config.get('case_sensitive', False)
            
            extracted_value = self.extract_value_between_text(
                text_content, before_text, after_text, case_sensitive
            )
            
            if extracted_value is None:
                extracted_value = self.settings.get('default_value_if_not_found', 'NOT_FOUND')
            
            results[config_name] = extracted_value
        
        return results


class ExcelExporter:
    """Handles Excel file creation and data export."""
    
    def __init__(self, export_folder: str, excel_filename: str):
        """Initialize with export settings."""
        self.export_folder = Path(export_folder)
        self.excel_filename = excel_filename
        self.excel_path = self.export_folder / excel_filename
        
        # Create export folder if it doesn't exist
        self.export_folder.mkdir(parents=True, exist_ok=True)
    
    def export_to_excel(self, extraction_results: List[Dict[str, Any]]) -> bool:
        """Export extraction results to Excel file."""
        if not extraction_results:
            print("❌ No data to export")
            return False
        
        try:
            # Create DataFrame
            df = pd.DataFrame(extraction_results)
            
            # Reorder columns: filename first, then config names, then values
            columns = ['filename'] + [col for col in df.columns if col != 'filename']
            df = df[columns]
            
            # Export to Excel
            df.to_excel(self.excel_path, index=False, engine='openpyxl')
            print(f"✅ Data exported successfully to: {self.excel_path}")
            return True
            
        except Exception as e:
            print(f"❌ Error exporting to Excel: {e}")
            return False


def process_text_files(text_folder: str, extractor: TextExtractor) -> List[Dict[str, Any]]:
    """Process all text files in the specified folder."""
    text_folder_path = Path(text_folder)
    
    if not text_folder_path.exists():
        print(f"❌ Text files folder not found: {text_folder}")
        return []
    
    results = []
    text_files = list(text_folder_path.glob("*.txt"))
    
    if not text_files:
        print(f"❌ No text files found in: {text_folder}")
        return []
    
    print(f"📁 Processing {len(text_files)} text files...")
    
    for text_file in text_files:
        print(f"📄 Processing: {text_file.name}")
        
        # Extract values from this file
        extracted_values = extractor.extract_from_text_file(str(text_file))
        
        # Create a row for each configuration
        for config_name, extracted_value in extracted_values.items():
            result_row = {
                'filename': text_file.stem,  # Use .stem to get filename without extension
                'config_name': config_name,
                'extracted_value': extracted_value
            }
            results.append(result_row)
    
    return results


def main():
    """Main function to run the text extraction and Excel export process."""
    print("🚀 Starting Text to Excel Extractor...")
    print("=" * 50)
    
    # Display current configuration
    print("📋 Current Configuration:")
    print(f"   Text Files Folder: {TEXT_FILES_FOLDER_PATH}")
    print(f"   Config File: {CONFIG_FILE_PATH}")
    print(f"   Export Folder: {EXPORT_FOLDER_PATH}")
    print(f"   Excel File: {EXCEL_FILE_NAME}")
    print("=" * 50)
    
    # Initialize extractor
    extractor = TextExtractor(CONFIG_FILE_PATH)
    
    if not extractor.config.get('configurations'):
        print("❌ No configurations found. Please check your config file.")
        return
    
    print(f"✅ Loaded {len(extractor.config['configurations'])} extraction configurations")
    
    # Process text files
    extraction_results = process_text_files(TEXT_FILES_FOLDER_PATH, extractor)
    
    if not extraction_results:
        print("❌ No data extracted. Please check your text files and configurations.")
        return
    
    print(f"✅ Extracted {len(extraction_results)} data points")
    
    # Export to Excel
    exporter = ExcelExporter(EXPORT_FOLDER_PATH, EXCEL_FILE_NAME)
    success = exporter.export_to_excel(extraction_results)
    
    if success:
        print("=" * 50)
        print("🎉 Process completed successfully!")
        print(f"📊 Excel file created: {exporter.excel_path}")
    else:
        print("❌ Export failed. Please check the error messages above.")
    
    print("\nPress Enter to exit...")
    input()


if __name__ == "__main__":
    main()