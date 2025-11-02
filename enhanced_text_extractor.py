#!/usr/bin/env python3
"""
Enhanced Text to Excel Extractor
Advanced text extraction with fuzzy matching, regex support, and improved accuracy.
"""

import os
import json
import re
import pandas as pd
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from difflib import SequenceMatcher
import logging

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
EXCEL_FILE_NAME = "extracted_data_enhanced.xlsx"

# ================================
# ENHANCED EXTRACTOR CLASS
# ================================

class EnhancedTextExtractor:
    """Enhanced text extraction with fuzzy matching and advanced features."""
    
    def __init__(self, config_path: str):
        """Initialize with configuration file."""
        self.config = self.load_config(config_path)
        self.settings = self.config.get('settings', {})
        self.setup_logging()
    
    def setup_logging(self):
        """Setup logging for debugging."""
        logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
        self.logger = logging.getLogger(__name__)
    
    def load_config(self, config_path: str) -> Dict[str, Any]:
        """Load configuration from JSON file."""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"❌ Configuration file not found: {config_path}")
            return {"configurations": [], "settings": {}}
        except json.JSONDecodeError as e:
            print(f"❌ Error parsing JSON: {e}")
            return {"configurations": [], "settings": {}}
    
    def fuzzy_find(self, text: str, pattern: str, threshold: float = 0.8) -> Optional[int]:
        """Find pattern in text using fuzzy matching."""
        if not pattern or not text:
            return None
        
        text_lower = text.lower()
        pattern_lower = pattern.lower()
        
        # First try exact match
        exact_pos = text_lower.find(pattern_lower)
        if exact_pos != -1:
            return exact_pos
        
        # Try fuzzy matching with sliding window
        pattern_len = len(pattern_lower)
        best_ratio = 0
        best_pos = None
        
        for i in range(len(text_lower) - pattern_len + 1):
            window = text_lower[i:i + pattern_len]
            ratio = SequenceMatcher(None, pattern_lower, window).ratio()
            
            if ratio > best_ratio and ratio >= threshold:
                best_ratio = ratio
                best_pos = i
        
        return best_pos
    
    def normalize_whitespace(self, text: str) -> str:
        """Normalize whitespace in text."""
        # Replace multiple whitespace with single space
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def validate_extraction(self, value: str, field_name: str) -> Tuple[str, float]:
        """Validate extracted value and return confidence score."""
        if not value:
            return value, 0.0
        
        confidence = 1.0
        
        # Field-specific validation
        if 'amount' in field_name.lower() or 'total' in field_name.lower():
            # Check for currency patterns
            if re.search(r'[\$£€¥][\d,]+\.?\d*', value):
                confidence = 0.95
            elif re.search(r'\d+\.?\d*', value):
                confidence = 0.8
            else:
                confidence = 0.3
        
        elif 'date' in field_name.lower():
            # Check for date patterns
            date_patterns = [
                r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}',  # MM/DD/YYYY or DD/MM/YYYY
                r'\d{2,4}[/-]\d{1,2}[/-]\d{1,2}',  # YYYY/MM/DD
                r'\w+ \d{1,2}, \d{4}'              # Month DD, YYYY
            ]
            
            for pattern in date_patterns:
                if re.search(pattern, value):
                    confidence = 0.9
                    break
            else:
                confidence = 0.4
        
        elif 'invoice' in field_name.lower() or 'number' in field_name.lower():
            # Check for alphanumeric patterns
            if re.search(r'[A-Z]{2,}-?\d+', value):
                confidence = 0.95
            elif re.search(r'\d+', value):
                confidence = 0.8
            else:
                confidence = 0.5
        
        return value, confidence
    
    def extract_with_regex(self, text: str, pattern: str) -> Optional[str]:
        """Extract using regex pattern."""
        try:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                # Return the first group if exists, otherwise the whole match
                return match.group(1) if match.groups() else match.group(0)
        except re.error as e:
            self.logger.warning(f"Invalid regex pattern '{pattern}': {e}")
        return None
    
    def smart_boundary_detection(self, text: str, start_pos: int, max_length: int) -> str:
        """Detect smart boundaries for extraction."""
        if start_pos >= len(text):
            return ""
        
        end_pos = min(start_pos + max_length, len(text))
        extracted = text[start_pos:end_pos]
        
        # Try to end at word boundary
        if end_pos < len(text) and not text[end_pos].isspace():
            # Find the last space within the extracted text
            last_space = extracted.rfind(' ')
            if last_space > len(extracted) * 0.7:  # Only if it's not too early
                extracted = extracted[:last_space]
        
        # Remove trailing punctuation except for currency and numbers
        while extracted and extracted[-1] in '.,;:!?' and not re.search(r'\d[.,]\d*$', extracted):
            extracted = extracted[:-1]
        
        return extracted.strip()
    
    def extract_value_between_text(self, text: str, before_text: str, after_text: str, 
                                 case_sensitive: bool = False, 
                                 fuzzy_threshold: float = 0.8,
                                 max_length: int = None,
                                 regex_pattern: str = None) -> Tuple[Optional[str], float]:
        """Enhanced extraction with fuzzy matching and validation."""
        if not text or not before_text:
            return None, 0.0
        
        # Use regex pattern if provided
        if regex_pattern:
            extracted = self.extract_with_regex(text, regex_pattern)
            if extracted:
                confidence = 0.9
                return self.normalize_whitespace(extracted), confidence
        
        # Normalize text
        normalized_text = self.normalize_whitespace(text)
        search_text = normalized_text if case_sensitive else normalized_text.lower()
        before_pattern = before_text if case_sensitive else before_text.lower()
        after_pattern = after_text if case_sensitive else after_text.lower() if after_text else ""
        
        # Find start position with fuzzy matching
        start_pos = self.fuzzy_find(search_text, before_pattern, fuzzy_threshold)
        if start_pos is None:
            return None, 0.0
        
        start_pos += len(before_pattern)
        
        # Determine max extraction length
        if max_length is None:
            max_length = self.settings.get('max_extraction_length', 100)
        
        # Find end position
        if not after_text:
            extracted = self.smart_boundary_detection(normalized_text, start_pos, max_length)
        else:
            # Find end position with fuzzy matching
            end_pos = self.fuzzy_find(search_text[start_pos:], after_pattern, fuzzy_threshold)
            if end_pos is None:
                extracted = self.smart_boundary_detection(normalized_text, start_pos, max_length)
            else:
                end_pos += start_pos
                extracted = normalized_text[start_pos:end_pos].strip()
        
        # Apply length limit
        if len(extracted) > max_length:
            extracted = self.smart_boundary_detection(extracted, 0, max_length)
        
        # Clean up the extracted text
        if self.settings.get('trim_whitespace', True):
            extracted = extracted.strip()
        
        if self.settings.get('remove_special_chars', False):
            extracted = re.sub(r'[^\w\s.-]', '', extracted)
        
        # Calculate confidence based on extraction quality
        confidence = 0.8  # Base confidence for successful extraction
        if start_pos == search_text.find(before_pattern):  # Exact match
            confidence += 0.1
        if after_text and after_pattern in search_text[start_pos:]:  # Found end pattern
            confidence += 0.1
        
        return extracted if extracted else None, min(confidence, 1.0)
    
    def extract_from_text_file(self, file_path: str) -> Dict[str, Any]:
        """Extract all configured values from a single text file with enhanced features."""
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
            extracted_value = None
            best_confidence = 0.0
            
            # Get field-specific settings
            field_max_length = config.get('max_length', self.settings.get('max_extraction_length', 100))
            fuzzy_threshold = config.get('fuzzy_threshold', 0.8)
            
            # Check if config uses patterns structure
            if 'patterns' in config:
                # Try each pattern until one succeeds or find the best one
                for pattern in config['patterns']:
                    before_text = pattern.get('before_text', '')
                    after_text = pattern.get('after_text', '')
                    case_sensitive = pattern.get('case_sensitive', False)
                    regex_pattern = pattern.get('regex_pattern')
                    pattern_max_length = pattern.get('max_length', field_max_length)
                    
                    value, confidence = self.extract_value_between_text(
                        text_content, before_text, after_text, case_sensitive,
                        fuzzy_threshold, pattern_max_length, regex_pattern
                    )
                    
                    if value and confidence > best_confidence:
                        extracted_value = value
                        best_confidence = confidence
                        
                        # If we have high confidence, stop trying other patterns
                        if confidence >= 0.9:
                            break
            
            # Validate the extracted value
            if extracted_value:
                validated_value, validation_confidence = self.validate_extraction(extracted_value, config_name)
                final_confidence = (best_confidence + validation_confidence) / 2
                
                results[config_name] = {
                    'value': validated_value,
                    'confidence': round(final_confidence, 2)
                }
            else:
                results[config_name] = {
                    'value': self.settings.get('default_value_if_not_found', 'NOT_FOUND'),
                    'confidence': 0.0
                }
        
        return results
    
    def process_all_files(self) -> pd.DataFrame:
        """Process all text files and return results as DataFrame."""
        if not os.path.exists(TEXT_FILES_FOLDER_PATH):
            print(f"❌ Text files folder not found: {TEXT_FILES_FOLDER_PATH}")
            return pd.DataFrame()
        
        text_files = [f for f in os.listdir(TEXT_FILES_FOLDER_PATH) if f.endswith('.txt')]
        
        if not text_files:
            print(f"❌ No text files found in: {TEXT_FILES_FOLDER_PATH}")
            return pd.DataFrame()
        
        print(f"📄 Found {len(text_files)} text files to process")
        
        all_results = []
        
        for file_name in text_files:
            file_path = os.path.join(TEXT_FILES_FOLDER_PATH, file_name)
            print(f"🔍 Processing: {file_name}")
            
            # Extract data from the file
            extracted_data = self.extract_from_text_file(file_path)
            
            # Prepare row for DataFrame
            row_data = {'File_Name': file_name}
            
            # Add extracted values and confidence scores
            for field_name, field_data in extracted_data.items():
                if isinstance(field_data, dict):
                    row_data[field_name] = field_data.get('value', 'NOT_FOUND')
                    row_data[f"{field_name}_Confidence"] = field_data.get('confidence', 0.0)
                else:
                    # Backward compatibility
                    row_data[field_name] = field_data
                    row_data[f"{field_name}_Confidence"] = 0.5
            
            all_results.append(row_data)
        
        return pd.DataFrame(all_results)
    
    def export_to_excel(self, df: pd.DataFrame) -> bool:
        """Export DataFrame to Excel with enhanced formatting."""
        try:
            # Create export folder if it doesn't exist
            os.makedirs(EXPORT_FOLDER_PATH, exist_ok=True)
            
            excel_path = os.path.join(EXPORT_FOLDER_PATH, EXCEL_FILE_NAME)
            
            # Create Excel writer with formatting
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                # Write main data
                df.to_excel(writer, sheet_name='Extracted_Data', index=False)
                
                # Create summary sheet
                self.create_summary_sheet(writer, df)
                
                # Format the main sheet
                self.format_main_sheet(writer, df)
            
            print(f"✅ Enhanced data exported to: {excel_path}")
            return True
            
        except Exception as e:
            print(f"❌ Error exporting to Excel: {e}")
            return False
    
    def create_summary_sheet(self, writer, df: pd.DataFrame):
        """Create a summary sheet with extraction statistics."""
        summary_data = []
        
        # Get confidence columns
        confidence_cols = [col for col in df.columns if col.endswith('_Confidence')]
        
        for conf_col in confidence_cols:
            field_name = conf_col.replace('_Confidence', '')
            if field_name in df.columns:
                avg_confidence = df[conf_col].mean()
                success_rate = (df[conf_col] > 0).sum() / len(df) * 100
                
                summary_data.append({
                    'Field': field_name,
                    'Average_Confidence': round(avg_confidence, 2),
                    'Success_Rate_%': round(success_rate, 1),
                    'Total_Extractions': len(df),
                    'Successful_Extractions': (df[conf_col] > 0).sum()
                })
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Extraction_Summary', index=False)
    
    def format_main_sheet(self, writer, df: pd.DataFrame):
        """Apply formatting to the main Excel sheet."""
        try:
            from openpyxl.styles import PatternFill, Font
            from openpyxl.formatting.rule import ColorScaleRule
            
            workbook = writer.book
            worksheet = writer.sheets['Extracted_Data']
            
            # Header formatting
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
            
            # Apply conditional formatting to confidence columns
            confidence_cols = [col for col in df.columns if col.endswith('_Confidence')]
            
            for i, col_name in enumerate(df.columns):
                if col_name in confidence_cols:
                    col_letter = chr(65 + i)  # Convert to Excel column letter
                    col_range = f"{col_letter}2:{col_letter}{len(df) + 1}"
                    
                    # Color scale: red (0) to green (1)
                    rule = ColorScaleRule(
                        start_type='num', start_value=0, start_color='FF6B6B',
                        end_type='num', end_value=1, end_color='51CF66'
                    )
                    worksheet.conditional_formatting.add(col_range, rule)
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
                
        except ImportError:
            print("⚠️  openpyxl formatting features not available")
        except Exception as e:
            print(f"⚠️  Error applying Excel formatting: {e}")


# ================================
# ENHANCED CONFIGURATION CREATOR
# ================================

def create_enhanced_config():
    """Create an enhanced configuration with fuzzy matching and regex support."""
    enhanced_config = {
        "settings": {
            "max_extraction_length": 100,
            "trim_whitespace": True,
            "remove_special_chars": False,
            "default_value_if_not_found": "NOT_FOUND",
            "fuzzy_matching_threshold": 0.8
        },
        "configurations": [
            {
                "name": "Invoice_Number",
                "max_length": 50,
                "fuzzy_threshold": 0.8,
                "patterns": [
                    {
                        "before_text": "Invoice Number:",
                        "after_text": "Date:",
                        "case_sensitive": False,
                        "max_length": 30
                    },
                    {
                        "before_text": "Invoice #:",
                        "after_text": "\n",
                        "case_sensitive": False,
                        "max_length": 30
                    },
                    {
                        "regex_pattern": r"(?:Invoice|INV)[-#:\s]*([A-Z0-9-]+)",
                        "max_length": 30
                    }
                ]
            },
            {
                "name": "Total_Amount",
                "max_length": 30,
                "fuzzy_threshold": 0.7,
                "patterns": [
                    {
                        "before_text": "Total:",
                        "after_text": "\n",
                        "case_sensitive": False,
                        "max_length": 20
                    },
                    {
                        "before_text": "Amount Due:",
                        "after_text": "\n",
                        "case_sensitive": False,
                        "max_length": 20
                    },
                    {
                        "regex_pattern": r"(?:Total|Amount Due|Grand Total)[:$\s]*(\$?[\d,]+\.?\d*)",
                        "max_length": 20
                    }
                ]
            },
            {
                "name": "Company_Name",
                "max_length": 80,
                "fuzzy_threshold": 0.8,
                "patterns": [
                    {
                        "before_text": "Company:",
                        "after_text": "Invoice",
                        "case_sensitive": False,
                        "max_length": 60
                    },
                    {
                        "before_text": "Bill To:",
                        "after_text": "\n",
                        "case_sensitive": False,
                        "max_length": 60
                    },
                    {
                        "regex_pattern": r"Company:\s*([A-Za-z\s&.,'-]+?)(?:\s*Invoice|\s*$)",
                        "max_length": 60
                    }
                ]
            },
            {
                "name": "Date",
                "max_length": 25,
                "fuzzy_threshold": 0.8,
                "patterns": [
                    {
                        "before_text": "Date:",
                        "after_text": "\n",
                        "case_sensitive": False,
                        "max_length": 20
                    },
                    {
                        "before_text": "Invoice Date:",
                        "after_text": "\n",
                        "case_sensitive": False,
                        "max_length": 20
                    },
                    {
                        "regex_pattern": r"(?:Date|Invoice Date)[:]\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",
                        "max_length": 15
                    }
                ]
            }
        ]
    }
    
    config_path = os.path.join(os.path.dirname(__file__), "enhanced_extraction_config.json")
    
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(enhanced_config, f, indent=2, ensure_ascii=False)
        print(f"✅ Enhanced configuration created: {config_path}")
        return config_path
    except Exception as e:
        print(f"❌ Error creating enhanced config: {e}")
        return None


# ================================
# MAIN EXECUTION
# ================================

def main():
    """Main function to run the enhanced text extraction."""
    print("🚀 Starting Enhanced Text to Excel Extraction...")
    print("=" * 60)
    
    # Check if enhanced config exists, create if not
    enhanced_config_path = os.path.join(os.path.dirname(__file__), "enhanced_extraction_config.json")
    
    if not os.path.exists(enhanced_config_path):
        print("📝 Creating enhanced configuration...")
        enhanced_config_path = create_enhanced_config()
        if not enhanced_config_path:
            print("❌ Failed to create enhanced configuration. Using original config.")
            enhanced_config_path = CONFIG_FILE_PATH
    else:
        print(f"📋 Using existing enhanced configuration: {enhanced_config_path}")
    
    # Initialize enhanced extractor
    extractor = EnhancedTextExtractor(enhanced_config_path)
    
    # Process all files
    print("\n🔄 Processing text files...")
    results_df = extractor.process_all_files()
    
    if results_df.empty:
        print("❌ No data extracted. Please check your text files and configuration.")
        return
    
    # Display summary
    print(f"\n📊 Extraction Summary:")
    print(f"   • Files processed: {len(results_df)}")
    print(f"   • Fields extracted: {len([col for col in results_df.columns if not col.endswith('_Confidence') and col != 'File_Name'])}")
    
    # Show confidence statistics
    confidence_cols = [col for col in results_df.columns if col.endswith('_Confidence')]
    if confidence_cols:
        avg_confidence = results_df[confidence_cols].mean().mean()
        print(f"   • Average confidence: {avg_confidence:.2f}")
    
    # Export to Excel
    print(f"\n💾 Exporting to Excel...")
    success = extractor.export_to_excel(results_df)
    
    if success:
        print(f"\n✅ Enhanced extraction completed successfully!")
        print(f"📁 Output file: {os.path.join(EXPORT_FOLDER_PATH, EXCEL_FILE_NAME)}")
        
        # Show top results
        print(f"\n🔍 Sample Results:")
        for i, row in results_df.head(3).iterrows():
            print(f"   📄 {row['File_Name']}")
            for col in results_df.columns:
                if not col.endswith('_Confidence') and col != 'File_Name':
                    conf_col = f"{col}_Confidence"
                    confidence = row.get(conf_col, 0)
                    value = row[col]
                    print(f"      {col}: {value} (confidence: {confidence:.2f})")
    else:
        print("❌ Export failed. Please check the error messages above.")


if __name__ == "__main__":
    main()