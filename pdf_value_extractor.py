"""
Enhanced PDF Value Extraction System
====================================

This system extracts specific values from PDF documents using configurable search patterns
and exports results to Excel files with comprehensive formatting and statistics.
"""

import os
import re
from typing import List, Optional, Dict, Any
import PyPDF2
from config import ExtractionConfig, get_pdf_files, EXTRACTION_CONFIGS
from excel_exporter import ExcelExporter

class PDFValueExtractor:
    """
    Enhanced PDF text extraction tool that can find values based on configurable patterns.
    Supports multiple extraction modes and exports results to Excel.
    """
    
    def __init__(self, config: ExtractionConfig = None):
        self.config = config
        self.results = []
        self.excel_exporter = ExcelExporter()
        
    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """Extract all text from a PDF file"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
                return text
        except Exception as e:
            print(f"Error reading PDF {pdf_path}: {str(e)}")
            return ""
    
    def validate_data_type(self, text: str) -> bool:
        """Validate if extracted text matches expected data type"""
        # Default to "any" if no config
        expected_type = self.config.expected_type if self.config and hasattr(self.config, 'expected_type') else "any"
        
        if expected_type == "any":
            return True
        elif expected_type == "numbers":
            # Check if text contains only numbers, spaces, and common number symbols
            return bool(re.match(r'^[\d\s\.,\-\+\$\%]+$', text.strip()))
        elif expected_type == "letters":
            # Check if text contains only letters and spaces
            return bool(re.match(r'^[a-zA-Z\s]+$', text.strip()))
        elif expected_type == "both":
            # Check if text contains both letters and numbers
            text_clean = text.strip()
            has_letters = bool(re.search(r'[a-zA-Z]', text_clean))
            has_numbers = bool(re.search(r'\d', text_clean))
            return has_letters and has_numbers
        return False
    
    def clean_extracted_text(self, text: str) -> str:
        """Clean and format extracted text"""
        text = text.strip()
        # Remove excessive whitespace
        text = re.sub(r'\s+', ' ', text)
        return text
    
    def find_text_positions(self, full_text: str, search_terms: List[str]) -> List[int]:
        """Find all positions of search terms in the text"""
        positions = []
        # Default to case insensitive if no config
        case_sensitive = self.config.case_sensitive if self.config and hasattr(self.config, 'case_sensitive') else False
        text_to_search = full_text if case_sensitive else full_text.lower()
        
        for term in search_terms:
            search_term = term if case_sensitive else term.lower()
            start = 0
            while True:
                pos = text_to_search.find(search_term, start)
                if pos == -1:
                    break
                positions.append((pos, pos + len(term)))
                start = pos + 1
        
        return sorted(positions)
    
    def extract_value_after_text(self, full_text: str, search_terms: List[str]) -> Optional[str]:
        """Extract value that appears after any of the search terms"""
        positions = self.find_text_positions(full_text, search_terms)
        
        for start_pos, end_pos in positions:
            # Extract text after the search term
            after_text = full_text[end_pos:end_pos + self.config.max_distance]
            
            # Look for the first meaningful content (skip whitespace and common separators)
            match = re.search(r'[\s:]*([^\n\r]+)', after_text)
            if match:
                extracted = match.group(1).strip()
                if extracted and self.validate_data_type(extracted):
                    return self.clean_extracted_text(extracted)
        
        return None
    
    def extract_value_before_text(self, full_text: str, search_terms: List[str]) -> Optional[str]:
        """Extract value that appears before any of the search terms"""
        positions = self.find_text_positions(full_text, search_terms)
        
        for start_pos, end_pos in positions:
            # Extract text before the search term
            before_start = max(0, start_pos - self.config.max_distance)
            before_text = full_text[before_start:start_pos]
            
            # Look for the last meaningful content before the search term
            lines = before_text.split('\n')
            for line in reversed(lines):
                line = line.strip()
                if line and self.validate_data_type(line):
                    return self.clean_extracted_text(line)
        
        return None
    
    def extract_value_around_text(self, full_text: str, search_terms: List[str]) -> Optional[str]:
        """Extract value that appears around any of the search terms"""
        positions = self.find_text_positions(full_text, search_terms)
        
        for start_pos, end_pos in positions:
            # Extract text around the search term
            around_start = max(0, start_pos - self.config.max_distance // 2)
            around_end = min(len(full_text), end_pos + self.config.max_distance // 2)
            around_text = full_text[around_start:around_end]
            
            # Split into lines and look for valid content
            lines = around_text.split('\n')
            for line in lines:
                line = line.strip()
                if line and line not in search_terms and self.validate_data_type(line):
                    return self.clean_extracted_text(line)
        
        return None
    
    def extract_value_between_text(self, full_text: str, search_terms: List[str], end_text: str = None) -> Optional[str]:
        """Extract value that appears between start and end text markers"""
        if not end_text:
            # If no end_text specified, use the second search term if available
            if len(search_terms) >= 2:
                start_terms = [search_terms[0]]
                end_text = search_terms[1]
            else:
                return None
        else:
            start_terms = search_terms
        
        # Find positions of start text
        start_positions = self.find_text_positions(full_text, start_terms)
        
        for start_pos, start_end_pos in start_positions:
            # Look for end text after the start position
            remaining_text = full_text[start_end_pos:]
            
            # Find end text (case sensitive or not based on config)
            if self.config and hasattr(self.config, 'case_sensitive') and self.config.case_sensitive:
                end_pos = remaining_text.find(end_text)
            else:
                end_pos = remaining_text.lower().find(end_text.lower())
            
            if end_pos != -1:
                # Extract text between start and end markers
                between_text = remaining_text[:end_pos].strip()
                
                # Clean up the extracted text
                between_text = re.sub(r'^\s*[:;,.-]*\s*', '', between_text)  # Remove leading punctuation
                between_text = re.sub(r'\s*[:;,.-]*\s*$', '', between_text)  # Remove trailing punctuation
                
                if between_text and self.validate_data_type(between_text):
                    return self.clean_extracted_text(between_text)
        
        return None
    
    def extract_value_from_text(self, text: str) -> Optional[str]:
        """Extract value from text using the configured extraction method"""
        if not self.config:
            return None
        
        if self.config.extraction_mode == "after":
            return self.extract_value_after_text(text, self.config.search_terms)
        elif self.config.extraction_mode == "before":
            return self.extract_value_before_text(text, self.config.search_terms)
        elif self.config.extraction_mode == "between":
            # Use end_text if available, otherwise use second search term
            end_text = getattr(self.config, 'end_text', None)
            return self.extract_value_between_text(text, self.config.search_terms, end_text)
        elif self.config.extraction_mode == "around":
            return self.extract_value_around_text(text, self.config.search_terms)
        
        return None
    
    def extract_value(self, pdf_path: str) -> Optional[str]:
        """Extract value from a single PDF using the configured extraction method"""
        if not self.config:
            return None
            
        full_text = self.extract_text_from_pdf(pdf_path)
        if not full_text:
            return None
        
        if self.config.extraction_mode == "after":
            return self.extract_value_after_text(full_text, self.config.search_terms)
        elif self.config.extraction_mode == "before":
            return self.extract_value_before_text(full_text, self.config.search_terms)
        elif self.config.extraction_mode == "between":
            # Use end_text if available, otherwise use second search term
            end_text = getattr(self.config, 'end_text', None)
            return self.extract_value_between_text(full_text, self.config.search_terms, end_text)
        elif self.config.extraction_mode == "around":
            return self.extract_value_around_text(full_text, self.config.search_terms)
        
        return None
    
    def extract_from_multiple_pdfs(self, pdf_paths: List[str], config_name: str = None) -> List[Dict[str, Any]]:
        """Extract values from multiple PDFs and return results"""
        results = []
        
        for pdf_path in pdf_paths:
            pdf_name = os.path.basename(pdf_path)
            extracted_value = self.extract_value(pdf_path)
            
            if extracted_value:
                status = "Success"
                value = extracted_value
            else:
                status = "No Match"
                value = "No match found"
            
            result = {
                'pdf_name': pdf_name,
                'config_name': config_name or self.config.name,
                'extracted_value': value,
                'status': status
            }
            results.append(result)
            
            # Add to Excel exporter
            self.excel_exporter.add_result(
                pdf_name=pdf_name,
                config_name=config_name or self.config.name,
                extracted_value=value,
                status=status
            )
        
        return results
    
    def extract_with_multiple_configs(self, pdf_paths: List[str], configs: Dict[str, ExtractionConfig]) -> List[Dict[str, Any]]:
        """Extract values using multiple configurations"""
        all_results = []
        
        for config_name, config in configs.items():
            self.config = config
            results = self.extract_from_multiple_pdfs(pdf_paths, config_name)
            all_results.extend(results)
        
        return all_results
    
    def export_to_excel(self, filename: str = None) -> str:
        """Export all results to Excel file"""
        if filename and not filename.endswith('.xlsx'):
            filename = f"{filename}.xlsx"
        return self.excel_exporter.save_to_file(filename)
    
    def print_results_summary(self):
        """Print a summary of extraction results"""
        self.excel_exporter.print_summary()
    
    def clear_results(self):
        """Clear all stored results"""
        self.excel_exporter.clear_results()