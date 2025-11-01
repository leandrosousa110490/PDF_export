#!/usr/bin/env python3
"""
PDF Text Extractor

This script extracts text from all PDF files in a specified folder and saves
the extracted text to separate text files with the same name as the original PDFs.
Supports multi-page PDFs and handles various PDF formats.

Requirements:
- PyPDF2 library (install with: pip install PyPDF2)

Author: AI Assistant
Date: 2024
"""

import os
import sys
from pathlib import Path

# =============================================================================
# CONFIGURATION VARIABLES - MODIFY THESE PATHS AS NEEDED
# =============================================================================

# Path to the folder containing PDF files
PDF_FOLDER_PATH = r"C:\Users\nbaba\Desktop\PDF to Excel\folder"

# Path to the folder where extracted text files will be saved
# Set to None to save in the same folder as PDFs
OUTPUT_FOLDER_PATH = r"C:\Users\nbaba\Desktop\PDF to Excel\extracted_texts"

# =============================================================================
try:
    import PyPDF2
except ImportError:
    try:
        import pypdf as PyPDF2
    except ImportError:
        print("Error: PyPDF2 or pypdf library is required.")
        print("Install it using: pip install PyPDF2")
        print("Or: pip install pypdf")
        sys.exit(1)


def extract_text_from_pdf(pdf_path):
    """
    Extract text from a PDF file, handling multi-page documents.
    Returns all text as one continuous line with spaces between words.
    
    Args:
        pdf_path (str): Path to the PDF file
        
    Returns:
        str: Extracted text as one continuous line, or None if extraction fails
    """
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            # Check if PDF is encrypted
            if pdf_reader.is_encrypted:
                print(f"⚠️  Warning: {os.path.basename(pdf_path)} is encrypted and cannot be processed")
                return None
            
            # Extract text from all pages
            all_text = ""
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                page_text = page.extract_text()
                if page_text:
                    all_text += page_text
            
            # Format text as one continuous line
            # Replace multiple whitespace characters (including newlines) with single spaces
            formatted_text = ' '.join(all_text.split())
            
            return formatted_text if formatted_text.strip() else None
            
    except Exception as e:
        print(f"❌ Error extracting text from {os.path.basename(pdf_path)}: {str(e)}")
        return None


def process_pdfs_in_folder(folder_path, output_folder=None):
    """
    Process all PDF files in the specified folder.
    
    Args:
        folder_path (str): Path to the folder containing PDF files
        output_folder (str): Path to the folder where text files will be saved.
                           If None, saves in the same folder as PDFs.
    """
    folder_path = Path(folder_path)
    
    # Set output folder - use same folder as PDFs if not specified
    if output_folder is None:
        output_path = folder_path
    else:
        output_path = Path(output_folder)
    
    pdf_files = list(folder_path.glob("*.pdf"))
    
    if not pdf_files:
        print(f"No PDF files found in '{folder_path}'")
        return
    
    print(f"Found {len(pdf_files)} PDF file(s) in '{folder_path}'")
    if output_folder:
        print(f"Text files will be saved to: '{output_path}'")
    else:
        print("Text files will be saved in the same folder as PDFs")
    print("Processing files...")
    
    successful_extractions = 0
    failed_extractions = 0
    
    for pdf_file in pdf_files:
        print(f"\nProcessing: {pdf_file.name}")
        
        # Extract text from PDF
        extracted_text = extract_text_from_pdf(pdf_file)
        
        if extracted_text is not None:
            # Create output text file with same name as PDF
            output_file = output_path / f"{pdf_file.stem}.txt"
            
            try:
                with open(output_file, 'w', encoding='utf-8') as text_file:
                    text_file.write(f"Text extracted from: {pdf_file.name}\n")
                    text_file.write("=" * 50 + "\n\n")
                    text_file.write(extracted_text)
                
                print(f"✓ Text saved to: {output_file}")
                successful_extractions += 1
                
            except Exception as e:
                print(f"✗ Error saving text file for {pdf_file.name}: {str(e)}")
                failed_extractions += 1
        else:
            failed_extractions += 1
    
    # Summary
    print(f"\n" + "=" * 50)
    print(f"Processing complete!")
    print(f"Successfully processed: {successful_extractions} files")
    print(f"Failed to process: {failed_extractions} files")
    
    if successful_extractions > 0:
        print(f"Text files saved in: {output_path}")


def main():
    """
    Main function to orchestrate the PDF text extraction process.
    Uses the configured path variables instead of user input.
    """
    print("=" * 50)
    print("PDF Text Extractor")
    print("=" * 50)
    
    # Use the configured paths
    pdf_folder = PDF_FOLDER_PATH
    output_folder = OUTPUT_FOLDER_PATH
    
    print(f"📁 PDF Folder: {pdf_folder}")
    print(f"📄 Output Folder: {output_folder if output_folder else 'Same as PDF folder'}")
    print()
    
    # Validate PDF folder exists
    if not os.path.exists(pdf_folder):
        print(f"❌ Error: PDF folder does not exist: {pdf_folder}")
        print("Please update the PDF_FOLDER_PATH variable in the script.")
        input("\nPress Enter to exit...")
        return
    
    # Create output folder if specified and doesn't exist
    if output_folder and not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
            print(f"✅ Created output folder: {output_folder}")
        except Exception as e:
            print(f"❌ Error creating output folder: {str(e)}")
            input("\nPress Enter to exit...")
            return
    
    # Process PDFs
    process_pdfs_in_folder(pdf_folder, output_folder)
    
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()