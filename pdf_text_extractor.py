#!/usr/bin/env python3
"""
Enhanced PDF Text Extractor

This script extracts text from all PDF files in a specified folder using multiple
extraction methods and libraries to maximize text recovery, including from redacted
pages and image-based PDFs. Saves extracted text to separate text files.

Features:
- Multiple extraction libraries (PyPDF2, pdfplumber, PyMuPDF)
- OCR support for image-based content
- Enhanced text recovery from redacted pages
- Fallback methods for difficult PDFs
- Multi-page PDF support

Requirements:
- PyPDF2 or pypdf
- pdfplumber
- PyMuPDF (fitz)
- pytesseract (for OCR)
- Pillow (PIL)

Author: AI Assistant
Date: 2024
"""

import os
import sys
import io
from pathlib import Path
import warnings
warnings.filterwarnings("ignore")

# =============================================================================
# CONFIGURATION VARIABLES - MODIFY THESE PATHS AS NEEDED
# =============================================================================

# Path to the folder containing PDF files
PDF_FOLDER_PATH = r"C:\Users\nbaba\Desktop\PDF to Excel\folder"

# Path to the folder where extracted text files will be saved
# Set to None to save in the same folder as PDFs
OUTPUT_FOLDER_PATH = r"C:\Users\nbaba\Desktop\PDF to Excel\extracted_texts"

# =============================================================================

# Import libraries with fallbacks
libraries_available = {
    'pypdf': False,
    'pdfplumber': False,
    'fitz': False,
    'tesseract': False
}

# Try PyPDF2/pypdf
try:
    import PyPDF2
    libraries_available['pypdf'] = True
except ImportError:
    try:
        import pypdf as PyPDF2
        libraries_available['pypdf'] = True
    except ImportError:
        PyPDF2 = None

# Try pdfplumber
try:
    import pdfplumber
    libraries_available['pdfplumber'] = True
except ImportError:
    pdfplumber = None

# Try PyMuPDF (fitz)
try:
    import fitz  # PyMuPDF
    libraries_available['fitz'] = True
except ImportError:
    fitz = None

# Try OCR libraries
try:
    import pytesseract
    from PIL import Image
    libraries_available['tesseract'] = True
except ImportError:
    pytesseract = None
    Image = None


def print_library_status():
    """Print the status of available libraries."""
    print("📚 Available extraction libraries:")
    for lib, available in libraries_available.items():
        status = "✅ Available" if available else "❌ Not installed"
        print(f"   {lib}: {status}")
    print()


def extract_with_pypdf(pdf_path):
    """Extract text using PyPDF2/pypdf."""
    if not PyPDF2:
        return None
    
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            if pdf_reader.is_encrypted:
                # Try to decrypt with empty password
                try:
                    pdf_reader.decrypt("")
                except:
                    return None
            
            all_text = ""
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                page_text = page.extract_text()
                if page_text:
                    all_text += page_text + " "
            
            return all_text.strip() if all_text.strip() else None
            
    except Exception as e:
        print(f"   PyPDF2 extraction failed: {str(e)}")
        return None


def extract_with_pdfplumber(pdf_path):
    """Extract text using pdfplumber - better for complex layouts."""
    if not pdfplumber:
        return None
    
    try:
        all_text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                # Extract text
                page_text = page.extract_text()
                if page_text:
                    all_text += page_text + " "
                
                # Also try to extract text from tables
                try:
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            if row:
                                table_text = " ".join([cell for cell in row if cell])
                                all_text += table_text + " "
                except:
                    pass
        
        return all_text.strip() if all_text.strip() else None
        
    except Exception as e:
        print(f"   pdfplumber extraction failed: {str(e)}")
        return None


def extract_with_fitz(pdf_path):
    """Extract text using PyMuPDF (fitz) - good for redacted content."""
    if not fitz:
        return None
    
    try:
        doc = fitz.open(pdf_path)
        all_text = ""
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            
            # Extract text
            page_text = page.get_text()
            if page_text:
                all_text += page_text + " "
            
            # Try to extract text from annotations and forms
            try:
                # Get text from annotations
                for annot in page.annots():
                    if annot.content:
                        all_text += annot.content + " "
                
                # Get text from form fields
                widgets = page.widgets()
                for widget in widgets:
                    if widget.field_value:
                        all_text += str(widget.field_value) + " "
            except:
                pass
            
            # Try OCR on images if tesseract is available
            if pytesseract and Image:
                try:
                    # Get images from page
                    image_list = page.get_images()
                    for img_index, img in enumerate(image_list):
                        # Extract image
                        xref = img[0]
                        pix = fitz.Pixmap(doc, xref)
                        if pix.n - pix.alpha < 4:  # GRAY or RGB
                            img_data = pix.tobytes("png")
                            img_pil = Image.open(io.BytesIO(img_data))
                            # OCR the image
                            ocr_text = pytesseract.image_to_string(img_pil)
                            if ocr_text.strip():
                                all_text += ocr_text + " "
                        pix = None
                except:
                    pass
        
        doc.close()
        return all_text.strip() if all_text.strip() else None
        
    except Exception as e:
        print(f"   PyMuPDF extraction failed: {str(e)}")
        return None


def extract_text_from_pdf_enhanced(pdf_path):
    """
    Enhanced text extraction using multiple methods and libraries.
    
    Args:
        pdf_path (str): Path to the PDF file
        
    Returns:
        str: Extracted text, or None if all methods fail
    """
    print(f"   Trying multiple extraction methods...")
    
    extracted_texts = []
    methods_tried = []
    
    # Method 1: PyMuPDF (best for redacted content and images)
    if libraries_available['fitz']:
        methods_tried.append("PyMuPDF")
        fitz_text = extract_with_fitz(pdf_path)
        if fitz_text:
            extracted_texts.append(("PyMuPDF", fitz_text))
    
    # Method 2: pdfplumber (best for complex layouts)
    if libraries_available['pdfplumber']:
        methods_tried.append("pdfplumber")
        plumber_text = extract_with_pdfplumber(pdf_path)
        if plumber_text:
            extracted_texts.append(("pdfplumber", plumber_text))
    
    # Method 3: PyPDF2 (fallback)
    if libraries_available['pypdf']:
        methods_tried.append("PyPDF2")
        pypdf_text = extract_with_pypdf(pdf_path)
        if pypdf_text:
            extracted_texts.append(("PyPDF2", pypdf_text))
    
    print(f"   Methods tried: {', '.join(methods_tried)}")
    print(f"   Successful extractions: {len(extracted_texts)}")
    
    if not extracted_texts:
        return None
    
    # Combine results, prioritizing the longest extraction
    best_extraction = max(extracted_texts, key=lambda x: len(x[1]))
    best_method, best_text = best_extraction
    
    print(f"   Best result from: {best_method} ({len(best_text)} characters)")
    
    # Combine unique content from all methods
    combined_text = best_text
    
    # Add unique content from other methods
    for method, text in extracted_texts:
        if method != best_method:
            # Simple check for unique content
            words_in_best = set(best_text.lower().split())
            words_in_current = set(text.lower().split())
            unique_words = words_in_current - words_in_best
            
            if len(unique_words) > 10:  # If significant unique content
                combined_text += "\n\n--- Additional content from " + method + " ---\n"
                combined_text += text
    
    # Format text as one continuous line
    text_with_spaces = combined_text.replace('\n', ' ').replace('\r', ' ')
    formatted_text = ' '.join(text_with_spaces.split())
    
    return formatted_text if formatted_text.strip() else None


def process_pdfs_in_folder(folder_path, output_folder=None):
    """
    Process all PDF files in the specified folder using enhanced extraction.
    
    Args:
        folder_path (str): Path to the folder containing PDF files
        output_folder (str): Path to the folder where text files will be saved.
                           If None, saves in the same folder as PDFs.
    """
    folder_path = Path(folder_path)
    
    # Set output folder
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
    
    print_library_status()
    print("Processing files with enhanced extraction...")
    
    successful_extractions = 0
    failed_extractions = 0
    
    for pdf_file in pdf_files:
        print(f"\n📄 Processing: {pdf_file.name}")
        
        # Extract text using enhanced method
        extracted_text = extract_text_from_pdf_enhanced(pdf_file)
        
        if extracted_text is not None:
            # Create output text file
            output_file = output_path / f"{pdf_file.stem}.txt"
            
            try:
                with open(output_file, 'w', encoding='utf-8') as text_file:
                    text_file.write(f"Text extracted from: {pdf_file.name}\n")
                    text_file.write(f"Enhanced extraction with multiple methods\n")
                    text_file.write("=" * 50 + "\n\n")
                    text_file.write(extracted_text)
                
                print(f"   ✅ Text saved to: {output_file.name}")
                print(f"   📊 Extracted {len(extracted_text)} characters")
                successful_extractions += 1
                
            except Exception as e:
                print(f"   ❌ Error saving text file: {str(e)}")
                failed_extractions += 1
        else:
            print(f"   ❌ No text could be extracted")
            failed_extractions += 1
    
    # Summary
    print(f"\n" + "=" * 60)
    print(f"🎯 EXTRACTION COMPLETE!")
    print(f"✅ Successfully processed: {successful_extractions} files")
    print(f"❌ Failed to process: {failed_extractions} files")
    
    if successful_extractions > 0:
        print(f"📁 Text files saved in: {output_path}")


def main():
    """
    Main function to orchestrate the enhanced PDF text extraction process.
    """
    print("=" * 60)
    print("🚀 ENHANCED PDF TEXT EXTRACTOR")
    print("=" * 60)
    
    # Use the configured paths
    pdf_folder = PDF_FOLDER_PATH
    output_folder = OUTPUT_FOLDER_PATH
    
    print(f"📁 PDF Folder: {pdf_folder}")
    print(f"📄 Output Folder: {output_folder if output_folder else 'Same as PDF folder'}")
    print()
    
    # Check if any extraction library is available
    if not any(libraries_available.values()):
        print("❌ ERROR: No PDF extraction libraries are installed!")
        print("\nPlease install at least one of the following:")
        print("   pip install PyPDF2")
        print("   pip install pdfplumber")
        print("   pip install PyMuPDF")
        print("   pip install pytesseract pillow  # For OCR support")
        input("\nPress Enter to exit...")
        return
    
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
    
    # Process PDFs with enhanced extraction
    process_pdfs_in_folder(pdf_folder, output_folder)
    
    input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()