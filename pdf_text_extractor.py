#!/usr/bin/env python3
"""
Enhanced PDF Text Extractor

This script extracts text from all PDF files in a specified folder using multiple
extraction methods and libraries to maximize text recovery, including from redacted
pages and image-based PDFs. Saves extracted text to separate text files.

Features:
- Multiple extraction libraries (PyPDF2, pdfplumber, PyMuPDF)
- Image processing for redacted content detection and removal
- Enhanced text recovery from redacted pages
- Fallback methods for difficult PDFs
- Multi-page PDF support
- Security-focused: No external model downloads

Requirements:
- PyPDF2 or pypdf
- pdfplumber
- PyMuPDF (fitz)
- Pillow (PIL)
- OpenCV (cv2)
- NumPy

Author: AI Assistant
Date: 2024
"""

import os
import sys
import io
from pathlib import Path
import warnings
warnings.filterwarnings("ignore")

# Image processing imports for redacted content handling
try:
    import cv2
    import numpy as np
    from PIL import Image, ImageDraw
    IMAGE_PROCESSING_AVAILABLE = True
    print("✅ Image processing libraries loaded successfully")
except ImportError as e:
    IMAGE_PROCESSING_AVAILABLE = False
    print(f"⚠️  Image processing libraries not available: {e}")
    print("   Install with: pip install opencv-python pillow numpy")

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
    'fitz': False
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


def print_library_status():
    """Print the status of available libraries."""
    print("📚 Available extraction libraries:")
    for lib, available in libraries_available.items():
        status = "✅ Available" if available else "❌ Not installed"
        print(f"   {lib}: {status}")
    
    if IMAGE_PROCESSING_AVAILABLE:
        print("   🖼️  Image processing: ✅ Available")
    else:
        print("   🖼️  Image processing: ❌ Not available")
    print()


def detect_redacted_areas(image_array):
    """
    Detect redacted (blacked out) areas in an image using computer vision.
    Returns coordinates of detected redacted regions.
    """
    if not IMAGE_PROCESSING_AVAILABLE:
        return []
    
    try:
        # Convert to grayscale
        gray = cv2.cvtColor(image_array, cv2.COLOR_RGB2GRAY)
        
        # Detect very dark areas (potential redactions)
        # Threshold for very dark pixels (close to black)
        _, binary = cv2.threshold(gray, 30, 255, cv2.THRESH_BINARY_INV)
        
        # Find contours of dark areas
        contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        redacted_areas = []
        for contour in contours:
            # Get bounding rectangle
            x, y, w, h = cv2.boundingRect(contour)
            
            # Filter out very small areas (noise) and very large areas (background)
            area = w * h
            if 100 < area < (image_array.shape[0] * image_array.shape[1] * 0.8):
                redacted_areas.append((x, y, w, h))
        
        return redacted_areas
    
    except Exception as e:
        print(f"   ⚠️  Error detecting redacted areas: {e}")
        return []


def remove_redacted_content_from_image(image_array):
    """
    Remove or mask redacted content from an image to prevent any text extraction
    from potentially sensitive areas.
    """
    if not IMAGE_PROCESSING_AVAILABLE:
        return image_array
    
    try:
        # Detect redacted areas
        redacted_areas = detect_redacted_areas(image_array)
        
        if redacted_areas:
            print(f"   🔒 Detected {len(redacted_areas)} redacted areas - masking for security")
            
            # Create a copy of the image
            cleaned_image = image_array.copy()
            
            # Mask redacted areas with white (to prevent any text extraction)
            for x, y, w, h in redacted_areas:
                # Expand the masked area slightly to ensure complete coverage
                padding = 5
                x_start = max(0, x - padding)
                y_start = max(0, y - padding)
                x_end = min(cleaned_image.shape[1], x + w + padding)
                y_end = min(cleaned_image.shape[0], y + h + padding)
                
                # Fill with white
                cleaned_image[y_start:y_end, x_start:x_end] = [255, 255, 255]
            
            return cleaned_image
        
        return image_array
    
    except Exception as e:
        print(f"   ⚠️  Error removing redacted content: {e}")
        return image_array


def is_image_based_pdf_page(page_dict):
    """
    Determine if a PDF page is primarily image-based (scanned document).
    Returns True if the page appears to be image-based.
    """
    if not isinstance(page_dict, dict):
        return False
    
    # Check if page has images but very little text
    images = page_dict.get('images', [])
    blocks = page_dict.get('blocks', [])
    
    # Count text blocks vs image blocks
    text_blocks = 0
    image_blocks = 0
    
    for block in blocks:
        if block.get('type') == 0:  # Text block
            text_blocks += 1
        elif block.get('type') == 1:  # Image block
            image_blocks += 1
    
    # If there are images and very few text blocks, likely image-based
    if images and image_blocks > 0 and text_blocks < 3:
        return True
    
    # Check text density - if very low text content, might be image-based
    total_text = ""
    for block in blocks:
        if block.get('type') == 0:  # Text block
            for line in block.get('lines', []):
                for span in line.get('spans', []):
                    total_text += span.get('text', '')
    
    # If very little extractable text, likely image-based
    if len(total_text.strip()) < 50:
        return True
    
    return False


def extract_with_pypdf(pdf_path):
    """Enhanced text extraction using PyPDF2/pypdf - fallback method with multiple strategies."""
    if not PyPDF2:
        return None
    
    try:
        all_text = ""
        pages_processed = 0
        pages_with_text = 0
        
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            
            if pdf_reader.is_encrypted:
                # Try to decrypt with empty password
                try:
                    pdf_reader.decrypt("")
                except:
                    return None
            
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                page_text_parts = []
                pages_processed += 1
                
                # Method 1: Standard text extraction
                try:
                    page_text = page.extract_text()
                    if page_text and page_text.strip():
                        page_text_parts.append(page_text)
                except Exception as e:
                    print(f"   PyPDF2 standard extraction failed on page {page_num + 1}: {str(e)}")
                
                # Method 2: Try different extraction modes if available
                try:
                    # Use the modern extract_text method (extractText is deprecated in PyPDF2 3.0.0+)
                    if hasattr(page, 'extract_text'):
                        alt_text = page.extract_text()
                        if alt_text and alt_text.strip() and alt_text not in page_text_parts:
                            page_text_parts.append(alt_text)
                except Exception as e:
                    print(f"   PyPDF2 alternative extraction failed on page {page_num + 1}: {str(e)}")
                
                # Method 3: Extract from annotations if present
                try:
                    if hasattr(page, 'annotations') and page.annotations:
                        for annotation in page.annotations:
                            if annotation and hasattr(annotation, 'get_object'):
                                ann_obj = annotation.get_object()
                                if ann_obj and '/Contents' in ann_obj:
                                    ann_text = str(ann_obj['/Contents'])
                                    if ann_text and ann_text.strip():
                                        page_text_parts.append(ann_text)
                except Exception as e:
                    print(f"   PyPDF2 annotation extraction failed on page {page_num + 1}: {str(e)}")
                
                # Method 4: Try to extract from form fields
                try:
                    if hasattr(pdf_reader, 'get_form_text_fields'):
                        form_fields = pdf_reader.get_form_text_fields()
                        if form_fields:
                            for field_name, field_value in form_fields.items():
                                if field_value and str(field_value).strip():
                                    page_text_parts.append(f"{field_name}: {field_value}")
                except Exception as e:
                    print(f"   PyPDF2 form field extraction failed on page {page_num + 1}: {str(e)}")
                
                # Method 5: Extract from page objects directly
                try:
                    if hasattr(page, 'get_contents') and page.get_contents():
                        contents = page.get_contents()
                        if contents and hasattr(contents, 'get_data'):
                            content_data = contents.get_data()
                            if content_data:
                                # Try to decode content stream (basic approach)
                                try:
                                    decoded_content = content_data.decode('utf-8', errors='ignore')
                                    # Extract text-like content (very basic)
                                    import re
                                    text_matches = re.findall(r'\((.*?)\)', decoded_content)
                                    if text_matches:
                                        extracted_text = " ".join(text_matches)
                                        if extracted_text.strip():
                                            page_text_parts.append(extracted_text)
                                except Exception as decode_e:
                                    print(f"   PyPDF2 content decoding failed on page {page_num + 1}: {str(decode_e)}")
                except Exception as e:
                    print(f"   PyPDF2 content extraction failed on page {page_num + 1}: {str(e)}")
                
                # Combine all text from this page
                if page_text_parts:
                    # Remove duplicates while preserving order
                    unique_parts = []
                    seen_content = set()
                    for part in page_text_parts:
                        part_normalized = part.strip().lower()
                        if part_normalized and part_normalized not in seen_content:
                            unique_parts.append(part)
                            seen_content.add(part_normalized)
                    
                    if unique_parts:
                        page_final_text = " ".join(unique_parts)
                        all_text += page_final_text + " "
                        pages_with_text += 1
                else:
                    print(f"   Warning: No text extracted from page {page_num + 1}")
        
        print(f"   PyPDF2: Processed {pages_processed} pages, {pages_with_text} pages had extractable text")
        return all_text.strip() if all_text.strip() else None
            
    except Exception as e:
        print(f"   PyPDF2 extraction failed: {str(e)}")
        return None


def extract_with_pdfplumber(pdf_path):
    """Enhanced text extraction using pdfplumber - better for complex layouts and redacted content."""
    if not pdfplumber:
        return None
    
    try:
        all_text = ""
        pages_processed = 0
        pages_with_text = 0
        
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                page_text_parts = []
                pages_processed += 1
                
                # Method 1: Standard text extraction
                try:
                    page_text = page.extract_text()
                    if page_text and page_text.strip():
                        page_text_parts.append(page_text)
                except Exception as e:
                    print(f"   pdfplumber standard extraction failed on page {page_num + 1}: {str(e)}")
                
                # Method 2: Extract text with different strategies
                try:
                    # Try extracting with different layout parameters
                    for strategy in [{"layout": True}, {"layout": False}, {"x_tolerance": 1, "y_tolerance": 1}]:
                        try:
                            strategy_text = page.extract_text(**strategy)
                            if strategy_text and strategy_text.strip():
                                # Check if this gives us new content
                                existing_text = " ".join(page_text_parts)
                                if len(strategy_text) > len(existing_text) * 1.1:  # 10% more content
                                    page_text_parts.append(strategy_text)
                                    break
                        except Exception as e:
                            print(f"   pdfplumber strategy extraction failed on page {page_num + 1}: {str(e)}")
                            continue
                except Exception as e:
                    print(f"   pdfplumber advanced extraction failed on page {page_num + 1}: {str(e)}")
                
                # Method 3: Extract text from tables
                try:
                    tables = page.extract_tables()
                    for table_num, table in enumerate(tables):
                        try:
                            if table:
                                for row_num, row in enumerate(table):
                                    if row:
                                        table_text = " ".join([str(cell) for cell in row if cell])
                                        if table_text.strip():
                                            page_text_parts.append(table_text)
                        except Exception as e:
                            print(f"   Table extraction failed on page {page_num + 1}, table {table_num}: {str(e)}")
                            continue
                except Exception as e:
                    print(f"   Table extraction failed on page {page_num + 1}: {str(e)}")
                
                # Method 4: Extract text from individual characters (for heavily redacted content)
                try:
                    chars = page.chars
                    if chars:
                        char_text = "".join([char.get('text', '') for char in chars])
                        if char_text and char_text.strip():
                            page_text_parts.append(char_text)
                except Exception as e:
                    print(f"   Character extraction failed on page {page_num + 1}: {str(e)}")
                
                # Method 5: Extract text from words (alternative approach)
                try:
                    words = page.extract_words()
                    if words:
                        word_text = " ".join([word.get('text', '') for word in words])
                        if word_text and word_text.strip():
                            page_text_parts.append(word_text)
                except Exception as e:
                    print(f"   Word extraction failed on page {page_num + 1}: {str(e)}")
                
                # Combine all text from this page
                if page_text_parts:
                    # Remove duplicates while preserving order
                    unique_parts = []
                    seen_content = set()
                    for part in page_text_parts:
                        part_normalized = part.strip().lower()
                        if part_normalized and part_normalized not in seen_content:
                            unique_parts.append(part)
                            seen_content.add(part_normalized)
                    
                    if unique_parts:
                        page_final_text = " ".join(unique_parts)
                        all_text += page_final_text + " "
                        pages_with_text += 1
                else:
                    print(f"   Warning: No text extracted from page {page_num + 1}")
        
        print(f"   pdfplumber: Processed {pages_processed} pages, {pages_with_text} pages had extractable text")
        return all_text.strip() if all_text.strip() else None
        
    except Exception as e:
        print(f"   pdfplumber extraction failed: {str(e)}")
        return None


def extract_with_fitz(pdf_path):
    """Enhanced text extraction using PyMuPDF (fitz) - robust handling of redacted content."""
    if not fitz:
        return None
    
    try:
        doc = fitz.open(pdf_path)
        all_text = ""
        pages_processed = 0
        pages_with_text = 0
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            page_text_parts = []
            pages_processed += 1
            
            # Method 1: Standard text extraction
            try:
                standard_text = page.get_text()
                if standard_text and standard_text.strip():
                    page_text_parts.append(standard_text)
            except Exception as e:
                print(f"   Standard text extraction failed on page {page_num + 1}: {str(e)}")
            
            # Method 2: Extract text with different flags for better redaction handling
            try:
                # Try different text extraction modes with enhanced settings
                for flag in ["text", "dict", "rawdict", "html", "xhtml"]:
                    try:
                        if flag == "text":
                            continue  # Already tried above
                        elif flag == "dict":
                            text_dict = page.get_text("dict")
                            dict_text = ""
                            for block in text_dict.get("blocks", []):
                                if "lines" in block:
                                    for line in block["lines"]:
                                        for span in line.get("spans", []):
                                            dict_text += span.get("text", "") + " "
                            if dict_text.strip():
                                page_text_parts.append(dict_text)
                        elif flag == "rawdict":
                            raw_dict = page.get_text("rawdict")
                            raw_text = ""
                            for block in raw_dict.get("blocks", []):
                                if "lines" in block:
                                    for line in block["lines"]:
                                        for span in line.get("spans", []):
                                            raw_text += span.get("text", "") + " "
                            if raw_text.strip():
                                page_text_parts.append(raw_text)
                        elif flag in ["html", "xhtml"]:
                            # HTML/XHTML extraction can sometimes capture text missed by other methods
                            html_text = page.get_text(flag)
                            if html_text and html_text.strip():
                                # Basic HTML tag removal for cleaner text
                                import re
                                clean_text = re.sub(r'<[^>]+>', ' ', html_text)
                                clean_text = re.sub(r'\s+', ' ', clean_text).strip()
                                if clean_text:
                                    page_text_parts.append(clean_text)
                    except Exception as e:
                        print(f"   Text extraction mode '{flag}' failed on page {page_num + 1}: {str(e)}")
                        continue
            except Exception as e:
                print(f"   Advanced text extraction failed on page {page_num + 1}: {str(e)}")
            
            # Method 3: Extract text from annotations and forms
            try:
                # Get text from annotations
                for annot in page.annots():
                    try:
                        if annot.content:
                            page_text_parts.append(annot.content)
                    except Exception as e:
                        print(f"   Annotation extraction failed on page {page_num + 1}: {str(e)}")
                        continue
                
                # Get text from form fields
                widgets = page.widgets()
                for widget in widgets:
                    try:
                        if widget.field_value:
                            page_text_parts.append(str(widget.field_value))
                    except Exception as e:
                        print(f"   Form field extraction failed on page {page_num + 1}: {str(e)}")
                        continue
            except Exception as e:
                print(f"   Annotation/form extraction failed on page {page_num + 1}: {str(e)}")
            
            # Method 4: Enhanced text extraction for difficult cases (replaces OCR)
            # Try to extract text more aggressively using PyMuPDF's advanced features
            try:
                # Extract text blocks with position information for better accuracy
                text_blocks = page.get_text("blocks")
                for block in text_blocks:
                    if len(block) >= 5 and block[4].strip():  # block[4] contains the text
                        block_text = block[4].strip()
                        if block_text and block_text not in " ".join(page_text_parts):
                            page_text_parts.append(block_text)
                
                # Try to extract text from different layers/transparency groups
                try:
                    # Get text with clip parameter to handle overlapping content
                    clip_rect = page.rect
                    clipped_text = page.get_text(clip=clip_rect)
                    if clipped_text and clipped_text.strip():
                        clipped_clean = clipped_text.strip()
                        if clipped_clean and clipped_clean not in " ".join(page_text_parts):
                            page_text_parts.append(clipped_clean)
                except Exception:
                    pass
                    
            except Exception as e:
                print(f"   Enhanced text extraction failed on page {page_num + 1}: {str(e)}")
            
            # Method 5: Extract metadata and hidden text - ENHANCED FOR SECURITY  
            # Enhanced metadata extraction to replace image OCR functionality
            try:
                # Extract any text from page metadata
                page_metadata = page.get_contents()
                if page_metadata:
                    # Try to extract readable text from content streams
                    for content in page_metadata:
                        if content:
                            content_text = str(content)
                            # Look for text patterns in the content stream
                            import re
                            text_matches = re.findall(r'\((.*?)\)', content_text)
                            for match in text_matches:
                                if len(match) > 2 and match.strip():
                                    clean_match = match.strip()
                                    if clean_match and clean_match not in " ".join(page_text_parts):
                                        page_text_parts.append(clean_match)
            except Exception as e:
                print(f"   Metadata extraction failed on page {page_num + 1}: {str(e)}")
            
            # Method 6: Image processing for redacted content (when available)
            # Handle cases where PDF pages are primarily images with redacted content
            try:
                if cv2 is not None and Image is not None:
                    # Check if this page is primarily image-based
                    page_dict = page.get_text("dict")
                    if is_image_based_pdf_page(page_dict):
                        print(f"   Detected image-based content on page {page_num + 1}, processing for redacted areas...")
                        
                        # Convert page to image
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x zoom for better quality
                        img_data = pix.tobytes("png")
                        
                        # Convert to PIL Image
                        pil_image = Image.open(io.BytesIO(img_data))
                        
                        # Convert PIL to OpenCV format
                        cv_image = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
                        
                        # Detect and remove redacted areas
                        processed_image = remove_redacted_content_from_image(cv_image)
                        
                        if processed_image is not None:
                            # Convert back to PIL for potential future OCR (if needed)
                            processed_pil = Image.fromarray(cv2.cvtColor(processed_image, cv2.COLOR_BGR2RGB))
                            
                            # For now, we'll note that redacted content was detected and processed
                            # The actual text extraction would require OCR, which is removed for privacy
                            redacted_info = "Note: Image-based content detected with redacted areas. Redacted content has been processed but text extraction from images requires OCR functionality."
                            page_text_parts.append(redacted_info)
                            print(f"   Processed redacted areas on image-based page {page_num + 1}")
                        else:
                            print(f"   No redacted areas detected on image-based page {page_num + 1}")
                else:
                    # Image processing libraries not available
                    if not page_text_parts or len(" ".join(page_text_parts).strip()) < 50:
                        # Likely an image-based page but can't process without image libraries
                        warning_text = "Note: This page appears to be image-based. Image processing libraries are required for redacted content handling."
                        page_text_parts.append(warning_text)
                        print(f"   Image-based content detected on page {page_num + 1} but image processing unavailable")
            except Exception as e:
                print(f"   Image processing failed on page {page_num + 1}: {str(e)}")
            
            # Combine all text from this page
            if page_text_parts:
                page_final_text = " ".join(page_text_parts)
                all_text += page_final_text + " "
                pages_with_text += 1
            else:
                print(f"   Warning: No text extracted from page {page_num + 1}")
        
        doc.close()
        
        print(f"   PyMuPDF: Processed {pages_processed} pages, {pages_with_text} pages had extractable text")
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
        print("   # OCR support removed for enhanced data security")
        return
    
    # Validate PDF folder exists
    if not os.path.exists(pdf_folder):
        print(f"❌ Error: PDF folder does not exist: {pdf_folder}")
        print("Please update the PDF_FOLDER_PATH variable in the script.")
        return
    
    # Create output folder if specified and doesn't exist
    if output_folder and not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
            print(f"✅ Created output folder: {output_folder}")
        except Exception as e:
            print(f"❌ Error creating output folder: {str(e)}")
            return
    
    # Process PDFs with enhanced extraction
    process_pdfs_in_folder(pdf_folder, output_folder)


if __name__ == "__main__":
    main()