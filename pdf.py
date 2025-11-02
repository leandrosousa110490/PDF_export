import os
import fitz  # PyMuPDF
from pathlib import Path
import pytesseract
from PIL import Image
import io

# ===== CONFIGURATION: Set your paths here =====
PDF_INPUT_PATH = "pdfs"  # Folder containing your PDF files
TEXT_OUTPUT_PATH = "text_files"  # Folder where text files will be saved
USE_OCR = True  # Set to True to extract text from image-based pages
# ==============================================

def extract_pdf_to_text(pdf_path, output_folder, use_ocr=True):
    """
    Extract all text from a PDF and save it as a text file.
    Uses OCR for pages that are images or have no extractable text.
    
    Args:
        pdf_path: Path to the PDF file
        output_folder: Folder where text files will be saved
        use_ocr: Whether to use OCR on image-based pages
    """
    try:
        # Open the PDF
        pdf_document = fitz.open(pdf_path)
        
        # Extract text from all pages
        full_text = ""
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            
            # Try to extract text normally first
            page_text = page.get_text()
            
            # If no text found and OCR is enabled, try OCR
            if use_ocr and len(page_text.strip()) < 50:  # Less than 50 chars suggests image
                print(f"  Page {page_num + 1}: No text found, attempting OCR...")
                try:
                    # Convert page to image
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x zoom for better OCR
                    img_data = pix.tobytes("png")
                    img = Image.open(io.BytesIO(img_data))
                    
                    # Perform OCR
                    ocr_text = pytesseract.image_to_string(img)
                    
                    if ocr_text.strip():
                        page_text = f"[OCR EXTRACTED TEXT]\n{ocr_text}"
                    else:
                        page_text = "[No text could be extracted from this page]"
                        
                except Exception as ocr_error:
                    page_text = f"[OCR failed: {str(ocr_error)}]"
            
            full_text += page_text
            full_text += f"\n\n--- Page {page_num + 1} ---\n\n"
        
        pdf_document.close()
        
        # Create output filename (same name but .txt extension)
        pdf_filename = os.path.basename(pdf_path)
        txt_filename = os.path.splitext(pdf_filename)[0] + ".txt"
        output_path = os.path.join(output_folder, txt_filename)
        
        # Save the extracted text
        with open(output_path, 'w', encoding='utf-8') as txt_file:
            txt_file.write(full_text)
        
        print(f"✓ Converted: {pdf_filename} -> {txt_filename}")
        return True
        
    except Exception as e:
        print(f"✗ Error processing {pdf_path}: {str(e)}")
        return False

def process_pdf_folder(input_folder, output_folder, use_ocr=True):
    """
    Process all PDFs in a folder and convert them to text files.
    
    Args:
        input_folder: Folder containing PDF files
        output_folder: Folder where text files will be saved
        use_ocr: Whether to use OCR on image-based pages
    """
    # Create output folder if it doesn't exist
    Path(output_folder).mkdir(parents=True, exist_ok=True)
    
    # Get all PDF files in the input folder
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"No PDF files found in {input_folder}")
        return
    
    print(f"Found {len(pdf_files)} PDF file(s) to process")
    if use_ocr:
        print("OCR is ENABLED - will extract text from image-based pages\n")
    else:
        print("OCR is DISABLED - only extracting native text\n")
    
    # Process each PDF
    successful = 0
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file)
        if extract_pdf_to_text(pdf_path, output_folder, use_ocr):
            successful += 1
    
    print(f"\n{'='*50}")
    print(f"Processing complete: {successful}/{len(pdf_files)} files converted successfully")
    print(f"Text files saved to: {output_folder}")

# Main execution
if __name__ == "__main__":
    process_pdf_folder(PDF_INPUT_PATH, TEXT_OUTPUT_PATH, USE_OCR)
