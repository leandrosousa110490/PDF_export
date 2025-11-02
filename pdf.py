import os
import fitz  # PyMuPDF
from pathlib import Path

# ===== CONFIGURATION: Set your paths here =====
PDF_INPUT_PATH = "pdfs"  # Folder containing your PDF files
TEXT_OUTPUT_PATH = "text_files"  # Folder where text files will be saved
# ==============================================

def extract_pdf_to_text(pdf_path, output_folder):
    """
    Extract all text from a PDF and save it as a text file.
    
    Args:
        pdf_path: Path to the PDF file
        output_folder: Folder where text files will be saved
    """
    try:
        # Open the PDF
        pdf_document = fitz.open(pdf_path)
        
        # Extract text from all pages
        full_text = ""
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            full_text += page.get_text()
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

def process_pdf_folder(input_folder, output_folder):
    """
    Process all PDFs in a folder and convert them to text files.
    
    Args:
        input_folder: Folder containing PDF files
        output_folder: Folder where text files will be saved
    """
    # Create output folder if it doesn't exist
    Path(output_folder).mkdir(parents=True, exist_ok=True)
    
    # Get all PDF files in the input folder
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"No PDF files found in {input_folder}")
        return
    
    print(f"Found {len(pdf_files)} PDF file(s) to process\n")
    
    # Process each PDF
    successful = 0
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file)
        if extract_pdf_to_text(pdf_path, output_folder):
            successful += 1
    
    print(f"\n{'='*50}")
    print(f"Processing complete: {successful}/{len(pdf_files)} files converted successfully")
    print(f"Text files saved to: {output_folder}")

# Main execution
if __name__ == "__main__":
    process_pdf_folder(PDF_INPUT_PATH, TEXT_OUTPUT_PATH)
