import fitz  # PyMuPDF
import pdfplumber
import os
import pandas as pd

def extract_text_and_tables_from_pdf(pdf_path):
    """Extract text and tables from a PDF file"""
    results = {
        'text_content': [],
        'table_content': []
    }
    
    # Extract text using PyMuPDF (faster)
    try:
        pdf_document = fitz.open(pdf_path)
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            text = page.get_text()
            # Replace newlines with spaces and clean up extra spaces
            text_single_line = ' '.join(text.split())
            if text_single_line.strip():
                results['text_content'].append(f"[Page {page_num + 1}] {text_single_line}")
        pdf_document.close()
    except Exception as e:
        results['text_content'].append(f"[Error extracting text] {str(e)}")
    
    # Extract tables using pdfplumber
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for table_num, table in enumerate(tables):
                    if table:
                        # Convert table to string format
                        table_str = f"[Table {table_num + 1} on Page {page_num + 1}] "
                        for row in table:
                            if row:
                                # Join cells with spaces, handle None values
                                row_str = " ".join([str(cell) if cell is not None else "" for cell in row])
                                table_str += row_str + " | "
                        results['table_content'].append(table_str.rstrip(" | "))
    except Exception as e:
        results['table_content'].append(f"[Error extracting tables] {str(e)}")
    
    return results

def main():
    # ===== CONFIGURATION VARIABLES =====
    # Change these paths as needed:
    pdf_source_folder = "folder"  # Folder containing the PDF files to process
    export_destination_folder = "extracted_text_files"  # Folder where text files will be saved
    
    # Create export destination folder if it doesn't exist
    if not os.path.exists(export_destination_folder):
        os.makedirs(export_destination_folder)
        print(f"Created export folder: {export_destination_folder}")
    
    # Check if PDF source folder exists
    if not os.path.exists(pdf_source_folder):
        print(f"Error: PDF source folder '{pdf_source_folder}' not found!")
        return
    
    # Get all PDF files
    pdf_files = [f for f in os.listdir(pdf_source_folder) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"No PDF files found in '{pdf_source_folder}' folder!")
        return
    
    print(f"Found {len(pdf_files)} PDF files to process...")
    print(f"Source folder: {pdf_source_folder}")
    print(f"Export folder: {export_destination_folder}")
    print()
    
    # Process each PDF and create individual text files
    for pdf_file in sorted(pdf_files):
        pdf_path = os.path.join(pdf_source_folder, pdf_file)
        print(f"Processing: {pdf_file}")
        
        # Create output filename (same as PDF but with .txt extension)
        output_filename = os.path.splitext(pdf_file)[0] + ".txt"
        output_path = os.path.join(export_destination_folder, output_filename)
        
        # Extract content
        content = extract_text_and_tables_from_pdf(pdf_path)
        
        # Write to individual file in export folder
        with open(output_path, 'w', encoding='utf-8') as output:
            output.write("=" * 80 + "\n")
            output.write(f"FILE: {pdf_file}\n")
            output.write("=" * 80 + "\n\n")
            
            # Write text content
            output.write("--- TEXT CONTENT ---\n")
            if content['text_content']:
                for text_line in content['text_content']:
                    output.write(text_line + "\n")
            else:
                output.write("[No text content found]\n")
            
            output.write("\n--- TABLE CONTENT ---\n")
            if content['table_content']:
                for table_line in content['table_content']:
                    output.write(table_line + "\n")
            else:
                output.write("[No tables found]\n")
            
            output.write("\n")
        
        print(f"  → Created: {output_path}")
    
    print(f"\nExtraction complete! Created {len(pdf_files)} individual text files.")
    print(f"All text files saved in: {export_destination_folder}")
    print("Each text file has the same name as its corresponding PDF.")

if __name__ == "__main__":
    main()
