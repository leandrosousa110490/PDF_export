"""Enhanced PDF Value Extraction Runner
===================================

This script provides a comprehensive PDF value extraction system with:
- Multiple predefined extraction configurations
- Excel export functionality
- Automatic PDF file discovery
- Batch processing capabilities

Usage:
    python run_extraction.py

The script will:
1. Find all PDF files in the configured folder
2. Apply all extraction configurations
3. Export results to Excel with formatting
4. Display a summary of results
"""

import os
from datetime import datetime
from pdf_value_extractor import PDFValueExtractor
from config import PDF_FOLDER, EXTRACTION_CONFIGS, EXCEL_OUTPUT_FOLDER

def get_pdf_files(folder_path: str = None) -> list:
    """
    Get all PDF files from the specified folder or current directory.
    
    Args:
        folder_path: Path to folder containing PDFs (defaults to PDF_FOLDER from config)
    
    Returns:
        List of full paths to PDF files
    """
    if folder_path is None:
        folder_path = PDF_FOLDER
    
    pdf_files = []
    
    # Check if specified folder exists
    if os.path.exists(folder_path) and os.path.isdir(folder_path):
        search_path = folder_path
        print(f"Looking for PDF files in: {folder_path}")
    else:
        search_path = "."
        print(f"Folder '{folder_path}' not found. Looking in current directory.")
    
    # Get all PDF files
    for file in os.listdir(search_path):
        if file.lower().endswith('.pdf'):
            full_path = os.path.join(search_path, file)
            pdf_files.append(full_path)
    
    print(f"Found {len(pdf_files)} PDF files")
    return pdf_files

def run_comprehensive_extraction():
    """Run extraction with all predefined configurations"""
    print("=" * 70)
    print("COMPREHENSIVE PDF VALUE EXTRACTION")
    print("=" * 70)
    
    # Get all PDF files
    pdf_files = get_pdf_files()
    
    if not pdf_files:
        print("No PDF files found!")
        print("Please ensure PDF files are in the configured folder or current directory.")
        return
    
    print(f"\nProcessing {len(pdf_files)} PDF files with {len(EXTRACTION_CONFIGS)} configurations...")
    
    # Create extractor
    extractor = PDFValueExtractor()
    
    # Extract using all configurations
    results = extractor.extract_with_multiple_configs(
        pdf_paths=pdf_files,
        configs=EXTRACTION_CONFIGS
    )
    
    # Display results summary
    print("\n" + "=" * 70)
    print("EXTRACTION RESULTS SUMMARY")
    print("=" * 70)
    
    # Group results by configuration
    config_results = {}
    for result in results:
        config_name = result['config_name']
        if config_name not in config_results:
            config_results[config_name] = {'success': 0, 'no_match': 0, 'results': []}
        
        config_results[config_name]['results'].append(result)
        if result['status'] == 'Success':
            config_results[config_name]['success'] += 1
        else:
            config_results[config_name]['no_match'] += 1
    
    # Print summary for each configuration
    for config_name, data in config_results.items():
        print(f"\n--- {config_name} ---")
        print(f"Success: {data['success']}, No Match: {data['no_match']}")
        
        # Show successful extractions
        for result in data['results']:
            if result['status'] == 'Success':
                print(f"  ✓ {result['pdf_name']}: {result['extracted_value']}")
    
    # Export to Excel
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"comprehensive_extraction_{timestamp}"
    excel_file = extractor.export_to_excel(excel_filename)
    
    print("\n" + "=" * 70)
    print("EXPORT COMPLETE")
    print("=" * 70)
    print(f"Results exported to: {excel_file}")
    
    # Print detailed summary
    extractor.print_results_summary()
    
    return excel_file

def run_targeted_extraction(config_names: list = None):
    """Run extraction with specific configurations only"""
    if config_names is None:
        config_names = ['Invoice Total', 'Invoice Number']
    
    print("=" * 70)
    print("TARGETED PDF VALUE EXTRACTION")
    print("=" * 70)
    print(f"Using configurations: {', '.join(config_names)}")
    
    # Get PDF files
    pdf_files = get_pdf_files()
    
    if not pdf_files:
        print("No PDF files found!")
        return
    
    # Filter configurations
    selected_configs = {name: config for name, config in EXTRACTION_CONFIGS.items() 
                       if name in config_names}
    
    if not selected_configs:
        print(f"No matching configurations found for: {config_names}")
        return
    
    # Create extractor and run extraction
    extractor = PDFValueExtractor()
    results = extractor.extract_with_multiple_configs(pdf_files, selected_configs)
    
    # Display results
    print(f"\nResults for {len(pdf_files)} PDF files:")
    current_config = None
    for result in results:
        if result['config_name'] != current_config:
            current_config = result['config_name']
            print(f"\n--- {current_config} ---")
        
        status_symbol = "✓" if result['status'] == 'Success' else "✗"
        print(f"  {status_symbol} {result['pdf_name']}: {result['extracted_value']}")
    
    # Export to Excel
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"targeted_extraction_{timestamp}"
    excel_file = extractor.export_to_excel(excel_filename)
    
    print(f"\nResults exported to: {excel_file}")
    return excel_file

def interactive_extraction():
    """Interactive mode for custom extractions"""
    print("=" * 70)
    print("INTERACTIVE PDF VALUE EXTRACTION")
    print("=" * 70)
    
    from config import ExtractionConfig
    
    # Get PDF files
    pdf_files = get_pdf_files()
    
    if not pdf_files:
        print("No PDF files found!")
        return
    
    print("\nAvailable extraction modes:")
    print("1. after - Extract text that appears after search terms")
    print("2. before - Extract text that appears before search terms")
    print("3. between - Extract text that appears between two text markers")
    print("4. around - Extract text that appears around search terms")
    
    print("\nAvailable data types:")
    print("1. any - Accept any text")
    print("2. numbers - Only numeric values")
    print("3. letters - Only alphabetic text")
    print("4. both - Text containing both letters and numbers")
    
    # Get user input
    try:
        search_terms = input("\nEnter search terms (comma-separated): ").split(',')
        search_terms = [term.strip() for term in search_terms if term.strip()]
        
        if not search_terms:
            print("No search terms provided!")
            return
        
        mode = input("Enter extraction mode (after/before/between/around): ").strip().lower()
        if mode not in ['after', 'before', 'between', 'around']:
            mode = 'after'
            print(f"Invalid mode, using 'after'")
        
        # Get end text for between mode
        end_text = None
        if mode == 'between':
            end_text = input("Enter end text marker (what text should extraction stop at): ").strip()
            if not end_text:
                print("End text is required for 'between' mode!")
                return
        
        data_type = input("Enter expected data type (any/numbers/letters/both): ").strip().lower()
        if data_type not in ['any', 'numbers', 'letters', 'both']:
            data_type = 'any'
            print(f"Invalid data type, using 'any'")
        
        # Create custom configuration
        custom_config = ExtractionConfig(
            name="Interactive Extraction",
            search_terms=search_terms,
            extraction_mode=mode,
            expected_type=data_type,
            case_sensitive=False,
            max_distance=100,
            end_text=end_text
        )
        
        # Run extraction
        extractor = PDFValueExtractor(config=custom_config)
        results = extractor.extract_from_multiple_pdfs(pdf_files, "Interactive Extraction")
        
        # Display results
        print(f"\nResults for search terms: {', '.join(search_terms)}")
        print(f"Mode: {mode}, Data type: {data_type}")
        print("-" * 50)
        
        for result in results:
            status_symbol = "✓" if result['status'] == 'Success' else "✗"
            print(f"{status_symbol} {result['pdf_name']}: {result['extracted_value']}")
        
        # Export to Excel
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"interactive_extraction_{timestamp}"
        excel_file = extractor.export_to_excel(excel_filename)
        
        print(f"\nResults exported to: {excel_file}")
        
    except KeyboardInterrupt:
        print("\nInteractive extraction cancelled.")
    except Exception as e:
        print(f"Error during interactive extraction: {e}")

def main():
    """Main function with menu options"""
    print("Enhanced PDF Value Extraction System")
    print("=" * 70)
    
    while True:
        print("\nSelect extraction mode:")
        print("1. Comprehensive extraction (all configurations)")
        print("2. Targeted extraction (specific configurations)")
        print("3. Interactive extraction (custom configuration)")
        print("4. List available configurations")
        print("5. Exit")
        
        try:
            choice = input("\nEnter your choice (1-5): ").strip()
            
            if choice == '1':
                run_comprehensive_extraction()
            elif choice == '2':
                print("\nAvailable configurations:")
                for i, name in enumerate(EXTRACTION_CONFIGS.keys(), 1):
                    print(f"  {i}. {name}")
                
                config_input = input("\nEnter configuration names (comma-separated) or numbers: ").strip()
                if config_input:
                    # Parse input (could be names or numbers)
                    config_names = []
                    for item in config_input.split(','):
                        item = item.strip()
                        if item.isdigit():
                            idx = int(item) - 1
                            if 0 <= idx < len(EXTRACTION_CONFIGS):
                                config_names.append(list(EXTRACTION_CONFIGS.keys())[idx])
                        else:
                            if item in EXTRACTION_CONFIGS:
                                config_names.append(item)
                    
                    if config_names:
                        run_targeted_extraction(config_names)
                    else:
                        print("No valid configurations selected!")
            elif choice == '3':
                interactive_extraction()
            elif choice == '4':
                print("\nAvailable extraction configurations:")
                for name, config in EXTRACTION_CONFIGS.items():
                    print(f"\n{name}:")
                    print(f"  Search terms: {', '.join(config.search_terms)}")
                    print(f"  Mode: {config.extraction_mode}")
                    print(f"  Expected type: {config.expected_type}")
            elif choice == '5':
                print("Goodbye!")
                break
            else:
                print("Invalid choice! Please enter 1-5.")
                
        except KeyboardInterrupt:
            print("\nGoodbye!")
            break
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    main()