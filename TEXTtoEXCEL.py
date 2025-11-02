import json
import os
import re
import pandas as pd
from pathlib import Path

def load_config(config_file):
    """Load extraction configuration from JSON file"""
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config
    except Exception as e:
        print(f"Error loading config file: {e}")
        return None

def read_text_file(file_path):
    """Read content from a text file"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        return content
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return None

def get_value_pattern(value_type):
    """Generate regex pattern based on value type"""
    patterns = {
        'number': r'[\d$,.()\-\s%/]+',  # Numbers, currency, dates, phone numbers, percentages
        'text': r'[A-Za-z\s&.,\'-]+',   # Alphabetic text, names, companies
        'both': r'[A-Za-z0-9\s@._\-#&,.()\'\'/]+' # Alphanumeric combinations, emails, IDs
    }
    return patterns.get(value_type, r'.+?')  # Default to any character if type not recognized

def extract_value(text, rule):
    """Extract value from text using the specified rule"""
    try:
        before_text = rule.get('before_text', '')
        after_text = rule.get('after_text', '')
        value_type = rule.get('value_type', 'both')
        case_sensitive = rule.get('case_sensitive', False)
        
        # Set regex flags based on case sensitivity
        flags = 0 if case_sensitive else re.IGNORECASE
        
        # Get the appropriate pattern for the value type
        value_pattern = get_value_pattern(value_type)
        
        # Method 1: Use before_text and after_text as anchors
        if before_text and after_text:
            pattern = re.escape(before_text) + r'\s*(' + value_pattern + r')\s*' + re.escape(after_text)
            match = re.search(pattern, text, flags)
            if match:
                return match.group(1).strip()
        
        # Method 2: Use only before_text
        elif before_text:
            pattern = re.escape(before_text) + r'\s*(' + value_pattern + r')(?:\s|$|\n)'
            match = re.search(pattern, text, flags)
            if match:
                extracted = match.group(1).strip()
                # Clean up common trailing characters
                extracted = re.sub(r'[,.\s]+$', '', extracted)
                return extracted
        
        # Method 3: Use only value_type pattern to search entire document
        elif value_type:
            # This is less precise but can be used as fallback
            target_sentence = rule.get('target_sentence', '')
            if target_sentence:
                # Try to find similar patterns based on target sentence
                words = target_sentence.split()
                for word in words:
                    if len(word) > 3:  # Look for meaningful words
                        pattern = re.escape(word) + r'\s*[:\-]?\s*(' + value_pattern + r')'
                        match = re.search(pattern, text, flags)
                        if match:
                            return match.group(1).strip()
        
        # Method 4: Fallback - use target_sentence as reference
        target_sentence = rule.get('target_sentence', '')
        if target_sentence and before_text:
            # Try to find similar patterns in the text
            pattern = re.escape(before_text) + r'\s*(' + value_pattern + r')(?:\s|$|\n)'
            match = re.search(pattern, text, flags)
            if match:
                extracted = match.group(1).strip()
                # Clean up common trailing characters
                extracted = re.sub(r'[,.\s]+$', '', extracted)
                return extracted
        
        return None
        
    except Exception as e:
        print(f"Error extracting value with rule {rule.get('name', 'unknown')}: {e}")
        return None

def process_text_files(config):
    """Process all text files and extract values according to configuration"""
    settings = config.get('settings', {})
    text_folder = settings.get('text_files_folder', 'extracted_text_files')
    extraction_rules = config.get('extraction_rules', [])
    
    results = []
    
    # Get all text files from the specified folder
    text_files = []
    if os.path.exists(text_folder):
        for file in os.listdir(text_folder):
            if file.endswith('.txt'):
                text_files.append(os.path.join(text_folder, file))
    
    print(f"Found {len(text_files)} text files to process")
    
    # Process each text file
    for file_path in text_files:
        filename = Path(file_path).stem  # Get filename without extension
        print(f"Processing: {filename}")
        
        # Read the text file
        text_content = read_text_file(file_path)
        if not text_content:
            continue
        
        # Apply each extraction rule
        for rule in extraction_rules:
            rule_name = rule.get('name', 'unknown')
            extracted_value = extract_value(text_content, rule)
            
            # Always add a result for each rule, whether successful or not
            if extracted_value:
                results.append({
                    'Filename': filename,
                    'Config_Name': rule_name,
                    'Extracted_Value': extracted_value
                })
                print(f"  ✓ {rule_name}: {extracted_value}")
            else:
                results.append({
                    'Filename': filename,
                    'Config_Name': rule_name,
                    'Extracted_Value': 'Not Found'
                })
                print(f"  ✗ {rule_name}: Not Found")
    
    return results

def save_to_excel(results, output_file):
    """Save extraction results to Excel file"""
    try:
        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(output_file)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Create DataFrame and save to Excel
        df = pd.DataFrame(results)
        df.to_excel(output_file, index=False)
        
        # Print detailed summary
        print(f"\n📊 Extraction Summary:")
        print(f"Total files processed: {len(df['Filename'].unique())}")
        print(f"Total extraction rules: {len(df['Config_Name'].unique())}")
        print(f"Total extraction attempts: {len(df)}")
        
        # Count successful vs failed extractions
        successful = len(df[df['Extracted_Value'] != 'Not Found'])
        failed = len(df[df['Extracted_Value'] == 'Not Found'])
        print(f"Successful extractions: {successful}")
        print(f"Failed extractions: {failed}")
        
        # Show breakdown by config name
        print(f"\nBreakdown by extraction rule:")
        for config_name in df['Config_Name'].unique():
            config_data = df[df['Config_Name'] == config_name]
            successful_count = len(config_data[config_data['Extracted_Value'] != 'Not Found'])
            total_count = len(config_data)
            print(f"  {config_name}: {successful_count}/{total_count} successful")
        
        print(f"\n✅ Results saved to: {output_file}")
        
    except Exception as e:
        print(f"Error saving to Excel: {e}")

def main():
    """Main function to run the value extraction process"""
    config_file = 'extraction_config.json'
    
    # Load configuration
    config = load_config(config_file)
    if not config:
        print("Failed to load configuration. Exiting.")
        return
    
    print("🔍 Starting value extraction process...")
    print(f"Configuration loaded from: {config_file}")
    
    # Process text files and extract values
    results = process_text_files(config)
    
    if results:
        # Save results to Excel
        output_file = config.get('settings', {}).get('output_excel_file', 'extracted_values.xlsx')
        save_to_excel(results, output_file)
    else:
        print("No values were extracted. Please check your configuration and text files.")

if __name__ == "__main__":
    main()
