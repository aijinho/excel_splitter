import os
import pandas as pd
from dotenv import load_dotenv
import json

def load_environment_variables():
    """Load environment variables from .env file"""
    load_dotenv('.env')
    
    source_folder = os.getenv('SOURCE_FOLDER', './source_files')
    output_folder = os.getenv('OUTPUT_FOLDER', './output')
    split_rows_max = int(os.getenv('SPLIT_ROWS_MAX', '200000'))
    split_suffix = os.getenv('SPLIT_SUFFIX', 'part')
    
    return source_folder, output_folder, split_rows_max, split_suffix

def find_excel_file(source_folder):
    """Find the first Excel file in the source folder"""
    excel_extensions = ['.xlsx', '.xls']
    
    for file in os.listdir(source_folder):
        if any(file.lower().endswith(ext) for ext in excel_extensions):
            return os.path.join(source_folder, file)
    
    return None

def create_output_folder(output_folder, excel_filename):
    """Create output folder with the same name as the Excel file (without extension)"""
    # Get filename without extension
    base_name = os.path.splitext(os.path.basename(excel_filename))[0]
    output_path = os.path.join(output_folder, base_name)
    
    # Create folder if it doesn't exist
    if not os.path.exists(output_path):
        os.makedirs(output_path, exist_ok=True)
        print(f"Created output folder: {output_path}")
    else:
        print(f"Output folder already exists, skipping creation: {output_path}")
    
    return output_path

def split_dataframe_to_csv(df, output_path, base_name, split_suffix, split_rows_max):
    """Split DataFrame into smaller CSV files"""
    total_rows = len(df)
    num_splits = (total_rows + split_rows_max - 1) // split_rows_max  # Ceiling division
    
    print(f"Splitting {total_rows} rows into {num_splits} files with max {split_rows_max} rows each")
    
    split_files = []
    for i in range(num_splits):
        start_idx = i * split_rows_max
        end_idx = min((i + 1) * split_rows_max, total_rows)
        
        # Extract the chunk
        chunk = df.iloc[start_idx:end_idx]
        
        # Create filename
        csv_filename = f"{base_name}_{split_suffix}{i+1}.csv"
        csv_path = os.path.join(output_path, csv_filename)
        
        # Save to CSV
        chunk.to_csv(csv_path, index=False)
        split_files.append(csv_path)
        
        print(f"Created: {csv_filename} (rows {start_idx+1}-{end_idx})")
    
    return split_files

def convert_to_json(df, output_path, base_name):
    """Convert DataFrame to JSON with column names as keys"""
    json_filename = f"{base_name}.json"
    json_path = os.path.join(output_path, json_filename)
    
    # Convert DataFrame to dictionary with column names as keys
    json_data = {}
    for column in df.columns:
        json_data[column] = df[column].tolist()
    
    # Save to JSON file
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, indent=2, ensure_ascii=False)
    
    print(f"Created: {json_filename}")
    return json_path

def process_excel_file():
    """Main function to process Excel file according to specifications"""
    print("Loading environment variables...")
    source_folder, output_folder, split_rows_max, split_suffix = load_environment_variables()
    
    print(f"Source folder: {source_folder}")
    print(f"Output folder: {output_folder}")
    print(f"Split rows max: {split_rows_max}")
    print(f"Split suffix: {split_suffix}")
    
    # Find Excel file
    print(f"\nLooking for Excel files in {source_folder}...")
    excel_file = find_excel_file(source_folder)
    
    if not excel_file:
        print("No Excel file found in the source folder!")
        return False
    
    print(f"Found Excel file: {excel_file}")
    
    # Load Excel file into DataFrame
    print("Loading Excel file into DataFrame...")
    try:
        df = pd.read_excel(excel_file)
        print(f"Loaded DataFrame with shape: {df.shape}")
        print(f"Columns: {list(df.columns)}")
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return False
    
    # Create output folder
    print(f"\nCreating output folder...")
    output_path = create_output_folder(output_folder, excel_file)
    
    # Get base filename without extension
    base_name = os.path.splitext(os.path.basename(excel_file))[0]
    
    # Split DataFrame into CSV files
    print(f"\nSplitting DataFrame into CSV files...")
    split_files = split_dataframe_to_csv(df, output_path, base_name, split_suffix, split_rows_max)
    
    # Convert DataFrame to JSON
    print(f"\nConverting DataFrame to JSON...")
    json_file = convert_to_json(df, output_path, base_name)
    
    # Summary
    print(f"\n=== Processing Complete ===")
    print(f"Original file: {excel_file}")
    print(f"DataFrame shape: {df.shape}")
    print(f"Output folder: {output_path}")
    print(f"CSV files created: {len(split_files)}")
    for file in split_files:
        print(f"  - {os.path.basename(file)}")
    print(f"JSON file created: {os.path.basename(json_file)}")
    
    return True

if __name__ == "__main__":
    process_excel_file()
