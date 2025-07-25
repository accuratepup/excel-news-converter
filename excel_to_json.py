import pandas as pd
import json
import sys
import os
from datetime import datetime

def excel_to_json(excel_file_path, output_file_path=None):
    """
    Read Excel file and convert table data to JSON format.
    
    Args:
        excel_file_path (str): Path to the Excel file
        output_file_path (str, optional): Path for the output JSON file. 
                                        If None, will use the same name as Excel file
    
    Returns:
        dict: JSON object containing the table data
    """
    try:
        # Read the Excel file
        print(f"Reading Excel file: {excel_file_path}")
        df = pd.read_excel(excel_file_path)
        
        # Check if required columns exist
        required_columns = ['Date', 'Title', 'Description']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"Warning: Missing columns: {missing_columns}")
            print(f"Available columns: {list(df.columns)}")
            print("Proceeding with available columns...")
        
        # Convert DataFrame to list of dictionaries
        data = []
        for index, row in df.iterrows():
            row_dict = {}
            for column in df.columns:
                # Handle date conversion
                if column == 'Date' and pd.notna(row[column]):
                    if isinstance(row[column], datetime):
                        row_dict[column] = row[column].strftime('%Y-%m-%d')
                    else:
                        row_dict[column] = str(row[column])
                else:
                    # Handle NaN values
                    if pd.isna(row[column]):
                        row_dict[column] = None
                    else:
                        row_dict[column] = str(row[column])
            data.append(row_dict)
        
        # Create JSON object
        json_data = {
            "source_file": excel_file_path,
            "export_date": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "total_records": len(data),
            "data": data
        }
        
        # Determine output file path
        if output_file_path is None:
            base_name = os.path.splitext(excel_file_path)[0]
            output_file_path = f"{base_name}.json"
        
        # Write JSON to file
        with open(output_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(json_data, json_file, indent=2, ensure_ascii=False)
        
        print(f"Successfully exported {len(data)} records to: {output_file_path}")
        
        return json_data
        
    except FileNotFoundError:
        print(f"Error: File '{excel_file_path}' not found.")
        return None
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        return None

def main():
    """Main function to handle command line arguments."""
    if len(sys.argv) < 2:
        print("Usage: python excel_to_json.py <excel_file_path> [output_json_path]")
        print("Example: python excel_to_json.py data.xlsx")
        print("Example: python excel_to_json.py data.xlsx output.json")
        return
    
    excel_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Convert Excel to JSON
    result = excel_to_json(excel_file, output_file)
    
    if result:
        print("\nPreview of exported data:")
        print(json.dumps(result, indent=2)[:500] + "..." if len(json.dumps(result)) > 500 else json.dumps(result, indent=2))

if __name__ == "__main__":
    main() 