# Excel to JSON Converter

A Python script that reads Excel files (.xlsx) and converts table data to JSON format. The script is designed to work with tables that have columns for Date, Title, and Description.

## Features

- Reads Excel files (.xlsx format)
- Converts table data to structured JSON
- Handles date formatting
- Supports custom output file names
- Provides detailed error messages
- Works with any number of columns (not just Date, Title, Description)

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage
```bash
python excel_to_json.py your_file.xlsx
```

### Specify Output File
```bash
python excel_to_json.py your_file.xlsx output.json
```

### Using as a Module
```python
from excel_to_json import excel_to_json

# Convert Excel to JSON
result = excel_to_json('your_file.xlsx', 'output.json')
```

## Example

### Input Excel File Structure
| Date       | Title              | Description                    |
|------------|-------------------|--------------------------------|
| 2024-01-15 | Project Launch    | Successfully launched project  |
| 2024-01-14 | Team Meeting      | Weekly team meeting            |

### Output JSON Structure
```json
{
  "source_file": "your_file.xlsx",
  "export_date": "2024-01-15 10:30:00",
  "total_records": 2,
  "data": [
    {
      "Date": "2024-01-15",
      "Title": "Project Launch",
      "Description": "Successfully launched project"
    },
    {
      "Date": "2024-01-14",
      "Title": "Team Meeting",
      "Description": "Weekly team meeting"
    }
  ]
}
```

## Testing

1. Create a sample Excel file:
```bash
python create_sample_excel.py
```

2. Convert the sample file to JSON:
```bash
python excel_to_json.py sample_data.xlsx
```

## Requirements

- Python 3.6+
- pandas
- openpyxl
- xlrd

## Notes

- The script expects the first row to contain column headers
- Date columns are automatically formatted as YYYY-MM-DD
- Empty cells are converted to `null` in JSON
- The script will work with any column structure, not just Date/Title/Description
- If the expected columns are missing, the script will show a warning but continue processing 