# Excel to News Directory Converter

A Python script that reads Excel files (.xlsx) and converts table data to a structured news directory format with HTML files. The script is designed to work with tables that have columns for Date, Source, Title, Link, and Description.

## Features

- Reads Excel files (.xlsx format)
- Converts table data to HTML news articles
- Creates organized directory structure with date-based file naming
- Generates configuration files for easy navigation
- Handles date formatting and content structuring
- Supports metadata including source and link information
- Provides detailed conversion summaries
- Works with the complete 5-column structure (Date, Source, Title, Link, Description)

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage
```bash
python excel_to_news_directory.py your_file.xlsx output_directory
```

### Example
```bash
python excel_to_news_directory.py sheets/robinhood.xlsx robinhood-news-output
```

### Using as a Module
```python
from excel_to_news_directory import excel_to_news_directory

# Convert Excel to news directory
result = excel_to_news_directory('your_file.xlsx', 'output_directory')
```

## Input Excel File Structure

The script expects an Excel file with the following columns:

| Date       | Source        | Title              | Link                    | Description                    |
|------------|---------------|-------------------|-------------------------|--------------------------------|
| 2024-01-15 | Benzinga      | Project Launch    | https://example.com     | Successfully launched project  |
| 2024-01-14 | Yahoo Finance | Team Meeting      | https://example2.com    | Weekly team meeting            |

## Output Directory Structure

```
output_directory/
├── 2024-01-15-01.html
├── 2024-01-14-01.html
├── article-config.js
└── conversion-summary.json
```

### HTML File Structure
Each HTML file contains:
```html
<h2 class="text-32 mb-4 font-700 elite-bold">[Article Title]</h2>
<div class="article-meta">
  <p class="source"><strong>Source:</strong> [Source Name]</p>
  <p class="link"><strong>Link:</strong> <a href="[URL]" target="_blank">[URL]</a></p>
</div>
<div class="description">
  [Article content with proper formatting]
</div>
```

### Configuration Files

**article-config.js**: Contains a list of all HTML files in chronological order
**conversion-summary.json**: Contains conversion statistics and metadata

## Example Output

### Conversion Summary
```json
{
  "source_file": "sheets/robinhood.xlsx",
  "export_date": "2025-07-25 13:45:30",
  "total_records": 2835,
  "date_range": {
    "start": "2024-05-13",
    "end": "2025-07-25"
  },
  "files_created": 2835
}
```

## Requirements

- Python 3.6+
- pandas
- openpyxl
- xlrd

## Project Structure

```
dataentryhelp/
├── excel_to_news_directory.py    # Main conversion script
├── requirements.txt              # Python dependencies
├── README.md                     # This file
├── .gitignore                    # Git ignore rules
├── sheets/                       # Input Excel files directory
│   └── robinhood.xlsx
├── robinhood-news-output/        # Generated news directory
│   ├── *.html                    # Individual news articles
│   ├── article-config.js         # File listing
│   └── conversion-summary.json   # Conversion statistics
└── venv/                         # Virtual environment
```

## Notes

- The script expects the first row to contain column headers
- Date columns are automatically formatted as YYYY-MM-DD
- Files are named using the pattern: YYYY-MM-DD-XX.html
- If the expected columns are missing, the script will show a warning but continue processing
- The script creates a complete news directory structure suitable for web publishing
- All HTML files include proper metadata and formatting 