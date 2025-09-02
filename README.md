# Excel to News Directory Converter

A Python script that reads Excel files (.xlsx) and converts table data to a structured news directory format with HTML files. The script is designed to work with tables that have columns for Date, Source, Title, Link, and Description.

**NEW: Smart Article Selection** - The script now intelligently selects only the most important and recent news articles (default: 12) based on source credibility, importance keywords, and recency.

## Features

- Reads Excel files (.xlsx format)
- **Smart article selection**: Automatically selects the most important and recent news articles
- **Configurable limit**: Choose how many articles to extract (default: 12)
- **Intelligent prioritization**: Ranks articles by importance score and recency
- Converts table data to HTML news articles
- Creates organized directory structure with date-based file naming
- Generates configuration files for easy navigation
- Handles date formatting and content structuring
- Supports metadata including source and link information
- Provides detailed conversion summaries
- Works with the complete 5-column structure (Date, Source, Title, Link, Description)

## Article Selection Algorithm

The script uses a sophisticated scoring system to select the most important articles:

1. **Recency Score**: Newer articles get higher scores (up to 30 points)
2. **Source Credibility**: Articles from reputable sources get +20 points
   - Reuters, Bloomberg, CNBC, Wall Street Journal, Financial Times, Yahoo Finance, MarketWatch, Seeking Alpha, Benzinga, Investing.com
3. **Importance Keywords**: Articles with keywords like "BREAKING", "EXCLUSIVE", "UPDATE", "ALERT" get +10 points each
4. **Content Quality**: Longer, more detailed descriptions get bonus points

Articles are then sorted by total score and recency, with only the top N articles selected.

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage (extracts 12 most important articles)
```bash
python excel_to_news_directory.py your_file.xlsx output_directory
```

### Specify Output Directory
```bash
python excel_to_news_directory.py your_file.xlsx robinhood-news-output
```

### Specify Number of Articles
```bash
python excel_to_news_directory.py your_file.xlsx output_directory 20
```

### Examples
```bash
# Extract 12 most important articles (default)
python excel_to_news_directory.py sheets/robinhood.xlsx robinhood-news-output

# Extract only 5 most important articles
python excel_to_news_directory.py sheets/robinhood.xlsx robinhood-news-output 5

# Extract 15 most important articles
python excel_to_news_directory.py sheets/robinhood.xlsx robinhood-news-output 15
```

### Using as a Module
```python
from excel_to_news_directory import excel_to_news_directory

# Convert Excel to news directory (12 articles)
result = excel_to_news_directory('your_file.xlsx', 'output_directory')

# Convert Excel to news directory (custom number of articles)
result = excel_to_news_directory('your_file.xlsx', 'output_directory', max_articles=8)
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
  "total_articles_processed": 2835,
  "articles_selected": 12,
  "max_articles_limit": 12,
  "date_range": {
    "start": "2024-05-13",
    "end": "2025-07-25"
  },
  "files_created": 12,
  "selection_criteria": "Most recent and most important news based on source credibility, keywords, and content quality"
}
```

### Console Output Example
```
🎯 Extracting up to 12 most important and recent news articles...
Reading Excel file: sheets/robinhood.xlsx
Creating news directory: robinhood-news-output
Selected 12 articles out of 2835 total articles
Articles selected based on importance score and recency
Created: 2025-07-25-01.html (Score: 45, Date: 2025-07-25)
Created: 2025-07-25-02.html (Score: 42, Date: 2025-07-25)
...
Created: 2025-07-24-01.html (Score: 38, Date: 2025-07-24)

✅ Successfully converted 12 articles to news directory structure!
📁 Output directory: robinhood-news-output
📄 Configuration file: article-config.js
📊 Summary file: conversion-summary.json
🎯 Selected 12 most important articles out of 2835 total

📈 Conversion Summary:
   Total articles processed: 2835
   Articles selected: 12
   Max articles limit: 12
   Date range: 2024-05-13 to 2025-07-25
   Files created: 12
   Selection criteria: Most recent and most important news based on source credibility, keywords, and content quality
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