import pandas as pd
import json
import sys
import os
from datetime import datetime
import re
from pathlib import Path

def clean_filename(title, max_length=50):
    """
    Clean and format title for use as filename.
    
    Args:
        title (str): The title to clean
        max_length (int): Maximum length for filename
    
    Returns:
        str: Cleaned filename-safe string
    """
    # Remove special characters and replace spaces with hyphens
    cleaned = re.sub(r'[^\w\s-]', '', title)
    cleaned = re.sub(r'[-\s]+', '-', cleaned)
    cleaned = cleaned.strip('-')
    
    # Truncate if too long
    if len(cleaned) > max_length:
        cleaned = cleaned[:max_length].rstrip('-')
    
    return cleaned.lower()

def format_date_for_filename(date_str):
    """
    Format date string for filename use.
    
    Args:
        date_str (str): Date string in various formats
    
    Returns:
        str: Formatted date string (YYYY-MM-DD)
    """
    try:
        # Try to parse the date
        if isinstance(date_str, str):
            # Handle different date formats
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
                try:
                    parsed_date = datetime.strptime(date_str, fmt)
                    return parsed_date.strftime('%Y-%m-%d')
                except ValueError:
                    continue
        
        # If it's already a datetime object
        if isinstance(date_str, datetime):
            return date_str.strftime('%Y-%m-%d')
        
        # If it's a date object
        if hasattr(date_str, 'strftime'):
            return date_str.strftime('%Y-%m-%d')
            
    except Exception:
        pass
    
    # Default to today's date if parsing fails
    return datetime.now().strftime('%Y-%m-%d')

def create_html_content(title, source, link, description):
    """
    Create HTML content for a news article.
    
    Args:
        title (str): Article title
        source (str): Article source
        link (str): Article link/URL
        description (str): Article description/content
    
    Returns:
        str: Formatted HTML content
    """
    # Clean and format the description
    if description:
        # Split into paragraphs if there are line breaks
        paragraphs = description.split('\n\n')
        if len(paragraphs) == 1:
            paragraphs = description.split('. ')
            # Rejoin sentences that were split incorrectly
            formatted_paragraphs = []
            current_para = ""
            for para in paragraphs:
                if para.strip():
                    if current_para:
                        current_para += ". " + para
                    else:
                        current_para = para
                    if len(current_para) > 200:  # Reasonable paragraph length
                        formatted_paragraphs.append(current_para)
                        current_para = ""
            if current_para:
                formatted_paragraphs.append(current_para)
        else:
            formatted_paragraphs = [p.strip() for p in paragraphs if p.strip()]
    else:
        formatted_paragraphs = ["No description available."]
    
    # Create HTML content
    html_content = f'<h2 class="text-32 mb-4 font-700 elite-bold">{title}</h2>\n'
    
    # Add source and link information if available
    if source or link:
        html_content += '<div class="article-meta">\n'
        if source:
            html_content += f'  <p class="source"><strong>Source:</strong> {source}</p>\n'
        if link:
            html_content += f'  <p class="link"><strong>Link:</strong> <a href="{link}" target="_blank">{link}</a></p>\n'
        html_content += '</div>\n'
    
    html_content += '<div class="description">\n'
    
    for paragraph in formatted_paragraphs:
        if paragraph.strip():
            # Add emphasis to certain phrases
            paragraph = re.sub(r'\b(What Happened|Why It Matters|Price Action|EXCLUSIVE|Breaking|Update)\b', 
                             r'<span class="font-700">\1</span>', paragraph, flags=re.IGNORECASE)
            html_content += f'  <p>{paragraph}</p>\n'
    
    html_content += '</div>'
    
    return html_content

def excel_to_news_directory(excel_file_path, output_directory="news-articles"):
    """
    Convert Excel file to news directory structure.
    
    Args:
        excel_file_path (str): Path to the Excel file
        output_directory (str): Directory to create the news structure
    
    Returns:
        dict: Summary of the conversion process
    """
    try:
        # Read the Excel file
        print(f"Reading Excel file: {excel_file_path}")
        df = pd.read_excel(excel_file_path)
        
        # Create output directory
        output_path = Path(output_directory)
        output_path.mkdir(exist_ok=True)
        
        print(f"Creating news directory: {output_path}")
        
        # Check if required columns exist
        required_columns = ['Date', 'Source', 'Title', 'Link', 'Description']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"Warning: Missing columns: {missing_columns}")
            print(f"Available columns: {list(df.columns)}")
            print("Proceeding with available columns...")
        
        # Group articles by date
        articles_by_date = {}
        file_list = []
        
        for index, row in df.iterrows():
            try:
                # Get date and format it
                date_str = row.get('Date', '')
                formatted_date = format_date_for_filename(date_str)
                
                # Get title, source, link, and description
                title = str(row.get('Title', f'Article {index + 1}'))
                source = str(row.get('Source', ''))
                link = str(row.get('Link', ''))
                description = str(row.get('Description', ''))
                
                # Skip rows with "Social Media Post" in description
                if 'Social Media Post' in description:
                    print(f"Skipping row {index + 1}: Social Media Post detected in description")
                    continue
                
                # Clean title for filename
                clean_title = clean_filename(title)
                
                # Create filename
                if formatted_date not in articles_by_date:
                    articles_by_date[formatted_date] = 1
                else:
                    articles_by_date[formatted_date] += 1
                
                sequence_num = articles_by_date[formatted_date]
                filename = f"{formatted_date}-{sequence_num:02d}.html"
                
                # Create HTML content
                html_content = create_html_content(title, source, link, description)
                
                # Write HTML file
                file_path = output_path / filename
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(html_content)
                
                file_list.append(filename)
                print(f"Created: {filename}")
                
            except Exception as e:
                print(f"Error processing row {index + 1}: {str(e)}")
                continue
        
        # Create article configuration file
        config_content = f"""// Article configuration for {Path(excel_file_path).stem}
const articleConfigs = {{
  files: [
    {', '.join([f"'{file}'" for file in sorted(file_list, reverse=True)])}
  ]
}};"""
        
        config_file_path = output_path / "article-config.js"
        with open(config_file_path, 'w', encoding='utf-8') as f:
            f.write(config_content)
        
        print(f"\nCreated configuration file: article-config.js")
        
        # Create summary
        summary = {
            "source_file": excel_file_path,
            "output_directory": str(output_path),
            "total_articles": len(file_list),
            "date_range": {
                "earliest": min(articles_by_date.keys()) if articles_by_date else None,
                "latest": max(articles_by_date.keys()) if articles_by_date else None
            },
            "files_created": file_list,
            "export_date": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        # Save summary as JSON
        summary_file_path = output_path / "conversion-summary.json"
        with open(summary_file_path, 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)
        
        print(f"\n✅ Successfully converted {len(file_list)} articles to news directory structure!")
        print(f"📁 Output directory: {output_path}")
        print(f"📄 Configuration file: article-config.js")
        print(f"📊 Summary file: conversion-summary.json")
        
        return summary
        
    except FileNotFoundError:
        print(f"Error: File '{excel_file_path}' not found.")
        return None
    except Exception as e:
        print(f"Error processing Excel file: {str(e)}")
        return None

def main():
    """Main function to handle command line arguments."""
    if len(sys.argv) < 2:
        print("Usage: python excel_to_news_directory.py <excel_file_path> [output_directory]")
        print("Example: python excel_to_news_directory.py robinhood.xlsx")
        print("Example: python excel_to_news_directory.py robinhood.xlsx my-news")
        return
    
    excel_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "news-articles"
    
    # Convert Excel to news directory
    result = excel_to_news_directory(excel_file, output_dir)
    
    if result:
        print(f"\n📈 Conversion Summary:")
        print(f"   Total articles: {result['total_articles']}")
        print(f"   Date range: {result['date_range']['earliest']} to {result['date_range']['latest']}")
        print(f"   Files created: {len(result['files_created'])}")

if __name__ == "__main__":
    main() 