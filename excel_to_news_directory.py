import pandas as pd
import json
import sys
import os
from datetime import datetime
import re
from pathlib import Path

def load_config(config_file="config.json"):
    """
    Load configuration from JSON file.
    
    Args:
        config_file (str): Path to configuration file
    
    Returns:
        dict: Configuration dictionary
    """
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        print(f"✅ Configuration loaded from {config_file}")
        return config
    except FileNotFoundError:
        print(f"⚠️  Configuration file {config_file} not found. Using default settings.")
        return get_default_config()
    except json.JSONDecodeError as e:
        print(f"❌ Error parsing {config_file}: {e}. Using default settings.")
        return get_default_config()

def get_default_config():
    """
    Get default configuration if config file is not available.
    
    Returns:
        dict: Default configuration
    """
    return {
        "algorithm_settings": {
            "max_articles": 12,
            "recency_max_days": 30,
            "recency_max_points": 30
        },
        "important_sources": [
            "Reuters", "Bloomberg", "CNBC", "Wall Street Journal", "Financial Times",
            "Yahoo Finance", "MarketWatch", "Seeking Alpha", "Benzinga", "Investing.com"
        ],
        "importance_keywords": [
            "BREAKING", "EXCLUSIVE", "UPDATE", "ALERT", "CRITICAL", "URGENT",
            "MAJOR", "SIGNIFICANT", "IMPORTANT", "KEY", "CRUCIAL", "VITAL"
        ],
        "scoring_weights": {
            "source_credibility": 20,
            "keyword_importance": 10,
            "content_quality": {
                "long_description": 5,
                "medium_description": 3,
                "long_threshold": 200,
                "medium_threshold": 100
            }
        },
        "output_settings": {
            "default_output_directory": "news-articles",
            "create_summary": True,
            "create_config_file": True
        }
    }

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

def create_html_content(title, source, link, description, config):
    """
    Create HTML content for a news article.
    
    Args:
        title (str): Article title
        source (str): Article source
        link (str): Article link/URL
        description (str): Article description/content
        config (dict): Configuration dictionary
    
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
    
    # Get important sources from config
    important_sources = config.get("important_sources", [])
    
    # Create HTML content
    html_content = f'<h2 class="text-32 mb-4 font-700 elite-bold">{title}</h2>\n'
    
    # Add source and link information if available
    if source or link:
        html_content += '<div class="article-meta">\n'
        if source:
            # Check if this is an important source for special styling
            is_important = any(imp_source.lower() in source.lower() for imp_source in important_sources)
            source_class = 'source-important' if is_important else 'source'
            source_icon = '🔴' if is_important else '📰'
            html_content += f'  <p class="{source_class}">{source_icon} <strong>Source:</strong> <span class="source-name">{source}</span></p>\n'
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

def excel_to_news_directory(excel_file_path, output_directory="news-articles", max_articles=None, config_file="config.json"):
    """
    Convert Excel file to news directory structure.
    
    Args:
        excel_file_path (str): Path to the Excel file
        output_directory (str): Directory to create the news structure
        max_articles (int): Maximum number of articles to extract (overrides config)
        config_file (str): Path to configuration file
    
    Returns:
        dict: Summary of the conversion process
    """
    try:
        # Load configuration
        config = load_config(config_file)
        
        # Use provided max_articles or config value
        if max_articles is None:
            max_articles = config["algorithm_settings"]["max_articles"]
        
        # Read the Excel file
        print(f"Reading Excel file: {excel_file_path}")
        df = pd.read_excel(excel_file_path)
        
        # Create output directory
        output_path = Path(output_directory)
        output_path.mkdir(exist_ok=True)
        
        print(f"Creating news directory: {output_path}")
        print(f"🎯 Algorithm configured for {max_articles} articles")
        
        # Check if required columns exist
        required_columns = ['Date', 'Source', 'Title', 'Link', 'Description']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"Warning: Missing columns: {missing_columns}")
            print(f"Available columns: {list(df.columns)}")
            print("Proceeding with available columns...")
        
        # Get configuration values
        important_sources = config["important_sources"]
        importance_keywords = config["importance_keywords"]
        scoring_weights = config["scoring_weights"]
        recency_settings = config["algorithm_settings"]
        
        print(f"📰 Important sources: {len(important_sources)} configured")
        print(f"🔑 Importance keywords: {len(importance_keywords)} configured")
        
        # Process and score articles
        articles_data = []
        
        for index, row in df.iterrows():
            try:
                # Get date and format it
                date_str = row.get('Date', '')
                formatted_date = format_date_for_filename(date_str)
                
                # Parse date for sorting
                try:
                    if isinstance(date_str, str):
                        for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
                            try:
                                parsed_date = datetime.strptime(date_str, fmt)
                                break
                            except ValueError:
                                continue
                        else:
                            parsed_date = datetime.now()
                    elif isinstance(date_str, datetime):
                        parsed_date = date_str
                    else:
                        parsed_date = datetime.now()
                except:
                    parsed_date = datetime.now()
                
                # Get title, source, link, and description
                title = str(row.get('Title', f'Article {index + 1}'))
                source = str(row.get('Source', ''))
                link = str(row.get('Link', ''))
                description = str(row.get('Description', ''))
                
                # Skip rows with "Social Media Post" in description
                if 'Social Media Post' in description:
                    print(f"Skipping row {index + 1}: Social Media Post detected in description")
                    continue
                
                # Calculate importance score
                importance_score = 0
                
                # Date score (more recent = higher score)
                days_old = (datetime.now() - parsed_date).days
                max_days = recency_settings["recency_max_days"]
                max_points = recency_settings["recency_max_points"]
                date_score = max(0, max_points - days_old)
                importance_score += date_score
                
                # Source importance score
                source_upper = source.upper()
                if any(imp_source.upper() in source_upper for imp_source in important_sources):
                    importance_score += scoring_weights["source_credibility"]
                
                # Keyword importance score
                title_upper = title.upper()
                desc_upper = description.upper()
                keyword_count = sum(1 for keyword in importance_keywords if keyword in title_upper or keyword in desc_upper)
                importance_score += keyword_count * scoring_weights["keyword_importance"]
                
                # Content length score (longer descriptions might indicate more detailed/important news)
                content_length = len(description)
                long_threshold = scoring_weights["content_quality"]["long_threshold"]
                medium_threshold = scoring_weights["content_quality"]["medium_threshold"]
                
                if content_length > long_threshold:
                    importance_score += scoring_weights["content_quality"]["long_description"]
                elif content_length > medium_threshold:
                    importance_score += scoring_weights["content_quality"]["medium_description"]
                
                # Store article data with score
                articles_data.append({
                    'index': index,
                    'date': parsed_date,
                    'formatted_date': formatted_date,
                    'title': title,
                    'source': source,
                    'link': link,
                    'description': description,
                    'importance_score': importance_score
                })
                
            except Exception as e:
                print(f"Error processing row {index + 1}: {str(e)}")
                continue
        
        # Sort articles by importance score (descending) and then by date (descending)
        articles_data.sort(key=lambda x: (x['importance_score'], x['date']), reverse=True)
        
        # Take only the top max_articles
        selected_articles = articles_data[:max_articles]
        
        print(f"Selected {len(selected_articles)} articles out of {len(articles_data)} total articles")
        print("Articles selected based on importance score and recency")
        
        # Group articles by date for file naming
        articles_by_date = {}
        file_list = []
        
        for article in selected_articles:
            # Create filename
            if article['formatted_date'] not in articles_by_date:
                articles_by_date[article['formatted_date']] = 1
            else:
                articles_by_date[article['formatted_date']] += 1
            
            sequence_num = articles_by_date[article['formatted_date']]
            filename = f"{article['formatted_date']}-{sequence_num:02d}.html"
            
            # Create HTML content
            html_content = create_html_content(article['title'], article['source'], article['link'], article['description'], config)
            
            # Write HTML file
            file_path = output_path / filename
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            file_list.append(filename)
            print(f"Created: {filename} (Score: {article['importance_score']}, Date: {article['formatted_date']})")
        
        # Create article configuration file
        config_content = f"""const articleConfigs = {{
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
            "total_articles_processed": len(articles_data),
            "articles_selected": len(selected_articles),
            "max_articles_limit": max_articles,
            "date_range": {
                "earliest": min(articles_by_date.keys()) if articles_by_date else None,
                "latest": max(articles_by_date.keys()) if articles_by_date else None
            },
            "files_created": file_list,
            "export_date": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "selection_criteria": "Most recent and most important news based on source credibility, keywords, and content quality",
            "config_used": config_file,
            "algorithm_settings": {
                "important_sources_count": len(important_sources),
                "importance_keywords_count": len(importance_keywords),
                "scoring_weights": scoring_weights
            }
        }
        
        # Save summary as JSON
        summary_file_path = output_path / "conversion-summary.json"
        with open(summary_file_path, 'w', encoding='utf-8') as f:
            json.dump(summary, f, indent=2, ensure_ascii=False)
        
        print(f"\n✅ Successfully converted {len(file_list)} articles to news directory structure!")
        print(f"📁 Output directory: {output_path}")
        print(f"📄 Configuration file: article-config.js")
        print(f"📊 Summary file: conversion-summary.json")
        print(f"🎯 Selected {len(selected_articles)} most important articles out of {len(articles_data)} total")
        print(f"⚙️  Configuration loaded from: {config_file}")
        
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
        print("Usage: python excel_to_news_directory.py <excel_file_path> [output_directory] [max_articles] [config_file]")
        print("Example: python excel_to_news_directory.py robinhood.xlsx")
        print("Example: python excel_to_news_directory.py robinhood.xlsx my-news")
        print("Example: python excel_to_news_directory.py robinhood.xlsx my-news 12")
        print("Example: python excel_to_news_directory.py robinhood.xlsx my-news 12 custom-config.json")
        return
    
    excel_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "news-articles"
    max_articles = int(sys.argv[3]) if len(sys.argv) > 3 else None
    config_file = sys.argv[4] if len(sys.argv) > 4 else "config.json"
    
    if max_articles:
        print(f"🎯 Extracting up to {max_articles} most important and recent news articles...")
    else:
        print(f"🎯 Extracting articles based on configuration file: {config_file}")
    
    # Convert Excel to news directory
    result = excel_to_news_directory(excel_file, output_dir, max_articles, config_file)
    
    if result:
        print(f"\n📈 Conversion Summary:")
        print(f"   Total articles processed: {result['total_articles_processed']}")
        print(f"   Articles selected: {result['articles_selected']}")
        print(f"   Max articles limit: {result['max_articles_limit']}")
        print(f"   Date range: {result['date_range']['earliest']} to {result['date_range']['latest']}")
        print(f"   Files created: {len(result['files_created'])}")
        print(f"   Selection criteria: {result['selection_criteria']}")
        print(f"   Configuration used: {result['config_used']}")
        print(f"   Algorithm settings: {result['algorithm_settings']['important_sources_count']} sources, {result['algorithm_settings']['importance_keywords_count']} keywords")

if __name__ == "__main__":
    main() 