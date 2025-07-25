import pandas as pd
from datetime import datetime, timedelta

def create_sample_excel_with_source_link():
    """Create a sample Excel file with Date, Source, Title, Link, and Description columns."""
    
    # Sample data with the new structure
    data = {
        'Date': [
            datetime.now().date(),
            (datetime.now() - timedelta(days=1)).date(),
            (datetime.now() - timedelta(days=2)).date(),
            (datetime.now() - timedelta(days=3)).date(),
            (datetime.now() - timedelta(days=4)).date()
        ],
        'Source': [
            'Benzinga',
            'Yahoo Finance',
            'Reuters',
            'Bloomberg',
            'CNBC'
        ],
        'Title': [
            'Robinhood Markets, Inc. to Announce Third Quarter 2024 Results',
            'Robinhood Shares Are Trading Higher Today: What You Need To Know',
            'Robinhood Markets, Inc. to Present at Goldman Sachs Conference',
            'Robinhood 2024: Five Reasons The Financial Platform Is Still the Best',
            'Robinhood app purchase to sustain healthy competition'
        ],
        'Link': [
            'https://www.benzinga.com/news/24/10/01/robinhood-markets-inc-to-announce-third-quarter-2024-results',
            'https://finance.yahoo.com/news/robinhood-shares-trading-higher-today-123456789',
            'https://www.reuters.com/technology/robinhood-markets-present-goldman-sachs-conference',
            'https://www.bloomberg.com/news/articles/2024-05-13/robinhood-2024-five-reasons-financial-platform-best',
            'https://www.cnbc.com/2024/10/02/robinhood-app-purchase-sustain-healthy-competition'
        ],
        'Description': [
            'Robinhood Markets, Inc. (NASDAQ: HOOD) has announced that it will release its third quarter 2024 financial results on Wednesday, October 30, 2024, after market close. An earnings conference call will be held at 2:00 PM PT / 5:00 PM ET on the same day.',
            'Robinhood Markets, Inc. (NASDAQ:HOOD) shares are on the rise Monday amid possible optimism ahead of Fed President Powell\'s discussion and possibly as a result of investors assessing the impact of the attempted assassination of former President Trump.',
            'Robinhood Markets, Inc. ("Robinhood") (NASDAQ: HOOD) today announced that it will be participating in the upcoming Goldman Sachs Communacopia + Technology Conference on Tuesday, September 10, 2024.',
            'In today\'s ever-changing financial landscape, the pursuit of wealth through investment can be both tantalizing and daunting – especially for beginners. But technological advancements in recent years have ushered in a new area of accessibility.',
            'Thailand\'s on-demand delivery sector should see healthy competition with the Robinhood food delivery app remaining in the market, say industry observers.'
        ]
    }
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Save to Excel file
    output_file = 'sample_data_with_source_link.xlsx'
    df.to_excel(output_file, index=False)
    
    print(f"Sample Excel file created: {output_file}")
    print("\nSample data:")
    print(df.to_string(index=False))
    
    return output_file

if __name__ == "__main__":
    create_sample_excel_with_source_link() 