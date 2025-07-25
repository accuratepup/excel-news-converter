import pandas as pd
from datetime import datetime, timedelta

def create_sample_excel():
    """Create a sample Excel file with Date, Title, and Description columns."""
    
    # Sample data
    data = {
        'Date': [
            datetime.now().date(),
            (datetime.now() - timedelta(days=1)).date(),
            (datetime.now() - timedelta(days=2)).date(),
            (datetime.now() - timedelta(days=3)).date(),
            (datetime.now() - timedelta(days=4)).date()
        ],
        'Title': [
            'Project Alpha Launch',
            'Team Meeting',
            'Client Presentation',
            'Code Review',
            'Documentation Update'
        ],
        'Description': [
            'Successfully launched the new project with all features working as expected.',
            'Weekly team meeting to discuss progress and upcoming tasks.',
            'Presented the quarterly results to the client team.',
            'Completed code review for the new authentication module.',
            'Updated project documentation with latest API changes.'
        ]
    }
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Save to Excel file
    output_file = 'sample_data.xlsx'
    df.to_excel(output_file, index=False)
    
    print(f"Sample Excel file created: {output_file}")
    print("\nSample data:")
    print(df.to_string(index=False))
    
    return output_file

if __name__ == "__main__":
    create_sample_excel() 