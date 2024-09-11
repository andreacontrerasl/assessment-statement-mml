import os
import pandas as pd
from datetime import datetime

def load_data(path):
    combined_df = pd.DataFrame()

    for root, dirs, files in os.walk(path):
        for file in files:
            full_path = os.path.join(root, file)
            
            # Extract date from the path (assumes date is part of the folder name)
            date_part = root.split(os.sep)[-1]  # Gets the last part of the path
            try:
                # Parse the date assuming format like '26-01-24' for 24th January 2026
                transaction_date = datetime.strptime(date_part, "%d-%m-%y").date()
            except ValueError:
                # If the date is not in the expected format or if it's the root directory which might not be a date
                continue

            # Ignore Excel temporary files and unsupported files
            if file.startswith('~$') or not (file.endswith('.xlsx') or file.endswith('.csv')):
                print(f"Ignoring file: {file}")
                continue

            try:
                # Determine file format and load data
                if file.endswith('.xlsx'):
                    print(f"Loading XLSX file: {full_path}")
                    temp_df = pd.read_excel(full_path, engine='openpyxl')
                elif file.endswith('.csv'):
                    print(f"Loading CSV file: {full_path}")
                    temp_df = pd.read_csv(full_path, on_bad_lines='skip', delimiter=',')
                    
                if 'Country' in temp_df.columns:
                    temp_df['Country'] = temp_df['Country'].str.upper()
                
                # Add a new column for the date
                temp_df['Date'] = transaction_date
                
                # Concatenate the temporary DataFrame to the combined DataFrame
                combined_df = pd.concat([combined_df, temp_df], ignore_index=True)
                print(f"Successfully loaded file: {file}")
            
            except Exception as e:
                print(f"Error loading file {file}: {e}")

    return combined_df