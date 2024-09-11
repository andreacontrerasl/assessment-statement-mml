import os
import pandas as pd

def load_data(path):
    combined_df = pd.DataFrame()

    for root, dirs, files in os.walk(path):
        for file in files:
            full_path = os.path.join(root, file)
            
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
                else:
                    continue
                
                # Concatenate the temporary DataFrame to the combined DataFrame
                combined_df = pd.concat([combined_df, temp_df], ignore_index=True)
                print(f"Successfully loaded file: {file}")
            
            except Exception as e:
                print(f"Error loading file {file}: {e}")

    return combined_df