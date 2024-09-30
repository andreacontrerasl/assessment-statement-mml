import os
import pandas as pd
from datetime import datetime
import logging

# Setup logging
logging.basicConfig(filename='process.log', level=logging.INFO, format='%(asctime)s %(levelname)s:%(message)s')

def load_data(path):
    logging.info("Starting data loading process")
    combined_df = pd.DataFrame()

    expected_columns = {'Client', 'Country', 'Currency', 'Transaction'}
    supported_extensions = ('.xlsx', '.csv')

    for root, dirs, files in os.walk(path):
        for file in files:
            full_path = os.path.join(root, file)
            date_part = root.split(os.sep)[-1]  # Assumes date is part of the folder name

            # Skip temporary or system files
            if file.startswith('~$') or file.startswith('.'):
                logging.debug(f"Skipping temporary or system file: {file}")
                continue

            # Check if the file has a supported extension
            if not file.endswith(supported_extensions):
                logging.warning(f"Unsupported file format: {file}. Skipping file.")
                continue

            # Extract transaction date from the folder name
            try:
                transaction_date = datetime.strptime(date_part, "%d-%m-%y").date()
            except ValueError as ve:
                logging.error(f"Unable to parse date from folder name '{date_part}' for file '{file}': {ve}")
                continue

            # Check if the file exists
            if not os.path.exists(full_path):
                logging.error(f"File not found: {full_path}")
                continue

            try:
                # Read the file
                if file.endswith('.xlsx'):
                    temp_df = pd.read_excel(full_path, engine='openpyxl')
                elif file.endswith('.csv'):
                    temp_df = pd.read_csv(full_path, on_bad_lines='skip', delimiter=',', encoding='utf-8')
                else:
                    logging.warning(f"Unsupported file extension for file: {file}. Skipping file.")
                    continue

                # Validate columns
                actual_columns = set(temp_df.columns.str.strip())
                if not expected_columns.issubset(actual_columns):
                    missing_columns = expected_columns - actual_columns
                    logging.error(f"Missing expected columns in file '{file}': {missing_columns}. Skipping file.")
                    continue

                # Standardize 'Country' column
                temp_df['Country'] = temp_df['Country'].astype(str).str.strip().str.upper()

                # Validate data types
                temp_df['Transaction'] = pd.to_numeric(temp_df['Transaction'], errors='coerce')
                num_invalid = temp_df['Transaction'].isnull().sum()
                if num_invalid > 0:
                    logging.warning(f"File '{file}' contains {num_invalid} invalid 'Transaction' values set to NaN.")

                # Remove rows with NaN in critical columns
                temp_df.dropna(subset=['Client', 'Country', 'Currency', 'Transaction'], inplace=True)

                # Add transaction date
                temp_df['Date'] = transaction_date

                # Concatenate to combined DataFrame
                combined_df = pd.concat([combined_df, temp_df], ignore_index=True)
                logging.info(f"Successfully loaded file: {file}")

            except pd.errors.EmptyDataError as ede:
                logging.error(f"No data in file '{file}': {ede}")
            except pd.errors.ParserError as pe:
                logging.error(f"Parsing error in file '{file}': {pe}")
            except UnicodeDecodeError as ude:
                logging.error(f"Encoding error in file '{file}': {ude}")
            except PermissionError as perr:
                logging.error(f"Permission error accessing file '{file}': {perr}")
            except Exception as e:
                logging.error(f"Unexpected error loading file '{file}': {e}", exc_info=True)

    # Remove duplicates
    combined_df.drop_duplicates(inplace=True)
    logging.info(f"Removed duplicates, resulting in {len(combined_df)} total records.")

    logging.info("Data loading process completed")
    return combined_df