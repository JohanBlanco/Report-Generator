import os
import pandas as pd
import shutil
from datetime import datetime
import xlwings as xw
import chardet
import subprocess
from variables import *
import time

def move_files_info_to_template(combined_dataframes: pd.DataFrame, template: pd.DataFrame):

    common_columns = [col for col in combined_dataframes.columns if col in template.columns]
    relevant_information: pd.DataFrame = combined_dataframes[common_columns]

    if len(template) < len(relevant_information):
        # If the template DataFrame is smaller than the relevant_information, extend it
        additional_rows = len(relevant_information) - len(template)
        additional_data = pd.DataFrame(index=range(len(template), len(template) + additional_rows), columns=common_columns)
        template = pd.concat([template, additional_data], ignore_index=True)
    
    # Fill in the updated template
    for column_name in common_columns:
        template[column_name] = relevant_information[column_name].values

    return template

    

def copy_file(source_file_path, destination_dir, file_name):
    try:
        # Check if source file exists
        if not os.path.exists(source_file_path):
            raise FileNotFoundError(f"The source file '{source_file_path}' does not exist.")
        
        # Create destination directory if it does not exist
        if not os.path.exists(destination_dir):
            os.makedirs(destination_dir)
        
        # Construct full destination path
        destination_path = os.path.join(destination_dir, file_name)

        aux = destination_path.split('.xlsx')
        destination_path = aux[0] + ' ' + str(datetime.now().strftime("%Y-%m-%d_%H-%M-%S")) + aux[1] + '.xlsx'
        
        # Copy the file
        shutil.copy2(source_file_path, destination_path)
        
        # Return the path of the new file
        return destination_path
        
    except Exception as e:
        print(f"Failed to copy file: {e}")
        return None

def read_excel_file(file_path):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)
        
        # Drop rows with any NaN values
        df_cleaned = df.fillna("")
        
        return df_cleaned
    except Exception as e:
        print(f"Failed to read {file_path}: {e}")
        return pd.DataFrame()  # 
    
def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read(10000))  # Read the first 10000 bytes for detection
    return result['encoding']

def clean_description(text: str) -> str:
    """Remove unwanted characters while preserving newlines."""
    if pd.isna(text):
        return " "
    # Replace carriage returns (\r) with spaces, while preserving newlines (\n)
    text = text.lower().replace('\r', " ").replace('\t', " ").replace('_x000d_', " ")
    return text

def read_files(files_directory_path: str):
    sites_dict = {}
    data_frames = []

    if not os.path.exists(files_directory_path):
        print(f"The directory '{files_directory_path}' does not exist.")
        return data_frames, sites_dict

    for filename in os.listdir(files_directory_path):
        file_path = os.path.join(files_directory_path, filename)

        if os.path.isfile(file_path) and filename.endswith('.csv'):
            try:
                # Read the CSV file into a DataFrame with UTF-8 encoding
                df = pd.read_csv(file_path, encoding='utf-8')

                # Ensure column names are cleaned and set properly
                if not df.empty:
                    df.columns = df.columns.str.strip()  # Clean column names

                    # Clean the 'Description' column if it exists
                    if 'Description' in df.columns:
                        df['Description'] = df['Description'].apply(clean_description)

                    data_frames.append(df)

                    # Extract Task IDs if available
                    if 'Task ID' in df.columns:
                        sites_dict[filename] = list(df['Task ID'].dropna().astype(str))

            except Exception as e:
                print(f"Failed to read {filename}: {e}")

    return data_frames, sites_dict

def combine_dataframes(data_frames):
    if data_frames:
        combined_df = pd.concat(data_frames, ignore_index=True)
        return combined_df
    else:
        return pd.DataFrame()

def column_index_to_letter(index):
    """Convert a 1-based column index to an Excel column letter."""
    letter = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter

def get_last_column_letter(sheet, table_name):
    """Helper function to get the last column letter of a table in an Excel sheet."""
    tables = [tbl for tbl in sheet.tables if tbl.name == table_name]
    if not tables:
        raise ValueError(f"The sheet does not contain a table named '{table_name}'.")
    
    table = tables[0]
    last_col_index = table.range.columns.count
    return column_index_to_letter(last_col_index)  # Convert column index to letter

def open_file(file_path: str, app:str):
    if app == "excel":
        app = xw.App(visible=True)
        try:
            # Open the workbook in read-only mode
            wb = app.books.open(file_path, read_only=True)
            # Perform operations here, if needed
            print("File opened in read-only mode.")
            # You can add code here to read data, etc.
        except Exception as e:
            print(f"An error occurred: {e}")
        finally:
            wb.close()
            app.quit()
    else:
        try:
            abs_path = os.path.abspath(file_path)
            excel_command = f'start {app} "{abs_path}"'
            subprocess.run(excel_command, shell=True, check=True)
        except Exception as e:
            print(f"Error opening Excel file: {e}")

def delete_all_files_in_folder(folder_path):
    # Ensure the provided path is a directory
    if not os.path.isdir(folder_path):
        raise ValueError(f"The path {folder_path} is not a directory.")

    # Iterate over all files in the directory
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        # Check if it's a file and not .gitkeep, then delete it
        if os.path.isfile(file_path) and filename != '.gitkeep':
            os.remove(file_path)
            print(f"Deleted file: {file_path}")

def store_excel_as_csv(file_paths):
    #empty the folder first
    delete_all_files_in_folder(files_directory_path)
    # Ensure the destination folder exists, create if it doesn't
    if not os.path.exists(files_directory_path):
        os.makedirs(files_directory_path)
    
    for file_path in file_paths:
        if os.path.isfile(file_path) and file_path.lower().endswith('.xlsx'):
            # Get the file name without the extension
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            # Define the destination CSV path
            destination_csv_path = os.path.join(files_directory_path, f"{file_name}.csv")
            
            try:
                # Read the Excel file
                df = pd.read_excel(file_path)
                # Save the DataFrame to a CSV file
                df.to_csv(destination_csv_path, index=False)
                print(f"Converted {file_path} to {destination_csv_path}")
            except Exception as e:
                print(f"Failed to convert {file_path}: {e}")
        else:
            print(f"Skipping {file_path}: Not an Excel file or does not exist")
