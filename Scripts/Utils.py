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

def refresh_pivot_tables(sheet=None):
    if sheet:
        # Refresh pivot tables in the specific sheet
        for pt in sheet.api.PivotTables():
            pt.PivotCache().Refresh()
    else:
        # Refresh all pivot tables in the workbook
        for s in xw.books.active.sheets:
            for pt in s.api.PivotTables():
                pt.PivotCache().Refresh()

def create_excel_with_content(source_file_path: str, destination_path: str, content: pd.DataFrame, new_file_name: str, lattest_report_dir: str, tasks_to_be_painted) -> str:
    # Combine the destination path and new file name to form the full path
    lattest_report_path = os.path.join(lattest_report_dir, f"{new_file_name}.xlsx")
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    new_file_path = os.path.join(destination_path, f"{new_file_name}_{timestamp}.xlsx")
    
    # Copy the source file to the new destination
    shutil.copy2(source_file_path, new_file_path)  # Using copy2 to preserve metadata

    # Open the workbook with xlwings (invisible)
    app = xw.App(visible=False)
    try:
        # Open the source file to determine the last column of the table
        source_app = xw.App(visible=False)
        source_wb = source_app.books.open(source_file_path)
        source_sheet = source_wb.sheets['Tasks']
        last_col_letter = get_last_column_letter(source_sheet, 'Tasks')
        source_wb.close()
        source_app.quit()

        # Now open the new workbook
        wb = app.books.open(new_file_path)
        sheet = None

        # Check if 'Tasks' sheet exists
        for s in wb.sheets:
            if s.name == 'Tasks':
                sheet = s
                break
        
        if sheet is None:
            raise ValueError("The source file does not contain a sheet named 'Tasks'.")

        # Get the existing table in the sheet
        table_name = 'Tasks'
        tables = [tbl for tbl in sheet.tables if tbl.name == table_name]
        if not tables:
            raise ValueError(f"The sheet 'Tasks' does not contain a table named '{table_name}'.")

        table = tables[0]
        
        # Find the first empty row in the table
        last_row = sheet.range(f"A{sheet.cells.last_cell.row}").end('up').row
        first_empty_row = last_row + 1
        
        # Replace NaN values with None
        content_cleaned = content.where(pd.notna(content), "")
        
        # Write the cleaned content to the sheet starting from the first empty row
        sheet.range(f"A{first_empty_row}").value = content_cleaned.values
        
        # Define the new table range
        table_range = sheet.range(f"A1:{last_col_letter}{first_empty_row + content_cleaned.shape[0] - 1}")
        
        # Resize the table to include the new data
        table.source_range = table_range

        # Get the range of the table
        table_range = table.data_body_range

        # Extract the values from the table range
        table_values = table_range.value

        # Replace cells with the string 'nan' with an empty string
        for row_idx, row in enumerate(table_values):
            for col_idx, value in enumerate(row):
                if isinstance(value, str) and value.lower() == 'nan':
                    table_range[row_idx, col_idx].value = ''

        # Find the indexes of the rows to be painted
        task_id_col_index = content.columns.get_loc('Task ID')
        task_ids = content['Task ID'].values
        row_indexes_to_paint = [i + first_empty_row for i, task_id in enumerate(task_ids) if task_id in tasks_to_be_painted]

        # Paint the rows yellow if the Task ID is in tasks_to_be_painted
        yellow_color = (255, 255, 0)  # RGB value for yellow
        for row_index in row_indexes_to_paint:
            sheet.range(f"A{row_index}:{last_col_letter}{row_index}").color = yellow_color

        # Adjust first_empty_row for possible deletion
        first_empty_row -= 1
        # Delete the first empty row
        if first_empty_row <= sheet.cells.last_cell.row:
            sheet.range(f"{first_empty_row}:{first_empty_row}").api.Delete()
        
        # Refresh pivot tables
        pivot_sheet = None
        for s in wb.sheets:
            if s.name == 'Pivot Tables':
                pivot_sheet = s
                break

        if pivot_sheet:
            refresh_pivot_tables(pivot_sheet)
        else:
            refresh_pivot_tables()
        
        # Save and close the workbook
        wb.save()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        wb.close()
        app.quit()

    # Copy the file to the latest report directory
    delete_all_files_in_folder(lattest_report_dir)
    shutil.copy2(new_file_path, lattest_report_path)  # Using copy2 to preserve metadata

    return lattest_report_path

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
        # Check if it's a file and then delete it
        if os.path.isfile(file_path):
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
