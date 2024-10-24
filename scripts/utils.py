import os
import pandas as pd
from snowflake_connector import create_snowflake_connection
import openpyxl

# Function to read SQL queries from files
def read_query(file_path):
    try:
        with open(file_path, 'r') as file:
            query = file.read()
        return query
    except FileNotFoundError:
        print(f"Error: File not found - {file_path}")
        return None

# Function to fetch data from Snowflake and return a DataFrame
def fetch_data(query):
    try:
        conn = create_snowflake_connection()
    except Exception as e:
        print(f"Error connecting to Snowflake: {e}")
        return None

    try:
        cur = conn.cursor()
        cur.execute(query)
        results = cur.fetchall()

        if not results:
            print("Warning: No data returned from query.")
            return pd.DataFrame()

        df = pd.DataFrame(results, columns=[desc[0] for desc in cur.description])
    except Exception as e:
        print(f"Error executing query: {e}")
        return pd.DataFrame() 
    finally:
        cur.close()
        conn.close()

    return df

# Function to save data into an Excel file with multiple sheets
def save_to_excel(excel_file_path, query_files):
    os.makedirs(os.path.dirname(excel_file_path), exist_ok=True)

    # Initialize an Excel writer
    try:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            for sheet_name, query_file in query_files.items():
                query = read_query(query_file)
                if query:
                    df = fetch_data(query)
                    if not df.empty:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        print(f"Warning: No data to write for {sheet_name}.")
                else:
                    print(f"Error: Could not read query file for {sheet_name}. Skipping.")
    except Exception as e:
        print(f"Error saving data to Excel: {e}")
    else:
        print(f"Data saved to '{excel_file_path}' successfully!")
        

def copy_table(source_ws, target_ws, source_range, target_start_cell):
    min_row, min_col, max_row, max_col = source_range
    target_row = target_ws[target_start_cell].row
    target_col = target_ws[target_start_cell].column

    for i, row in enumerate(source_ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col), start=target_row):
        for j, cell in enumerate(row, start=target_col):
            target_ws.cell(row=i, column=j, value=cell.value)
            

def replace_table(source_file, source_sheet_name, source_table_range, target_file, target_sheet_name, target_start_cell):
    source_wb = openpyxl.load_workbook(source_file)
    target_wb = openpyxl.load_workbook(target_file)

    source_ws = source_wb[source_sheet_name]
    target_ws = target_wb[target_sheet_name]

    copy_table(source_ws, target_ws, source_table_range, target_start_cell)

    target_wb.save(target_file)
    print(f"Table updated in '{target_sheet_name}' of '{target_file}' successfully.")


def remove_trailing_zeros(file_path, sheet_name, column_name):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]
    
    stat_col = None
    for cell in sheet[1]:
        if cell.value == column_name:
            stat_col = cell.column
            break

    if not stat_col:
        print("Column not found!")
        return
    
    for row in sheet.iter_rows(min_row=2, min_col=stat_col, max_col=stat_col):
        cell = row[0]
        if isinstance(cell.value, str):
            try:
                cell.value = float(cell.value)
            except ValueError:
                continue
        
        if isinstance(cell.value, (int, float)):
            if cell.value.is_integer():
                cell.value = int(cell.value)
                cell.number_format = '0'
            else:
                cell.number_format = '0.######'

    wb.save(file_path)
    wb.close()
