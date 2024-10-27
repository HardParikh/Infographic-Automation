import os
import pandas as pd
from snowflake_connector import create_snowflake_connection
import xlwings as xw
from datetime import datetime

def get_previous_month():
    today = datetime.today()
    first = today.replace(day=1)
    prev_month = first.replace(month=first.month - 1 if first.month > 1 else 12, 
                               year=first.year if first.month > 1 else first.year - 1)
    return prev_month.strftime("%B")


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
    # Select the source range and target range
    source_cells = source_ws.range((source_range[0], source_range[1]), (source_range[2], source_range[3]))
    target_start = target_ws.range(target_start_cell)
    
    # Copy the table from source to target
    source_cells.copy(target_start)

# Function to replace table
def replace_table(source_file, source_sheet_name, source_table_range, target_file, target_sheet_name, target_start_cell):
    previous_month = get_previous_month()
    app = xw.App(visible=False)
    
    try:
        source_wb = app.books.open(source_file)
        target_wb = app.books.open(target_file)

        try:
            source_ws = source_wb.sheets[source_sheet_name]
            target_ws = target_wb.sheets[target_sheet_name]

            # Copy the table from source to target
            copy_table(source_ws, target_ws, source_table_range, target_start_cell)
            
            # Replace the content in cell I7 with the previous month
            target_ws.range('I7').value = previous_month

            # Save the changes to the target workbook
            target_wb.save(target_file)
            print(f"Table updated in '{target_sheet_name}' of '{target_file}' successfully.")
        finally:
            # Close the workbooks after operations
            source_wb.close()
            target_wb.close()
    except Exception as e:
        if "Cannot access" in str(e) and "xlmain11.chm" in str(e):
            print("Encountered known file access error, but operations completed successfully.")
        else:
            print(f"Unexpected error while replacing the table: {e}")
    finally:
        app.quit()

# Function to remove trailing zeros
def remove_trailing_zeros(file_path, sheet_name, column_name):
    app = xw.App(visible=False)
    
    try:
        # Open the workbook and select the sheet
        wb = app.books.open(file_path)
        sheet = wb.sheets[sheet_name]

        # Find the column index for the specified column name
        col_range = sheet.range(1, 1).expand('right').value
        stat_col = None
        for i, col_name in enumerate(col_range):
            if col_name == column_name:
                stat_col = i + 1
                break

        if not stat_col:
            print(f"Column '{column_name}' not found!")
            return

        # Get the last row with data in the column
        last_row = sheet.range(1, stat_col).end('down').row

        # Iterate over each cell in the column to apply formatting
        for row in range(2, last_row + 1):
            cell = sheet.range((row, stat_col))
            value = cell.value
            
            # Convert text-based numbers to float
            if isinstance(value, str):
                try:
                    value = float(value)
                except ValueError:
                    continue

            # Format value as an Excel text formula with thousands separator
            if isinstance(value, (int, float)):
                if value.is_integer():
                    # Format integer as ="#,###"
                    cell.value = f'="{int(value):,}"'
                else:
                    # Format float with up to 6 decimals and remove unnecessary trailing zeros
                    formatted_value = f"{value:,.6f}".rstrip('0').rstrip('.')
                    cell.value = f'="{formatted_value}"'

        wb.save()
        print(f"Values formatted as text formulas with commas in column '{column_name}' of sheet '{sheet_name}' successfully.")
    except Exception as e:
        print(f"Error while formatting column '{column_name}': {e}")
    finally:
        wb.close()
        app.quit()