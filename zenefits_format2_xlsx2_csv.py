import numpy as np
import pandas as pd
import locale
import openpyxl
import csv
import sys
import os
import re


# Set the low level number formatting
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

"""
This Python script accepts user input and checks if it's a non-empty string. 
If the input is not a valid string or is empty, it will ask for input up to 3 
times before raising an error

This script defines a get_filepath_input and a get_filename_input function that:
Prompts the user for input. Checks if the input is a non-empty string.
If the input is invalid, it asks for input again up to 3 more times.
If the input is still invalid after 3 more attempts, it raises a ValueError.
"""
def get_filepath_input(prompt):
    attempts = 0
    max_attempts = 3
    
    while attempts <= max_attempts:
        user_input = input(prompt)
        
        if isinstance(user_input, str) and user_input.strip():
            return user_input
        else:
            attempts += 1
            print(f"Invalid input. Please enter a non-empty string. Attempts remaining: {max_attempts - attempts + 1}")
            
    raise ValueError("Maximum attempts exceeded. Valid input not provided.")

def get_filename_input(prompt):
    attempts = 0
    max_attempts = 3
    
    while attempts <= max_attempts:
        user_input = input(prompt)
        
        if isinstance(user_input, str) and user_input.strip():
            return user_input
        else:
            attempts += 1
            print(f"Invalid input. Please enter a non-empty string. Attempts remaining: {max_attempts - attempts + 1}")
            
    raise ValueError("Maximum attempts exceeded. Valid input not provided.")

def get_excel_range(prompt):
    attempts = 0
    max_attempts = 3

    while attempts < max_attempts:
        try:
            input_range = input(prompt)
            input_range = input_range.strip().replace(" ", "").split(",")
            if len(input_range) != 4:
                raise ValueError("You must enter exactly 4 numbers.")

            first_range = tuple(map(int, input_range))
            if all(isinstance(i, int) for i in first_range):
                print(f"Valid range entered: {first_range}")
                return first_range
            else:
                raise ValueError("All values must be integers.")
        
        except ValueError as e:
            attempts += 1
            print(f"Invalid input: {e}")
            if attempts < max_attempts:
                print(f"You have {max_attempts - attempts} attempts left.")
            else:
                print("No more attempts left. Please restart the program.")


def get_sort_column_index(prompt):
    attempts = 0
    max_attempts = 3

    while attempts < max_attempts:
        try:
            input_index = input(prompt)
            input_index = input_index.strip()

            if not input_index.isdigit():
                raise ValueError("You must enter a single numeric value.")

            sort_column_index = int(input_index)
            print(f"Valid sort column index entered: {sort_column_index}")
            return sort_column_index
        
        except ValueError as e:
            attempts += 1
            print(f"Invalid input: {e}")
            if attempts < max_attempts:
                print(f"You have {max_attempts - attempts} attempts left.")
            else:
                print("No more attempts left. Please restart the program.")
                
"""
The script reads the specified ranges from the Excel file, 
combines them into a single DataFrame, sorts the DataFrame by the given column, 
and writes the sorted data to a CSV file in the specified output directory.
    
In this script:

    1. file_path: The path to your Excel file.
    2. first_range: A tuple containing the start and end rows and columns for the first range 
       (start_row, start_col, end_row, end_col).
    3. second_range: A tuple containing the start and end rows and columns for the second range 
       (start_row, start_col, end_row, end_col).
    4. sort_column: The 0-based index of the column to sort the DataFrame by.
    5. dir_path: The path to the directory where the CSV file will be saved.
"""

def read_excel_ranges(file, first_range, sort_column, dir_path):
    
    # Helper function to format as currency
    def format_currency(amount):
        return '${:,.2f}'.format(amount)
    
    # Helper function to read a range of rows and columns into a DataFrame
    def read_range(start_row, start_col, end_row, end_col):
        data = []
        for row in range(start_row, end_row + 1):
            print(row)
            data.append([str(sheet.cell(row=row, column=col).value) for col in range(start_col, end_col+1)])
        return pd.DataFrame(data)
    
    # Create an output file with the name of the original
    # Update filename replace spaces with underscores
    output_file = re.sub("xlsx", "csv", file).replace(" ", "_")
    
    # Get the current working directory
    current_directory = os.getcwd()
    print(f"Current directory: {current_directory}")
    
    # Change the current working directory
    try:
        # Get the current working directory
        directory_list = current_directory.split("\\")
        if dir_path not in directory_list:
            print(f"No {dir_path} is in directory_list")
            os.chdir(dir_path)
            print(f"Directory changed to: {os.getcwd()}")
        else:
            print(f"Yes {current_directory} is set")
        
        # Load the workbook and select the active worksheet
        workbook = openpyxl.load_workbook(file)
        sheet = workbook.active  
        # Set Pandas dataframe display options
        # Display all rows
        pd.set_option('display.max_rows', None)

        # Display all columns
        pd.set_option('display.max_columns', None)

        # If you also want to adjust the column width
        pd.set_option('display.max_colwidth', None)
    
        # Create the pandas DataFrame 
        # Read the first range
        first_df = read_range(*first_range)

        # Replacing None with NaN for missing values
        df = first_df.replace({None: np.nan})  
        print(df[1])
        
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        # # Replace all column names using a list
        # df.columns = ['Last Name', 'First Name', 'W.E. Date', 'Hrs', 'Sub Total', 'Gross', 'Adjusted Wage', 'Fed. tax', 'State. tax', 'City Tax', 'Social Sec', 'Medicare', 'SDI', 'Misc Net', 'Net Pay']
        
        # # Drop multiple columns 'A' and 'C'
        # drop_columns = ['Hrs', 'Sub Total', 'Gross', 'Adjusted Wage', 'City Tax', 'Misc Net']
        # df.drop(drop_columns, axis=1, inplace=True)

        # # Sort the combined DataFrame by the specified column
        # sort_column = 'Last Name' 
        # if sort_column in df.columns:
        #     df.sort_values(by=sort_column, inplace=True)
        # else:
        #     print(f"Error: Column '{sort_column}' not found in the DataFrame.")
        
        # # Update date format to exclude time
        # df['W.E. Date'] = pd.to_datetime(df['W.E. Date']).dt.date
        
        # # make positive and format as dollar and cents
        # col_to_filter = ['Fed. tax','State. tax','Social Sec','Medicare','SDI','Net Pay']
        # for col in col_to_filter:
        #     df[col] = df[col].replace(r'-', '', regex=True)
        #     df[col] = df[col].astype(float)
        #     df[col] = df[col].apply(lambda x: '${:,.2f}'.format(x))
        
        # # print dataframe. 
        # df.to_csv(output_file, index=False)
    
    except FileNotFoundError:
        print(f"Error: Directory not found: {dir_path}")
    except NotADirectoryError:
        print(f"Error: Not a directory: {dir_path}")
    except PermissionError:
        print(f"Error: Permission denied to access: {dir_path}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    return output_file

def sort_csv(dir_path, input_csv_file, output_csv_file, sort_column_index, reverse=False):
    """
    Sorts rows in a CSV file based on a specified column.

    Args:
        input_file (str): Path to the input CSV file.
        output_file (str): Path to the output CSV file.
        sort_column_index (int): Index of the column to sort by (0-based).
        reverse (bool, optional): Sort in descending order if True. Defaults to False.
    """
    # Get the current working directory
    current_directory = os.getcwd()
    
    # Change the current working directory
    try:
        
        # Get the current working directory
        directory_list = current_directory.split("\\")
        if dir_path not in directory_list:
            print(f"No {dir_path} is in directory_list")
            os.chdir(dir_path)
            print(f"Directory changed to: {os.getcwd()}")
        else:
            print(f"Yes {current_directory} is set")
        
        with open(input_csv_file, 'r') as infile, open(output_csv_file, 'w', newline='') as outfile:
            reader = csv.reader(infile)
            header = next(reader)
            rows = list(reader)

            # Sort rows based on the specified column
            rows.sort(key=lambda row: row[sort_column_index], reverse=reverse)

            writer = csv.writer(outfile)
            writer.writerow(header)
            writer.writerows(rows)
            
    except FileNotFoundError:
        print(f"Error: Directory not found: {dir_path}")
        sys.exit(1)
    except NotADirectoryError:
        print(f"Error: Not a directory: {dir_path}")
        sys.exit(1)
    except PermissionError:
        print(f"Error: Permission denied to access: {dir_path}")
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        sys.exit(1)
    return output_csv_file

if __name__ == '__main__':
    
    file_path_input_prompt =  '''
    Please enter the folder were the files
    are stored relative to this script 
    (e.g. C:\\<script directory>\\data): 
    '''
    file_input_prompt = '''
    Please enter a <Zenefits>.xlsx files.
    Zenefits file name: 
    '''
    excel_table1_range_prompt = '''
    Enter the range (start_row, start_col, end_row, end_col) for Table 1.
    (e.g  7, 2, 44, 15):
    '''
    excel_table2_range_prompt = '''
    Enter the range (start_row, start_col, end_row, end_col) for Table 2.
    (e.g. 52, 2, 66, 15):
    '''
    sort_column_index_prompt = '''
    Enter the sort column index (e.g., 2 for the third column): 
    '''
    try:
        dir_path = get_filepath_input(file_path_input_prompt)
        print(f"You entered: {dir_path}")
    except ValueError as e:
        print(e)
    
    try:
        file = get_filename_input(file_input_prompt)
        print(f"You entered: {file}")
    except ValueError as e:
        print(e)
        
    try:
       tbl_range = get_excel_range(excel_table1_range_prompt)
       print(tbl_range)
    except ValueError as e:
        print(e)
        
    try:
        sort_column_index = get_sort_column_index(sort_column_index_prompt)
        print(sort_column_index)
    except ValueError as e:
        print(e)    
        
    try:
        read_excel_ranges(file, tbl_range, sort_column_index, dir_path)
        # output_file = read_excel_ranges(file, tbl_range, sort_column_index, dir_path)
        # print(f"CSV file {output_file} has been processed and written to {dir_path}!")
    except ValueError as e:
        print(e)
    
 