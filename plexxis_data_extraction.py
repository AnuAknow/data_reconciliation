import numpy as np
import pandas as pd
import openpyxl
import os
import re

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
    5. output_path: The path to the directory where the CSV file will be saved.
"""

def read_excel_ranges(file_path, first_range, second_range, sort_column, output_path):
    
    # Combine file_path (name) with output_path (directory)
    file_name = output_path + '\\' + file_path
    # Load the workbook and select the active worksheet
    workbook = openpyxl.load_workbook(file_name)
    sheet = workbook.active

    # Helper function to read a range of rows and columns into a DataFrame
    def read_range(start_row, start_col, end_row, end_col):
        data = []
        for row in range(start_row, end_row + 1):
            data.append([str(sheet.cell(row=row, column=col).value) for col in range(start_col, end_col+1)])
        return pd.DataFrame(data)
    
    # Create an output file with the name of the original
    filename = re.sub("xlsx", "csv", file_path)
    
    # Get the current working directory
    current_directory = os.getcwd()
    print(f"Current directory: {current_directory}")
    
    # Change the current working directory
    try:
        os.chdir(output_path)
        print(f"Directory changed to: {os.getcwd()}")
        
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

        # Read the second range
        second_df = read_range(*second_range)

        # Combine the two DataFrames
        combined_df = pd.concat([first_df, second_df], ignore_index=True)    

        # Replacing None with NaN for missing values
        df = combined_df.replace({None: np.nan})  
        
        # Replace all column names using a list
        df.columns = ['Last Name', 'First Name', 'W.E. Date', 'Hrs', 'Sub Total', 'Gross', 'Adjusted Wage', 'Fed. tax', 'State. tax', 'City Tax', 'Social Sec', 'Medicare', 'SDI', 'Misc Net', 'Net Pay']
        
        # Drop multiple columns 'A' and 'C'
        drop_columns = ['Hrs', 'Sub Total', 'Gross', 'Adjusted Wage', 'City Tax', 'Misc Net']
        df.drop(drop_columns, axis=1, inplace=True)d

        # Sort the combined DataFrame by the specified column
        sort_column = 'Last Name' 
        if sort_column in df.columns:
            df.sort_values(by=sort_column, inplace=True)
        else:
            print(f"Error: Column '{sort_column}' not found in the DataFrame.")

        # print dataframe. 
        df.to_csv(filename, index=False)
    
    except FileNotFoundError:
        print(f"Error: Directory not found: {output_path}")
    except NotADirectoryError:
        print(f"Error: Not a directory: {output_path}")
    except PermissionError:
        print(f"Error: Permission denied to access: {output_path}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    

if __name__ == '__main__':
    try:
        output_path = get_filepath_input("Please enter the file path: ")
        print(f"You entered: {output_path}")
    except ValueError as e:
        print(e)
    
    try:
        file_path = get_filename_input("Please enter the file name: ")
        print(f"You entered: {file_path}")
    except ValueError as e:
        print(e)
    # Example usage
    first_range = (7, 2, 44, 16)  # (start_row, start_col, end_row, end_col)
    second_range = (53, 2, 82, 16)  # (start_row, start_col, end_row, end_col)
    sort_column = 2  # Index of the column to sort by (0-based index)

    read_excel_ranges(file_path, first_range, second_range, sort_column, output_path)
    
    print(f"CSV file {file_path} has been processed and written to {output_path}!")