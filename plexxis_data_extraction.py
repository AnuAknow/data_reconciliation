import pandas as pd
import openpyxl
import os

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
    # Load the workbook and select the active worksheet
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Helper function to read a range of rows and columns into a DataFrame
    def read_range(start_row, start_col, end_row, end_col):
        data = []
        for row in range(start_row, end_row + 1):
            data.append([sheet.cell(row=row, column=col).value for col in range(start_col, end_col + 1)])
        return pd.DataFrame(data)

    # Read the first range
    first_df = read_range(*first_range)

    # Read the second range
    second_df = read_range(*second_range)

    # Combine the two DataFrames
    combined_df = pd.concat([first_df, second_df], ignore_index=True)

    # Sort the combined DataFrame by the specified column
    combined_df.sort_values(by=sort_column, inplace=True)

    # Change the directory to the specified output path
    os.chdir(output_path)

    # Write the DataFrame to a CSV file
    combined_df.to_csv('output.csv', index=False)

# Example usage
file_path = 'your_file.xlsx'
first_range = (2, 1, 6, 4)  # (start_row, start_col, end_row, end_col)
second_range = (8, 1, 12, 4)  # (start_row, start_col, end_row, end_col)
sort_column = 0  # Index of the column to sort by (0-based index)
output_path = '/path/to/output/directory'

read_excel_ranges(file_path, first_range, 