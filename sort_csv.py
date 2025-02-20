import csv
import os
import re

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
    print(f"Current directory: {current_directory}")
    
    # Change the current working directory
    try:
        os.chdir(dir_path)
        print(f"Directory changed to: {os.getcwd()}")
        
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
    except NotADirectoryError:
        print(f"Error: Not a directory: {dir_path}")
    except PermissionError:
        print(f"Error: Permission denied to access: {dir_path}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        
if __name__ == '__main__':
    # Example usage:
    dir_path = 'data'
    input_csv_file = 'Plexxis_Tax Details_Ck Date 121523.csv'
    output_csv_file = 'sorted_Plexxis_Tax Details_Ck_Date_121523.csv'
    sort_column_index = 2  # Sort by the third column (index 2)
    sort_csv(dir_path, input_csv_file, output_csv_file, sort_column_index)