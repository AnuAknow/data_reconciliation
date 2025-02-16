import pandas as pd 
import os
import re




def write_file(dirpath, filename):
    
    # Define the target directory path
    new_directory = dirpath
    
    # Define the target filepath path
    filepath = dirpath + filename
    
    # Create an output file with the name of the original
    filename = re.sub("xlsx", "csv", filename)
    
    # Get the current working directory
    current_directory = os.getcwd()
    print(f"Current directory: {current_directory}")
    
    # Change the current working directory
    try:
        os.chdir(new_directory)
        print(f"Directory changed to: {os.getcwd()}")
        
        # Set Pandas dataframe display options
        # Display all rows
        pd.set_option('display.max_rows', None)

        # Display all columns
        pd.set_option('display.max_columns', None)

        # If you also want to adjust the column width
        pd.set_option('display.max_colwidth', None)
    
        # Read from csv
        level5dw_emp_data = pd.read_csv(filepath)
    
        # print dataframe. 
        print(level5dw_emp_data)
    
    except FileNotFoundError:
        print(f"Error: Directory not found: {new_directory}")
    except NotADirectoryError:
        print(f"Error: Not a directory: {new_directory}")
    except PermissionError:
        print(f"Error: Permission denied to access: {new_directory}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        
if __name__ == '__main__':
    v_path = 'data\\'
    v_file = 'Zenefits_payroll_detail_Ck Date 121523.xlsx'
