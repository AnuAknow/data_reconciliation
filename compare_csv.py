
from tabulate import tabulate
import time
import csv
import os
import sys


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

def get_output_filename_input(prompt):
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

def create_diff_report(dir_path, csv_file1, csv_file2, output_file):
    def compare_row_to_csv(row, csv_filepath):
        """
        Compares a given row (list) against all rows in a CSV file.

        Args:
            row (list): The row to compare.
            csv_filepath (str): Path to the CSV file.

        Returns:
            list: List of rows from the CSV that match the input row.
        """
        __matching_rows = []
        __unmatching_rows = []
        with open(csv_filepath, 'r') as file:
            csv_reader = csv.reader(file)
            for csv_row in csv_reader:
                __diff = ['diffs']
                __row_contents = []
                if row == csv_row:
                    # print(f"matching:{row}, {csv_row}")
                    __row_contents.append("matching") 
                    __row_contents.append(row)
                    __row_contents.append(csv_row)
                    diff1 = [x for x in row if x not in csv_row]
                    if len(diff1) > 0:
                        # print(f"Diff1 {diff1}")
                        __diff.append(diff1)
                    diff2 = [x for x in csv_row if x not in row]
                    if len(diff2) > 0:
                        # print(f"Diff2 {diff2}")
                        __diff.append(diff2)
                    if len(__diff) > 0: 
                        __row_contents.append(__diff)
                        __matching_rows.append(__row_contents) 
                 
                elif row[0] == csv_row[0] and row[1] == csv_row[1] and row != csv_row:
                    # print(f"unmatched:{row}, {csv_row}")
                    __row_contents.append("unmatching") 
                    __row_contents.append(row)
                    __row_contents.append(csv_row)
                    diff1 = [x for x in row if x not in csv_row]
                    if len(diff1) > 0:
                        __diff.append(diff1) 
                    diff2 = [x for x in csv_row if x not in row]
                    if len(diff2) > 0:
                        __diff.append(diff2)
                    if len(__diff) > 0: 
                        __row_contents.append(__diff)
                        __unmatching_rows.append(__row_contents) 
                    
        return __matching_rows, __unmatching_rows
    
    # Get the current working directory
    current_directory = os.getcwd()
    print(f"Current directory: {current_directory}")
    
    # Change the current working directory
    try:
        # Change directory
        os.chdir(dir_path)
        print(f"Directory changed to: {os.getcwd()}")
        
        with open(csv_file2, 'r') as file, open(output_file, 'w') as f:
            __csv_reader = csv.reader(file)
            __row_count = 0

            for __csv_row in __csv_reader:
                __matching_list, __unmatching_list = compare_row_to_csv(__csv_row, csv_file1)
                if len(__matching_list) > 0:
                    
                    ''' 
                    Print matching output on screen
                    '''
                    print("".join(__matching_list[0][0]))
                    print("|".join(__matching_list[0][1]))
                    print("|".join(__matching_list[0][2]))
                    if len(__matching_list[0][3]) > 1:
                        for index in range(len(__matching_list[0][3])):
                            if index == 0:
                                print("".join(__matching_list[0][3][index]))
                            else:
                                print("|".join(__matching_list[0][3][index]))
                    print("\n")
                    
                    ''' 
                    Write matching output to file
                    '''
                    f.write("".join(__matching_list[0][0]))
                    f.write("\n")
                    f.write("|".join(__matching_list[0][1]))
                    f.write("\n")
                    f.write("|".join(__matching_list[0][2]))
                    f.write("\n")
                    if len(__matching_list[0][3]) > 1:
                        for index in range(len(__matching_list[0][3])):
                            if index == 0:
                                f.write("".join(__matching_list[0][3][index]))
                            else:
                                f.write("|".join(__matching_list[0][3][index]))
                    f.write("\n")
                    
                            
                elif len(__unmatching_list) > 0:
                    
                    ''' 
                    Print matching output on screen
                    '''
                    print("".join(__unmatching_list[0][0]))
                    print("|".join(__unmatching_list[0][1]))
                    print("|".join(__unmatching_list[0][2]))
                    if len(__unmatching_list[0][3]) > 1:
                        for index in range(len(__unmatching_list[0][3])):
                            if index == 0:
                                print("".join(__unmatching_list[0][3][index]))
                            else:
                                print("|".join(__unmatching_list[0][3][index]))
                    print("\n")
                
                    ''' 
                    Write matching output to file
                    '''
                    f.write("".join(__unmatching_list[0][0]))
                    f.write("\n")
                    f.write("|".join(__unmatching_list[0][1]))
                    f.write("\n")
                    f.write("|".join(__unmatching_list[0][2]))
                    f.write("\n")
                    if len(__unmatching_list[0][3]) > 1:
                        for index in range(len(__unmatching_list[0][3])):
                            if index == 0:
                                f.write("".join(__unmatching_list[0][3][index]))
                            else:
                                f.write("|".join(__unmatching_list[0][3][index]))
                    f.write("\n")
                
                else:
                    '''
                    Print and write Header
                    '''
                    if "First Name" in __csv_row:
                        print("|".join(__csv_row))
                        f.write("|".join(__csv_row))
                    else:
                        print("No Match")   
                        print("|".join(__csv_row))
                        f.write("|".join(__csv_row))
                    print("\n")
                    f.write("\n")
                __row_count += 1
                time.sleep(3) #only for debugging
    
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
    
if __name__ == '__main__':
    
    try:
        user_input_filepath = get_filepath_input("Please enter the file path (e.g. C:\\data): ")
        print(f"You entered: {user_input_filepath}")
    except ValueError as e:
        print(e)
    
    try:
        user_input_filename1 = get_filename_input("Please enter the file1 name: ")
        print(f"You entered: {user_input_filename1}")
    except ValueError as e:
        print(e) 
    
    try:
        user_input_filename2 = get_filename_input("Please enter the file2 name: ")
        print(f"You entered: {user_input_filename2}")
    except ValueError as e:
        print(e) 
    
    try:
        user_output_filename_input = get_output_filename_input("Please enter the output file name: ")
        print(f"You entered: {user_output_filename_input}")
    except ValueError as e:
        print(e) 
    
    '''Args'''    
    dir_path = user_input_filepath
    csv_file1 = user_input_filename1
    csv_file2 = user_input_filename2
    output_file = user_output_filename_input
    
    try:
        create_diff_report(dir_path, csv_file1, csv_file2, output_file)
        print(f"CSV file {dir_path} has been processed and written to {output_file}!")
    except ValueError as e:
        print(e) 
    
    
