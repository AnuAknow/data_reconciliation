
from tabulate import tabulate
import time
import csv
import os
import re
import sys

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
        os.chdir(dir_path)
        print(f"Directory changed to: {os.getcwd()}")

        with open(csv_file2, 'r') as file, open('plexxis_zenefits_table.txt', 'w') as f:
            __csv_reader = csv.reader(file)
            __row_count = 0

            for __csv_row in __csv_reader:
                __matching_list, __unmatching_list = compare_row_to_csv(__csv_row, csv_file1)
                if len(__matching_list) > 0:
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
                    
                 
                
                    # f.write(__matching_list[0][0])
                    # f.write(__matching_list[0][1])
                    # f.write(__matching_list[0][2])
                    # f.write(__matching_list[0][3])
                    # if len(__unmatching_list[0][3]) > 1:
                    #     for list in __unmatching_list[0][3]:
                    #         f.write(list)
                    
                            
                elif len(__unmatching_list) > 0:
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
                
                    # f.write(__unmatching_list[0][0])
                    # f.write(__unmatching_list[0][1])
                    # f.write(__unmatching_list[0][2])
                    # f.write(__unmatching_list[0][3]) 
                    # if len(__unmatching_list[0][3]) > 1:
                    #     for list in __unmatching_list[0][3]:
                    #         f.write(list)
                        
                        
                    
                else:
                    if "First Name" in __csv_row:
                        print("|".join(__csv_row))
                    else:
                        print("No Match")   
                        print("|".join(__csv_row))
                    print("\n")
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
    dir_path = 'data'
    csv_file1 = 'sorted_Plexxis_Tax_Details_Ck_Date_121523.csv'
    csv_file2 = 'Zenefits_payroll_detail_Ck_Date_121523.csv'
    output_file = 'Zenefits_Plexxis_Tax_merge.txt'
    create_diff_report(dir_path, csv_file1, csv_file2, output_file)
    
    print(f"CSV file {dir_path} has been processed and written to {output_file}!")
