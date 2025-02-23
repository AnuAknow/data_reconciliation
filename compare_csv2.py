
# from pretty_html_table import build_table
# import pandas as pd
import time
import csv
import os
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
        matching_rows = []
        unmatching_rows = []
        with open(csv_filepath, 'r') as file:
            csv_reader = csv.reader(file)
            for csv_row in csv_reader:
                diff = ['diffs']
                row_contents = []
                if row == csv_row:
                    # print(f"matching:{row}, {csv_row}")
                    row_contents.append("matching") 
                    row_contents.append(row)
                    row_contents.append(csv_row)
                    diff1 = [x for x in row if x not in csv_row]
                    if len(diff1) != 0:
                        # print(f"Diff1 {diff1}")
                        diff.append(diff1)
                    diff2 = [x for x in csv_row if x not in row]
                    if len(diff2) != 0:
                        # print(f"Diff2 {diff2}")
                        diff.append(diff2)
                    if len(diff) != 0: 
                        row_contents.append(diff)
                        matching_rows.append(row_contents) 
                 
                elif row[0] == csv_row[0] and row[1] == csv_row[1] and row != csv_row:
                    # print(f"unmatched:{row}, {csv_row}")
                    row_contents.append("unmatching") 
                    row_contents.append(row)
                    row_contents.append(csv_row)
                    diff1 = [x for x in row if x not in csv_row]
                    if len(diff1) != 0:
                        diff.append(diff1) 
                    diff2 = [x for x in csv_row if x not in row]
                    if len(diff2) != 0:
                        diff.append(diff2)
                    if len(diff) != 0: 
                        row_contents.append(diff)
                        unmatching_rows.append(row_contents) 
                    
        return matching_rows, unmatching_rows
    
    # Get the current working directory
    current_directory = os.getcwd()
    print(f"Current directory: {current_directory}")
    
    # Change the current working directory
    try:
        os.chdir(dir_path)
        print(f"Directory changed to: {os.getcwd()}")

        with open(csv_file2, 'r') as file, open('plexxis_zenefits_table.txt', 'w') as f:
            csv_reader = csv.reader(file)
            row_count = 0
            for csv_row in csv_reader:
                matching_list, unmatching_list = compare_row_to_csv(csv_row, csv_file1)
                if len(matching_list) > 0:
                    print(matching_list[0][0])
                    print(matching_list[0][1])
                    print(matching_list[0][2])
                    print(matching_list[0][3])
    
                    f.write(f"{matching_list}\n")
                elif len(unmatching_list) > 0:
                    print(unmatching_list[0][0])
                    print(unmatching_list[0][1])
                    print(unmatching_list[0][2])
                    print(unmatching_list[0][3])
                    
                    f.write(f"{unmatching_list}\n")
                else:
                    print(f"No match: {csv_row}")
                    f.write(f"{csv_row}\n")
                print("\n")
                row_count += 1
                time.sleep(5)
            print(row_count)
            
            # html_table_blue_light = build_table(df_diff_report_out, 'grey_light')
            # html_table = df_diff_report_out.to_html()
            
            # Save to html file
            # with open('plexxis_zenefits_table.html', 'w') as f:
            #     f.write(html_table_blue_light)
        
            
            # diff_df.to_csv(outfile, index=False)
    
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
