
from pretty_html_table import build_table
import pandas as pd
import os

def create_diff_report(dir_path, csv_file1, csv_file2, output_file):
    
    # Get the current working directory
    current_directory = os.getcwd()
    print(f"Current directory: {current_directory}")
    
    # Change the current working directory
    try:
        if dir_path not in current_directory:
            os.chdir(dir_path)
            print(f"Directory changed to: {os.getcwd()}")
        
        with open(csv_file1, 'r') as infile1, open(csv_file2, 'r') as infile2, open(output_file, 'w', newline='') as outfile:

            df1 = pd.read_csv(infile1)
            df2 = pd.read_csv(infile2)
    
            column_dict = dict()
            for df1_col,df2_col in zip(df1.columns, df2.columns):
                column_dict[df2_col] = df1_col

            df_w_updated_cols = df2.rename(columns=column_dict)
    
            print(df1.columns)
            # print(len(df1))
            # print("\nNumber of Rows: ")
            print(df_w_updated_cols.columns)
            # print(len(renamed_df))
            df_diff_report_out = df1.merge(df_w_updated_cols, on=['Last Name','First Name'], how='outer', indicator=True)
            df_diff_report_left = df1.merge(df_w_updated_cols, on=['Last Name','First Name'], how='left', indicator=True)
            df_diff_report_right = df1.merge(df_w_updated_cols, on=['Last Name','First Name'], how='right', indicator=True)
            # Find differences (full outer join)
            # diff_df = df1.merge(df2, indicator=True, how='outer')
            # Show only differences
            # diff_df = diff_df[diff_df['_merge'] != 'both']
            # print(diff_df)
            
            print(df_diff_report_out)
            print(df_diff_report_left)
            print(df_diff_report_right)
            
            html_table_blue_light = build_table(df_diff_report_out, 'grey_light')
            # html_table = df_diff_report_out.to_html()
            
            # Save to html file
            with open('plexxis_zenefits_table.html', 'w') as f:
                f.write(html_table_blue_light)
        
            
            # diff_df.to_csv(outfile, index=False)
    
    except FileNotFoundError:
        print(f"Error: Directory not found: {dir_path}")
    except NotADirectoryError:
        print(f"Error: Not a directory: {dir_path}")
    except PermissionError:
        print(f"Error: Permission denied to access: {dir_path}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")    
    
if __name__ == '__main__':
    dir_path = 'data'
    csv_file1 = 'sorted_Plexxis_Tax Details_Ck_Date_121523.csv'
    csv_file2 = 'Zenefits_payroll_detail_Ck Date 121523.csv'
    # output_file = 'Zenefits_Plexxis_Tax_comparison.xlsx'
    output_file = 'Zenefits_Plexxis_Tax_merge.csv'
    create_diff_report(dir_path, csv_file1, csv_file2, output_file)
    
    print(f"CSV file {dir_path} has been processed and written to {output_file}!")
