import pandas as pd
import os

def create_diff_report(dir_path, csv_file1, csv_file2, output_file):

 # Get the current working directory
    current_directory = os.getcwd()
    print(f"Current directory: {current_directory}")
    
    df1 = pd.read_csv(infile1)
    df2 = pd.read_csv(infile2)
    
    column_dict = dict()
    for df1_col,df2_col in zip(df1.columns, df2.columns):
        column_dict[df2_col] = df1_col

    renamed_df = df2.rename(columns=column_dict)
    
    print(df1.columns)
    print(len(df1))
    print("\nNumber of Rows: ")
    print(renamed_df.columns)
    print(len(renamed_df))
    df_diff_report = df1.merge(renamed_df, on='Last Name', how='left', indicator=True)
    print(df_diff_report)
    
    # Change the current working directory
    # try:
    #     os.chdir(dir_path)
    #     print(f"Directory changed to: {os.getcwd()}")
        
    #     with open(csv_file1, 'r') as infile1, open(csv_file2, 'r') as infile2, open(output_file, 'w', newline='') as outfile:

    #         df1 = pd.read_csv(infile1)
    #         df2 = pd.read_csv(infile2)
    
    #         column_dict = dict()
    #         for df1_col,df2_col in zip(df1.columns, df2.columns):
    #             column_dict[df2_col] = df1_col

    #         renamed_df = df2.rename(columns=column_dict)
    
    #         # print(df1.columns)
    #         # print(len(df1))
    #         # print("\nNumber of Rows: ")
    #         # print(renamed_df.columns)
    #         # print(len(renamed_df))
    #         df_diff_report = df1.merge(renamed_df, on='Last Name', how='left', indicator=True)
    #         print(df_diff_report)
    #         # df_diff_report.to_excel(outfile, engine='xlsxwriter', index=False)
    
    # except FileNotFoundError:
    #     print(f"Error: Directory not found: {dir_path}")
    # except NotADirectoryError:
    #     print(f"Error: Not a directory: {dir_path}")
    # except PermissionError:
    #     print(f"Error: Permission denied to access: {dir_path}")
    # except Exception as e:
    #     print(f"An unexpected error occurred: {e}")    
    
if __name__ == '__main__':
    dir_path = 'data'
    csv_file1 = 'sorted_Plexxis_Tax Details_Ck_Date_121523.csv'
    csv_file2 = 'Zenefits_payroll_detail_Ck Date 121523.csv'
    output_file = 'Zenefits_Plexxis_Tax_comparison.xlsx'
    create_diff_report(dir_path, csv_file1, csv_file2, output_file)
