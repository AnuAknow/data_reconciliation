import csv
from openpyxl import Workbook

def compare_csv(file1_path, file2_path, output_file_path="differences.xlsx"):
    """
    Compares two CSV files and writes the differences to an Excel file.

    Args:
        file1_path (str): Path to the first CSV file.
        file2_path (str): Path to the second CSV file.
        output_file_path (str, optional): Path to save the Excel output. Defaults to "differences.xlsx".
    """
    with open(file1_path, 'r') as file1, open(file2_path, 'r') as file2:
        reader1 = csv.reader(file1)
        reader2 = csv.reader(file2)

        # Store differences as tuples (row_index, row_from_file1, row_from_file2)
        differences = []
        row_index = 0
        for row1, row2 in zip(reader1, reader2):
            if row1 != row2:
                differences.append((row_index, row1, row2))
            row_index += 1

    # Create a new workbook and select the active sheet
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Differences"

    # Write headers
    sheet.append(["Row Index", "File 1 Row", "File 2 Row"])

    # Write differences to the sheet
    for diff in differences:
        # Convert lists to strings before appending
        diff = (diff[0], str(diff[1]), str(diff[2]))
        sheet.append(diff)

    # Save the workbook
    workbook.save(output_file_path)

if __name__ == "__main__":
    # Example usage:
    data_path = "data\\"
    file1_path = data_path + "Plexxis_Tax Details_Ck Date 121523.csv"
    file2_path = data_path + "Zenefits_payroll_detail_Ck Date 121523.csv"
    
    compare_csv(file1_path, file2_path)