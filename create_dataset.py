from decimal import Decimal, getcontext
import locale
import openpyxl
import re

# Set the desired precision (e.g., for typical currency with 2 decimal places)
getcontext().prec = 2

locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')
l_headings = list()
l_taxes = list()
l_taxes_paid = list()
l_emp_info = list()
d_employee_info = dict()


def extract_employee_info(dirpath, filename, heading):
    filepath = dirpath + filename
    
    col_b_index = openpyxl.utils.column_index_from_string('B')
    # Load the workbook and select the active sheet
    wb = openpyxl.load_workbook(filepath)
    sheet = wb.active

    # Find the "Employees" heading and get the row number
    employees_heading = None
    for row in sheet.iter_rows(min_row=1, max_row=698, min_col=1, max_col=col_b_index):
        for cell in row:
            if cell.value == heading:
                employees_heading = cell.row
                break
        if employees_heading:
            break

    # Check if "Employees" heading was found
    if not employees_heading:
        print("The 'Employees' heading was not found in the worksheet.")
        return

    # Extract employee information
    # heading_list = list()
    split_string_list = []
    employee_data = {}
    filtered_employee_data = []
    # current_employee = []
    start_extracting = False

    # Iterate over rows in the worksheet and return each row as a tuple of cell objects. 
    for row in sheet.iter_rows(min_row=employees_heading, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        for cell in row:
            current_employee = []
            
            # Stop extracting data if "NetPay" or 'Grand Totals
            if cell.value == "NetPay" and cell.value != 'Grand Totals':
                start_extracting = False
                employee_data[f"Employee_{len(employee_data) + 1}"] = re.sub(r"\n+", ", ", current_employee, count=0)
                break
            if start_extracting:
                merged_cell_value = cell.value
                if merged_cell_value != None:
                    if cell.coordinate in sheet.merged_cells:
                        if sheet[cell.coordinate].value != None:
                            add_fname_title = "First Name: " + sheet[cell.coordinate].value
                            add_lname_title = re.sub(r",", " ,Last Name:", add_fname_title, count=1)
                            merged_cell_value = re.sub(r"\n+", ", ", add_lname_title, count=0)
                            
                            if merged_cell_value != 'First Name: Grand Totals':
                                split_string_list = merged_cell_value.split(",")
                                current_employee.append(split_string_list)
                                for info in current_employee:
                                    heading_list =list()
                                    filtered_data = list()
                                    emp_info_heading = list()
                                    for i in range(len(info)):
                                        emp_info_heading=info[i].split(":") 
                                        if len(emp_info_heading) == 2:
                                            if emp_info_heading[0] not in heading_list:
                                                heading_list.append(emp_info_heading[0])
                                                filtered_data.append(emp_info_heading[1])
                                                # filtered_employee_data.append(filtered_data)
                                    filtered_employee_data.append(filtered_data)            
        if cell.value:
            start_extracting = True

        # Stop extracting if an empty cell is encountered
        if not any(cell.value for cell in row):
            break
    return heading_list,filtered_employee_data            


def payroll_taxes(dirpath, filename, heading):
    filepath = dirpath + filename
    
    __taxes = list()
    __all_taxes = list()
    __tax_headings = list()
    taxes = str()

    
    # Load the workbook and select the active sheet
    wb = openpyxl.load_workbook(filepath)
    sheet = wb.active

    # Find the row containing the heading
    heading_row = None
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == heading:
                heading_row = cell.row
                break
        if heading_row:
            break
    
    if heading_row is None:
        print(f"Heading '{heading}' not found.")

    # Print the values from every 10 cells in column J after the heading
    col_j_index = openpyxl.utils.column_index_from_string('J')
    col_i_index = openpyxl.utils.column_index_from_string('I')
    row_index = heading_row + 1
    while True:
        cell_value = sheet.cell(row=row_index, column=col_j_index).value
        cell_tax_title = sheet.cell(row=row_index, column=col_i_index).value
        
        if cell_value is None:
            return __tax_headings, __all_taxes
            
        if cell_tax_title == 'CA California state tax':
            __all_taxes.append(__taxes)
            __taxes = list()
            row_index += 4
            
        if cell_value is not None: 
            if cell_tax_title not in __tax_headings:
                __tax_headings.append(cell_tax_title)
            __taxes.append(f"${cell_value:,.2f}") 
            row_index += 1
            
            
def extract_taxes_paid(dirpath, filename, heading):
    filepath = dirpath + filename
    
    # Array for taxes paid
    __taxes = list()
    # Load the workbook and select the active sheet
    wb = openpyxl.load_workbook(filepath)
    sheet = wb.active

    # Find the row containing the heading
    heading_row = None
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == heading:
                heading_row = cell.row
                break
        if heading_row:
            break
    
    if heading_row is None:
        print(f"Heading '{heading}' not found.")

    # Print the values from every 10 cells in column J after the heading
    col_j_index = openpyxl.utils.column_index_from_string('J')
    row_index = heading_row + 10
    
    while True:
        cell_value = sheet.cell(row=row_index, column=col_j_index).value
        if cell_value is None:
            break
        
        if cell_value is not heading:
            taxes = locale.currency(cell_value, grouping=True) 
            __taxes.append(taxes) 
        row_index += 10
        
    return __taxes

def combine_list(x, y, z):
    x.extend(y)
    x.append(z)
    return x

def combine_data_lists(x, y, z):
    for i in range(len(x)):
        x[i].extend(y[i])
        x[i].append(z[i])
    return x

def write_to_file_in_directory(dirpath, filename, headings, emp_data):
    filename = re.sub("xlsx", "csv", filename)
    filepath = dirpath + filename
    emp_data.insert(0, headings)
        
    with open(filepath, 'w') as file:
        for line in emp_data:
            file.write(str(line) + '\n')

if __name__ == '__main__':
    v_path = 'data\\'
    v_file = 'Zenefits_payroll_detail_Ck Date 121523.xlsx'

    # Example usage:
    l_emp_info_heading, l_emp_info = extract_employee_info(v_path, v_file, heading='Employees')
    # print(l_emp_info) 
    # print(l_emp_info_heading)  
   
    l_payroll_taxes_headings, l_payroll_taxes = payroll_taxes(v_path, v_file, heading='Amount')
    # print(l_payroll_taxes)
    # print(l_payroll_taxes_headings)
    
    # Example usage
    l_taxes_paid = extract_taxes_paid(v_path, v_file, heading='Amount')
    v_taxes_paid_heading= 'Taxes Paid'
    # print(v_taxes_paid_heading)

    # Example usage
    l_combined_headings = combine_list(l_emp_info_heading, l_payroll_taxes_headings, v_taxes_paid_heading)
    # print(l_combined_headings)

    # Example usage
    l_combined_emp_tax = combine_data_lists(l_emp_info, l_payroll_taxes, l_taxes_paid)
    # print(l_combined_emp_tax)
    
    # Example usage
    write_to_file_in_directory(v_path, v_file, l_combined_headings, l_combined_emp_tax)
    