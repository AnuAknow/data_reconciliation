# Employee Payroll & Data Reconciliation Scripts  

This repository contains Python scripts designed for **efficient payroll data processing and reconciliation of financial records**. These tools streamline the extraction, transformation, and validation of employee payroll data while comparing and identifying discrepancies between datasets.  

## Features  
✔ **Automated Payroll Data Extraction** – Extract employee details, payroll taxes, and paid taxes from Excel files.  
✔ **Data Integration** – Combine extracted data into a single structured dataset, ensuring accuracy.  
✔ **File Reconciliation** – Compare CSV files with matching identifiers to detect discrepancies.  
✔ **Security & Compliance** – Supports data integrity checks and GDPR/HIPAA compliance validation.  

## Table of Contents  
- [Requirements](#requirements)  
- [Usage](#usage)  
- [Functions](#functions)  
- [Installation](#installation)  
- [License](#license)  
- [Contact](#contact)  

## Requirements  
Ensure you have the following installed:  
- **Python 3.x**  
- **openpyxl** (for Excel data handling)  
- **locale** (for localized financial processing)  
- **pandas** (for CSV data comparison)  

Install dependencies using:  
```sh
pip install openpyxl pandas
```

## Usage  

### Payroll Data Processing  
1. Clone the repository.  
2. Place your **Excel payroll file** inside the `data` directory.  
3. Update the script's **file path** (`v_path`) and **file name** (`v_file`).  
4. Run the payroll processing script:  
```sh
python payroll_processing.py
```  
This will extract employee details, payroll taxes, and paid taxes, then compile them into a structured CSV file.  

### Reconciliation Script  
1. Place CSV files into the target directory following the naming convention:  
   - `plexus<number>.csv`  
   - `lumber<number>.csv`  
2. Run the reconciliation script:  
```sh
python compare_files.py
```  
The script will compare matching files and highlight differences between datasets.  

## Functions  

### Payroll Processing  
- **extract_employee_info(dirpath, filename, heading)** – Extracts employee details.  
- **payroll_taxes(dirpath, filename, heading)** – Retrieves payroll tax data.  
- **extract_taxes_paid(dirpath, filename, heading)** – Extracts tax payment records.  
- **combine_list(x, y, z)** – Merges multiple lists into a single dataset.  
- **combine_data_lists(x, y, z)** – Consolidates structured data into a unified format.  
- **write_to_file_in_directory(dirpath, filename, data)** – Saves processed data as a CSV file.  

### Data Reconciliation  
- **compare_files(dir_path)** – Identifies mismatches in CSV datasets and prints discrepancies.  

## License  
This project is licensed under the **MIT License**.  

## Contact  
For inquiries, reach out via email:  
📧 **everettaknowledge50@gmail.com**  
