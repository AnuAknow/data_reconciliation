Reconciliation Script
This Python script compares all files in a specified directory with corresponding numbers in their filenames (e.g., "plexus<number>.csv" and "lumber<number>.csv") and performs a data comparison. The script lists the files, filters the ones that start with "plexus" and "lumber," extracts the numbers from the filenames, and identifies differences between the pairs.

Features
Compares files with corresponding numbers in their filenames.

Identifies and prints differences between the files.

Supports CSV files with comma-delimited data.

Prerequisites
Python 3.x

Pandas library

Installation
Clone the repository or download the script files.

Ensure you have Python 3.x installed on your system.

Install the Pandas library using pip:

sh
pip install pandas
Usage
Place your CSV files in the specified directory.

Ensure the CSV files follow the naming convention "plexus<number>.csv" and "lumber<number>.csv".

Run the script:

sh
python compare_files.py
The script will compare the files in the specified directory and print the differences.

Script Details
python
import os
import pandas as pd

def compare_files(dir_path):
    # List of files in the directory
    files = os.listdir(dir_path)
    
    # Filter files that start with "plexus" and "lumber"
    plexus_files = [f for f in files if f.startswith("plexus") and f.endswith(".csv")]
    lumber_files = [f for f in files if f.startswith("lumber") and f.endswith(".csv")]

    for plexus_file in plexus_files:
        # Extract the number from the file name
        number = plexus_file[len("plexus"):-len(".csv")]
        lumber_file = f"lumber{number}.csv"

        if lumber_file in lumber_files:
            print(f"Comparing {plexus_file} and {lumber_file}")
            data1 = pd.read_csv(os.path.join(dir_path, plexus_file))
            data2 = pd.read_csv(os.path.join(dir_path, lumber_file))

            # Merge the DataFrames to identify differences
            compare_data = data1.compare(data2, align_axis=0, keep_shape=True, keep_equal=True, result_names=(plexus_file, lumber_file))

            print(compare_data)
        else:
            print(f"File {lumber_file} not found for comparison with {plexus_file}")

# Example usage
compare_files('data')
Example
To run the script with example data:

Create a directory named data.

Place the following sample CSV files in the data directory:

plexus1.csv

ID,Name,Age,Department
1,John Doe,28,Engineering
2,Jane Smith,34,Marketing
3,Robert Johnson,25,Sales
4,Alice Brown,30,Engineering
5,Tom Clark,40,HR
lumber1.csv

ID,Name,Age,Department
1,John Doe,28,Engineering
2,Jane Smith,35,Marketing
3,Robert Johnson,25,Sales
4,Alice Brown,30,Engineering
6,Nancy Davis,45,Finance
Run the script:

sh
python compare_files.py
The script will compare the files and print the differences.

License
This project is licensed under the MIT License.

Contact
If you have any questions or need further assistance, please feel free to contact me at:

Email: everettaknowledge50@gmail.com