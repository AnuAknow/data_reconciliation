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
            compare_data = data1.compare(data2, align_axis=0, keep_shape=True, keep_equal=False, result_names=(plexus_file, lumber_file))

            print(compare_data)
        else:
            print(f"File {lumber_file} not found for comparison with {plexus_file}")

if __name__ == '__main__':
    # Example usage
    compare_files('data')
