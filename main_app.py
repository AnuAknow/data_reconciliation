# main_app.py (Tkinter UI file)

import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess  # For running the processing script

def browse_file():
    filepath = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
    if filepath:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, filepath)

def process_file():
    filepath = file_entry.get()
    if not filepath:
        messagebox.showerror("Error", "Please select an Excel file.")
        return

    start_row = start_row_entry.get()
    stop_row = stop_row_entry.get()

    if not start_row or not stop_row:
        messagebox.showerror("Error", "Please enter start and stop rows.")
        return

    try:
        start_row = int(start_row)
        stop_row = int(stop_row)
    except ValueError:
        messagebox.showerror("Error", "Invalid start or stop row. Must be integers.")
        return


    try:
        # Construct the command to run your processing script.
        # It's crucial to pass the parameters correctly.
        command = [
            "python",  # Or "python3" depending on your setup.
            "process_excel.py",  # The name of your processing script.
            filepath,
            str(start_row),
            str(stop_row)
        ]

        # Use subprocess.run to execute the script and capture output
        result = subprocess.run(command, capture_output=True, text=True, check=True)

        # Display the output in a text box
        output_text.delete("1.0", tk.END)  # Clear previous output
        output_text.insert(tk.END, result.stdout)

        messagebox.showinfo("Success", "File processed successfully!")

    except FileNotFoundError:
        messagebox.showerror("Error", "Processing script not found (process_excel.py).")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"Error during processing:\n{e.stderr}")
    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")



root = tk.Tk()
root.title("Excel Data Extractor")

# File selection
file_label = tk.Label(root, text="Excel File:")
file_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")  # Align left

file_entry = tk.Entry(root, width=50)
file_entry.grid(row=0, column=1, padx=5, pady=5)

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=5, pady=5)

# Row input fields
start_row_label = tk.Label(root, text="Start Row:")
start_row_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")

start_row_entry = tk.Entry(root, width=10)
start_row_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

stop_row_label = tk.Label(root, text="Stop Row:")
stop_row_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")

stop_row_entry = tk.Entry(root, width=10)
stop_row_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")


# Process button
process_button = tk.Button(root, text="Process", command=process_file)
process_button.grid(row=3, column=0, columnspan=3, padx=5, pady=10)

# Output Text Box
output_text = tk.Text(root, wrap=tk.WORD, height=10) # Set a height
output_text.grid(row=4, column=0, columnspan=3, padx=5, pady=5)

root.mainloop()