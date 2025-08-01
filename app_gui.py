import tkinter as tk
from tkinter import filedialog, messagebox
import os
import csv

def compare_csv_files(file1, file2, output_file):
    with open(file1, 'r') as f1, open(file2, 'r') as f2:
        reader1 = list(csv.reader(f1))
        reader2 = list(csv.reader(f2))

    header = reader1[0]
    rows1 = reader1[1:]
    rows2 = reader2[1:]

    diffs = []
    for row1 in rows1:
        if row1 not in rows2:
            diffs.append(["Missing in File 2"] + row1)
    for row2 in rows2:
        if row2 not in rows1:
            diffs.append(["Missing in File 1"] + row2)

    with open(output_file, 'w', newline='') as f_out:
        writer = csv.writer(f_out)
        writer.writerow(["Source"] + header)
        writer.writerows(diffs)

class TestFireReconcilerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Test Fire: Data Reconciler")
        self.root.geometry("600x300")
        
        tk.Label(root, text="Select File 1 (CSV):").pack()
        self.entry_file1 = tk.Entry(root, width=80)
        self.entry_file1.pack()
        tk.Button(root, text="Browse", command=self.browse_file1).pack()

        tk.Label(root, text="Select File 2 (CSV):").pack()
        self.entry_file2 = tk.Entry(root, width=80)
        self.entry_file2.pack()
        tk.Button(root, text="Browse", command=self.browse_file2).pack()

        tk.Label(root, text="Select Output File:").pack()
        self.entry_output = tk.Entry(root, width=80)
        self.entry_output.pack()
        tk.Button(root, text="Save As", command=self.save_output_file).pack()

        tk.Button(root, text="Run Reconciliation", command=self.run_comparison).pack(pady=10)

    def browse_file1(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.entry_file1.delete(0, tk.END)
            self.entry_file1.insert(0, file_path)

    def browse_file2(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.entry_file2.delete(0, tk.END)
            self.entry_file2.insert(0, file_path)

    def save_output_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, file_path)

    def run_comparison(self):
        file1 = self.entry_file1.get()
        file2 = self.entry_file2.get()
        output_file = self.entry_output.get()

        if not file1 or not file2 or not output_file:
            messagebox.showerror("Error", "Please select all files.")
            return

        try:
            compare_csv_files(file1, file2, output_file)
            messagebox.showinfo("Success", f"Comparison complete.\nReport saved to:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = TestFireReconcilerApp(root)