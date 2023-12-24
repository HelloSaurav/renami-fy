import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


class FileRenamerApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Rename-fy by Saurav")
        self.file_directory = ""
        self.excel_file_path = ""
        self.create_widgets()

    def create_widgets(self):
        # Label and Entry for selecting file directory
        tk.Label(self.master, text="Choose File Directory:").grid(row=0, column=0, padx=10, pady=10)
        self.dir_entry = tk.Entry(self.master, width=50, state='disabled')
        self.dir_entry.grid(row=0, column=1, padx=10, pady=10)
        tk.Button(self.master, text="Browse", command=self.browse_directory).grid(row=0, column=2, padx=10, pady=10)

        # Button to generate Excel
        tk.Button(self.master, text="Generate Excel", command=self.generate_excel).grid(row=1, column=1, pady=20)

        # Label and Entry for selecting Excel file
        tk.Label(self.master, text="Choose Excel File:").grid(row=2, column=0, padx=10, pady=10)
        self.excel_entry = tk.Entry(self.master, width=50, state='disabled')
        self.excel_entry.grid(row=2, column=1, padx=10, pady=10)
        tk.Button(self.master, text="Browse", command=self.browse_excel).grid(row=2, column=2, padx=10, pady=10)

        # Button to start renaming
        tk.Button(self.master, text="Rename Files", command=self.rename_files).grid(row=3, column=1, pady=20)

    def browse_directory(self):
        self.file_directory = filedialog.askdirectory()
        self.dir_entry.config(state='normal')
        self.dir_entry.delete(0, 'end')
        self.dir_entry.insert(0, self.file_directory)
        self.dir_entry.config(state='disabled')

    def generate_excel(self):
        try:
            # Get the list of files in the directory
            files = os.listdir(self.file_directory)

            # Create a DataFrame with files in Row A and add a header for Column B
            df = pd.DataFrame({'old_name': files, 'new_name': ""})

            # Save the DataFrame to Excel
            self.excel_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                                filetypes=[("Excel files", "*.xlsx;*.xls")])
            df.to_excel(self.excel_file_path, index=False)

            messagebox.showinfo("Success", f"Excel file generated successfully at {self.excel_file_path}")

            # Display a pop-up message with instructions
            messagebox.showinfo("Instructions", "Please fill in the 'Custom Names' column with desired names "
                                                "before using the 'Rename Files' button.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def browse_excel(self):
        self.excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        self.excel_entry.config(state='normal')
        self.excel_entry.delete(0, 'end')
        self.excel_entry.insert(0, self.excel_file_path)
        self.excel_entry.config(state='disabled')

    def rename_files(self):
        try:
            # Read the Excel file
            df = pd.read_excel(self.excel_file_path)
            files = df.iloc[:, 0].tolist()
            new_names = df.iloc[:, 1].tolist()

            # Rename files
            for file, new_name in zip(files, new_names):
                old_path = os.path.join(self.file_directory, file)
                new_path = os.path.join(self.file_directory, new_name)

                os.rename(old_path, new_path)

            messagebox.showinfo("Success", "Files renamed successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = FileRenamerApp(root)
    root.mainloop()