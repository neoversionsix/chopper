import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class ChopperApp:
    def __init__(self, master):
        self.master = master
        master.title("File Chopper")
        
        # Variables to hold file paths and user input
        self.input_file = ''
        self.output_dir = ''
        self.num_rows = tk.IntVar()
        
        # Widgets
        self.create_widgets()
        
    def create_widgets(self):
        # Input file selection
        self.label1 = tk.Label(self.master, text="Select input file:")
        self.label1.grid(row=0, column=0, padx=5, pady=5, sticky='e')
        
        self.input_file_entry = tk.Entry(self.master, width=50)
        self.input_file_entry.grid(row=0, column=1, padx=5, pady=5)
        
        self.browse_input_button = tk.Button(self.master, text="Browse", command=self.browse_input_file)
        self.browse_input_button.grid(row=0, column=2, padx=5, pady=5)
        
        # Number of rows per file
        self.label2 = tk.Label(self.master, text="Number of rows per file:")
        self.label2.grid(row=1, column=0, padx=5, pady=5, sticky='e')
        
        self.num_rows_entry = tk.Entry(self.master, textvariable=self.num_rows)
        self.num_rows_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # Output directory selection
        self.label3 = tk.Label(self.master, text="Select output directory:")
        self.label3.grid(row=2, column=0, padx=5, pady=5, sticky='e')
        
        self.output_dir_entry = tk.Entry(self.master, width=50)
        self.output_dir_entry.grid(row=2, column=1, padx=5, pady=5)
        
        self.browse_output_button = tk.Button(self.master, text="Browse", command=self.browse_output_dir)
        self.browse_output_button.grid(row=2, column=2, padx=5, pady=5)
        
        # Start Chopping button
        self.start_button = tk.Button(self.master, text="Start Chopping", command=self.start_chopping)
        self.start_button.grid(row=3, column=1, padx=5, pady=20)
        
    def browse_input_file(self):
        filetypes = [("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
        filename = filedialog.askopenfilename(title="Select input file", filetypes=filetypes)
        if filename:
            self.input_file = filename
            self.input_file_entry.delete(0, tk.END)
            self.input_file_entry.insert(0, filename)
        
    def browse_output_dir(self):
        directory = filedialog.askdirectory(title="Select output directory")
        if directory:
            self.output_dir = directory
            self.output_dir_entry.delete(0, tk.END)
            self.output_dir_entry.insert(0, directory)
        
    def start_chopping(self):
        if not self.input_file:
            messagebox.showerror("Error", "Please select an input file.")
            return
        if not self.output_dir:
            messagebox.showerror("Error", "Please select an output directory.")
            return
        try:
            num_rows = int(self.num_rows_entry.get())
            if num_rows <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number of rows per file.")
            return
        
        self.chop_file()
        
    def chop_file(self):
        input_file = self.input_file
        output_dir = self.output_dir
        num_rows = int(self.num_rows_entry.get())
        base_filename = os.path.basename(input_file)
        filename_no_ext, ext = os.path.splitext(base_filename)
        
        # Determine file type and process accordingly
        if ext.lower() == '.csv':
            self.chop_csv(input_file, output_dir, filename_no_ext, num_rows)
        elif ext.lower() == '.xlsx':
            self.chop_excel(input_file, output_dir, filename_no_ext, num_rows)
        else:
            messagebox.showerror("Error", "Unsupported file type.")
            return
        messagebox.showinfo("Success", "File has been chopped successfully.")
        
    def chop_csv(self, input_file, output_dir, filename_no_ext, num_rows):
        try:
            chunk_iter = pd.read_csv(input_file, chunksize=num_rows)
            for i, chunk in enumerate(chunk_iter):
                output_file = os.path.join(output_dir, f"{filename_no_ext}_{i+1}.csv")
                chunk.to_csv(output_file, index=False)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while processing CSV file: {e}")
        
    def chop_excel(self, input_file, output_dir, filename_no_ext, num_rows):
        try:
            df = pd.read_excel(input_file)
            total_rows = df.shape[0]
            for i in range(0, total_rows, num_rows):
                chunk = df.iloc[i:i+num_rows]
                output_file = os.path.join(output_dir, f"{filename_no_ext}_{(i // num_rows)+1}.xlsx")
                chunk.to_excel(output_file, index=False)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while processing Excel file: {e}")
        
if __name__ == '__main__':
    root = tk.Tk()
    app = ChopperApp(root)
    root.mainloop()