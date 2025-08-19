import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from openpyxl import load_workbook
import os
from srscalculator import SRSCALCULATOR

# --- Helper Functions ---
def update_status(message):
    """Updates the text of the status label in the GUI."""
    status_label.config(text=message)

def open_excel_file():
    """
    Opens a file dialog for the user to select an Excel file.
    Loads the selected file into a pandas DataFrame, performs calculations,
    and then attempts to update the original Excel file with the results.
    """
    # Open a file dialog to select an Excel file
    filepath = filedialog.askopenfilename(
        title = "Select an Excel File",
        filetypes=(("Excel Files", "*.xlsx *.xls"), ("All files", "*.*"))
    )
    if filepath:
        try:
            # Load the Excel file into a DataFrame, specifying the header row (0-indexed)
            # and the specific sheet name.
            srs_excel = SRSCALCULATOR(filepath)
            srs_excel.sum()
            srs_excel.write_to_file()

            update_status(f"File '{os.path.basename(filepath)}' loaded successfully. Performing calculations...")
            update_status("Calculations complete. Ready to save the modified file.")
            messagebox.showinfo("File Loaded & Calculated", f"File '{os.path.basename(filepath)}' loaded and calculations performed.")

        except FileNotFoundError:
            update_status(f"Error: File not found at '{os.path.basename(filepath)}'.")
            messagebox.showerror("File Error", f"File not found: {os.path.basename(filepath)}")
        except pd.errors.EmptyDataError:
            update_status(f"Error: File '{os.path.basename(filepath)}' is empty or not a valid Excel file.")
            messagebox.showerror("File Error", f"File is empty or not a valid Excel file: {os.path.basename(filepath)}")
        except KeyError as e:
            update_status(f"Error: Sheet 'srs201_work05_sample' not found in file. {e}")
            messagebox.showerror("Sheet Error", f"The required sheet 'srs201_work05_sample' was not found in the Excel file. Please ensure the sheet name is correct. Error: {e}")
        except Exception as e:
            update_status(f"Error during file loading or calculation: {e}")
            messagebox.showerror("Processing Error", f"An error occurred: {e}")




# --- GUI Setup ---
root = tk.Tk()
root.title("SRS File Reader")
root.geometry("500x300") # Set initial window size

# Frame for buttons to organize them visually
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

# Button to open and process the Excel file
open_button = tk.Button(button_frame, text="Open and Modify Excel File", command=open_excel_file)
open_button.pack(side=tk.LEFT, padx=10, pady=5)

# Button to exit the application
exit_button = tk.Button(button_frame, text = "Exit", command = root.destroy)
exit_button.pack(side=tk.RIGHT, padx = 10, pady=5)

# Status label to provide feedback to the user
status_label = tk.Label(root, text="Ready to load an Excel file...", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_label.pack(side=tk.BOTTOM, fill=tk.X)

# Start the Tkinter event loop
root.mainloop()