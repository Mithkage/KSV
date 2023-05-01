import tkinter as tk
from tkinter import filedialog
import pandas as pd
import Timesheet_Functions as TF
from fpdf import FPDF

# Create a Tkinter window to browse for the CSV file
root = tk.Tk()
root.withdraw()  # Hide the root window

# Ask the user to select the CSV file
# file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
file_path = 'wsp-march.csv'

# cleaned_df = pd.DataFrame()
cleaned_df = TF.clean_data(file_path)

# Print the resulting dataframe
print(cleaned_df)


TF.export_data(cleaned_df,'CSV_Time.csv', 'XLSX_Time.xlsx')

