import tkinter as tk
from tkinter import filedialog
import pandas as pd

# Create a Tkinter window to browse for the CSV file
root = tk.Tk()
root.withdraw()  # Hide the root window

# Ask the user to select the CSV file
file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])

# Load the CSV file as a dataframe
df = pd.read_csv(file_path)

# Remove the financial columns
df = df.drop(['Hourly Rate', 'Billable Amount', 'Billable Currency'], axis=1)

# remove the "Invoice" and "ID" columns
df = df.drop(["Invoice", "ID"], axis=1)

# Print the resulting dataframe
print(df)

# Export to an XLSX file
output_file = 'output.xlsx'
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()