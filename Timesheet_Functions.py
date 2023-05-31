import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Create a Tkinter window to browse for the CSV file
root = tk.Tk()
root.withdraw()  # Hide the root window
# Ask the user to select the CSV file
file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])

def clean_data(file_path):
    # Load the CSV file as a dataframe
    df = pd.read_csv(file_path)

    # Remove the financial columns and unnecessary columns
    df = df.drop(['ID', 'Invoice', 'Hourly Rate', 'Duration (mins)', 'Billable Amount', 'Billable Currency'], axis=1)

    # Add a "Day of Week" column based on the "Date" column
    df['Day of Week'] = pd.to_datetime(df['Date']).dt.day_name()

    # Reorder the columns
    df = df[['Contact','Project Name','Tasks','Date', 'Day of Week', 'Duration (hours)', 'Description']]

    # Sort the dataframe in ascending order based on the "Date" column
    df = df.sort_values(by='Date')

    return df


def export_data(df, csv_file_name, xlsx_file_name):
    # Create or override the CSV file with the current data
    df.to_csv(csv_file_name, index=False)
    
    # Export data to XLSX format
    df.to_excel(xlsx_file_name, index=False)
    
    print(f"Data exported to {csv_file_name} and {xlsx_file_name} successfully.")




