import pandas as pd

# Get the file path from user input
file_path = input("Enter the path of the CSV file: ")

# Load the CSV file as a DataFrame
df = pd.read_csv(file_path)

# Remove the financial columns
df = df.drop(columns=["Hourly Rate", "Billable Amount", "Billable Currency"])

# Print the updated DataFrame
print(df)