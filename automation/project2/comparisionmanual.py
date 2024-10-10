import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog
import os

# Function to get the file path using a file dialog
def get_file_path(file_description):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(title=f"Select {file_description} file",
                                           filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")])
    return file_path

# Function to get GW Code from user
def get_gw_code():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    gw_code = simpledialog.askstring("Input", "Please enter the GW code:")
    root.destroy()  # Close the hidden window
    return gw_code

# Function to load data from either an Excel or CSV file based on the file extension
def load_data(file_path):
    file_extension = os.path.splitext(file_path)[-1].lower()
    try:
        if file_extension == '.xlsx':
            return pd.read_excel(file_path, engine='openpyxl')
        elif file_extension == '.csv':
            return pd.read_csv(file_path)
        else:
            print(f"Unsupported file format: {file_extension}")
            return None
    except Exception as e:
        print(f"Error reading file: {e}")
        return None

# Get the MySQL data file path using file dialog
mysql_file_path = get_file_path("MySQL configuration")

# Get the Excel file path using file dialog
excel_file_path = get_file_path("Default Excel parameters")

# Load the MySQL and Excel data into DataFrames
mysql_data = load_data(mysql_file_path)
excel_data = load_data(excel_file_path)

if mysql_data is None or excel_data is None:
    print("Error loading one or both data files. Exiting.")
    exit()

# Get GW Code from user
gw_code = get_gw_code()

# Check if the GW code exists in Excel columns
if gw_code not in excel_data.columns:
    print(f"GW code '{gw_code}' not found in Excel columns.")
    exit()

# Sort dataframes by 'SettingName' for easier comparison
mysql_data.sort_values(by='SettingName', inplace=True)
excel_data.sort_values(by='SettingName', inplace=True)

# Reset index for both to avoid index mismatches
mysql_data.reset_index(drop=True, inplace=True)
excel_data.reset_index(drop=True, inplace=True)

# Create a new DataFrame from the Excel data, keeping 'SettingName' and the GW code
excel_filtered = excel_data[['SettingName', gw_code]].rename(columns={gw_code: 'value_excel'})

# Debug: Check the structure of the DataFrames before merging
print("MySQL DataFrame columns before merge:", mysql_data.columns.tolist())
print("Filtered Excel DataFrame columns before merge:", excel_filtered.columns.tolist())

# Merge the MySQL data with the filtered Excel data based on the 'SettingName' column
comparison = pd.merge(mysql_data, excel_filtered, on='SettingName', how='outer', suffixes=('_mysql', '_excel'), indicator=True)

# Debug: Check the columns of the merged DataFrame
print("Comparison DataFrame columns after merge:", comparison.columns.tolist())

# Ensure the 'value_mysql' column exists in the MySQL data
if 'value_mysql' not in mysql_data.columns:
    print("Error: 'value_mysql' column not found in the MySQL data.")
    exit()

# Find discrepancies
differences = []

# For rows that exist in both but with different values
if 'value_mysql' in comparison.columns and 'value_excel' in comparison.columns:
    diff_in_value = comparison[(comparison['_merge'] == 'both') & (comparison['value_mysql'] != comparison['value_excel'])]
else:
    print("Error: 'value_mysql' or 'value_excel' not found in the comparison DataFrame.")
    exit()

# For rows that exist in MySQL but not in Excel
only_in_mysql = comparison[comparison['_merge'] == 'left_only']

# For rows that exist in Excel but not in MySQL
only_in_excel = comparison[comparison['_merge'] == 'right_only']

# Prepare differences output
for idx, row in diff_in_value.iterrows():
    differences.append(f"Value mismatch for '{row['SettingName']}': MySQL = {row['value_mysql']}, Excel = {row['value_excel']}")

for idx, row in only_in_mysql.iterrows():
    differences.append(f"Name '{row['SettingName']}' exists in MySQL but not in Excel")

for idx, row in only_in_excel.iterrows():
    differences.append(f"Name '{row['SettingName']}' exists in Excel but not in MySQL")

# Print the differences found
if differences:
    print("Differences found:")
    for diff in differences:
        print(diff)
else:
    print("No differences found.")

# Specify the path where you want to save the differences report
output_path = filedialog.asksaveasfilename(title="Save Differences Report", defaultextension=".csv", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])

# Optionally, you can save the differences to a file
if output_path:
    differences_df = pd.DataFrame(differences, columns=['Differences'])
    differences_df.to_csv(output_path, index=False)
    print(f"Differences report saved to: {output_path}")
else:
    print("Save operation was cancelled.")
