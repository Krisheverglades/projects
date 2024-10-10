import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog
import os
import subprocess  # To open the file automatically

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

# Define the default Excel file (it should be in the same folder as this script)
script_dir = os.path.dirname(os.path.abspath(__file__))
default_excel_path = os.path.join(script_dir, 'Defaultparameters.xlsx')

# Check if the default Excel file exists
if not os.path.exists(default_excel_path):
    print(f"Default Excel file not found at: {default_excel_path}")
    exit()

# Get the MySQL data file path using file dialog
mysql_file_path = get_file_path("MySQL configuration")

# Load the MySQL and Excel data into DataFrames
mysql_data = load_data(mysql_file_path)
excel_data = load_data(default_excel_path)

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
excel_filtered = excel_data[['SettingName', gw_code]].rename(columns={gw_code: 'value_default'})

# Merge the MySQL data with the filtered Excel data based on the 'SettingName' column
comparison = pd.merge(mysql_data, excel_filtered, on='SettingName', how='outer', suffixes=('_mysql', '_default'))

# Sort the comparison DataFrame by 'SettingName'
comparison.sort_values(by='SettingName', inplace=True)

# Reset index after sorting
comparison.reset_index(drop=True, inplace=True)

# Prepare final DataFrame for output, referencing the correct column for MySQL values
final_output = comparison[['SettingName', 'SettingValue', 'value_default']].rename(columns={'SettingValue': 'value_mysql'})

# Function to apply cell-wise highlighting for only wrong values and missing parameters
def highlight_row(row):
    if pd.isna(row['value_mysql']) or pd.isna(row['value_default']):
        # Missing parameters highlighted in red
        return ['background-color: red'] * len(row)
    elif row['value_mysql'] != row['value_default']:
        # Wrong values highlighted in yellow
        return ['background-color: yellow'] * len(row)
    return [''] * len(row)  # No highlight if values are the same

# Apply the highlight function row-wise
final_styled = final_output.style.apply(highlight_row, axis=1)

# Specify the path where you want to save the final output (replace if it exists)
output_dir = 'C:/automation_files/comparasion_files/'  # Change this to your desired path
output_filename = 'GWJ3404Ck_default_factreset_comparison.xlsx'
output_path = os.path.join(output_dir, output_filename)

# Check if the directory exists, if not create it
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Save the styled DataFrame to an Excel file with highlights, replacing any existing file
final_styled.to_excel(output_path, index=False, engine='openpyxl')

# Automatically open the Excel file after saving
subprocess.Popen(['start', output_path], shell=True)

print(f"Side-by-side comparison with highlights saved to: {output_path}")
