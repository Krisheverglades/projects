import pandas as pd

# Read MySQL data into a DataFrame
mysql_data = pd.read_csv('C:/automation_files/comparasion_files/configuration.xlsx')

# Read Excel data into a DataFrame
excel_data = pd.read_excel('C:\Users\BachuKP\Lib\site-packages\pythonwin\pywin\automation\Defaultparameters.xlsx')

# Assuming both datasets have 'name' and 'value' columns
# Sort dataframes by name for efficient comparison
mysql_data.sort_values(by='name', inplace=True)
excel_data.sort_values(by='name', inplace=True)

# Compare data
differences = []
for idx, row in mysql_data.iterrows():
    name = row['name']
    mysql_value = row['value']
    excel_row = excel_data.loc[excel_data['name'] == name]
    if excel_row.empty:
        differences.append(f"Name '{name}' exists in MySQL but not in Excel")
    else:
        excel_value = excel_row['value'].iloc[0]
        if mysql_value != excel_value:
            differences.append(f"Value for name '{name}' differs between MySQL ({mysql_value}) and Excel ({excel_value})")

for idx, row in excel_data.iterrows():
    name = row['name']
    excel_value = row['value']
    mysql_row = mysql_data.loc[mysql_data['name'] == name]
    if mysql_row.empty:
        differences.append(f"Name '{name}' exists in Excel but not in MySQL")

# Print differences
if differences:
    print("Differences found:")
    for diff in differences:
        print(diff)
else:
    print("No differences found.")