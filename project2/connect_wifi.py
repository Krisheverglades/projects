import subprocess

# Path to your PowerShell script
script_path = r"C:/Users/BachuKP/Lib/site-packages/pythonwin/pywin/automation/project2/connect.ps1"

# Enclose the script path in double quotes to handle spaces
command = f'powershell.exe -ExecutionPolicy Bypass -File "{script_path}"'

# Run the PowerShell script
try:
    result = subprocess.run(command, check=True, shell=True, capture_output=True, text=True)
    print("Wi-Fi connection script executed successfully.")
    print(result.stdout)  # Output of the script
except subprocess.CalledProcessError as e:
    print(f"An error occurred while running the script: {e}")
    print(e.stdout)  # Script output before error
    print(e.stderr)  # Error message output
