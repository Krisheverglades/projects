import pandas as pd
import os
import openpyxl
import subprocess

# Define the given parameter list
parameter_list = [
    "GWCODE", "WIFI_MAC", "WIFI_SSID", "WIFI_PASSWD", "WIFI_PASSWD_TYPE", "WIFI_STATIC", "WIFI_STATIC_IP", 
    "WIFI_STATIC_GATEWAY", "WIFI_STATIC_NETMASK", "WIFI_STATIC_DNS", "WIFI_CONNECTED", "WIFI_HOTSPOT_ENABLE", 
    "WIFI_HOTSPOT_SSID", "WIFI_HOTSPOT_PASSWORD", "WIFI_CONNECTION_MODE", "IOT_HOSTNAME", "IOT_DEVICE_ID", 
    "IOT_ACCESS_KEY", "MID_ENERGY_METER", "ETHERNET_CONNECTED", "WEB_PORTAL_PASSWORD", "CHARGE_CURRENT", 
    "POWER_SUPPLY", "OTA_UPDATE_SESSION_ID", "SERIAL_NUMBER", "SOCKET_NUMBER", "NOMINAL_CURRENT",
    "INPUT_CURRENT", "PP_MOBILE_CABLE", "TYPE_OF_SOCKET", "VANDAL_VERSION", "SOCKETMGMT_1", "RFID_AUTH", 
    "AUTH_TYPE", "LOG_LEVEL", "MID_ENERGY_METER", "MASTER_MODBUS_BAUDRATE", "MASTER_MODBUS_NODE", "SLAVE_MODBUS", 
    "SLAVE_MODBUS_NODE", "SLAVE_MODBUS_BAUDRATE", "REMOTE_CNT_1", "REMOTE_CNT_1_FUNC", "REMOTE_CNT_2", 
    "REMOTE_CNT_2_FUNC", "DISPLAY", "TYPE_OF_DISPLAY", "MAX_UNBALANCE_CURRENT", "HOME_FUNCTIONAL_MODE", 
    "HOME_PROGRAM_TIME", "PLANT_POWER", "PLANT_POWER_SUPPLY", "CT_PRESENCE", "PEN_LOGIC", "MS_MANAGEMENT", 
    "MS_ROLE", "ETH_DHCP", "MODEM_AVAILABLE", "OCPP_PROTOCOL", "OCPP_CENTRAL_SYSTEM_ENDPOINT", "OCPP_CHARGEBOX_IDENTITY", 
    "OCPP_BOOTNOTIFICATION_USERNAME", "OCPP_BOOTNOTIFICATION_PASSWORD", "OCPP_AUTHORIZATION_CACHE_ENABLED", 
    "OCPP_AUTHORIZE_REMOTE_TXREQUESTS", "OCPP_CLOCK_ALIGNED_DATA_INTERVAL", "OCPP_CONNECTION_TIMEOUT", 
    "OCPP_CONNECTOR_PHASE_ROTATION", "OCPP_GET_CONFIGURATION_MAX_KEYS", "OCPP_HEARTBEAT_INTERVAL", "OCPP_CHANGE_AVAILABILITY", 
    "OCPP_LOCAL_AUTHORIZE_OFFLINE", "OCPP_LOCAL_PRE_AUTHORIZE", "OCPP_METER_VALUES_ALIGNED_DATA", 
    "OCPP_METER_VALUES_SAMPLED_DATA", "OCPP_METER_VALUE_SAMPLE_INTERVAL", "OCPP_NUMBER_OF_CONNECTOR", 
    "OCPP_RESET_RETRIES", "OCPP_STOP_TRANSACTION_ON_EV_SIDE_DISCONNECT", "OCPP_STOP_TRANSACTION_ON_INVALID_ID", 
    "OCPP_STOP_TXN_ALIGNED_DATA", "OCPP_STOP_TXN_SAMPLED_DATA", "OCPP_SUPPORTED_FEATURE_PROFILES", 
    "OCPP_TRANSACTION_MESSAGE_ATTEMPTS", "OCPP_TRANSACTION_MESSAGE_RETRY_INTERVAL", "OCPP_UNLOCK_CONNECTOR_ON_EV_SIDE_DISCONNECT", 
    "OCPP_LOCAL_AUTH_LIST_ENABLED", "OCPP_LOCAL_AUTH_LIST_MAX_LENGTH", "OCPP_SEND_LOCAL_LIST_MAX_LENGTH", "EICHRECHT", 
    "OZEV", "OZEV_MAX_RANDOM_DELAY", "TIC", "TIC_PROTOCOL", "PEN_MIN_VALUE", "PEN_MAX_VALUE", "SHUTDOWN_WIFI_IOTLOCAL", 
    "IOT_CONNECTED", "PHASE_ROTATION", "ION_BOARD_SX_DX", "RFID_IOT_AUTH_LIST_VERSION", "RFID_IOT_AUTH_LIST_ERROR_MESSAGE", 
    "RFID_IOT_AUTH_LIST_BLOB_URL", "GROUPING_OCPP", "GROUPING_OCPP_ROLE", "GROUPING_OCPP_MAC", "OCPP_OTA_RETRIEVE_DATE", 
    "OCPP_OTA_LINK", "ETHERNET_MAC_ADDRESS", "MAX"
]

# Load the factory table (Excel) that contains the SettingName column
excel_file_path = 'C:\Users\BachuKP\Lib\site-packages\pythonwin\pywin\automation\project2\configurationfactGWJ3404CK.csv'  # Replace with your file path
df_factory = pd.read_excel(excel_file_path, engine='openpyxl')

# Ensure the 'SettingName' column is in the factory table
if 'SettingName' not in df_factory.columns:
    print("Error: 'SettingName' column not found in the Excel file.")
    exit()

# Extract the 'SettingName' column from the Excel file
setting_names = df_factory['SettingName'].tolist()

# Create a DataFrame for comparison
comparison_df = pd.DataFrame({'ParameterList': parameter_list})

# Add a new column to check if the parameter is present in the Excel's 'SettingName'
comparison_df['InExcel'] = comparison_df['ParameterList'].apply(lambda x: 'Yes' if x in setting_names else 'No')

# Save the comparison to an Excel file and highlight missing parameters
output_path = 'C:\Users\BachuKP\OneDrive - Gewiss S.p.A\Documenti\krish\Joinon\SW and guides\docs\aftercomparison\rte_comparison.xlsx'  # Change this to your desired path

# Write to Excel with highlighting
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    # Save comparison DataFrame
    comparison_df.to_excel(writer, sheet_name='Comparison', index=False)
    
    # Get the workbook and sheet for styling
    workbook = writer.book
    worksheet = writer.sheets['Comparison']
    
    # Define formats for highlighting
    red_fill = openpyxl.styles.PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    
    # Highlight cells in the 'InExcel' column where the value is 'No'
    for row in range(2, len(comparison_df) + 2):  # Start at row 2 since row 1 is the header
        if worksheet[f'B{row}'].value == 'No':  # Check if 'InExcel' is 'No'
            worksheet[f'A{row}'].fill = red_fill
            worksheet[f'B{row}'].fill = red_fill

# Automatically open the Excel file after saving
subprocess.Popen(['start', output_path], shell=True)

print(f"Parameter comparison saved to: {output_path}")
