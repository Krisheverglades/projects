import subprocess
import pyautogui
import time



def powershell_script():
    # Path to the PowerShell script
    ps_script_path = r'C:/Users/BachuKP/Lib/site-packages/pythonwin/pywin/automation/project2/open_powershell.ps1'
    
    # Execute the PowerShell script
    subprocess.Popen(['powershell.exe', '-File', ps_script_path], shell=True)

    time.sleep(5)
    # Send the SCP command to MobaXterm
    pyautogui.write('scp root@192.168.0.1:/root/.joinon-configuration-manager/configuration.db C:/automation_files/porataledibordo_temp/configuration.db', interval=0)
    
    # Press Enter to execute the command
    pyautogui.press('enter')

    # Wait for the command to complete
    time.sleep(5)

    # Send the next command to MobaXterm
    pyautogui.write('Joinon.ValeBarak', interval=0.1)

    # Press Enter to execute the command
    pyautogui.press('enter')

if __name__ == "__main__":
    powershell_script()

    