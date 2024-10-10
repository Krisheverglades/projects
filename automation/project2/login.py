from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def login_to_website(url, username, password, driver_path):
    # Set up Edge WebDriver with the pre-downloaded driver
    options = Options()
    options.add_argument('--start-maximized')  # Optional: Open browser in maximized mode
    options.add_argument('--ignore-certificate-errors')

    # Directly specify the path to the Edge WebDriver
    driver = webdriver.Edge(service=EdgeService(executable_path=driver_path), options=options)
    
    try:
        driver.get(url)
        
        # Wait for the username field to be present
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "username")))
        
        username_field = driver.find_element(By.ID, "username")
        username_field.send_keys(username)
        
        password_field = driver.find_element(By.ID, "password")
        password_field.send_keys(password)
        
        # Press Enter to submit the form
        password_field.send_keys(Keys.RETURN)
        
        # Wait for some element that indicates successful login or page load
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "username")))  # Replace with an actual element ID
        
        print("Login attempted successfully.")
    except Exception as ex:
        print(f"An error occurred: {ex}")
    finally:
        # Close the browser
        driver.quit()

if __name__ == "__main__":
    url = "https://192.168.0.1:8080"
    username = "Installer"
    password = "YjhlNmM5M2"
    driver_path = "C:/Users/BachuKP/Lib/site-packages/pythonwin/pywin/automation/project2/edgedriver_win64/msedgedriver.exe"  # Replace with the actual path to msedgedriver.exe
    
    login_to_website(url, username, password, driver_path)