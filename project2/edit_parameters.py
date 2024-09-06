from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

def login_to_website(url, username, password):
    print("Setting up options...")
    options = Options()
    options.add_argument('--start-maximized')
    #options.add_argument('--ignore-certificate-errors')
    options.set_capability("acceptInsecureCerts", True)

    print("Starting Edge WebDriver...")
    driver = webdriver.Edge(service=EdgeService(executable_path=driver_path), options=options)
    
    try:
        print(f"Navigating to {url}...")
        driver.get(url)

        print("Locating username field...")
        username_field = driver.find_element(By.ID, "username")
        username_field.send_keys(username)

        print("Locating password field...")
        password_field = driver.find_element(By.ID, "password")
        password_field.send_keys(password)

        print("Submitting the form...")
        password_field.send_keys(Keys.RETURN)

        print("Login attempted successfully.")
        
    except Exception as ex:
        print(f"An error occurred: {ex}")
    finally:
        print("Quitting the driver...")
        driver.quit()

if __name__ == "__main__":
    url = "https://192.168.0.1:8080"
    username = "Installer"
    password = "YjhlNmM5M2"
    driver_path = "C:/Users/BachuKP/Lib/site-packages/pythonwin/pywin/automation/project2/edgedriver_win64/msedgedriver.exe"  # Replace with the actual path to msedgedriver.exe
    login_to_website(url, username, password)
 