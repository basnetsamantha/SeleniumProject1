from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import ElementNotInteractableException
import pandas as pd

try:
    # Open the website
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
    driver.maximize_window()
    driver.get("http://imis.innovativesolution.com.np/")

    # Find and click the sign-in button
    sign_in_button = driver.find_element(By.XPATH, "//button[normalize-space()='Sign In']")
    sign_in_button.click()

    # Fill in username and password
    username_field = driver.find_element(By.NAME, "username")
    password_field = driver.find_element(By.ID, "password")

    username_field.send_keys("insoldev100@gmail.com")
    password_field.send_keys("admin@321")

    # Click the login button
    login_button = driver.find_element(By.XPATH, "//button[normalize-space()='Log In']")
    login_button.click()

    # Wait for the sidebar link to be clickable
    sidebar_link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//i[@class='fas fa-bars']")))
    sidebar_link.click()

    # Navigate to sidebar and click module
    sidebar_link = driver.find_element(By.XPATH, "//i[@class='fas fa-bars']")
    sidebar_link.click()

    # Find and click the module link (with error handling)
    try:
        module_link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[@class='nav-link "
                                                                                            "active']")))
        module_link.click()
    except ElementNotInteractableException as e:
        print("Sub-module link is not intractable:", e)

    # Scrape text and titles
    elements = driver.find_elements(By.CLASS_NAME, "container-fluid")

    # Extract data into a list
    data = [element.text.split("\n") for element in elements]

    # Convert data to DataFrame
    df = pd.DataFrame(data)

    # Transpose the DataFrame
    df = df.transpose()

    # Write DataFrame to Excel file
    df.to_excel("scraped_data.xlsx", index=False, header=False)

finally:
    # Close the browser
    driver.quit()
