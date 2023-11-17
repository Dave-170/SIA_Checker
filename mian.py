import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from tkinter import Tk, filedialog
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Define the XPath expressions for data extraction
XPATH_SIA_NUMBER = "/html/body/div[1]/div[9]/div/div[2]/div[1]/div/div"
XPATH_FIRST_NAME = "/html/body/div[1]/div[9]/div/div[1]/div[1]/div/div"
XPATH_SURNAME = "/html/body/div[1]/div[9]/div/div[1]/div[2]/div/div"
XPATH_ROLE_STATUS = "/html/body/div[1]/div[9]/div/div[2]/div[2]/div/div"
XPATH_EXPIRY_DATE = "/html/body/div[1]/div[9]/div/div[3]/div[1]/div/div"
XPATH_LICENCE_SECTOR = "/html/body/div[1]/div[9]/div/div[2]/div[3]/div/div"
XPATH_STATUS = "/html/body/div[1]/div[9]/div/div[3]/div[2]/div"


def select_spreadsheet_file():
    root = Tk()
    root.withdraw()  # Hide the main window

    filetypes = [
        ("All Files", "*.*"),
        ("Excel Files", "*.xlsx"),
        ("OpenDocument Spreadsheet", "*.ods"),
    ]

    file_path = filedialog.askopenfilename(
        title="Select Spreadsheet File",
        filetypes=filetypes,
    )

    return file_path


def setup_driver():
    chrome_options = Options()
    chrome_options.headless = True  # Set to True for headless mode
    service = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.implicitly_wait(10)

    return driver


def extract_data(driver, sia_numbers):
    sia_data = []

    for sia_number in sia_numbers:
        if pd.notna(sia_number):
            try:
                driver.get("https://services.sia.homeoffice.gov.uk/rolh")
                sia_input = driver.find_element(By.ID, "LicenseNo")
                sia_input.clear()
                sia_input.send_keys(str(sia_number))
                sia_input.send_keys(Keys.RETURN)

                # Wait for the SIA number element to be present
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, XPATH_SIA_NUMBER))
                )

                # Extract data using XPath
                sia_dict = {
                    "SIA Number": sia_number,
                    "First Name": driver.find_element(
                        By.XPATH, XPATH_FIRST_NAME
                    ).text.strip(),
                    "Surname": driver.find_element(
                        By.XPATH, XPATH_SURNAME
                    ).text.strip(),
                    "Role Status": driver.find_element(
                        By.XPATH, XPATH_ROLE_STATUS
                    ).text.strip(),
                    "Expiry Date": driver.find_element(
                        By.XPATH, XPATH_EXPIRY_DATE
                    ).text.strip(),
                    "Licence Sector": driver.find_element(
                        By.XPATH, XPATH_LICENCE_SECTOR
                    ).text.strip(),
                    "Status": driver.find_element(By.XPATH, XPATH_STATUS).text.strip(),
                }
                sia_data.append(sia_dict)
            except Exception as e:
                print(f"Error extracting data for SIA Number {sia_number}: {str(e)}")
                continue

    return sia_data


if __name__ == "__main__":
    spreadsheet_file_path = select_spreadsheet_file()

    if spreadsheet_file_path:
        df = pd.read_excel(spreadsheet_file_path)
        sia_numbers = df["SIA Number"].tolist()

        driver = setup_driver()
        sia_data = extract_data(driver, sia_numbers)
        driver.quit()

        df = pd.DataFrame(sia_data)
        df.to_excel("sia_data.xlsx", index=False)
