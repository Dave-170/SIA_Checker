# SIA_Checker
This Python script, named SIA_Checker, is designed to automate the extraction of information related to Security Industry Authority (SIA) numbers from a provided spreadsheet. The script utilizes the Selenium and Pandas libraries for web scraping and data manipulation, respectively.

## Table of Contents
- [Requirements](#requirements)
- [Usage](#usage)
- [Code Structure](#code-structure)
- [External Libraries](#external-libraries)
- [Notes](#notes)
- [Contribution](#contribution)

## Requirements
- Python 3.x
- Dependencies: pandas, selenium, webdriver_manager
- Chrome browser installed

## Usage
1. Run the script, and a file dialog will prompt you to select a spreadsheet file containing SIA numbers.
2. The script will then iterate through the SIA numbers, querying the official SIA website to extract relevant information.
3. Extracted data is saved to a new Excel file named 'sia_data.xlsx' in the script's directory.

## Code Structure
- `select_spreadsheet_file()`: Opens a file dialog for selecting the input spreadsheet file.
- `setup_driver()`: Configures and returns a headless Chrome WebDriver for web scraping.
- `extract_data(driver, sia_numbers)`: Extracts data for each SIA number using Selenium and XPaths.
- XPaths are defined for each data element to be extracted from the SIA website.

## External Libraries
- Pandas: For efficient data manipulation.
- Selenium: For automated web interactions.
- ChromeDriverManager: For managing the Chrome WebDriver.

## Notes
- The script uses headless mode by default. Modify `chrome_options.headless` to `False` in the `setup_driver()` function if you want to run it in a visible browser.
- Ensure the required libraries are installed using `pip install -r requirements.txt`.
- Security Industry Authority Register of Licence Holders data based used to check input Licences

## Contribution
Contributions are welcome! Feel free to fork the repository, make improvements, and create a pull request. For detailed information and to contribute, visit the [GitHub repository](#).

