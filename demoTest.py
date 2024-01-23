import time

import pytest
from selenium import webdriver

from Selenium import xlUtils


# Command-line option to specify the browser (e.g., --browser_name chrome)
def pytest_addoption(parser):
    parser.addoption(
        "--browser_name", action="store", default="chrome"
    )


# Stored the Excel Path
path = "C:\\Users\\HP\\Documents\\Book1.xlsx"
rows = xlUtils.get_rows_count(path, "Sheet2")


# Fixture to set up the WebDriver and other properties
@pytest.fixture(scope="class")
def set_up(request):
    for r in range(2, rows + 1):
        browser = request.Excel_Utils.read_data(path, "Sheet2", )
        if browser.lower() == "chrome":
            driver = webdriver.Chrome()
            print("Test case is running with Chrome browser")
        elif browser.lower() == "edge":
            driver = webdriver.Edge()
            print("Test case is running with Edge browser")
        else:
            raise ValueError("Unsupported browser: {browser_name}")

        # Set up base URL and test data properties
        # assigns the WebDriver instance to a class attribute, making it accessible to all test methods in the class.
        request.cls.driver = driver

        driver.get(xlUtils.read_data(path, "Sheet2", r, 2))
        driver.maximize_window()
        driver.implicitly_wait(15)
        # Yield control to the test case
        yield

        def tear_down():
            driver.close()
    # After the test case, close the WebDriver
