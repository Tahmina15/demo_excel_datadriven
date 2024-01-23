import time

from selenium import webdriver
from selenium.webdriver.common.by import By

from Selenium import xlUtils

driver = webdriver.Chrome()
driver.get("https://demo.nopcommerce.com/")
driver.maximize_window()
time.sleep(5)

path = "C:\\Users\\HP\\Documents\\Book1.xlsx"
rows = xlUtils.get_rows_count(path, "Sheet2")

# using for loop to get the data from Excel
for r in range(2, rows + 1):
    userName = xlUtils.read_data(path, "Sheet2", r, 1)
    password = xlUtils.read_data(path, "Sheet2", r, 2)
    time.sleep(10)
    driver.find_element(By.CSS_SELECTOR, "a.ico-login").click()
    driver.find_element(By.CSS_SELECTOR, "input#Email").send_keys(userName)
    driver.find_element(By.CSS_SELECTOR, "input#Password").send_keys(password)
    driver.find_element(By.CSS_SELECTOR, "button.button-1.login-button").click()
    if driver.title == "Welcome to our store":
        print("Login is Success")
        xlUtils.write_data(path, "Sheet1", r, 3, "Pass")
    else:
        print("Login is Failed")
        xlUtils.write_data(path, "Sheet1", r, 3, "Fail")


# driver.get_screenshot_as_file("get.png")
# time.sleep()
# driver.quit()
