import Login_Add_SENR

from Login_Add_SENR import webdriver
from Login_Add_SENR import Keys
from Login_Add_SENR import WebDriverWait
from Login_Add_SENR import By
from Login_Add_SENR import Select
from Login_Add_SENR import driver
from Login_Add_SENR import webdriver

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time

time.sleep(5)
driver.find_element_by_css_selector("a.stdhvid:nth-child(10)").click()

# Move to popup window to add username and password
print(driver.current_window_handle)  # 18 <-- Parent
handles = driver.window_handles

for handle in handles:
    driver.switch_to.window(handle)
    print(driver.title)
driver.maximize_window()

time.sleep(12)
driver.find_element_by_xpath("//*[@id='defaultButton']").click()

NullINDB_ID = driver.find_element_by_xpath(
    "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/table[2]/tbody/tr[1]/td[2]").text

# Copy INDB_ID in to Excel

import openpyxl

path = (r"C:\Users\AbhinavDixit\PycharmProjects\Skat-eIncomes\Online-INDB\1-Submit_Online-INDB\Online_INDB_Excel.xlsx")
workbook = openpyxl.load_workbook(path)  # Load Workbook
NullINDB_Sheet = workbook['NullINDB']  # Load active sheet and save in sheet object
NullINDB_Sheet.cell(2, 1).value = NullINDB_ID
workbook.save(path)

driver.find_element_by_xpath("/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td[1]/input").click()  # Save NullINDB
driver.find_element_by_xpath("/html/body/table[3]/tbody/tr/td[2]/input").click()  # Close window
