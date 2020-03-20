# Call Login python file
import driver as driver
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
from selenium.webdriver.support import expected_conditions as EC, wait
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import *

import time
import xlrd  # import package to read data from Excel
import openpyxl
from openpyxl import load_workbook


# ************ Method to Handle Current Window **************#
def handle_current_window_method():
    handles = driver.window_handles
    for handle in handles:
        driver.switch_to.window(handle)
        # print(driver.title)
        driver.maximize_window()


# Click on Forespørg/Kopiér/Tilbagefør indberetninger
time.sleep(5)
driver.find_element_by_xpath(
    "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/a[7]").click()

# Move to current new window
handle_current_window_method()
time.sleep(5)

# select search option Indberetnings-ID from drop down
aa = driver.find_element_by_xpath(
    "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[7]/td[2]/select").is_displayed()
print(aa)
driver.find_element_by_xpath(
    "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[7]/td[2]/select").send_keys(
    "IndberetningsID")
time.sleep(3)

# Pass entire file path as parameter
file_path = (
    r"C:\Users\AbhinavDixit\PycharmProjects\Skat-eIncomes\Online-INDB\1-Submit_Online-INDB\Online_INDB_Excel.xlsx")  # set file path
book = xlrd.open_workbook(file_path)
sh = book.sheet_by_index(0)
data = sh.cell_value(rowx=1, colx=0)
print(data)
time.sleep(3)

# Enter INDB ID to be searched
driver.find_element_by_xpath(
    "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[7]/td[2]/input").send_keys(
    data)

# Press Sog Button
driver.find_element_by_xpath("//*[@id='defaultButton']").click()
try:
    WebDriverWait(driver, 3).until(EC.alert_is_present(),
                                   'Timed out waiting for PA creation ' +
                                   'confirmation popup to appear.')

    alert = driver.switch_to.alert
    print(alert.text)
    time.sleep(5)
    alert.accept()
    print("alert accepted")
except TimeoutException:
    print("no popup alert")
time.sleep(5)

try:
    INDB_ID = driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[18]/td/table/tbody/tr[2]/td[3]").text
    ART_VALUE = driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[18]/td/table/tbody/tr[2]/td[7]").text
    TILBAGEFORT_VALUE = driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[18]/td/table/tbody/tr[2]/td[8]").text
    print("INDB ID is :", INDB_ID)
    print("ART Value is:", ART_VALUE)
    print("TILBAGEFORT VALUE is:", TILBAGEFORT_VALUE)

    if driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[18]/td/table/tbody/tr[2]/td[3]").is_displayed():
        if INDB_ID == data and ART_VALUE == 'I' and TILBAGEFORT_VALUE == 'Nej':
            driver.find_element_by_xpath(
                "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[18]/td/table/tbody/tr[2]/td[11]/input").click()  # select Radio Button
            driver.find_element_by_xpath(
                "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[16]/td[2]/input").click()  # click Tilbagefør button
            time.sleep(3)

            New_INDBREV_ID = driver.find_element_by_xpath(
                "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[7]/td[2]/input").is_displayed()
            print(New_INDBREV_ID)  # Check whether New INDB REV ID field is dispalyed

            New_INDB_REV_ID = driver.find_element_by_xpath(
                "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[7]/td[2]/input").get_attribute(
                "value")
            print("New_INDB_REV_ID is:-", New_INDB_REV_ID)  # print New INDB REV ID

            Hovedindberetningsident = driver.find_element_by_xpath(
                "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[4]/td[2]/input").get_attribute(
                "value")
            print("Hovedindberetningsident is:", Hovedindberetningsident)

            # Copy INDB ID to Excel File
            time.sleep(5)
            workbook = openpyxl.load_workbook(file_path)  # Load Workbook
            sheet = workbook['REV_INDBID']
            print(New_INDBREV_ID)
            # Copy INDB_ID in to Excel
            sheet.cell(2, 1).value = INDB_ID
            sheet.cell(2, 2).value = New_INDB_REV_ID
            sheet.cell(2, 3).value = Hovedindberetningsident
            workbook.save(file_path)
    else:
        print('  ')
except NoSuchElementException:
    print('  ')

# Click bekraft button to commit Tilbage (Reversal)
driver.find_element_by_xpath("//*[@id='defaultButton']").click()
try:
    WebDriverWait(driver, 3).until(EC.alert_is_present(),
                                   'Timed out waiting for PA creation ' +
                                   'confirmation popup to appear.')

    alert = driver.switch_to.alert
    print(alert.text)
    time.sleep(5)
    alert.accept()
    print("alert accepted")
except TimeoutException:
    print("no popup alert")
time.sleep(5)

try:
    if driver.find_element_by_partial_link_text('dk.lec.jroad.exceptions.LECRuntimeException').is_displayed():
        print("RunTime Exception: dk.lec.jroad.exceptions.LECRuntimeException")
    else:
        print("No RunTime Exception and INDB submitted successfully")
except NoSuchElementException:
    print(' ')
