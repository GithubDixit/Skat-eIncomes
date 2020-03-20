# Call Login python file
import driver as driver
import re
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


# ************ Method to Copy Data in Excel **************#
def copy_data_excel_method():
    # ************** Load Excel **************#
    path = (
        r"C:\Users\AbhinavDixit\PycharmProjects\Skat-eIncomes\Online-INDB\1-Submit_Online-INDB\Online_INDB_Excel.xlsx")
    workbook = openpyxl.load_workbook(path)  # Load Workbook
    sheet = workbook['INDBID']
    time.sleep(5)
    INDB_ID = driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[3]/tbody/tr[3]/td[2]").text
    print(INDB_ID)
    # Copy INDB_ID in to Excel
    sheet.cell(2, 1).value = INDB_ID
    workbook.save(path)


# ************ Method to Create Online INDB **************#
def create_onlineINDB_method():
    driver.find_element_by_id("loenPeriodeStartDato").send_keys("2019.12.01")
    time.sleep(3)
    driver.find_element_by_id("loenPeriodeSlutDato").send_keys("2019.12.31")
    time.sleep(3)
    driver.find_element_by_id("dispositionsdato").send_keys("2019.12.15")
    time.sleep(3)
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/table/tbody/tr[15]/td[2]/input").click()

    # Enter value in Fields
    time.sleep(3)
    # **************** Field 13
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/table/tbody/tr[15]/td[2]/input").clear()
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/table/tbody/tr[15]/td[2]/input").send_keys(
        "1000")
    # **************** Field 14
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/table/tbody/tr[15]/td[4]/input").clear()
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/table/tbody/tr[15]/td[4]/input").send_keys(
        "500")
    driver.find_element_by_xpath("/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td[2]/input").click()
    time.sleep(5)


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
    time.sleep(3)
    alert.accept()
    print("alert accepted")
except TimeoutException:
    print("no popup alert")
time.sleep(5)

try:
    INDB_ID = driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[18]/td/table/tbody/tr[2]/td[3]").text
    print("INDB ID is :", INDB_ID)

    if driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[18]/td/table/tbody/tr[2]/td[3]").is_displayed():
        if INDB_ID == data:
            driver.find_element_by_xpath(
                "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[18]/td/table/tbody/tr[2]/td[11]/input").click()  # select Radio Button
            driver.find_element_by_xpath(
                "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[17]/td/input").click()  # click Vis button
            time.sleep(3)

            # CLick on Kopier option
            driver.find_element_by_xpath(
                "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[12]/td/input").click()
            time.sleep(3)

            driver.find_element_by_xpath(
                "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/table/tbody/tr[6]/td[2]/input").is_displayed()

            New_INDBCOPY_ID = driver.find_element_by_xpath(
                "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/table/tbody/tr[6]/td[2]/input").get_attribute(
                "value")
            print("New_INDB_COPY_ID is:-", New_INDBCOPY_ID)  # Check whether New INDB REV ID field is dispalyed

            # Copy INDB ID to Excel File
            time.sleep(5)
            workbook = openpyxl.load_workbook(file_path)  # Load Workbook
            sheet = workbook['COPY_INDB']
            print(New_INDBCOPY_ID)
            # Copy INDB_ID in to Excel
            sheet.cell(2, 1).value = INDB_ID
            sheet.cell(2, 2).value = New_INDBCOPY_ID
            workbook.save(file_path)
            create_onlineINDB_method()
            time.sleep(5)
    else:
        print('  ')
except NoSuchElementException:
    print('  ')

# ******************* Execption handling if case Pincode Validaton Dynamic Rule popus
try:
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/div/div/div[2]/input[1]").is_displayed()
    Dynamic_Rule_Text = driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/div/div/div[1]/div[1]").text
    print("Pincode Dynamic Validation Rule Error text message:", Dynamic_Rule_Text)
    Dynamic_Rule = re.search('\(([^)]+)', Dynamic_Rule_Text).group(1)  # Pick rule Name
    print("The Error is due to Dynamic Rule Number :", Dynamic_Rule)
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/div/div/div[2]/input[1]").click()  # click Tilbage

    # Search the Rule and fill value of Pinkode to proceed
    driver.find_element_by_xpath("/html/body/div[2]/span[1]/a").click()  # Click Indkomst menu
    time.sleep(3)

    # Click Systemadministration menu
    handle_current_window_method()
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/a[24]").click()

    # Move to current new window
    handle_current_window_method()
    time.sleep(5)

    # click Dynamiske valideringsregler menu
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/a[13]").click()

    # Move to current new window
    handle_current_window_method()
    time.sleep(5)

    # Enter Fejlnr in the Dynamic Rule Search field
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[3]/td[2]/input").send_keys(
        Dynamic_Rule)
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[4]/td[6]/input").click()  # Click Sog
    time.sleep(3)

    # Select the Rule
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr[2]/td[3]").click()
    time.sleep(5)

    # Copy value of Pincode
    Pinkode = driver.find_element_by_xpath("//*[@id='pinkode']").get_attribute("value")
    print(Pinkode)

    # Click on Indkomst menu
    driver.find_element_by_xpath("/html/body/div[2]/span[1]/a").click()
    handle_current_window_method()
    # click Afslut button
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/input").click()

    # Click on Indberet lønoplysninger - online to create INDB
    handle_current_window_method()
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/a[3]").click()
    handle_current_window_method()
    time.sleep(3)

    # Select EIncome Radio Button
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/table/tbody/tr[4]/td[2]/span[2]/input").click()
    # Provide CPR-no
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/table/tbody/tr[8]/td[2]/input").send_keys(
        "1909580060")
    driver.find_element_by_xpath("//*[@id='defaultButton']").click()
    time.sleep(3)
    create_onlineINDB_method()

    # Enter Pincode and click ok
    driver.find_element_by_xpath("//*[@id='pinkode']").send_keys(Pinkode)
    driver.find_element_by_xpath(
        "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/div/div/div[2]/input[2]").click()
except NoSuchElementException:
    print(' ')

try:
    if driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/table[2]/tbody/tr[2]/td[3]").is_displayed():
        Fejl_Advis = driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/div/table[2]/tbody/tr[2]/td[3]").text
        if Fejl_Advis == 'Fejl':
            print("Error found proceeding to submit INDB")
        else:
            print("Only Adivce found proceeding to submit INDB")
            driver.find_element_by_xpath("/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td[2]/input").click()
            time.sleep(5)
            copy_data_excel_method()  # Call Method to
            # Click Afslut Button to close INDB Creation Screen
            driver.find_element_by_xpath(
                "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[5]/tbody/tr/td[2]/input[1]").click()
except NoSuchElementException:
    time.sleep(5)

try:
    if driver.find_element_by_partial_link_text('dk.lec.jroad.exceptions.LECRuntimeException').is_displayed():
     print("RunTime Exception: dk.lec.jroad.exceptions.LECRuntimeException")
    else:
     print("No RunTime Exception and INDB submitted successfully")
except NoSuchElementException:
    print(" ")
