# Call Login python file
import re

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

global switcher


class PythonSwitch:
    # ************************************SWITCH CASE********************************
    def switch(self, value):
        default = "Incorrect value"
        return getattr(self, 'case_' + str(value), lambda: default)()

    def case_13(self): return '//option[@value="6001.13"] '

    def case_14(self): return '//option[@value="6001.14"] '

    def case_15(self): return '//option[@value="6001.15"] '

    def case_16(self): return '//option[@value="6001.16"] '

    def case_17(self): return '//option[@value="6001.17"] '

    def case_19(self): return '//option[@value="6001.19"] '

    def case_20(self): return '//option[@value="6001.20"] '

    def case_21(self): return '//option[@value="6001.21"] '

    def case_22(self): return '//option[@value="6001.22"] '

    def case_23(self): return '//option[@value="6001.23"] '

    def case_36(self): return '//option[@value="6001.36"] '

    def case_38(self): return '//option[@value="6001.38"] '

    def case_39(self): return '//option[@value="6001.39"] '

    def case_46(self): return '//option[@value="6001.46"] '

    def case_48(self): return '//option[@value="6001.48"] '

    def case_50(self): return '//option[@value="6001.50"] '

    def case_51(self): return '//option[@value="6001.51"] '

    def case_52(self): return '//option[@value="6001.52"] '

    def case_53(self): return '//option[@value="6001.53"] '

    def case_64(self): return '//option[@value="6001.64"] '

    def case_65(self): return '//option[@value="6001.65"] '

    def case_66(self): return '//option[@value="6001.66"] '

    def case_69(self): return '//option[@value="6001.69"] '

    def case_70(self): return '//option[@value="6001.70"] '

    def case_71(self): return '//option[@value="6001.71"] '

    def case_88(self): return '//option[@value="6001.88"] '

    def case_98(self): return '//option[@value="6001.98"] '

    def case_147(self): return '//option[@value="6001.147"]'

    def case_148(self): return '//option[@value="6001.148"]'

    def case_200(self): return '//option[@value="6001.200"]'

    def case_201(self): return '//option[@value="6001.201"]'

    def case_202(self): return '//option[@value="6001.202"]'

    def case_249(self): return '//option[@value="6001.249"]'


# *********************************************************************************************************************
class dynamicruleclass:

    # ************ Method to Handle Current Window **************#
    def handle_current_window_method(self):
        handles = driver.window_handles
        for handle in handles:
            driver.switch_to.window(handle)
            # print(driver.title)
            driver.maximize_window()

    def dynamicscreennavigation(self):
        time.sleep(4)
        # click on Systemadministration link
        driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[1]/td/a[24]").click()
        dynamicruleclass.handle_current_window_method(self)
        time.sleep(5)
        # Click on Dynamiske valideringsregler link
        driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/a[13]").click()
        dynamicruleclass.handle_current_window_method(self)
        time.sleep(3)
        dynamicruleclass.createdynamicrule(self)

    def createdynamicrule(self):
        # Click New Rule Button
        driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr/td/input").click()
        time.sleep(5)

        # Pass entire file path as parameter
        file_path = (
            r"C:\Users\AbhinavDixit\PycharmProjects\Skat-eIncomes\Online-INDB\Dynamic_Validation_Rule\Dynamic_Rule_creation.xlsx")  # set file path
        book = xlrd.open_workbook(file_path)
        sh = book.sheet_by_index(0)

        # Enter value in Fejlnr Field
        Fejlnr = sh.cell_value(rowx=1, colx=0)
        print("Fejlnr:-", Fejlnr)
        driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[1]/td[2]/input").send_keys(
            Fejlnr)

        # Enter value in Fejltekst Field
        Fejltekst = sh.cell_value(rowx=1, colx=1)
        print("Fejltekst:-", Fejltekst)
        driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[1]/td[4]/textarea").send_keys(
            Fejltekst)

        # Enter value in Fejlbeskrivelse Field
        Fejlbeskrivelse = sh.cell_value(rowx=1, colx=2)
        print("Fejlbeskrivelse :-", Fejlbeskrivelse)
        driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[1]/tbody/tr[2]/td[4]/textarea").send_keys(
            Fejlbeskrivelse)

        # Enter value in Gældende fra Field
        Valid_from = sh.cell_value(rowx=1, colx=3)
        print("Valid_from :-", Valid_from)
        driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr[1]/td[1]/input").send_keys(
            Valid_from)

        # Enter value in Gældende til Field
        Valid_Till = sh.cell_value(rowx=1, colx=4)
        driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr[1]/td[2]/input").send_keys(
            Valid_Till)

        # ******************************** CONDITION **********************************************************
        # Select Radio Button for condition
        driver.find_element_by_xpath("//*[@id='betingelse_6001']").click()
        time.sleep(3)

        # Enter value of Vælg betingelse
        # Select Condition Value from Excel
        conditon = int(sh.cell_value(rowx=1, colx=5))
        print("condition value", conditon)
        a = PythonSwitch()
        a.switch(conditon)
        print("CONDITON VALUE IS", a.switch(conditon))
        driver.find_element_by_xpath(a.switch(conditon)).click()

        # Select Operator from Excel
        Operator = sh.cell_value(rowx=1, colx=6)
        print("Operator value", Operator)
        driver.find_element_by_xpath(Operator).click()
        # driver.find_element_by_css_selector("table.clContentTable:nth-child(4) > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(2) > select:nth-child(1)"). \    send_keys(conditon)
        time.sleep(2)

        # Enter value of Condition
        Condition_Value = int(sh.cell_value(rowx=1, colx=7))
        print("Operator value", Condition_Value)
        driver.find_element_by_name("betingelse_veardi").send_keys(Condition_Value)
        time.sleep(3)
        driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[4]/tbody/tr[5]/td/input[2]").click()
        time.sleep(3)
        driver.find_element_by_id("betingelse_5000_6000").click()
        time.sleep(2)

        # ******************************** RULE **********************************************************
        # Select Radio Button for Rule(Vælg regel)
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        driver.find_element_by_id("regel_6001").click()
        time.sleep(3)

        # Enter value of Vælg regel
        # #Select Rule Value from Excel
        Rule = int(sh.cell_value(rowx=1, colx=8))
        print("Rule value", Rule)
        a = PythonSwitch()
        a.switch(Rule)
        print("RULE VALUE IS", a.switch(Rule))
        time.sleep(2)
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        time.sleep(2)
        # driver.find_element_by_xpath("/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[6]/tbody/tr[3]/td[1]/select/option[1]").click()
        driver.find_element_by_name("regel_record_selected").find_element_by_xpath(a.switch(Rule)).click()
        # driver.find_element_by_xpath("//option[@value='6001.66']").click()

        # Select Operator from Excel
        Operator = sh.cell_value(rowx=1, colx=6)
        print("Operator value", Operator)
        driver.find_element_by_xpath(Operator).click()
        # driver.find_element_by_css_selector("table.clContentTable:nth-child(4) > tbody:nth-child(1) > tr:nth-child(3) > td:nth-child(2) > select:nth-child(1)"). \    send_keys(conditon)
        time.sleep(2)

        # Enter value of Rule
        Rule_Value = int(sh.cell_value(rowx=1, colx=9))
        print("Operator value", Rule_Value)
        driver.find_element_by_name("regel_veardi").send_keys(Rule_Value)
        time.sleep(3)
        driver.find_element_by_xpath(
            "/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[6]/tbody/tr[5]/td/input[2]").click()
        time.sleep(3)


A1 = dynamicruleclass()
A1.dynamicscreennavigation()