import self
from click._compat import raw_input
import tkinter as tk
from tkinter import simpledialog
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import driver as driver
import time

global USER_INP_SENR

def new_window_handle_method(self):
     # Move to popup window to add username and password
    print(driver.current_window_handle)  # 18 <-- Parent
    handles = driver.window_handles
    for handle in handles:
        driver.switch_to.window(handle)
        print(driver.title)
        driver.maximize_window()

# open Firefox Browser
driver = webdriver.Firefox(executable_path="C:\geckodriver.exe")
driver.get("http://10.9.182.15/")
        #driver.get("https://dev.ei-admin.skat.dk/")
time.sleep(3)
new_window_handle_method(self)

    # On the popup screen select/click/activate the frame to provide login credentials
driver.switch_to.frame("letloen")
time.sleep(5)
driver.find_element_by_css_selector('#defaultFocusElementId').send_keys("Admin1032")
        #driver.find_element_by_css_selector('#defaultFocusElementId').send_keys("ABHI")
driver.find_element_by_name("password").send_keys("Test1234")
      #driver.find_element_by_name("password").send_keys("1234Test")
driver.find_element_by_xpath("//*[@id='defaultButton']").click()
        # Enter CVR/SE-nr Number
time.sleep(3)
driver.switch_to.__class__("clContentTable")
time.sleep(2)
driver.find_element_by_class_name("inputButton").click()
time.sleep(3)

ROOT = tk.Tk()
ROOT.withdraw()
      # the input dialog
#USER_INP_SENR = simpledialog.askstring(title="SENR-nr",
 #                               prompt="Please Enter SENR-nr : ")
         #print("Hello", USER_INP)
         #senr_number = raw_input("Enter SENR-nr and Press Enter : ")
driver.find_element_by_id("defaultFocusElementId").send_keys("15107286") #37172715 82000275 15107286
driver.find_element_by_xpath("//*[@id='defaultButton']").click()
time.sleep(18)

