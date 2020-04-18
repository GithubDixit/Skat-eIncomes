import self as self
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

import time

# open Firefox Browser
from methods import methods

driver = webdriver.Firefox(executable_path="C:\geckodriver.exe")
driver.get("http://10.9.182.12/")

time.sleep(3)

# Move to popup window to add username and password
print(driver.current_window_handle)  # 18 <-- Parent
handles = driver.window_handles

for handle in handles:
    driver.switch_to.window(handle)
    print(driver.title)
driver.maximize_window()

#def SwitchtoActiveWindow(self):
#    newWindow = Methods()
#    Methods.switchtonewwindow(newWindow)
#    #newWindow.switchtonewwindow()
#    #newWindow.logic_credentials()
#    #newWindow.enterCPRSENR()
#    return

o1 = methods
methods.logic_credentials(o1)
    #enterCPRSENR()
    #Enter_CPR_SENR.enterCPRSENR()





