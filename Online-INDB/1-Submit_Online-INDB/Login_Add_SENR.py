from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

import time

# open Firefox Browser
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

# On the popup screen select/click/activate the frame to provide login credentials
driver.switch_to.frame("letloen")
time.sleep(5)
driver.find_element_by_css_selector('#defaultFocusElementId').send_keys("Admin1032")
driver.find_element_by_name("password").send_keys("Test1234")
driver.find_element_by_xpath("//*[@id='defaultButton']").click()

# Enter CVR/SE-nr Number
time.sleep(5)
driver.switch_to.__class__("clContentTable")
time.sleep(5)
driver.find_element_by_class_name("inputButton").click()
time.sleep(5)
driver.find_element_by_id("defaultFocusElementId").send_keys("37172715")
driver.find_element_by_xpath("//*[@id='defaultButton']").click()
