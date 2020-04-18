import driver as driver
import time
from selenium import webdriver



# On the popup screen select/click/acativate the frame to provide login credentials
class methods:
    def logic_credentials(self):
        #time.sleep(3)
        #driver.switch_to.frame("letloen")
        time.sleep(5)
        driver.find_element_by_css_selector('#defaultFocusElementId').send_keys("Admin1032")
        driver.find_element_by_name("password").send_keys("Test1234")
        driver.find_element_by_xpath("//*[@id='defaultButton']").click()


    def enterCPRSENR(self):
    # Enter CVR/SE-nr Number
        time.sleep(5)
        driver.switch_to.__class__("clContentTable")
        time.sleep(5)
        driver.find_element_by_class_name("inputButton").click()
        time.sleep(5)
        driver.find_element_by_id("defaultFocusElementId").send_keys("37172715")
        driver.find_element_by_xpath("//*[@id='defaultButton']").click()
