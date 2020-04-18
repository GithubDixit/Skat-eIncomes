import driver as driver

# ************ Generic list of xpath used  **************#
class switchtopopupwindow:
    def __init__(self):
        pass

    def handle_current_window_method(self):
        handles = driver.window_handles
        print("Value of HANDLES IS:",handles)
        for handle in handles:
            driver.switch_to.window(handle)
            driver.maximize_window()

