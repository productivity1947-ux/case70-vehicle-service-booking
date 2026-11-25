import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

options = webdriver.ChromeOptions()
options.add_argument("--headless=new")
options.add_argument("--window-size=1440,900")
driver = webdriver.Chrome(options=options)
driver.get("http://127.0.0.1:5000/")
assert "Vehicle Service" in driver.page_source
driver.save_screenshot(r"C:\Users\guest pc\Desktop\suhas\Case70_Evidence_Screenshots\Application\Application_01_Dashboard.png")

driver.get("http://127.0.0.1:5000/book")
WebDriverWait(driver,10).until(EC.presence_of_element_located((By.NAME,"vehicle_id")))
driver.save_screenshot(r"C:\Users\guest pc\Desktop\suhas\Case70_Evidence_Screenshots\Module_3_Development\L3_2_Working_Booking_Module.png")

driver.get("http://127.0.0.1:5000/reminders")
WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,"//button[contains(.,'Run scheduler')]")))
driver.find_element(By.XPATH,"//button[contains(.,'Run scheduler')]").click()
time.sleep(1)
driver.save_screenshot(r"C:\Users\guest pc\Desktop\suhas\Case70_Evidence_Screenshots\Module_4_Testing\L4_3_Selenium_Reminder_Run.png")

driver.get(r"file:///C:/Users/guest%20pc/Desktop/suhas/Case70_Vehicle_Service_Booking/htmlcov/index.html")
time.sleep(1)
driver.save_screenshot(r"C:\Users\guest pc\Desktop\suhas\Case70_Evidence_Screenshots\Coverage\L4_2_Coverage_Report_90_Percent.png")

driver.get("http://127.0.0.1:5000/")
time.sleep(1)
driver.save_screenshot(r"C:\Users\guest pc\Desktop\suhas\Case70_Evidence_Screenshots\Final_Application\Final_Application_Dashboard.png")

driver.quit()
