from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait


options = Options()
options.add_argument("--start-maximized")
options.add_experimental_option("detach",True)

driver = webdriver.Chrome(options=options)

url = "C:/Users/admin/Desktop/김동환/Worksapce/Selenium/경비청구 자동화/exaple.html"

driver.get(url)
time.sleep(1)


button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//button[text()='등록']"))
)

# 버튼 클릭
button.click()

# 1. ID입력


