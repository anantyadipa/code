# scrapping

import undetected_chromedriver as uc
from selenium import webdriver
import time

#tambahan

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import re

#end


# non headless
options = uc.ChromeOptions()
options.add_argument("--user-data-dir=/Users/ruangguru/Library/Application Support/Google/Chrome/Profile 1")
driver = uc.Chrome(options=options, use_subprocess=True)
options.add_experimental_option("detach", True)

# # headless
# options = uc.ChromeOptions()
# options.add_argument("--user-data-dir=/Users/ruangguru/Library/Application Support/Google/Chrome/Profile 1")
# options.add_argument('--headless')
# driver = uc.Chrome(options=options, driver_executable_path='/Users/ruangguru/Documents/python-automation-test/src/chromedriver')
# options.add_experimental_option("detach", True)

# data store
moduleId = 'https://teman.sirogu.com/products/93?m=696'
sectionID = 'https://teman.sirogu.com/products/93?s=10836'
subsectionId = 'https://teman.sirogu.com/products/93?s=10839'

driver.get(sectionID)
driver.implicitly_wait(10)


WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH,"//div[@role='region' and @id='accordion-panel-25']")))
# container = driver.find_elements(By.XPATH, "//div[@role='region' and @id='accordion-panel-25']")

testcase_panel = driver.find_elements(By.XPATH, "//div[@role='region'][@id='accordion-panel-25']/a/div")
testcase_id = driver.find_elements(By.XPATH, "//div[@role='region'][@id='accordion-panel-25']/a/div/span")
testcase_title = driver.find_elements(By.XPATH, "//div[@role='region'][@id='accordion-panel-25']/a/div/p") 


    
# for tc_data in testcase_panel:
#     # print(re.sub(r'TC-', '', tc_data.text))
#     print(tc_data.text)

for id_element, title_element in zip(testcase_id, testcase_title):
    testcase_id_text = re.sub(r"TC-", "", id_element.text)
    testcase_title_text = title_element.text
    print(f'{testcase_id_text} {testcase_title_text}')
    
    

driver.quit()

