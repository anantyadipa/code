import undetected_chromedriver as uc
from openpyxl import load_workbook
import time

#tambahan

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
import re

#end

wb = load_workbook(filename='/Users/ruangguru/Documents/python-automation-test/src/automation.xlsx')
sheetRange = wb['Store Id']

# non headless
options = uc.ChromeOptions()
options.add_argument("--user-data-dir=/Users/ruangguru/Library/Application Support/Google/Chrome/Profile 1")
driver = uc.Chrome(options=options, use_subprocess=True)
options.add_experimental_option("detach", True)

# # headless
# options = uc.ChromeOptions()
# options.add_argument("--user-data-dir=/Users/ruangguru/Library/Application Support/Google/Chrome/Profile 1")
# options.add_argument('--headless')
# driver = uc.Chrome(options=options, use_subprocess=True)
# options.add_experimental_option("detach", True)

# data store
id = 93
empty_cell_found = False
message = ' - telah terhapus.'
#looping

i = 2

while i <= len(sheetRange['A']):
    # Title	Country	Platform	Stream	References	PRD	Design Link	Automation	Priority	Feature Flag	Precondition

    TC_ID = sheetRange['A'+str(i)].value
    driver.implicitly_wait(10)
    
    
    if TC_ID is None:
     TC_ID =[]
     break
    

    driver.get("https://teman.sirogu.com/products/" + str(id) + "/testcase/" + str(TC_ID) + "/delete")
    driver.implicitly_wait(10)


    
    try:
        WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH, '//button[@type=\'submit\']')))
        # driver.find_element(By.NAME, 'title').clear()
        if TC_ID is not None:
         driver.find_element(By.XPATH, '//button[@type=\'submit\']').click()
         testcase_title = driver.find_elements(By.XPATH, "//div[@role='region'][@id='accordion-panel-25']/a/div/p")
         for title_element in testcase_title:
            testcase_title_text = re.sub(r"Add case", "", title_element.text)
            # testcase_title_text = title_element.text
          
         print (f'{TC_ID}{message}')
        else:
           empty_cell_found = True
         
     
    except TimeoutException:
        print('gak cocok')
        pass

    
    time.sleep(1)
    i = i +1


if empty_cell_found:
    print('End of the list. Stopping the script.')
else:
    print('done')
driver.quit()
