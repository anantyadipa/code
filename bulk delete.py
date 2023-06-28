import undetected_chromedriver as uc
from openpyxl import load_workbook
import time

#tambahan

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys



#end


wb = load_workbook(filename='/Users/ruangguru/Documents/test QA automation test case.xlsx')
sheetRange = wb['Bulk Delete']

options = uc.ChromeOptions()
options.add_argument("--user-data-dir=/Users/ruangguru/Library/Application Support/Google/Chrome/Profile 1")
driver = uc.Chrome(options=options, use_subprocess=True)
options.add_experimental_option("detach", True)

# data store
id = 93


#looping

i = 2

while i <= len(sheetRange['A']):
    # Title	Country	Platform	Stream	References	PRD	Design Link	Automation	Priority	Feature Flag	Precondition
    TC_ID = sheetRange['A'+str(i)].value
    
   
    driver.get("https://teman.sirogu.com/products/" + str(id) + "/testcase/" + str(TC_ID) + "/delete")
    if TC_ID is None:
     print('End of Cell due to Empty value. Stopping the script')
     driver.quit()
    pass
    driver.implicitly_wait(10)

   

        
    try:
        WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH, '//button[@type=\'submit\']')))
        # driver.find_element(By.NAME, 'title').clear()
        driver.find_element(By.XPATH, '//button[@type=\'submit\']').click()
         
     
    except TimeoutException:
        print('gak cocok')
        pass

    
    time.sleep(1)
    i = i +1
    


print('udahan')
driver.quit()
