import undetected_chromedriver as uc
from openpyxl import load_workbook
import time

#tambahan

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
import re
from selenium.webdriver.common.keys import Keys

#end

wb = load_workbook(filename='/Users/ruangguru/Documents/python-automation-test/src/automation.xlsx')
sheetRange = wb['Bulk Copy']

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
message = ' - copied to '
#looping

i = 2

while i <= len(sheetRange['A']):
    # Title	Country	Platform	Stream	References	PRD	Design Link	Automation	Priority	Feature Flag	Precondition
    # Product	Module	Section	Sub Section
    TC_ID = sheetRange['A'+str(i)].value
    Product = sheetRange['B'+str(i)].value
    Module = sheetRange['C'+str(i)].value
    Section = sheetRange['D'+str(i)].value
    Sub_Section = sheetRange['E'+str(i)].value
    driver.implicitly_wait(10)
    
    
    if TC_ID is None:
     TC_ID =[]
     break
    

    driver.get("https://teman.sirogu.com/products/" + str(id) + "/testcase/" + str(TC_ID) + "/copy")
    driver.implicitly_wait(10)


    try:
        WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH, '//button[@type=\'submit\']')))
        
        if TC_ID is not None:
            print (f'Moving-{TC_ID} in progress..')
        else:
            empty_cell_found = True  # Sel kosong ditemukan di 'Title'
            break

        if Product:
            driver.implicitly_wait(5)
            driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[1]/div[2]/div[1]/div/div[2]').click()
            driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[1]/div[2]/div[1]/div/div[1]/input').send_keys(Product)
            driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[1]/div[2]/div[1]/div/div[1]/input').send_keys(Keys.ENTER)
            time.sleep(0.3)
            # driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[1]/div[2]/div[1]/div/div[1]/input').send_keys(Keys.TAB)
            # driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[1]/div[2]/div[1]/div/div[1]/input').send_keys(Keys.TAB)
            # driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[1]/div[2]/div[1]/ul/li').click()

        if Module:
            driver.implicitly_wait(5)
            driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[1]/div[1]/div/div[2]').click()
            driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[1]/div[1]/div/div[1]/input').send_keys(Module)
            driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[1]/div[1]/div/div[1]/input').send_keys(Keys.ENTER)
            time.sleep(0.3)
            # driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[1]/div[1]/div/div[1]/input').send_keys(Keys.TAB)
            # driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[1]/div[1]/div/div[1]/input').send_keys(Keys.TAB)
            driver.implicitly_wait(5)
            # driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[1]/div[1]/ul/li').click()

        if Section:
            driver.implicitly_wait(5)
            driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[2]/div[1]/div/div[2]').click()
            driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[2]/div[1]/div/div[1]/input').send_keys(Section)
            driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[2]/div[1]/div/div[1]/input').send_keys(Keys.ENTER)
            time.sleep(0.3)
            # driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[2]/div[1]/div/div[1]/input').send_keys(Keys.TAB)
            # driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[2]/div[1]/div/div[1]/input').send_keys(Keys.TAB)
            driver.implicitly_wait(5)
            # driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[2]/div[1]/ul/li').click()
     
        if Sub_Section:
            driver.implicitly_wait(5)
            driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[3]/div[1]/div/div[2]').click()
            driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[3]/div[1]/div/div[1]/input').send_keys(Sub_Section)
            driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[3]/div[1]/div/div[1]/input').send_keys(Keys.ENTER)
            time.sleep(0.3)
            # driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[3]/div[1]/div/div[1]/input').send_keys(Keys.TAB)
            driver.implicitly_wait(5)
            # driver.find_element(By.XPATH, '//div[3]/div[4]/div/section/form/div/div/div[2]/div[3]/div[1]/ul/li').click()
        
        
        driver.find_element(By.XPATH, "//div[3]/div[4]/div/section/form/footer/div/button[2]").click()
        print (f'{TC_ID}{message}{Product}>{Module}>{Section}>{Sub_Section}')
        time.sleep(1)

        p = driver.current_window_handle
        driver.implicitly_wait(5)
        # ngeklik link copy of TC di toast success copynya
        driver.find_element(By.CSS_SELECTOR, "li[class='chakra-toast'] a[class='chakra-link css-1e1c3k0']").click()

        #di new tab edit copy test casenya
        chwd = driver.window_handles
        for w in chwd:
            if(w!=p):
                driver.switch_to.window(w)
            
        
        time.sleep(0.9)
        WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH, "//div[1]/div[3]/div/div[2]/div/form/div[1]/div/input")))
        driver.find_element(By.XPATH, "//div[1]/div[3]/div/div[2]/div/form/div[1]/div/input").click()
        driver.find_element(By.XPATH, "//div[1]/div[3]/div/div[2]/div/form/div[1]/div/input").send_keys(Keys.HOME)
        for _ in range(8):
            driver.find_element(By.XPATH, "//div[1]/div[3]/div/div[2]/div/form/div[1]/div/input").send_keys(Keys.DELETE)
        driver.find_element(By.XPATH, "//button[@type=\'submit\']").click()
        time.sleep(1)
        driver.close()
        driver.switch_to.window(p)

    
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
