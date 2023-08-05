import undetected_chromedriver as uc
from openpyxl import load_workbook
import time

#tambahan

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException


#end

wb = load_workbook(filename='/Users/ruangguru/Documents/python-automation-test/src/automation.xlsx')
sheetRange = wb['Bulk Edit']

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
id = 14
empty_cell_found = False
message = ' - telah ter edit.'
#looping

i = 2

while i <= len(sheetRange['A']):
    # Title	Country	Platform	Stream	References	PRD	Design Link	Automation	Priority	Feature Flag	Precondition

    TC_ID = sheetRange['A'+str(i)].value
    TC_Title = sheetRange['B'+str(i)].value
    prereq = sheetRange['C'+str(i)].value
    precond = sheetRange['D'+str(i)].value
    action1 = sheetRange['E'+str(i)].value
    expected1 = sheetRange['F'+str(i)].value
    action2 = sheetRange['G'+str(i)].value
    expected2 = sheetRange['H'+str(i)].value
    action3 = sheetRange['I'+str(i)].value
    expected3 = sheetRange['J'+str(i)].value
    action4 = sheetRange['K'+str(i)].value
    expected4 = sheetRange['L'+str(i)].value
    action5 = sheetRange['M'+str(i)].value
    expected5 = sheetRange['N'+str(i)].value
    testData = sheetRange['O'+str(i)].value
    
    driver.implicitly_wait(0.3)
    
    
    if TC_ID is None:
     TC_ID =[]
     break
    

    driver.get("https://teman.sirogu.com/products/" + str(id) + "/testcase/" + str(TC_ID) + "/edit")
    driver.implicitly_wait(10)
    time.sleep(1)


    
    try:
        WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.XPATH, '//button[@type=\'submit\']')))
        # driver.find_element(By.NAME, 'title').clear()
        if TC_ID is not None:
            print (f'Editing TC-{TC_ID} in progress..')
        else:
            empty_cell_found = True  # Sel kosong ditemukan di 'Title'
            break
                
        prereqs = prereq.split('\n')
        testDatas = testData.split('\n\n')

        if TC_Title:
            driver.find_element(By.NAME, "title").clear()
            driver.find_element(By.NAME, "title").send_keys(TC_Title)    
       
        if prereq:
            # driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[2]/div/div/div/div[2]/button[1]').click()
            driver.find_element(By.XPATH, '//div[1]/div[3]/div[1]/div[2]/div[1]/form[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/input[1]').click()
            for _ in range(10):
                driver.find_element(By.XPATH, '//div[1]/div[3]/div[1]/div[2]/div[1]/form[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/input[1]').send_keys(Keys.BACKSPACE)
            driver.implicitly_wait(0.3)
                       

            for prline in prereqs:
                driver.implicitly_wait(0.3)
                driver.find_element(By.XPATH, '//div[1]/div[3]/div[1]/div[2]/div[1]/form[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/input[1]').click()
                driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[2]/div/div/div/div/input').send_keys(prline)
                driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[2]/div/div/div/div/input').send_keys(Keys.ENTER)
                driver.implicitly_wait(0.1)
                driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[2]/div/div/div/div/input').send_keys(Keys.ESCAPE)

            time.sleep(1)

        if precond:
            driver.find_element(By.NAME, ('precondition')).clear()
            driver.find_element(By.NAME, ('precondition')).send_keys(precond)

        
        if  testData:
            while True:
                try:
                    element = driver.find_element(By.CSS_SELECTOR, ".css-ipw2yt:nth-child(2) g > path")
                    element.click()
                    driver.implicitly_wait(0.3)

                except NoSuchElementException:
                    break

            for tdline in testDatas:
                driver.implicitly_wait(0.5)
                driver.find_element(By.CSS_SELECTOR, ".css-3ype3w").click()
                driver.implicitly_wait(0.5)  
                driver.find_element(By.XPATH, "//section/div/div/div/div/input").send_keys(tdline)
                time.sleep(0.5)
                driver.implicitly_wait(0.5)
                driver.find_element(By.XPATH, "//button/div/div/label/span").click()
                driver.implicitly_wait(0.5)
                driver.find_element(By.XPATH, "//footer/button[2]").click()

                if not testDatas:
                 print('')
            
      
                
        # add steps 1
        if action1 and expected1:
            while True:
                try:
                    element = driver.find_element(By.CSS_SELECTOR, ".css-yyzash > .chakra-stack > .chakra-icon path")
                    element.click()
                    driver.implicitly_wait(0.3)

                except NoSuchElementException:
                    break

            driver.find_element(By.CSS_SELECTOR, ".css-tpdm63").click()
            driver.find_element(By.XPATH, "//div[2]/div/textarea").send_keys(action1)
            driver.find_element(By.XPATH, "//div[2]/textarea").send_keys(expected1)

        # add steps 2
        if action2 and expected2:
            driver.find_element(By.CSS_SELECTOR, ".css-tpdm63").click()
            driver.find_element(By.XPATH, "//div[3]/div[2]/div/textarea").send_keys(action2)
            driver.find_element(By.XPATH, "//div[3]/div[2]/div[2]/textarea").send_keys(expected2)
           
    
        # add steps 3
        if action3 and expected3:
            driver.find_element(By.CSS_SELECTOR, ".css-tpdm63").click()
            driver.find_element(By.XPATH, "//div[4]/div[2]/div/textarea").send_keys(action3)
            driver.find_element(By.XPATH, "//div[4]/div[2]/div[2]/textarea").send_keys(expected3)
    
        # add steps 4
        if action4 and expected4:
            driver.find_element(By.CSS_SELECTOR, ".css-tpdm63").click()
            driver.find_element(By.XPATH, "//div[5]/div[2]/div/textarea").send_keys(action4)
            driver.find_element(By.XPATH, "//div[5]/div[2]/div[2]/textarea").send_keys(expected4)
    
        # add steps 5
        if action5 and expected5:
            driver.find_element(By.CSS_SELECTOR, ".css-tpdm63").click()
            driver.find_element(By.XPATH, "//div[6]/div[2]/div/textarea").send_keys(action5)
            driver.find_element(By.XPATH, "//div[6]/div[2]/div[2]/textarea").send_keys(expected5)
       
        driver.find_element(By.XPATH, '//button[@type=\'submit\']').click() 

        # submit

        finalMessage = f'TC-{TC_ID}{message}'
        print(finalMessage)

    except TimeoutException:
        print('something is wrong')
        pass

    
    time.sleep(1)
    i = i +1


if empty_cell_found:
    print('End of the list. Stopping the script.')
else:
    print('done')
driver.quit()
