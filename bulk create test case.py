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


wb = load_workbook(filename='/Users/ruangguru/Documents/automation.xlsx')
sheetRange = wb['Full Field input test case']

options = uc.ChromeOptions()
options.add_argument("--user-data-dir=/Users/ruangguru/Library/Application Support/Google/Chrome/Profile 1")
driver = uc.Chrome(options=options, use_subprocess=True)
options.add_experimental_option("detach", True)
# data store
id = 93
testCaseId = 10836
# https://teman.sirogu.com/products/93/10836/create
# driver.get("https://teman.sirogu.com/products/" + str(id) +"/" + str(testCaseId) + "/create")
# driver.implicitly_wait(10)

#looping

i = 2

while i <= len(sheetRange['A']):
    # Title	Country	Platform	Stream	References	PRD	Design Link	Automation	Priority	Feature Flag	Precondition
    Title = sheetRange['A'+str(i)].value
    Country = sheetRange['B'+str(i)].value
    Platform = sheetRange['C'+str(i)].value
    Stream = sheetRange['D'+str(i)].value
    References = sheetRange['E'+str(i)].value
    PRDn = sheetRange['F'+str(i)].value
    Design_Link = sheetRange['G'+str(i)].value
    Automation = sheetRange['H'+str(i)].value
    Priority = sheetRange['I'+str(i)].value
    Precondition = sheetRange['K'+str(i)].value
    Feature_Flag = sheetRange['J'+str(i)].value
    
    action1 = sheetRange['L'+str(i)].value
    expected1 = sheetRange['M'+str(i)].value
    action2 = sheetRange['N'+str(i)].value
    expected2 = sheetRange['O'+str(i)].value
    action3 = sheetRange['P'+str(i)].value
    expected3 = sheetRange['Q'+str(i)].value
    action4 = sheetRange['R'+str(i)].value
    expected4 = sheetRange['S'+str(i)].value
    action5 = sheetRange['T'+str(i)].value
    expected5 = sheetRange['U'+str(i)].value

    driver.get("https://teman.sirogu.com/products/" + str(id) +"/" + str(testCaseId) + "/create")
    driver.implicitly_wait(10)
    
    
    driver.find_element(By.NAME, "title").send_keys()
    
    try:
        WebDriverWait(driver,10).until(EC.visibility_of_element_located((By.NAME,'title')))
        # driver.find_element(By.NAME, 'title').clear()
        driver.find_element(By.NAME, 'title').send_keys(Title)
        driver.find_element(By.NAME, ('references')).send_keys(References)
        driver.find_element(By.NAME, ('prdLink')).send_keys(PRDn)
        driver.find_element(By.NAME, ('designLink')).send_keys(Design_Link)
        driver.find_element(By.NAME, ('isAutomation')).send_keys(Automation)
        driver.find_element(By.NAME, ('priorityId')).send_keys(Priority)
        driver.find_element(By.NAME, ('precondition')).send_keys(Precondition)
        driver.find_element(By.NAME, ('featureFlag')).send_keys(Feature_Flag)
        
          # stream
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[1]/div[6]/div/div[2]/div/div/div[2]/button').click()
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[1]/div[6]/div/div[2]/div/div/div[1]/input').send_keys(Stream)
        time.sleep(0.3)
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[1]/div[6]/div/div[2]/div/ul/li/div/div').click()
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[1]/div[6]/div/div[2]/div/div/div[2]/button[2]').click()
        #country
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[1]/div[1]/div/div[2]/div/div/div[2]/button').click()
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[1]/div[1]/div/div[2]/div/div/div[1]/input').send_keys(Country)
        time.sleep(0.3)
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[1]/div[1]/div/div[2]/div/div/div[1]/input').send_keys(Keys.ENTER)
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[1]/div[1]/div/div[2]/div/div/div[2]/button[2]').click()
        # platform
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[1]/div[3]/div/div[2]/div/div/div[2]/button').click()
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[1]/div[3]/div/div[2]/div/div/div[1]/input').send_keys(Platform)
        time.sleep(0.3)
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[1]/div[3]/div/div[2]/div/div/div[1]/input').send_keys(Keys.ENTER)
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[1]/div[1]/div[3]/div/div[2]/div/div/div[2]/button[2]').click()
       
        #add steps 1
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/button').click()
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/div[2]/div[2]/div[1]/textarea').send_keys(action1)
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/div[2]/div[2]/div[2]/textarea').send_keys(expected1)

         #add steps 2
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/button').click()
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/div[3]/div[2]/div[1]/textarea').send_keys(action2)
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/div[3]/div[2]/div[2]/textarea').send_keys(expected2)
         #add steps 3
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/button').click()
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/div[4]/div[2]/div[1]/textarea').send_keys(action3)
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/div[4]/div[2]/div[2]/textarea').send_keys(expected3)
         #add steps 4
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/button').click()
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/div[5]/div[2]/div[1]/textarea').send_keys(action4)
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/div[5]/div[2]/div[2]/textarea').send_keys(expected4)
         #add steps 5
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/button').click()
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/div[6]/div[2]/div[1]/textarea').send_keys(action5)
        driver.find_element(By.XPATH, '//div[1]/div[3]/div/div[2]/div/form/div[2]/div[3]/div/div[2]/div[6]/div[2]/div[2]/textarea').send_keys(expected5)

        #submit
        driver.find_element(By.XPATH, '//button[@type=\'submit\']').click()

    
      
     
    except TimeoutException:
        print('gak cocok')
        pass
    
    time.sleep(1)
    i = i +1

print('udahan')
driver.quit()