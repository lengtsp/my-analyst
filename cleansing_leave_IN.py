import pandas as pd
import time

import sys
sys.path.insert(0, '/usr/lib/chromium-browser/chromedriver')

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


 
class clssAccessWebsite():
#     base_url= my_url


    def setUp(self):
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        self.driver = webdriver.Chrome('chromedriver', chrome_options = options)
        self.driver.implicitly_wait(4)

    def load_home_page(self, base_url):
        driver = self.driver
#         driver.get(self.base_url)
        driver.get(base_url)
        
    def action_login(self, user_login, pass_login, P1, P2, P3):
        iduser=self.driver.find_element_by_id(P1)
        iduser.clear()
        iduser.send_keys(user_login)

        idpass=self.driver.find_element_by_id(P2)
        idpass.clear()
        idpass.send_keys(pass_login)

        id_submit=self.driver.find_element_by_class_name(P3)
        id_submit.send_keys(Keys.RETURN)

    def action_export_leaveadjustment(self, P1, P2, P3, P4):
        time.sleep(5) # Sleep for 3 seconds
        self.driver.find_element_by_xpath(P1).click()
        self.driver.find_element_by_xpath(P2).click()

        time.sleep(5) # Sleep for 3 seconds
        self.driver.find_element_by_xpath(P3).click()


        time.sleep(2) # Sleep for 3 seconds
        self.driver.find_element_by_xpath(P4).click()
        time.sleep(5)


        

        
    def action_export_empProfile(self, P1, P2, P3):
        time.sleep(2)
        self.driver.find_element_by_xpath(P1).click()
        self.driver.find_element_by_xpath(P2).click()
        self.driver.find_element_by_xpath(P3).click()
        time.sleep(5)


        
    def action_export_movement(self, P1, P2, P3, P4):
       time.sleep(2)
       self.driver.find_element_by_xpath(P1).click()
       self.driver.find_element_by_xpath(P2).click()
       self.driver.find_element_by_xpath(P3).click()
       self.driver.find_element_by_xpath(P4).click()
       time.sleep(5)

       
       


    def action_export_leainq(self, url2, P1, P2, P3, P4, P5):
       print(1)
       time.sleep(3) # Sleep for 3 seconds
       self.driver.find_element_by_xpath(P1).click()
       print(2)
       self.driver.find_element_by_xpath(P2).click()
       time.sleep(3)
       print(3)
       time.sleep(2)
       idtext=self.driver.find_element_by_id(P3)
       print(4)
       action = ActionChains(self.driver)    
       print(5)
       action.double_click(idtext).perform()

       time.sleep(3)
       print(6)

       button1 = self.driver.find_element_by_xpath(P4)
       self.driver.execute_script("arguments[0].click();", button1)
  
       time.sleep(70)
  
       print(7)
       button = self.driver.find_element_by_xpath(P5)
       self.driver.execute_script("arguments[0].click();", button)
       time.sleep(5)
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
def prepare_leaveadjust(filename):
    df            = pd.read_excel(filename)
    df            = df.rename(columns={df.columns[0] : 'emplid'}      )
    df['leave_type_k'] = 0
    df['leave_type_k'] = df['Leave Type'].apply(convert_leavetype)
    df['emplid'] = df['emplid'].astype(str)
    df = df[ (df['emplid'].str[:1] != 'K')] 
    return df

def prepare_empProfile(filename):
    df            = pd.read_excel(filename)
    df            = df.rename(columns={df.columns[0] : 'emplid'}      )
    df['emplid'] = df['emplid'].astype(str)
    df = df[ (df['emplid'].str[:1] != 'K')] 
    return df
   
def prepare_movement(filename):
    df            = pd.read_excel(filename)
    df            = df.rename(columns={df.columns[0] : 'emplid'}      )
    df['emplid'] = df['emplid'].astype(str)
    df = df[ (df['emplid'].str[:1] != 'K')] 
    return df
   
   
def prepare_inquiry(filename):
    df            = pd.read_excel(filename)
    df            = df.rename(columns={df.columns[0] : 'emplid'}      )
    df['emplid'] = df['emplid'].astype(str)
    df = df[ (df['emplid'].str[:1] != 'K')] 
    return df
   
   

def convert_leavetype(ty):
    if ty == 'ANL1':
        return 8
    elif ty == 'ANL2':
        return 8
    elif ty == 'ANLCARR':
        return 34
    elif ty == 'HJL1':
        return 13
    elif ty == 'HJL2':
        return 13
    elif ty == 'MSL':
        return 10
    elif ty == 'MTL_1':
        return 4
    elif ty == 'MTL_2':
        return 33
    elif ty == 'MTL_3':
        return 5

    elif ty == 'MTL_INF1':
        return 4
    elif ty == 'MTL_INF2':
        return 33
    elif ty == 'OLP1.1':
        return 6
    elif ty == 'OLP1.2':
        return 6
    elif ty == 'OLP2':
        return 6

    elif ty == 'PTL':
        return 31
    elif ty == 'PVL':
        return 2
    elif ty == 'SL':
        return 3
    elif ty == 'TNL':
        return 12

    elif ty == 'VPSNL':
        return 30
  
    elif ty == 'STL':
        return 19


    else:
        return 'N/A'

   
