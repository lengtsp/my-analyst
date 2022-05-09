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
        
    def action_login(self, user_login, pass_login):
        iduser=self.driver.find_element_by_id("ctl00_ContentHolder_user")
        iduser.clear()
        iduser.send_keys(user_login)

        idpass=self.driver.find_element_by_id("ctl00_ContentHolder_pwdbox")
        idpass.clear()
        idpass.send_keys(pass_login)

        id_submit=self.driver.find_element_by_class_name("LoginLink")
        id_submit.send_keys(Keys.RETURN)

    def action_export_leaveadjustment(self):
        time.sleep(5) # Sleep for 3 seconds
        self.driver.find_element_by_xpath('//span[contains(@class, "iconx_m_LV")]').click()
        self.driver.find_element_by_xpath('//a[contains(@href, "/HRM/Leave/eLeaveBulkAdjustment.aspx")]').click()

        time.sleep(5) # Sleep for 3 seconds
        self.driver.find_element_by_xpath('//td[@class="c_content" and  text() =" Get" ]').click()
        #สั่ง download file มันจะมาลง colab ให้เลย

        time.sleep(2) # Sleep for 3 seconds
        self.driver.find_element_by_xpath('//div[contains(@class, "ZG_TExportExcel")]').click()
        time.sleep(5)


    def action_export_empProfile(self):
        time.sleep(2)
        self.driver.find_element_by_xpath('//span[contains(@class, "iconx_m_WM")]').click()
        self.driver.find_element_by_xpath('//a[contains(@href, "/Routing/eEmployeeProfile.aspx")]').click()
        # self.driver.find_element_by_xpath('//div[contains(@class, "x8Hbtn")]').click()
        self.driver.find_element_by_xpath('//div[contains(@class, "x8Exportbtn")]').click()
        time.sleep(5)

    def action_export_movement(self):
       time.sleep(2)
       self.driver.find_element_by_xpath('//span[contains(@class, "iconx_m_WM")]').click()
       self.driver.find_element_by_xpath('//a[contains(@href, "/Routing/EHR_Movement.aspx")]').click()
       self.driver.find_element_by_xpath('//div[contains(@class, "x8Hbtn")]').click()
       self.driver.find_element_by_xpath('//div[contains(@class, "x8Exportbtn")]').click()
       time.sleep(5)

       
    def action_export_leainq(self):
       print(1)
       time.sleep(3) # Sleep for 3 seconds
       self.driver.find_element_by_xpath('//span[contains(@class, "iconx_m_LV")]').click()
       print(2)
       self.driver.find_element_by_xpath('//a[contains(@href, "/ESS/ELeave/eLeaveApplicationInquiry.aspx")]').click()
       time.sleep(3)
       print(3)
       
       idtext=self.driver.find_element_by_id("BtnEmpSetup_ctl00_ContentHolder_aceSearchEmployeeID")
       print(4)
       action = ActionChains(self.driver)    
       print(5)
       action.double_click(idtext).perform()

       time.sleep(1)
       print(6)
       self.driver.find_element_by_xpath('//*[@id="ctl00_ContentHolder_btnGetPop"]/a/table/tbody/tr/td[2]').click()
       time.sleep(10)

       print(7)
       self.driver.find_element_by_xpath('//div[contains(@class, "ZG_TExportExcel")]').click()
       time.sleep(5)
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
def prepare_leaveadjust(filename):
    df            = pd.read_excel(filename)
    df            = df.rename(columns={df.columns[0] : 'emplid'}      )
    df['leave_type_k'] = 0
    df['leave_type_k'] = df['Leave Type'].apply(convert_leavetype)
    df = df[ (df['emplid'].str[:1] != 'K')] 
    df['emplid'] = df['emplid'].astype(str)
    return df

def prepare_empProfile(filename):
    df            = pd.read_excel(filename)
    df            = df.rename(columns={df.columns[0] : 'emplid'}      )
    df = df[ (df['emplid'].str[:1] != 'K')] 
    df['emplid'] = df['emplid'].astype(str)
    return df
   
def prepare_movement(filename):
    df            = pd.read_excel(filename)
    df            = df.rename(columns={df.columns[0] : 'emplid'}      )
    df = df[ (df['emplid'].str[:1] != 'K')] 
    df['emplid'] = df['emplid'].astype(str)
    return df
   
   
def prepare_inquiry(filename):
    df            = pd.read_excel(filename)
    df            = df.rename(columns={df.columns[0] : 'emplid'}      )
    df = df[ (df['emplid'].str[:1] != 'K')] 
    df['emplid'] = df['emplid'].astype(str)
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

   
