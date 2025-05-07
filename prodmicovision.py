from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import json
from datetime import datetime
import pytz
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import subprocess
import pandas as pd
import requests
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.support.ui import Select
from datetime import datetime
def send_table_to_slack(webhook_url,table, message="Phase-1 Automation Report of Prod:"):
    payload = {
        "text": f"{message}\n{table}"  # Wrap in triple backticks for code block formatting in Slack
    }
    
    response = requests.post(webhook_url, json=payload)
    if response.status_code == 200:
        print("Table sent successfully!")
    else:
        print(f"Failed to send table. Status code: {response.status_code}, Response: {response.text}")

# Get current date and time


# def automation(url,email,password):
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

driver.get("https://prod.micovision.micoworks-system.jp/login")
time.sleep(3)
driver.maximize_window()


scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('micovision-cred.json', scope)
client = gspread.authorize(creds)


sheet = client.open_by_key("1l0m5lB4KdsBOh3u3UCT6Z0fW9BGrhfwwD6s5_v8zIMg")
worksheet = sheet.worksheet('module1_prod')
now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Convert to string
worksheet.update('B2', [[now]])  


def image_work():
    try:
        first_img = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[2]/div[1]/div/div[1]/div/span/img')
        first_img.click()
        time.sleep(5) 
        message_7 = "pass"
    except:
        message_7 = "fail"
        pass
    worksheet.update('B9', [[message_7]])  

    try:
        download_img = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[2]/div[1]/div[2]/div[1]/div/div[2]/div[1]/div[2]/ul/li[1]/a')
        download_img.click()
        time.sleep(5) 
        message_8 = "pass"
    except:
        message_8 = "fail"
        pass
    worksheet.update('B10', [[message_8]])  
    
    try:
        copy_link = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[2]/div[1]/div[2]/div[1]/div/div[2]/div[1]/div[2]/ul/li[2]/a')
        copy_link.click()
        time.sleep(5) 
        message_9 = "pass"
    except:
        message_9 = "fail"
        pass
    worksheet.update('B11', [[message_9]])  
    
    try:
        right_move = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[2]/div[1]/div[2]/div[1]/div/div[2]/div[1]/div[2]/ul/li[4]/a')
        right_move.click()
        time.sleep(5) 
        message_10 = "pass"
    except:
        message_10 = "fail"
        pass
    worksheet.update('B12', [[message_10]])  

    try:
        left_move = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[2]/div[1]/div[2]/div[1]/div/div[2]/div[1]/div[2]/ul/li[3]/a')
        left_move.click()
        time.sleep(5) 
        message_11 = "pass"
    except:
        message_11 = "fail"
        pass
    worksheet.update('B13', [[message_11]])  

    try:
        ceoss_mark = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[2]/div[1]/div[2]/div[1]/div/div[2]/div[1]/div[2]/ul/li[5]/a')
        ceoss_mark.click()
        time.sleep(10) 
        message_12 = "pass"
    except:
        message_12 = "fail"
        pass
    worksheet.update('B14', [[message_12]])  

def login(user_name,passw):
 
    login_button = driver.find_element(By.XPATH,'//*[@id="root"]/div[1]/div/div[1]/div/div[3]/div/div/span[3]/a')
    login_button.click()
    time.sleep(5)

    usernames = driver.find_elements(By.ID, 'signInFormUsername')
    # username = "a.chaudhary+systemadmin@micoworks.jp"   
    usernames[1].send_keys(user_name)
    time.sleep(5)
    password = driver.find_elements(By.ID, 'signInFormPassword')
    # pssw = "Avinash@9892"
    password[1].send_keys(passw)
    time.sleep(5)
    butn_name = driver.find_elements(By.NAME, 'signInSubmitButton')
    butn_name[1].click() 
    time.sleep(20) 


try:
    login("a.chaudhary+systemadmin@micoworks.jp","Avinash@9892")
    message_1 = "pass"
except:
    message_1 = "fail"

worksheet.update('B3', [[message_1]])  



def search_part():

    try:
        search_button = driver.find_element(By.XPATH,'//*[@id="root"]/div[1]/div[2]/a[1]')
        search_button.click()
        time.sleep(2) 
        message_2 = "pass"
    except:
        message_2 = "fail"
        pass
    worksheet.update('B4', [[message_2]])  

    try:
        send_key_search = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div/div/div/div/div[1]/input')
        send_key_search.send_keys("salon")
        time.sleep(2) 
        message_3 = "pass"
    except:
        message_3 = "fail"
        pass
    worksheet.update('B5', [[message_3]])  

    try:
        search_button_1 = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div/div/div/div/div[5]/button[1]')
        search_button_1.click()
        time.sleep(10) 
        message_4 = "pass"
    except:
        message_4 = "fail"
        pass
    worksheet.update('B6', [[message_4]])  

    try:
        group_sim_img = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[1]/div[2]/div/label/span[1]/span[1]/input')
        group_sim_img.click()
        time.sleep(10) 
        message_5 = "pass"
    except:
        message_5 = "fail"
        pass
    worksheet.update('B7', [[message_5]])  
    
    try:
        off_debug = driver.find_element(By.XPATH,'//*[@id="root"]/div[1]/div/header/div/div/div/div/div[2]/ul/div/li/div/div/input')
        off_debug.click()
        time.sleep(10) 
        message_6 = "pass"
    except:
        message_6 = "fail"
        pass
    worksheet.update('B8', [[message_6]])  

    image_work()

    try:
        plus_image = driver.find_element(By.XPATH,'//*[@id="root"]/div[1]/div/header/div/div/div/div/div[2]/ul/div/li/div/div/input')
        plus_image.click()
        time.sleep(10) 
        message_13 = "pass"
    except:
        message_13 = "fail"
        pass
    worksheet.update('B15', [[message_13]])  

    try:
        minus_image = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[1]/div[1]/div[2]/a[1]/img')
        minus_image.click()
        time.sleep(10) 
        message_14 = "pass"
    except:
        message_14 = "fail"
        pass
    worksheet.update('B16', [[message_14]])  

    try:
        dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "MuiAutocomplete-input"))
        )
        dropdown.click()
        time.sleep(10) 

        # Wait for the first option to be visible and select it
        first_option = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "li[role='option']"))
        )
        first_option.click()
        time.sleep(10) 
        message_15 = "pass"
    except:
        message_15 = "fail"
        pass
    worksheet.update('B17', [[message_15]])  
    #####################################################################################
    try:
        apply_button = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[1]/div[11]/button[2]')
        apply_button.click()
        time.sleep(10) 
        message_16 = "pass"
    except:
    #     message_16 = "fail"
        pass
    # worksheet.update('B18', [[message_16]])  

    try:
        clear_all_button = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[1]/div[11]/button[1]')
        clear_all_button.click()
        time.sleep(10) 
        message_17 = "pass"
    except:
    #     message_17 = "fail"
        pass
    # worksheet.update('B19', [[message_17]])  

    try:
        tags_button = driver.find_element(By.XPATH,'(//*[contains(@class, "MuiAutocomplete-input")])[3]')
        tags_button.click()
        time.sleep(10) 
        message_18 = "pass"
    except:
        message_18 = "fail"
        pass
    worksheet.update('B18', [[message_18]])  

    try:
    # Wait for the first option to be visible and click it
        first_option_tag = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, "//li[@role='option'][1]"))
        )
        first_option_tag.click()
        apply_button.click()
        time.sleep(10) 
        clear_all_button.click()
        time.sleep(10) 
        message_19 = "pass"
    except:
        message_19 = "fail"
        pass
    worksheet.update('B19', [[message_19]])  

    try:
        ctr = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[1]/div[5]/div[2]/div/div/input')
        ctr.send_keys(5)
        apply_button.click()
        time.sleep(10) 
        clear_all_button.click()
        time.sleep(10) 
        message_20 = "pass"
    except:
        message_20 = "fail"
        pass
    worksheet.update('B20', [[message_20]])  

    try:
        img_status = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[1]/div[6]/div[2]/div/div/label[2]/span[1]/input')
        img_status.click()
        time.sleep(10) 
        apply_button.click()
        time.sleep(10) 
        clear_all_button.click()
        time.sleep(10) 
        message_21 = "pass"
    except:
        message_21 = "fail"
        pass
    worksheet.update('B21', [[message_21]])  

    try:
        img_type = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[1]/div[7]/div[2]/div/div/label[2]/span[2]')
        img_type.click()
        time.sleep(10) 
        apply_button.click()
        time.sleep(10) 
        clear_all_button.click()
        time.sleep(10) 
        message_22 = "pass"
    except:
        message_22 = "fail"
        pass
    worksheet.update('B22', [[message_22]])  

    try:
        delivery_type = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[1]/div[8]/div[2]/div/div/label[2]/span[2]')
        delivery_type.click()
        time.sleep(10) 
        apply_button.click()
        time.sleep(10) 
        clear_all_button.click()
        time.sleep(10) 
        message_23 = "pass"
    except:
        message_23 = "fail"
        pass
    worksheet.update('B23', [[message_23]])  


    try:
        photo_ = driver.find_element(By.XPATH,'//*[@id="root"]/section/div/div[2]/div/div/div[1]/div[10]/div[2]/div/div/label[2]/span[1]/input')
        photo_.click()
        time.sleep(10) 
        apply_button.click()
        time.sleep(10) 
        clear_all_button.click()
        time.sleep(10) 
        message_24 = "pass"
    except:
        message_24 = "fail"
        pass
    worksheet.update('B24', [[message_24]])  

search_part()



data = worksheet.get_all_values()

# Convert the data to a DataFrame
df = pd.DataFrame(data[1:], columns=data[0])
table_str = df.to_markdown(index=False) 

status_list = df['status'].to_list()

# for i in status_list
if "Fail" in df['status'].values:
    print("Faillllllllllllllllllllllllllllll")
    send_table_to_slack("https://hooks.slack.com/services/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",table_str)

else:
    print("passsssssssssssssssssssssssssssss")

    send_table_to_slack("https://hooks.slack.com/services/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",table_str)


webhook_url = 'https://hooks.slack.com/services/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'

driver.quit()
