import schedule
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import Config
import Outlook
from datetime import date


def ScrapeData():
    print("\n 2. This ScrapeData func from Main file")
    driver = webdriver.Chrome("C:/Users/SourabhGupta/PycharmProjects/Selenium/chromedriver.exe")
    driver.maximize_window()
    driver.get("https://confluence.rampgroup.com/login.action?logout=true")

    #   Enter the user name.
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable(('id', "os_username"))).send_keys(Config.User_Name)

    #   Enter the password.
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable(('id', "os_password"))).send_keys(Config.Password)

    #   Click on login.
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable(('id', "loginButton"))).send_keys(Keys.ENTER)

    #   Click on Search button.
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable(('id', 'quick-search-query'))).send_keys(Keys.ENTER)

    #   Enter the page name.
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable(('id', 'search-filter-input'))).send_keys('Gerrit-Bitbucket Daily commit tracker')

    #   Select first option in search results.
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="search-result-container"]/div[2]/a[1]'))).click()

    #   Read the Table and conver it to ""DataFrame"".
    # table1 = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main-content"]/div[1]')))
    
    IPC_data = pd.read_html(WebDriverWait(driver, 20).until(lambda x: x.find_element(by = By.XPATH, value='//*[@id="main-content"]/div[1]')).get_attribute('outerHTML'), header=0, index_col= 0)[0]
    # IPC_data = pd.read_html(driver.find_element(by=By.XPATH, value='//*[@id="main-content"]/div[1]').get_attribute('outerHTML'), header=0, index_col= 0)[0]

    driver.execute_script("window.scrollTo(0, 10000)")
    time.sleep(10)

    RTOS_data = pd.read_html(WebDriverWait(driver, 20).until(lambda x: x.find_element(by = By.XPATH, value='/html/body/div[2]/div/div[2]/div[2]/div[4]/div[5]/div[2]/table')).get_attribute('outerHTML'), header=0, index_col= 0)[0]
    # RTOS_data = pd.read_html(driver.find_element(by=By.XPATH, value='//*[@id="main-content"]/div[2]/table').get_attribute('outerHTML'), header=0, index_col= 0)[0]

    print(f"<<<<<<<<<<<<<<<<<<<___________________________This is IPC_data______________________________________>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>\n {IPC_data}\n")
    print(type(IPC_data))
    print(f"<<<<<<<<<<<<<<<<<<<___________________________This is RTOS_data______________________________________>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>\n {RTOS_data}\n")
    print(type(RTOS_data))
    data = [IPC_data, RTOS_data]
    print("\n 3. This ScrapeData func done")
    return data

def event():

    Date = date.today()
    CurrentDate = date.strftime(Date, "%d-%b-%Y")
    t = time.localtime()
    x = int(time.strftime("%H", t))

    if (x > 0) and (x <= 12):
        Event = 'Morning'
        return Event
    elif (x > 12) and (x < 18):
        Event = 'Afternoon'
        return Event
    else:
        Event = 'Evening'
        return Event

def generateMaildata():

    print(f"\n\n\n\n\n\n\n\n                                                                 		    .......! ! ! Do Not Shut PC Down ! ! !......\n")
    print(f"                                                                         		      	.......Running main.py......\n\n\n\n\n\n\n")
    
    Date = date.today()
    CurrentDate = date.strftime(Date, "%d-%b-%Y")
    formate_date = CurrentDate + "-" + event()
    print(f"Test the issue {formate_date}")
    # formate_date = "12-Dec-2022-Morning" # This is for Test only.
    print("\n 1. From here the programe begins")
    mail_data_IPC = []
    mail_data_RTOS = []
    scrape_table = ScrapeData()
    print("\n 4. This ScrapeData func comes to generatemaildata func")
    # print('\n4.1 This is dataframe formate>>>\n', scrape_table)
    for index, row in scrape_table[0].iterrows():
        if index == formate_date:
            mail_data_IPC.append(row)
    for index, row in scrape_table[1].iterrows():
        if index == formate_date:
            mail_data_RTOS.append(row)

    # if (len(mail_data_IPC) == 0) and (len(mail_data_RTOS) == 0):
    #     print('No New Entries Found!!')
    #     # exit()
    
    mail_data = [mail_data_IPC, mail_data_RTOS]
    # print('\n4.2 \n', mail_data)
    return mail_data        

def dataframe():
    Maildata = generateMaildata()
    df_IPC = pd.DataFrame(Maildata[0])
    df_RTOS = pd.DataFrame(Maildata[1])
    print("\n 5. This generateMaildata func has come to dataframe func")
    df = [df_IPC, df_RTOS]
    Outlook.send_table(df)
    return "DF created!!"

if __name__ == '__main__':

    # dataframe()
    # ScrapeData()
    # Outlook.PR_Mail()
    for i in ["11:00", "22:00"]:
        if i == '22:00':
            schedule.every().monday.at("09:00").do(Outlook.PR_Mail)
            schedule.every().tuesday.at("09:00").do(Outlook.PR_Mail)
            schedule.every().wednesday.at("09:00").do(Outlook.PR_Mail)
            schedule.every().thursday.at("09:00").do(Outlook.PR_Mail)
            schedule.every().friday.at("09:00").do(Outlook.PR_Mail)
            
        schedule.every().monday.at(i).do(dataframe)
        schedule.every().tuesday.at(i).do(dataframe)
        schedule.every().wednesday.at(i).do(dataframe)
        schedule.every().thursday.at(i).do(dataframe)
        schedule.every().friday.at(i).do(dataframe)

    
    # Outlook.PR_Mail()
    # schedule.every().monday.tuesday.wednesday.thursday.friday.at('09:00').do(Outlook.PR_Mail)
    # schedule.every().monday.tuesday.wednesday.thursday.friday.at('10:00').do(dataframe)
    # schedule.every().monday.tuesday.wednesday.thursday.friday.at('22:00').do(dataframe)
    # schedule.every().day.at("10:00").do(dataframe)
    # schedule.every().day.at("22:00").do(dataframe)
    
    try:
        # This is here to simulate application activity (which keeps the main thread alive).
        while True:
            schedule.run_pending()
            time.sleep(60)
        
    except (KeyboardInterrupt, SystemExit):
        print('\nTerminating the process due to KeyboardInterrupt\n\n')

    except (ValueError):
        print("\nValue Error found>>>\n")
