from csv import excel
import time
from selenium import webdriver 
from selenium.webdriver.common.by import By 
from selenium.webdriver.firefox.options import Options
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.support import expected_conditions as EC # for Ajax see stackoverflow bookmark
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import xlwings as xw
import os
from apscheduler.schedulers.blocking import BlockingScheduler
import schedule
import win32com.client

def lastRow(idx, workbook, col=1):
    """ Find the last row in the worksheet that contains data.

    idx: Specifies the worksheet to select. Starts counting from zero.

    workbook: Specifies the workbook

    col: The column in which to look for the last cell containing data.
    """

    ws = workbook.sheets[idx]

    lwr_r_cell = ws.cells.last_cell      # lower right cell
    lwr_row = lwr_r_cell.row             # row of the lower right cell
    lwr_cell = ws.range((lwr_row, col))  # change to your specified column

    if lwr_cell.value is None:
        lwr_cell = lwr_cell.end('up')    # go up untill you hit a non-empty cell

    return lwr_cell.row

def timed_job():
    # print('This job is run every 120 seconds.')


    # instantiate options 
    # options = webdriver.ChromeOptions() 
    
    # # run browser in headless mode 
    # options.headless = True 
    
    # instantiate driver 
    # driver = webdriver.Firefox(options=Options()) # I CHANGED THIS CODE RECENTLY CAUSE OF GECKODRIVER PATH ERROR
    driver = webdriver.Firefox(options=Options(),executable_path=GeckoDriverManager().install())
    # driver.maximize_window() 
    # load website 

    url = 'https://www.uwindsor.ca/registrar/uw_transfer_credits'

    # get the entire website content 
    driver.get(url)

    # Create object of the Select class
    select_country = Select(driver.find_element(By.XPATH, "//*[@id='edit-country-select']"))
                
    # # Select the option with value "Canada"
    select_country.select_by_value("CAN")
    delay = 40 # seconds
    try:
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'edit-school-select--2')))
        print("Page is ready!")
    except TimeoutException:
        print("Loading took too much time!")
    # Create object of the Select class
    select_school = Select(driver.find_element(By.ID,'edit-school-select--2'))
                
    # # Select the option with St.Clar school
    select_school.select_by_value("108145741")

    try:
        myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.ID, 'edit-subject--2')))
        print("Page is ready!")
    except TimeoutException:
        print("Loading took too much time!")
    # # Create object of the Select class
    select_subject = Select(driver.find_element(By.ID,'edit-subject--2'))
                
    # # # Select all subjects
    options = select_subject.options
    # for index in range(0, len(options) - 1):

    # writer = pd.ExcelWriter("C:/Users/geodx/Desktop/Scraped_Table.xlsm")

    for index_val in range(1, 4):
        select_subject.select_by_index(index_val)
        subject_name = select_subject.first_selected_option.get_attribute("textContent")

    # select_subject.select_by_index(1)

    #Find search button
        search_button = driver.find_element(By.ID,'edit-submit')
        search_button.click()

        try:
            myElem = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.XPATH, '/html/body/div/div[3]/div/div/div[1]/div[2]/div/div[1]/div/div/form/div/div[2]/div[2]/table/tbody')))
            print("Table is ready!")
        except TimeoutException:
            print("Table Loading took too much time!")

        time.sleep(15)
        #Changed CODE!!!!!!!!!!!!!!!!!!!
        webtable_df = pd.read_html(driver.page_source.replace('<br>', '£'))[0]
        print(webtable_df)
        # webtable_df = pd.read_html(driver.find_element(By.XPATH, "/html/body/div/div[3]/div/div/div[1]/div[2]/div/div[1]/div/div/form/div/div[2]/div[2]/table").get_attribute('outerHTML'))[0]

        
        # print(webtable_df)
        print(index_val)
        # webtable_df.to_csv('C:/Users/geodx/Desktop/file2.csv')
        if index_val == 1:
            # with pd.ExcelWriter("C:/Users/geodx/Desktop/Scraped_Table.xlsm") as writer:
            #     webtable_df.to_excel(writer, sheet_name=subject_name, index=False)
            # try:
            #     wb = xw.Book("C:/Users/geodx/Desktop/Scraped_Table.xlsm")
            # except OSError as e:
                        
            wb = xw.Book("C:/Users/geodx/Desktop/Scraped_Table.xlsm")
            sheet = wb.sheets['Sheet1']
            sheet.clear()
            sheet['A1'].value = subject_name
            sheet['A1'].font.bold = True
            sheet['A2'].options(index=False).value = webtable_df
            wb.save()
        else:
            # with pd.ExcelWriter("C:/Users/geodx/Desktop/Scraped_Table.xlsm",mode='a') as writer:
            #     webtable_df.to_excel(writer, sheet_name=subject_name, index=False)    
            # sheet = wb.sheets['Sheet1']
            # num_row = sheet.range('A1').end('down').row
            # num_row = sheet.range('A1').cells.last_cell.row
            num_row = lastRow(0,wb,1)
            print(num_row)
            sheet['A'+ str(num_row+2)].value = subject_name
            sheet['A'+ str(num_row+2)].font.bold = True
            sheet['A'+ str(num_row+3)].options(index=False).value = webtable_df
            # sheet['A'+ str(num_row+3)].options(index=False).value = '*'
            # sheet['B'+ str(num_row+3)].options(index=False).value = '*'
            wb.save()
        time.sleep(6)    

    # os.startfile('C:/Users/geodx/Desktop/Scraped_Table.xlsm')


    #--- for selecting all courses later on

    # Select select = new Select(driver.findElement(By.id("oldSelectMenu")));
            
    # Get all the options of the dropdown
    # List<WebElement> options = select.getOptions();
    # try:
    #     book = xw.Book('C:/Users/geodx/Desktop/Scraped_Table.xlsm')
    #     book.close()
    # except Exception as e:
    #     print(e)
    # os.system("taskkill /im excel.exe /fi \"WINDOWTITLE eq " + "C:/Users/geodx/Desktop/Scraped_Table.xlsm" + "\" /f")
   
    #BELOW IS ORIGINAL LINE TO KILL EXCEL FILES
    # os.system("taskkill /im excel.exe /F")
    driver.close()

# xl = win32com.client.Dispatch("Excel.Application")  #instantiate excel app

# wb = xl.Workbooks.Open(r'C:/Users/geodx/Desktop/Scraped_Table.xlsm')
# xl.Application.Run('Scraped_Table.xlsm!Module2.ReplaceCharacter()')
# wb.Save()
# xl.Application.Quit()
    macro = wb.macro('ReplaceCharacter')
    macro()

# sched = BlockingScheduler()
# @sched.scheduled_job('interval', seconds=120)

# sched.configure(options_from_ini_file)
# sched.start()

schedule.every(0.01).minutes.do(timed_job)

while True:
    schedule.run_pending()
    time.sleep(1)

print("Table fetched!")