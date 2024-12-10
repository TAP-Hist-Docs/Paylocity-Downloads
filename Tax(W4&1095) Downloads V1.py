from selenium import webdriver #Web driver activities
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC #Error handling
from selenium.webdriver.support.ui import WebDriverWait #Web driver wait
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import Select
from configparser import ConfigParser #Configuration read
from pathlib import Path #Path conversions
from os import listdir
from os.path import isfile, join
import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
# import undetected_chromedriver as uc
import os #Path
import sys #System paths
import re
import csv
import ctypes #message box
import pyodbc
import logging #Activity logging
import time #Sleep
import pyautogui
import time
import pygetwindow
import shutil
import keyboard   #from webdriver_manager.chrome import ChromeDriverManager
 
server = "prod-di-db"
user = "tapadmin"
password = "Welcome@2021"
database = "HRG"
 
download_dir = os.path.dirname(os.path.realpath(__file__))+'\\temp_downloads'
ConfigPath = 'config.ini'
firefox_location='./'
 
 
global wait, driver, webpage_url, client_code, user_name, sq_1, sq_2, sq_3, sq_4, sq_5
global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection,min
 
 
def read_config_file():
    global webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4, sq_5
    global sql_server_name, sql_user_name, sql_password, sql_db
    try:
        config = ConfigParser()
        config.read(ConfigPath)
        webpage_url = config.get ("Data", "webpage_URL") 
        client_code = config.get ("Data", "company_id") 
        user_name = config.get ("Data", "user_name") 
        password = config.get ("Data", "password") 
        sql_server_name = config.get ("SQL", "server")
        sql_user_name = config.get ("SQL", "user")
        sql_password = config.get ("SQL", "password") 
        sql_db = config.get ("SQL", "database") 
        return(0)
    except Exception as ex:
        print('config file read error')
        return(-1)
 
#--------------------------------------------------------------------------------------------------------
 
 
def setup():
    global firefox_location, download_dir, webpage_url, driver, wait
    binary = FirefoxBinary(firefox_location)
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", download_dir)
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx, application/zip")
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "image/jpeg")
    profile.set_preference("browser.helperApps.saveToDisk.image/jpeg", "application/octet-stream")
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", download_dir)
    profile.set_preference("browser.download.useDownloadDir", True)
    profile.set_preference("browser.download.viewableInternally.enabledTypes", "")
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx, application/x-pdf, application/vnd.pdf, text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*")
    profile.set_preference("pdfjs.disabled", True)
    profile.set_preference("browser.download.manager.useWindow", False)
    profile.set_preference("browser.download.manager.closeWhenDone", True)
    profile.set_preference("print_printer", "Microsoft Print to PDF")
    profile.set_preference("print.always_print_silent", True)
    profile.set_preference("print.show_print_progress", False)
    profile.set_preference("print.save_as_pdf.links.enabled", True)
    profile.set_preference("browser.helperApps.alwaysAsk.force", False)
    profile.set_preference("plugin.disable_full_page_plugin_for_types", "text/css, application/vnd.openxmlformats-officedocument.wordprocessingml.document, application/octet-stream, application/word, application/wordpad, image/png, image/bmp,, image/jpeg, application/pdf, text/csv, text/html, text/plain, application/docx, application/x-pdf, application/vnd.pdf, text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*")
    # driver = webdriver.Firefox(service_log_path='NUL', firefox_profile=profile)
    driver = webdriver.Firefox()
    wait = WebDriverWait(driver, 10)
    driver.maximize_window()
    webpage_url = "https://access.paylocity.com/"
    driver.get(webpage_url)
    time.sleep(1)
    wait = WebDriverWait(driver, 20)
    return 0
 
#--------------------------------------------------------------------------------------------------------
 
# def addToDatabase(emp_id,emp_name,category,document_title,each_file):
#     global sql_db
#     try:
#         cursor = SQLconnection.cursor()
#         insert_statement = """INSERT INTO [Studio_C].[dbo].[Emp_Files1] (
#                     [Src_Id],[Employee_ Name],[Category],[File_Name],[Document_Title],[IsDownloaded],[DownloadedDateTime])
#                     VALUES (?, ?, ?, ?, ?, ?, ?);"""
#         cursor.execute(insert_statement,(emp_id, emp_name,category,each_file,document_title,1,time.strftime('%Y-%m-%d %H:%M:%S')))
#         SQLconnection.commit()
#         print(document_title)
#         cursor.execute("UPDATE [Studio_C].[dbo].[Emp_Files] SET [File_Name] = ?, [IsDownloaded] = ?, [DownloadedDateTime] = ? WHERE [Src_Id] = ? AND [Document_Title] = ? ;",
#         (each_file,1, time.strftime('%Y-%m-%d %H:%M:%S'), emp_id,document_title))
#         SQLconnection.commit()
#     except:
#         print('insert to database error')

def addToDatabase(src_id,emp_name,document_title,category):
    try:
            cursor = SQLconnection.cursor()
            insert_statement = """
                INSERT INTO [Orchestra].[dbo].[Emp_Files_emp_docs] ([Employee_Number],[Emp_Name],[Category],[File_Name],[IsDownloaded]
                ) VALUES (?, ?, ?, ?, ?)
            """
            values = (src_id,emp_name,category,document_title, 1)
            cursor.execute(insert_statement, values)
            SQLconnection.commit()
            print(src_id," :Added to Emp_Files")

    except Exception as d:
        print('Database error which is ', d)
 
#--------------------------------------------------------------------------------------------------------
 
def remove_files():
    files_to_delete = os.listdir(download_dir)
    for filename in files_to_delete:
        file_path = os.path.join(download_dir, filename)
        os.remove(file_path)  # Delete the file
 
#--------------------------------------------------------------------------------------------------------
 
def createFolder(sFolder):
    isExist = os.path.exists(sFolder)
    if not isExist:
        os.makedirs(sFolder)
    return(0)
 
#--------------------------------------------------------------------------------------------------------
 
def login():
    try:
        xpath = '//*[@id="CompanyId"]' #company code
        time.sleep(5)
        element = driver.find_element(By.XPATH,xpath)
        element.send_keys("115040")

        xpath = '//*[@id="Username"]' #user id
        element = driver.find_element(By.XPATH,xpath)
        element.send_keys("tapintegrationsP3")

        xpath  = '//*[@id="Password"]' #password
        element = driver.find_element(By.XPATH,xpath)
        element.send_keys("WelcomeTAP@2026")

        xpath = '/html/body/div[1]/div[1]/div[7]/div/form/div[8]/div/button' # sign in button
        element = driver.find_element(By.XPATH,xpath)
        element.click()

        # xpath = '//*[@id="Device.OtpType1"]' #email selector
        # element = driver.find_element(By.XPATH,xpath)
        # element.click()

        # xpath = '//*[@id="pcty-button-send"]' #send button for verification code.
        # element = driver.find_element(By.XPATH,xpath)
        # element.click()
        s=input("enter:")
        # time.sleep(40) 
        return 0
    except:
        return -1
    
#--------------------------------------------------------------------------------------------------------
 
def createFolder(sFolder):
     isExist = os.path.exists(sFolder)
     if not isExist:
        os.makedirs(sFolder)
        return(0)
 
import pyodbc

#-----------------------------------------#
    
def connectSQL(): 
    global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection
    try:
        SQLconnection = pyodbc.connect('Driver={SQL Server};'
                            'Server=' + sql_server_name + ';'
                            'Database=' + sql_db + ';'
                            'UID=' + sql_user_name + ';'
                            'PWD=' + sql_password + ';'
                            'Trusted_Connection=no;')
        return(0)
    except Exception as ex:
        print('SQL connection error ' ,ex)
        return(-1)

   
 

def rename_and_move_file(file_name, document_title, category, file_extension, target_dir, fld_name, download_dir, src_id, emp_name):
    try:
        # Define new folder path in the target directory based on category
        new_folder = os.path.join(target_dir, category)
        if not os.path.exists(new_folder):
            os.makedirs(new_folder)

        # Define the base new file path
        base_new_file_path = os.path.join(new_folder, document_title + file_extension)
        
        # Check if the file already exists and generate a new name if necessary
        new_file_path = base_new_file_path
        counter = 1
        while os.path.exists(new_file_path):
            new_file_path = os.path.join(new_folder, f"{document_title}_{counter}{file_extension}")
            counter += 1

        # Rename and move the file
        original_file = os.path.join(download_dir, file_name)
        shutil.move(original_file, new_file_path)
        print(f"Moved file '{file_name}' to '{new_file_path}'")

    # except Exception as e:
    #     print(f"Error: {e}")
        value = '"'+fld_name+'"'+","+'"'+document_title+'"'
        f= open('Berlin1_Downloaded_files.csv',"a")
        f.write(value)
        f.write('\n')
        f.close()
        addToDatabase(src_id,emp_name,document_title,category)
    except Exception as e:
        print(f"Error while renaming and moving the file: {e}")
 
def wait_for_download(download_dir, timeout=30):
    seconds = 0
    dl_wait = True
    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = any([filename.endswith('.part') for filename in os.listdir(download_dir)])
        seconds += 1
    return not dl_wait
   
#--------------------------------------------------------------------------------------------------------


import os

def rename_file_in_downloads(file_name, downloads_path):
    try:
        counter = 1
        file_count = os.listdir(downloads_path)
        if len(file_count) > 0:
            for file in os.listdir(downloads_path):
                
                base_name, extension = os.path.splitext(file)
                if base_name == file_name: 
                    print(base_name)
                    unique_file_name = file_name+'_'+str(counter)
                    print(unique_file_name)
                    counter += 1
                    return unique_file_name
                else:
                    unique_file_name = file_name
        else:
            unique_file_name = file_name
    except Exception as g:
        print("file rename error:",g)

    return unique_file_name


#--------------------------------------------------------------------------------------------------------


def rename_filename(download_dir, dest_dir,emp_id,employee_name,file_name):
    lis = os.listdir(download_dir)
    # file_name = lis[0]
    for file in lis:
        base_name, extension = os.path.splitext(file)
    
    new_file_name = file_name + extension
    destination_path = os.path.join(dest_dir, new_file_name)
    # if '/' in new_file_name:
    #         new_file_name = new_file_name.replace('/', '_')
    
    # Check if the file already exists in the destination directory
    if os.path.exists(destination_path):
        count = 1
        while os.path.exists(destination_path):
        
            file_name = new_file_name.split('.')[0]
            new_file_name = f"{file_name.split('_')[0]}_{count}{extension}"
            destination_path = os.path.join(dest_dir, new_file_name)
            count += 1
    
    os.rename(os.path.join(download_dir, lis[0]), os.path.join(download_dir, new_file_name))
    print("renamed and moving the file")
    # Move the file to the destination directory
    shutil.move(os.path.join(download_dir, new_file_name), dest_dir)
    h = open('downloaded files.csv', 'a')
    h.write(f'"{employee_name}",{emp_id},"{new_file_name}"\n')
    h.close()
    return 'downloaded'

#--------------------------------------------------------------------------------------------------------

 
def searchanddownload(fld_name,last_name, src_id, emp_name):
    global SQLconnection,min
    try:
        driver.get('https://login.paylocity.com/Escher/Escher_WebUI/EmployeeSearch/home/index?area=employees&view=EmployeeSearch')

        # print(last_name)
        time.sleep(5)
        try:
            xpath='//*[@id="breadcrumbs"]/ul/li[2]/ul/li/button/i' 
            element = driver.find_element(By.XPATH,xpath)
            element.click()
        except:
            pass

        button = driver.find_element(By.LINK_TEXT,"Advanced") # Clicks advanced
        driver.execute_script("arguments[0].click();", button)
        time.sleep(2)
                

        lastname_input = '//*[@id="Filters_LastNameFilter"]' #last name input
        element = driver.find_element(By.XPATH,lastname_input)
        element.clear()
        element.send_keys(last_name)
        time.sleep(2)

        emp_id_input  = '//*[@id="Filters_EmployeeIdFilter"]' #id input
        element = driver.find_element(By.XPATH, emp_id_input)
        element.clear()
        element.send_keys(src_id)
        time.sleep(2)
        
        # element.send_keys(Keys.ENTER)
        xpath = '//*[@id="quickSearch"]' #search button
        element = driver.find_element(By.XPATH,xpath)
        element.click()
        time.sleep(5)
        
        # Find the table rows
        rows_xpath = '/html/body/div[3]/div/form/div/div[3]/div/table/tbody/tr'
        rows_xpath = '//*[@role="rowgroup"][2]/tr' #number of rows after search
        rows = driver.find_elements(By.XPATH, rows_xpath)

        # If there are rows in the table
        if len(rows) > 0:
            # Loop through each row
            for i in range(1, len(rows) + 1):
                
                id_xpath = f'//*[@role="rowgroup"][2]/tr[{i}]/td[4]' #employee id
                element = driver.find_element(By.XPATH, id_xpath)
                id = element.text
                time.sleep(1)

                # If both id and company_id match, click the row and break the loop
                if str(id) == str(src_id):
                    createFolder("Z:/HRG/Documents/Paylocity/W2 & 1095C/"+fld_name)

                    click_xpath = f'//*[@role="rowgroup"][2]/tr[{i}]/td[2]/a'
                    element = driver.find_element(By.XPATH, click_xpath)
                    element.click()
                    time.sleep(5)
                    break
            else:
                # If no matching row is found, continue the process
                print("No matching ID and company ID found.")
        else:
            print("No rows found in the table.")
        time.sleep(1)

        #pay
        xpath="/html/body/div[4]/div[1]/div/div[2]/div[2]/div/div/div/div/div[2]/div[1]/div/div/div/div/div/ul/li[2]/button"
        xpath = "//span[normalize-space()='Pay']"
        element=driver.find_element(By.XPATH, xpath)
        driver.execute_script("arguments[0].click();", element)
        time.sleep(5)
       
        #tax form
        xpath='/html/body/div[4]/div[1]/div/div[2]/div[2]/div/div/div/div/div[3]/div/div/div/div/div/ul/li[3]/button'
        xpath = "//li[@data-automation-id='subtab-tax-forms']//button[@role='tab']"
        element=driver.find_element(By.XPATH, xpath)
        driver.execute_script("arguments[0].click();", element)
        time.sleep(5)
        download_path = r'C:/Users/RPATEAMADMIN/Downloads/'
        try:
            try:
                xpath = '//*[@id="ep-drawer-root"]/div[2]/div/div/div/div/div[4]/div/div/div[2]/div'
                xpath = "//body/div[@id='CoreHREmployeeProfile']/div[@class='ep-drawer']/div/div[@id='ep-drawer-root']/div[@class='css-1yb1fdg']/div[@class='css-0']/div[@class='pcty-row-flex css-0']/div[@class='pcty-col css-0']/div[@class='corehr-drawer-content']/div[@class='corehr-drawer-content-body ser-drawer-content-body']/div[@class='pcty-row-flex pcty-padding-horizontal ser-drawer-content-body-row css-0']/div[@class='pcty-col css-0']/div[@class='css-m47ybr-zero-state']/div[1]"
                txt_element=driver.find_element(By.XPATH,xpath).text
                print(txt_element)
                if txt_element=="No records found.":
                    f = open('No Record W2.csv', "a")
                    f.write(f'W2:"{txt_element}","{fld_name}"\n')
                    f.close()
                    print("element not found")
            except:
                xpath = '//*[@id="ep-drawer-root"]/div[2]/div/div/div/div/div[4]/div/div/div[2]/div[1]/div/table/tbody/tr/td[1]/a'
                link_paths=driver.find_elements(By.XPATH,xpath)
                print(len(link_paths))
                count = len(link_paths)
                f = open('doc count.csv', "a")
                f.write(f'W2_count:{count},"{fld_name}"\n')
                f.close()
                for link_path in link_paths:
                    file_yr = link_path.text
                    f = open('yrs in System W2.csv', "a")
                    f.write(f'W2:{file_yr},"{fld_name}"\n')
                    f.close()
                    time.sleep(2)
                    link_path.click()

                    try:
                        xpath = "//p[contains(text(),'Please authenticate before continuing. If you have')]"
                        element=driver.find_element(By.XPATH,xpath)
                        txt = element.text
                        if txt == 'Please authenticate before continuing. If you have previously been authenticated then you are past the 2 hour limit for authentication.':
                            xpath = "//div[@class='corehr-drawer-content-body ser-drawer-content-body']//button[1]//span[1]"
                            element = driver.find_element(By.XPATH,xpath)
                            element.click()
                            time.sleep(2)
                            xpath = "//div[contains(text(),'Send Pin')]"
                            element = driver.find_element(By.XPATH,xpath)
                            element.click()
                            time.sleep(2)
                            input('enter:')
                            # time.sleep(2)
                            # xpath = "//input[@id='sua-pin-input']"
                            # element = driver.find_element(By.XPATH,xpath)
                            # element.send_keys(otp)
                            # time.sleep(2)
                            # xpath = '//div[@class="css-dtw560" ]/div[text() = "Submit"]'
                            # driver.find_element(By.XPATH,xpath).click()
                            # time.sleep(2)
                    except:
                        pass
                    time.sleep(4)
                    driver.switch_to.window(driver.window_handles[-1])
                    xpath = "//input[@id='pdfProtect_pdfPwdOptOut']"
                    element=driver.find_element(By.XPATH,xpath)
                    element.click()
                    time.sleep(1)
                    xpath = "//a[@id='pdfProtect_viewPDF']"
                    element=driver.find_element(By.XPATH,xpath)
                    element.click()
                    time.sleep(10)
                    driver.close()
                    driver.switch_to.window(driver.window_handles[-1])  
                    download_dir = r'C:/Users/RPATEAMADMIN/Downloads/'
                    dst="Z:/HRG/Documents/Paylocity/W2 & 1095C/"+fld_name+'/'
                    files = os.listdir(download_dir)   
                    re = rename_filename(download_dir,dst,src_id,emp_name,file_yr+' - W2')


        except Exception as h:
            print("w2 error:",h)                      
                           
                        

        try:
            try:
                xpath = "//div[@class='pcty-row-flex pcty-padding-horizontal ser-drawer-content-body-row css-0']//div[@class='pcty-col css-0']//div//div[@class='css-m47ybr-zero-state']//div[@class='css-1a7bjfk-media-message'][normalize-space()='No records found.']"
                txt_element=driver.find_element(By.XPATH,xpath).text
                print(txt_element)
                if txt_element=="No records found.":
                    f = open('No Record 1095.csv', "a")
                    f.write(f'1095C:"{txt_element}","{fld_name}"\n')
                    f.close()
                    print("element not found")
            except:
                xpath = '//*[@id="ep-drawer-root"]/div[2]/div/div/div/div/div[4]/div/div/div[3]/div[2]/div[1]/div/table/tbody/tr/td[1]/a'
                link_paths = driver.find_elements(By.XPATH,xpath)
                print(len(link_paths))
                count = len(link_paths)
                f = open('doc count.csv', "a")
                f.write(f'"1095C_count":{count},"{fld_name}"\n')
                f.close()
                for link_path in link_paths:
                    file_yr = link_path.text
                    f = open('yrs in System 1095.csv', "a")
                    f.write(f'1095C:{file_yr},"{fld_name}"\n')
                    f.close()
                    time.sleep(2)
                    link_path.click()
                    try:
                        xpath = "//p[contains(text(),'Please authenticate before continuing. If you have')]"
                        element=driver.find_element(By.XPATH,xpath)
                        txt = element.text
                        if txt == 'Please authenticate before continuing. If you have previously been authenticated then you are past the 2 hour limit for authentication.':
                            xpath = "//div[@class='corehr-drawer-content-body ser-drawer-content-body']//button[1]//span[1]"
                            element = driver.find_element(By.XPATH,xpath)
                            element.click()
                            time.sleep(2)
                            xpath = "//div[contains(text(),'Send Pin')]"
                            element = driver.find_element(By.XPATH,xpath)
                            element.click()
                            time.sleep(2)
                            input('enter:')
                            # time.sleep(2)
                            # xpath = "//input[@id='sua-pin-input']"
                            # element = driver.find_element(By.XPATH,xpath)
                            # element.send_keys(otp)
                            # time.sleep(2)
                            # xpath = '//div[@class="css-dtw560" ]/div[text() = "Submit"]'
                            # driver.find_element(By.XPATH,xpath).click()
                            # time.sleep(2)
                    except:
                        pass
                    time.sleep(4)
                    driver.switch_to.window(driver.window_handles[-1])
                    time.sleep(2)
                    xpath = "//a[@id='ViewReportLink']"
                    element=driver.find_element(By.XPATH,xpath)
                    element.click()
                    time.sleep(15)
                    driver.switch_to.window(driver.window_handles[-1])
                    xpath = "//input[@id='pdfProtect_pdfPwdOptOut']"
                    element=driver.find_element(By.XPATH,xpath)
                    element.click()
                    time.sleep(1)
                    xpath = "//a[@id='pdfProtect_viewPDF']"
                    element=driver.find_element(By.XPATH,xpath)
                    element.click()
                    time.sleep(10)
                    driver.close()
                    driver.switch_to.window(driver.window_handles[-1])  
                    time.sleep(2)                        
                    driver.close()
                    driver.switch_to.window(driver.window_handles[-1]) 
                            
                    download_dir = r'C:/Users/RPATEAMADMIN/Downloads/'
                    dst="Z:/HRG/Documents/Paylocity/W2 & 1095C/"+fld_name+'/'
                    files = os.listdir(download_dir)   
                    re = rename_filename(download_dir,dst,src_id,emp_name,file_yr+' - 1095C')
                    print(re)
            
        except Exception as f:
            print('exception f is', f)
            pass
        
        
        # xpath="/html/body/div[4]/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/button[2]/div/div"
        # element=driver.find_element(By.XPATH,xpath)
        # element.click()
        # time.sleep(5)
    except Exception as e:
        # xpath="/html/body/div[4]/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/button[2]/div/div"
        # element=driver.find_element(By.XPATH,xpath)
        # element.click()
        # time.sleep(5)
        print("Employee documents Not downloaded\n",e)
        f= open('employee error.csv',"a")
        f.write(fld_name)
        f.write('\n')
        f.close()
        print("error ",e)
        return -1
    return 0
#--------------------------------------------------------------------------------------------------------
 
 
def main():
    global SQLconnection
    setup()
    read_config_file()
    connectSQL() 
    login()
 
    import csv
    # rows = []
    # with open("a.csv", 'r') as file:
    #     csvreader = csv.reader(file)
    #     header = next(csvreader)
    #     for row in csvreader:
    #         rows.append(row)
    # print(len(rows))
    rows = []
    select_statement = "SELECT [Employee_ID],[Last_Name],[First_Name],[IsDownloaded_W2_1095]"
    select_statement = select_statement + "FROM [HRG].[dbo].[Emp_Roster]"
    # select_statement += "WHERE [Company] = 'BerlinRosen';"
    cursor = SQLconnection.cursor()
    cursor.execute(select_statement)
    rows = cursor.fetchall()
    cursor.close()
    #266,267             #Term 448(Vm24)
    for i in range(0,190):
        each = rows[i]
        src_id = str(each[0]).strip()  
        # src_id = src_id_1.zfill(5)
        print("initial print:",src_id)
        #src_id = src_id.zfill(6)
        last_name = each[1].strip()
        first_name = each[2].strip()
        #emp_name = each[1].strip()
        emp_name = last_name+", "+first_name
        print(emp_name)
        isdownloaded = each[3]
        # path_1 = r'C:\Users\RPATEAMADMIN\Downloads'
        path_1 = r'C:/Users/RPATEAMADMIN/Downloads/'
        files = os.listdir(path_1)
        # Iterate over the files and delete them
        count = len(os.listdir(path_1))
        if count>0:
                print("document count in downloads:",count)
                break
                #os.remove(file_path)
                #print(f"Deleted {file_path}")
        
        path = "Z:/HRG/Documents/Paylocity/W2 & 1095C/"
        directory_contents = os.listdir(path)
        fld_name = emp_name+" ("+src_id+")"
        # if fld_name not in directory_contents:
        if not isdownloaded:
            print("entered_searchanddownload")
            destination_folder = 'Z:/HRG/Documents/Paylocity/W2 & 1095C/'+fld_name
            print(destination_folder)
            if os.path.exists(destination_folder):
                # Remove the entire folder and its contents
                shutil.rmtree(destination_folder)
                print('removed', destination_folder)
            # err_f = 1
            # count_flag = 1
            # while(err_f):
            res = searchanddownload(fld_name,last_name, src_id, emp_name)
            print("res=",res)
            if res == 0:
                cursor = SQLconnection.cursor()
                cursor.execute("UPDATE [HRG].[dbo].[Emp_Roster] SET [IsDownloaded_W2_1095] = ? WHERE [Employee_ID] = ?;",
                (1, src_id))
                print(src_id," :Updated in Emp_List")
                SQLconnection.commit()


            # elif res<0:
            #     print('Employee error '+src_id+" "+emp_name)
            #     continue

 
#--------------------------------------------------------------------------------------------------------
 
main()
