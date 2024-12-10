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
 
 
download_dir = os.path.dirname(os.path.realpath(__file__))+'\\temp_downloads'
ConfigPath = os.path.dirname(os.path.realpath(__file__)) + '\\config.ini'
firefox_location='./'
 
 
global wait, driver, webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4, sq_5
global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection,min
 
 
def read_config_file():
    global webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4, sq_5
    global sql_server_name, sql_user_name, sql_password, sql_db
    try:
        # configuration entries
        config = ConfigParser()
        config.read(ConfigPath)
        webpage_url = config.get ("Data", "webpage_URL")
        user_name = config.get ("Data", "user_name")
        password = config.get ("Data", "password")
        sql_server_name = config.get ("SQL", "server")
        sql_user_name = config.get ("SQL", "user")
        sql_password = config.get ("SQL", "password")
        sql_db = config.get ("SQL", "database")
        return(0)
    except Exception as e:
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
        element.send_keys("tapintegrations")

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
    except Exception as e:
        print('SQL connection error')
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
 
def searchanddownload(fld_name,last_name, src_id, emp_name):
    global SQLconnection,min
    driver.get('https://login.paylocity.com/Escher/Escher_WebUI/EmployeeSearch/home/index?area=employees&view=EmployeeSearch')
    time.sleep(4)
    try:
        xpath='//*[@id="breadcrumbs"]/ul/li[2]/ul/li/button/i' # removes filter
        element = driver.find_element(By.XPATH,xpath)
        element.click()
        time.sleep(2)

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
                    click_xpath = f'//*[@role="rowgroup"][2]/tr[{i}]/td[2]/a'
                    element = driver.find_element(By.XPATH, click_xpath)
                    element.click()
                    time.sleep(5)
                    break
            else:
                # If no matching row is found, continue the process
                print("No matching ID found.")
        else:
            print("No rows found in the table.")
        time.sleep(1)
        xpath = '//*[text()="Pay"]'
        element = driver.find_element(By.XPATH,xpath)
        element.click()
        time.sleep(1)
        xpath = '//span[text()="Checks"]'
        element = driver.find_element(By.XPATH,xpath)
        element.click()
        time.sleep(1)
       
        #show private data
        try:
            xpath = '//*[@class="pcty-margin-top css-14bc3ye css-qvnb8s"]'
            xpath = "//button[@class='pcty-margin-top css-1ke8rzj css-qvnb8s']"
            element = driver.find_element(By.XPATH,xpath)
            element.click()
            time.sleep(5)       
        except:
            pass
        try:
            xpath = "//button[@id='pendo-close-guide-8fec1d7e']"
            element = driver.find_element(By.XPATH,xpath)
            element.click()
        except:
            pass
        
        yrs = ['2024','2023','2022','2021','2020']
        y=1
        for yr in yrs:
            try: 
                xpath = '//*[@data-automation-id="btn-toggle-filter"]'
                element = driver.find_element(By.XPATH,xpath)
                element.click()
                time.sleep(1)
            except Exception as k:
                print("exception k is:",k)
                pass
            try:
                try:
                    xpath = '//*[@id="startDateFilter-button"]'
                    # element = driver.find_element(By.CSS_SELECTOR,'#startDateFilter-button')
                    element = driver.find_element(By.XPATH,xpath)
                    element.click()
                    time.sleep(2)
                except:
                    pass
                try:
                    # Clicks on year
                    element = driver.find_element(By.XPATH, '//span[@class="rdrYearPicker css-1frza2b"]')
                    element.click()
                except:
                    pass
                try:
                    xpath = '//option[@value="'+str(yr)+'"]' # year
                    element_yr = driver.find_element(By.XPATH,xpath)
                    element_yr.click()
                except:
                    print(yr,"not found")
                    with open('yrs not found.csv','a',encoding='utf-8') as f:
                        f.write(f'{src_id},"{emp_name}",{yr}\n')
                    try:
                        xpath = '//*[@id="startDateFilter-button"]'
                        # element = driver.find_element(By.CSS_SELECTOR,'#startDateFilter-button')
                        element = driver.find_element(By.XPATH,xpath)
                        element.click()
                        time.sleep(2)
                    except:
                        pass
                    continue
                time.sleep(2)
                element = driver.find_element(By.XPATH,'//span[@class="rdrMonthPicker css-lytxuw"]')
                element.click()
                time.sleep(2)
                xpath = "//option[normalize-space()='January']"
                element = driver.find_element(By.XPATH,xpath)
                element.click()
                time.sleep(2)
                xpaths = [
                    "//button[@class='rdrDay css-7w1la9 rdrDayStartOfMonth']//span[contains(text(),'1')]",
                    "//button[@class='rdrDay css-7w1la9 rdrDayStartOfMonth']",
                    "//button[@class='rdrDay css-7w1la9 rdrDayWeekend rdrDayEndOfWeek rdrDayStartOfMonth']//span[contains(text(),'1')]",
                    "//button[@class='rdrDay css-7w1la9 rdrDayWeekend rdrDayStartOfWeek rdrDayStartOfMonth']//span[contains(text(),'1')]"
                ]

                # Iterate through the XPath expressions
                clicked = False
                for xpath in xpaths:
                    try:
                        # Try to find the element
                        element = driver.find_element(By.XPATH, xpath)
                        element.click()
                        time.sleep(2)
                        break  # Exit loop after successful click
                    except :
                        print(f"No element found for XPath: {xpath}")

                
            except Exception as g:
                print("Exception in start date:",g)
            try:
                element = driver.find_element(By.CSS_SELECTOR,'#endDateFilter-button')
                element.click()
                time.sleep(2)
                # Clicks on year
                # element = driver.find_element(By.XPATH, '//span[@class="rdrYearPicker css-1frza2b"]')
                element= driver.find_element(By.XPATH,"//span[@class='rdrYearPicker css-1frza2b']//select")
                element.click()
                try:
                    xpath = '//option[@value="'+str(yr)+'"]' # year
                    element_yr = driver.find_element(By.XPATH,xpath)
                    element_yr.click()
                except:
                    print(yr,"not found")
                time.sleep(2)
                element = driver.find_element(By.XPATH,'//span[@class="rdrMonthPicker css-lytxuw"]')
                element.click()
                time.sleep(2)
                xpath = "//option[normalize-space()='December']"
                element = driver.find_element(By.XPATH,xpath)
                element.click()
                time.sleep(2)

                xpaths = [
                    "//button[@class='rdrDay css-7w1la9 rdrDayEndOfMonth']",
                    "//button[@class='rdrDay css-7w1la9 rdrDayWeekend rdrDayEndOfWeek rdrDayEndOfMonth']",
                    "//button[@class='rdrDay css-7w1la9 rdrDayWeekend rdrDayStartOfWeek rdrDayEndOfMonth']"
                    ]

                # Iterate through the XPaths and click the elements
                for xpath in xpaths:
                    try:
                        # Try to find the element
                        element = driver.find_element(By.XPATH, xpath)
                        element.click()
                        time.sleep(2)  # Add delay if needed
                        
                        print(f"Clicked element with XPath: {xpath}")
                    except :
                        print(f"No element found for XPath: {xpath}")
                

            except Exception as h:
                print("exception in end date:",h)
            xpath = '//*[text()="Apply Filter"]'
            element = driver.find_element(By.XPATH,xpath)
            element.click()
            time.sleep(6)
            
            xpath = '//*[@id="checks-sidebar"]/div[2]/div/div/div[2]/div/div/div[1]/p'
            element_names = driver.find_elements(By.XPATH,xpath)
            for element_name in element_names:
                element_name = element_name.text
                with open('File Names.csv','a',encoding = 'utf-8') as g:
                    g.write(f'{src_id},"{emp_name}",{yr},"{element_name}"\n')
            
            xpath = '//*[@id="checks-sidebar"]/div[2]/div/div'
            elements = driver.find_elements(By.XPATH, xpath)
            doc_count = len(elements)
            print(f"Document count: {doc_count}")
            with open('doc_count.csv','a',encoding = 'utf-8') as g:
                g.write(f'{src_id},"{emp_name}",{yr},{doc_count}\n')
            for i in range(doc_count):
                try:
                    # Re-fetch elements to avoid stale references
                    elements = driver.find_elements(By.XPATH, xpath)
                    parent = elements[i].find_element(By.XPATH, './div[2]')
                    cls = parent.get_attribute('class')
                    print(f"Processing element {i+1}/{doc_count}, Class: {cls}")

                    if i == 0:
                        print("Handling the first document directly.")
                        download_xpath = "//div[contains(text(),'Download Paystub')]"
                        download_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, download_xpath))
                        )
                        download_button.click()
                        time.sleep(10)  # Allow time for download
                        continue
                    
                    if cls in ['css-rlt06n', 'css-1snfe5f']:
                        parent.click()  # Perform click action
                        time.sleep(2)  # Optional: Adjust based on loading times
                        
                        download_xpath = "//div[contains(text(),'Download Paystub')]"
                        download_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, download_xpath))
                        )
                        download_button.click()
                        time.sleep(10)  # Allow time for download to process
                        # Check for the error message
                        error_xpath = "//*[text()='An error occurred while downloading the paystubs']"
                        try:
                            error_element = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.XPATH, error_xpath))
                            )
                            print("Error detected while downloading the paystub.")
                            
                            # Close the error dialog
                            close_button_xpath = '/html/body/div[4]/div[1]/div/div[2]/div[1]/div/div[3]/div[3]/button'
                            close_button = WebDriverWait(driver, 5).until(
                                EC.element_to_be_clickable((By.XPATH, close_button_xpath))
                            )
                            close_button.click()
                            txt = parent.text
                            with open('error file.csv','a',encoding = 'utf-8') as g:
                                g.write(f'{src_id},"{emp_name}",{yr},"{txt}"\n')
                            print("Error dialog closed.")
                        except:
                            print("No error occurred during this download.")
                except Exception as e:
                    print(f"Error processing element {i+1}: {e}")
                    
            
            createFolder("Z:/HRG/Documents/Paylocity/Paystubs/"+fld_name+"/"+yr)
            time.sleep(2)
            download_dir = r'C:/Users/RPATEAMADMIN/Downloads/'
            files = os.listdir(download_dir)   
            for file in files:
                old_path = os.path.join(download_dir, file)
                new_filename = file
                f = open('downloaded files.csv', "a")
                f.write(f'{{"paystub filenames": "{new_filename}", "{fld_name}"}}\n')
                f.close()
                dst="Z:/HRG/Documents/Paylocity/Paystubs/"+fld_name+'/'+yr+'/'
                new_path = os.path.join(dst, new_filename)
                shutil.move(old_path,new_path)  
                value = '"'+fld_name +'"'+", downloaded"
                print(value)

        # xpath="/html/body/div[4]/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/button[2]/div/div"
        # element=driver.find_element(By.XPATH,xpath)
        # element.click()
        # time.sleep(5)
    except Exception as e:
        # xpath="/html/body/div[4]/div[1]/div/div[2]/div[1]/div/div[2]/div[2]/div/div[2]/button[2]/div/div"
        # element=driver.find_element(By.XPATH,xpath)
        # element.click()
        # time.sleep(5)
        print("Employee documents Not downloaded\n")
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
 
    # import csv
    # rows = []
    # with open("a2.csv", 'r') as file:
    #     csvreader = csv.reader(file)
    #     header = next(csvreader)
    #     for row in csvreader:
    #         rows.append(row)
    # print(len(rows))
    rows = []
    select_statement = "SELECT [Employee_ID],[Last_Name],[First_Name],[IsDownloaded_Paystubs]"
    select_statement = select_statement + "FROM [HRG].[dbo].[Emp_Roster]"
    cursor = SQLconnection.cursor()
    cursor.execute(select_statement)
    rows = cursor.fetchall()
    cursor.close()
#266,267             #Term 448(Vm24)
    for i in range(0,len(rows)):
        each = rows[i]
        src_id = each[0].strip()
        print(src_id)
        #src_id = src_id.zfill(6)
        last_name = each[1].strip()
        first_name = each[2].strip()
        #emp_name = each[1].strip()
        emp_name = last_name+", "+first_name
        print(emp_name)
        path_1 = r'C:/Users/RPATEAMADMIN/Downloads/'
        files = os.listdir(path_1)
        # Iterate over the files and delete them
        for file in files:
            file_path = os.path.join(path_1, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted {file_path}")
                # path = "Z:/HRG/Documents/Paylocity/Paystubs/"
        
        # directory_contents = os.listdir(path)
        fld_name = emp_name+" ("+src_id+")"
        # remove folder if already exits in destination
        
        var=each[3]
        # if fld_name not in directory_contents:
        if not var:
            destination_folder = "Z:/HRG/Documents/Paylocity/Paystubs"+'/'+fld_name
            print(destination_folder)
            if os.path.exists(destination_folder):
                # Remove the entire folder and its contents
                shutil.rmtree(destination_folder)
                print('removed', destination_folder)
            print("entered_searchanddownload")
            # err_f = 1
            # count_flag = 1
            # while(err_f):
            res = searchanddownload(fld_name, last_name,src_id, emp_name)
            print("res=",res)
            if res == 0:
                cursor = SQLconnection.cursor()
                cursor = SQLconnection.cursor()
                cursor.execute("UPDATE [HRG].[dbo].[Emp_Roster] SET [IsDownloaded_Paystubs] = ? WHERE [Employee_ID] = ?;",
                (1, src_id))
                print(src_id," :Updated in Emp_List")
                SQLconnection.commit()


            elif res<0:
                print('Employee error '+src_id+" "+emp_name)
                continue

 
#--------------------------------------------------------------------------------------------------------
 
main()
