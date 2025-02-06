import logging
import configparser
import os
import sys
import re
import csv
import ctypes
import pyodbc
import pyautogui
import time
from time import sleep
import pygetwindow
import shutil
import keyboard
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from configparser import ConfigParser
from os import listdir
from os.path import isfile, join
from pathlib import Path


#----------------------------------------#

global download_dir
download_dir = r'C:\Users\HEMANTH.KUMAR.POLISE\Downloads'
ConfigPath ='./config.ini'
firefox_location='./'
global wait, driver, webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4, sq_5
global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection

#-----------------------------------------#



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
    
#-----------------------------------------#
    

def setup_firefox_driver():

    global firefox_location, download_dir, webpage_url, driver, wait
    firefox_options = Options()
    firefox_options.set_preference("browser.download.dir", download_dir)
    firefox_options.set_preference("browser.download.folderList", 2)
    firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")
    firefox_options.set_preference("browser.helperApps.neverAsk.saveToDisk", "*/*")

    firefox_options.set_preference("pdfjs.disabled", True)
    firefox_options.headless = False
    driver = webdriver.Firefox(options=firefox_options)
    driver.maximize_window()
    driver.get(webpage_url)
    return driver

#-----------------------------------------#


def remove_folder(folder_path):
    try:
        # Use shutil.rmtree to remove the folder and its contents
        shutil.rmtree(folder_path)
        print(f"Successfully removed the folder: {folder_path}")
    except Exception as e:
        print(f"Error removing the folder: {e}")
        
def remove_files():
    files_to_delete = os.listdir(download_dir)
    
    for filename in files_to_delete:
        if filename == 'Paylocity_files':
            remove_folder(os.path.join(download_dir,filename))
            continue
            
        file_path = os.path.join(download_dir, filename)
        os.remove(file_path)

#-----------------------------------------#

                    
def login():
    try:
        xpath = '//*[@id="CompanyId"]' #company code
        time.sleep(5)
        element = driver.find_element(By.XPATH,xpath)
        element.send_keys(client_code)

        xpath = '//*[@id="Username"]' #user id
        element = driver.find_element(By.XPATH,xpath)
        element.send_keys(user_name)

        xpath  = '//*[@id="Password"]' #password
        element = driver.find_element(By.XPATH,xpath)
        element.send_keys(password)

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
    
#-----------------------------------------#
    
    

def rename_filenamee(file_name):
    #download_dir=r"P:\MP\Downloads"
    print(file_name+"----------------")
    lis = listdir(download_dir)
    print(lis)
    if not file_name.endswith(".pdf"):
        if file_name.endswith(".docx"):
            new_file_name = file_name
        elif file_name.endswith(".jpeg"):
            new_file_name = file_name+".jpeg" 
        else:  
            new_file_name = file_name + ".pdf"
    else:
        new_file_name = file_name
    '''if new_file_name in lis:
        a = new_file_name.split('.')
        a[0] = a[0]+'1'
        new_file_name = '.'.join(a)'''

    os.rename(download_dir+"\\"+lis[0],download_dir+"\\"+new_file_name)
    

#-----------------------------------------#

def addToDatabase(paylocity_id,lastname,firstname,category,destination_file):
    try:
            cursor = SQLconnection.cursor()
            insert_statement = """
                INSERT INTO [Tobin Lucks].[dbo].[Emp_Files] ([Emp_Id],[Last_Name],[First_Name],[File_Name],[Category],[IsDownloaded],[DownloadedDateTime]
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """
            values = (paylocity_id,lastname,firstname,destination_file,category, 1, time.strftime('%Y-%m-%d %H:%M:%S'))
            cursor.execute(insert_statement, values)
            SQLconnection.commit()
            print(paylocity_id," :Added to Emp_Files")

    except Exception as d:
        print('Database error which is ', d)
        
#-------------------------------------------

def createFolder(sFolder):
    isExist = os.path.exists(sFolder)
    if not isExist:
        os.makedirs(sFolder)
    return(0)

#-------------------------------------------------------

def rename_filename(download_dir,company_id, new_file_name, dest_dir,emp_id,employee_name):
    lis = os.listdir(download_dir)
    file_name = lis[0]
    base_name, extension = os.path.splitext(file_name)
    #print(extension)
    new_file_name = re.sub(r'[^a-zA-Z0-9\s]', '_', new_file_name)
    # if '/' in new_file_name:
    #         new_file_name = new_file_name.replace('/', '_')
    new_file_name = new_file_name + extension
    #print(new_file_name)
    destination_path = os.path.join(dest_dir, new_file_name)
    #print("destination path:",destination_path)
    
    
    # Check if the file already exists in the destination directory
    if os.path.exists(destination_path):
        # If it exists, add numbers to the filename until a unique name is found
        # count = 1
        # while os.path.exists(destination_path):
        #     new_file_name = f"{new_file_name.split('.')[0]}_{count}{extension}"
        #     destination_path = os.path.join(dest_dir, new_file_name)
        #     count += 1
        count = 1
        while os.path.exists(destination_path):
        
            file_name = new_file_name.split('.')[0]
            new_file_name = f"{file_name.split('_')[0]}_{count}{extension}"
            destination_path = os.path.join(dest_dir, new_file_name)
            count += 1
    
    # Rename the file
    # print("renamed")
    os.rename(os.path.join(download_dir, lis[0]), os.path.join(download_dir, new_file_name))
    print("renamed and moving the file")
    # Move the file to the destination directory
    shutil.move(os.path.join(download_dir, new_file_name), dest_dir)
    h = open('downloaded files.csv', 'a')
    h.write(f'{company_id},"{employee_name}",{emp_id},"{new_file_name}"\n')
    h.close()
    return 'downloaded'
    
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

def remove_folder(folder_path):
    try:
        # Use shutil.rmtree to remove the folder and its contents
        shutil.rmtree(folder_path)
        print(f"Successfully removed the folder: {folder_path}")
    except Exception as e:
        print(f"Error removing the folder: {e}")

#--------------------------------------------------------------


def searchandDownload(paylocity_id,folder_name,paylocity_lastname,paylocity_firstname,emp_name,company_id):
    global err_f
    try:
        # remove_files()
        # destination_folder = os.path.join('Z:/FAG/Documents', str(company_id), folder_name)
        # if os.path.exists(destination_folder):
        #     # Remove the entire folder and its contents
        #     shutil.rmtree(destination_folder)
        time.sleep(5)
        # driver.get("https://login.paylocity.com/Escher/Escher_WebUI/EmployeeSearch/home/index?area=employees&view=EmployeeSearch")
        driver.get('https://login.paylocity.com/Escher/Escher_WebUI/EmployeeSearch/home/index?uniquecode=csEmployeeSearch&area=multico&view=EmployeeSearch&__viewedCompanyId=CS132001')

        print(paylocity_lastname)
        time.sleep(5)
        # xpath = '//*[@id="Filters_QuickSearchFilter"]' #employee id input bar
        # element = driver.find_element(By.XPATH,xpath)
        # element.send_keys(paylocity_id)

        # xpath = '//*[@id="breadcrumbs"]/ul/li[2]/ul/li/button'
        # element = driver.find_element(By.XPATH,xpath)
        # element.click()
        xpath='//*[@id="breadcrumbs"]/ul/li[2]/ul/li/button/i' 
        element = driver.find_element(By.XPATH,xpath)
        element.click()

        button = driver.find_element(By.LINK_TEXT,"Advanced") # Clicks advanced
        driver.execute_script("arguments[0].click();", button)
        time.sleep(2)
                

        lastname_input = '//*[@id="Filters_LastNameFilter"]' #last name input
        element = driver.find_element(By.XPATH,lastname_input)
        element.clear()
        element.send_keys(paylocity_lastname)
        time.sleep(2)

        emp_id_input  = '//*[@id="Filters_EmployeeIdFilter"]' #id input
        element = driver.find_element(By.XPATH, emp_id_input)
        element.clear()
        element.send_keys(paylocity_id)
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
                # Extract co_id and id for each row
                co_id_xpath = f'//*[@role="rowgroup"][2]/tr[{i}]/td[1]' #extracts company 
                element = driver.find_element(By.XPATH, co_id_xpath)
                co_id = element.text
                print(co_id)
                time.sleep(1)
                
                id_xpath = f'//*[@role="rowgroup"][2]/tr[{i}]/td[4]' #employee id
                element = driver.find_element(By.XPATH, id_xpath)
                id = element.text
                time.sleep(1)

                # If both id and company_id match, click the row and break the loop
                if str(id) == str(paylocity_id) and str(co_id) == str(company_id):
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
        try:
            # element = driver.find_element(By.LINK_TEXT,"Documents")
            element = driver.find_element(By.XPATH, "//span[text()='Documents']") # documents
            driver.execute_script("arguments[0].click();", element)
            time.sleep(5)
        except:
            pass
        try:
            xpath = '//div[text()="Sorry, we couldn\'t find any results. Try modifying your criteria."]'
            element=driver.find_element(By.XPATH,xpath)
            txt =element.text
            if txt == "Sorry, we couldn\'t find any results. Try modifying your criteria.":
                print(txt)
                with open(' no documents.csv','a+') as f:
                    f.write(f'{company_id},"{folder_name}","{txt}"\n')
            # xpath = '//div[contains(text(), "Sorry, we couldn\'t find any results")]'
        except:
            try:
                xpath = '//div[@class="css-0" and @data-automation-id="grid-documents-pager-lower"]' #taking count
                element=driver.find_element(By.XPATH,xpath)
                page_count =element.text
                count = page_count.split('of')[-1].strip()
                count = int(count.split('items')[0].strip())
                print(page_count, count)
            except: 
                # try:
                #     for i in range(1,5):
                #         xpath = f'//div[@class="css-0" and @data-automation-id="grid-documents-pager-lower"]/ul/li[{i}]'
                #         element=driver.find_element(By.XPATH,xpath)
                #         page_count =element.text
                #         count = page_count.split('of')[-1].strip()
                #         count = count.split('items')[0].strip()
                #         print(page_count, count)
                # except:
                    pass
            
            with open("employee_documentcount.txt", "a+") as info_file:
                info_file.write(f'{company_id},"{folder_name}",{count}\n')
            
            # destination_folder = os.path.join('Z:/FAG/Documents',folder_name)
            dest = os.path.join('Z:/FAG/Documents/Documents', str(company_id), folder_name)
            path = "C:/Users/RPATEAMADMIN/Downloads"
            c = count/20
            import math
            value = math.ceil(c)
            print("ceil value",value)
            # page_count = 1
            no_of_pages = 0
            for k in range(math.ceil(c)) :
                xpath = '//tbody[@class="css-0"]/tr' #file count per page
                elements  =driver.find_elements(By.XPATH,xpath)
                createFolder('Z:/FAG/Documents/Documents/'+str(company_id)+'/'+folder_name)
                for file in elements:
                    f_name = file.text
                    with open ("document Names.txt", "a+") as info_file:
                        info_file.write(f'{company_id},"{folder_name}",{f_name}\n')
                for p in range(1,len(elements)+1):
                    original_window = driver.current_window_handle
                    xpath = f'//tbody[@class="css-0"]/tr[{p}]/td/a' # clicks on each file
                    element = driver.find_element(By.XPATH,xpath)
                    file_name = element.text
                    print(file_name)
                    element.click()
                    time.sleep(20)
                    # keyboard.press_and_release('ctrl+s')
                    # time.sleep(5)
                    # pyautogui.press('enter')
                    # time.sleep(5)
                    # keyboard.write(file_name)
                    # time.sleep(2)
                    # keyboard.press_and_release('enter')
                    # # time.sleep(4)
                    # # pyautogui.hotkey('alt', 'f4')
                    # time.sleep(5)
                    # driver.close()
                    window_handles = driver.window_handles
                    print(len(driver.window_handles))
                    
                    if len(window_handles) > 1:
                        # driver.switch_to.window(window_handles[-1])
                        driver.switch_to.window(driver.window_handles[-1])
                        keyboard.press_and_release('ctrl+s')
                        time.sleep(2)
                        keyboard.press_and_release('enter')
                        time.sleep(5)
                        # driver.switch_to.window(driver.window_handles[-1])
                        driver.close()
                        time.sleep(2)
                        driver.switch_to.window(driver.window_handles[0])
                    value = rename_filename(path,company_id, file_name, dest,paylocity_id,emp_name)
                        

                try:
                    xpath = '//button[@title="Go to the next page" and @data-automation-id="grid-documents-pager-lower-next"]'
                    element = driver.find_element(By.XPATH,xpath)
                    element.click()
                except:
                    pass
      

    except Exception  as f:
        print('employee name error',f)
        f= open('employee name error.csv',"a")
        f.write(f'{company_id},"{emp_name}",{paylocity_id}')
        f.write('\n')
        f.close()
        return -1
    return 0
  


#------------------------------------------#
def main():
    global SQLconnection
    global err_f
    read_config_file()
    connectSQL()
    setup_firefox_driver()
    login()
 
    select_statement = "SELECT [First_Name],[Last_Name],[Company_Cone],[Employee_ID],[IsDownloaded],[DownloadedDateTime]  "
    select_statement = select_statement + " FROM [FAG].[dbo].[Emp_List];"
    cursor = SQLconnection.cursor()
    cursor.execute(select_statement)
    employee_details = cursor.fetchall()
    cursor.close()
    print("count is ")
    print(len(employee_details))
    
    
    for i in range(0,1100):
        each = employee_details[i]
        company_id=each[2]
        paylocity_id = each[3]
        paylocity_firstname = each[0].strip()
        paylocity_lastname = each[1].strip()
        # status = each[3]
        emp_name =paylocity_lastname+" "+paylocity_firstname
        is_downloaded = each[4]
        flg = 0
        flg = 0
        #path = './'
        path = 'C:/Users/RPATEAMADMIN/Downloads'
        directory_contents = os.listdir(path)
        fld_name=emp_name+" ("+str(paylocity_id)+")"
        c = len(os.listdir(path))
        if c>0:
            print("folder count in downloads:",c)
            break
        # if fld_name not in directory_contents:
        if not is_downloaded:
            print(fld_name)
            err_f = 1
            destination_folder = 'Z:/FAG/Documents/Documents/'+str(company_id)+'/'+fld_name
            print(destination_folder)
            count_flag = 1
            while(err_f):
                if not os.path.exists(destination_folder):
                    print(f"The folder {destination_folder} does not exist.")
                
                else:
                    # Loop through all files and subfolders in the specified folder
                    for root, dirs, files in os.walk(destination_folder):
                        for file in files:
                            file_path = os.path.join(root, file)
                            try:
                                os.remove(file_path)
                                print(f"Deleted file: {file_path}")
                            except Exception as e:
                                print(f"Error deleting file {file_path}: {e}")
                res = -1
                res = searchandDownload(paylocity_id,fld_name,paylocity_lastname,paylocity_firstname,emp_name,company_id)
                print("res=",res)
                if res == 0:
                    cursor = SQLconnection.cursor()
                    cursor.execute("UPDATE [FAG].[dbo].[Emp_List] SET [IsDownloaded] = ?, [DownloadedDateTime] = ? WHERE [Employee_ID] = ? AND [Company_Cone] = ?",
                    (1, time.strftime('%Y-%m-%d %H:%M:%S'), paylocity_id,company_id))
                    SQLconnection.commit()
                    print("updated in Emp_List")
                    err_f = 0
                elif res<0:
                    count_flag +=1
                    if count_flag>5:
                        err_f = 0
                    continue


main()
