import configparser
import os
import sys
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pyodbc
from os import listdir
import keyboard
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from configparser import ConfigParser
import csv

import re

# Define the path
global download_dir
download_dir = r'C:\Users\RPATEAMADMIN\Downloads\\'
configpath = r'Z:\Six Robblees\Scripts\Six\config.ini'
download_dir = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'temp_downloads')


# Check if the directory exists; if not, create it
if not os.path.exists(download_dir):
    os.makedirs(download_dir)
    print(f"Directory '{download_dir}' created successfully.")
else:
    print(f"Directory '{download_dir}' already exists.")

# Now, you can proceed with using download_dir as normal

# Set the path for the config file

#download_dir = os.path.dirname(os.path.realpath(__file__))+'\\temp_downloads'
ConfigPath = os.path.dirname(os.path.realpath(__file__)) + '\\config.ini'

# Global variables
global webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4, sq_5
global sql_server_name, sql_user_name, sql_password, sql_db

# def read_config_file():
#     global webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4, sq_5
#     global sql_server_name, sql_user_name, sql_password, sql_db
#     try:
#         # Read the config file
#         config = configparser.ConfigParser()
#         if not os.path.isfile(ConfigPath):
#             print("Configuration file not found!")
#             return -1

#         config.read(ConfigPath)

#         # Print the content of the config file for debugging purposes
#         with open(ConfigPath, 'r') as file:
#             print("Config file content:")
#             print(file.read())
        
#         # Read values from the config file
#         webpage_url = config.get("Data", "webpage_URL")
#         client_code = config.get("Data", "company_id")
#         user_name = config.get("Data", "user_name")
#         password = config.get("Data", "password")
        
#         sql_server_name = config.get("SQL", "server")
#         sql_user_name = config.get("SQL", "user")
#         sql_password = config.get("SQL", "password")
#         sql_db = config.get("SQL", "database")

#         print("Config values read successfully")

#         return 0
#     except Exception as ex:
#         print(f"Config file read error: {ex}")
#         return -1

# def read_config_file():
#     global webpage_url, client_code, user_name, password, sq_1, sq_2, sq_3, sq_4,  sq_5
#     global sql_server_name, sql_user_name, sql_password, sql_db
#     try:
#         config = ConfigParser()
#         config.read(ConfigPath)
#         webpage_url = config.get ("Data", "webpage_URL") 
#         client_code = config.get ("Data", "company_id") 
#         user_name = config.get ("Data", "user_name") 
#         password = config.get ("Data", "password") 
#         sql_server_name = config.get ("SQL", "server")
#         sql_user_name = config.get ("SQL", "user")
#         sql_password = config.get ("SQL", "password") 
#         sql_db = config.get ("SQL", "database") 
#         return(0)
#     except Exception as ex:
#         print('config file read error')
#         return(-1)

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
    driver.get("https://access.paylocity.com/")
    return driver

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

def login():
    try:
        xpath = '//*[@id="CompanyId"]'
        time.sleep(5)
        element = driver.find_element(By.XPATH,xpath)
        element.send_keys("53945")

        xpath = '//*[@id="Username"]'
        element = driver.find_element(By.XPATH,xpath)
        element.send_keys("HCMUnlocked")

        xpath  = '//*[@id="Password"]'
        element = driver.find_element(By.XPATH,xpath)
        element.send_keys("Welcome@2023")

        xpath = '/html/body/div[1]/div[1]/div[7]/div/form/div[8]/div/button'
        element = driver.find_element(By.XPATH,xpath)
        element.click()

        # xpath = '//*[@id="Device.OtpType1"]' #email selector
        # element = driver.find_element(By.XPATH,xpath)
        # element.click()

        # xpath = '//*[@id="pcty-button-send"]' #send button for verification code.
        # element = driver.find_element(By.XPATH,xpath)
        # element.click()
        s=input("enter:")
        # time.sleep(20) 
        return 0
    except:
        return -1
    
# def connectSQL(): 
#     global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection
#     try:
#         SQLconnection = pyodbc.connect('Driver={SQL Server};'
#                             'Server=' + sql_server_name + ';'
#                             'Database=' + sql_db + ';'
#                             'UID=' + sql_user_name + ';'
#                             'PWD=' + sql_password + ';'
#                             'Trusted_Connection=no;')
#         return(0)
#     except Exception as ex:
#         print('SQL connection error ' ,ex)
#         return(-1)
    
# def connectSQL():
#     global sql_server_name, sql_user_name, sql_password, sql_db, SQLconnection
#     sql_server_name = 'prod-di-db'
#     sql_user_name = 'tapadmin'
#     sql_password = 'Welcome@2021'
#     sql_db = 'Six_Robblees'
#     try:
#         SQLconnection = pyodbc.connect('Driver={SQL Server};'
#                             'Server=' + sql_server_name + ';'
#                             'Database=' + sql_db + ';'
#                             'UID=' + sql_user_name + ';'
#                             'PWD=' + sql_password + ';'
#                             'Trusted_Connection=no;')
#         return(0)
#     except Exception as ex:
#         print('SQL connection error',ex)
#         return(-1)

def connectSQL():
    global SQLconnection
    sql_server_name = 'prod-di-db'
    sql_user_name = 'tapadmin'
    sql_password = 'Welcome@2021'
    sql_db = 'Six_Robblees'
    try:
        SQLconnection = pyodbc.connect(
            'Driver={SQL Server};'
            f'Server={sql_server_name};'
            f'Database={sql_db};'
            f'UID={sql_user_name};'
            f'PWD={sql_password};'
            'Trusted_Connection=no;'
        )
        return 0
    except Exception as ex:
        print('SQL connection error:', ex)
        SQLconnection = None
        return -1
    
    
def read_config_file():
    global config, sql_server_name, sql_database, sql_username, sql_password
    config = configparser.ConfigParser()
    config.read('config.ini')  # Ensure your `config.ini` file is in the current working directory.
    
    sql_server_name = config['SQL']['server']
    sql_database = config['SQL']['database']
    sql_username = config['SQL']['user']
    sql_password = config['SQL']['password']
    print("Config values read successfully")


def rename_filename(file_name):
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

def addToDatabase(paylocity_id,lastname,firstname,category,destination_file):
    try:
            cursor = SQLconnection.cursor()
            insert_statement = """
                INSERT INTO [SIX ROBBLEES].[dbo].[Employee list] ([Employee_Id],[Last_Name],[First_Name],[IsDownloaded_1],[DownloadedDateTime]
                ) VALUES (?, ?, ?, ?, ?)
            """
            values = (paylocity_id,lastname,firstname,destination_file,category, 1, time.strftime('%Y-%m-%d %H:%M:%S'))
            cursor.execute(insert_statement, values)
            SQLconnection.commit()
            print(paylocity_id," :Added to Employee list")

    except Exception as d:
        print('Database error which is ', d)

def createFolder(sFolder):
    isExist = os.path.exists(sFolder)
    if not isExist:
        os.makedirs(sFolder)
    return(0)

def searchandDownload(paylocity_id,folder_name,paylocity_lastname,paylocity_firstname,emp_name):
    try:
        # driver.get("https://login.paylocity.com/Escher/Escher_WebUI/views/employees/empCheckHistory.aspx?view=empDemographicsHR&area=employees")
        driver.get("https://login.paylocity.com/Escher/Escher_WebUI/views/employees/empCheckDetail.aspx?view=empDemographicsHR&area=employees")
        time.sleep(10)
        #ID Dropdown
        dropdown = driver.find_element("id", "employeeNavControl_ddSortBy")
        select = Select(dropdown)
        select.select_by_visible_text("[ID] Name (Dept)")
        #All dropdown
        dropdown_element = driver.find_element(By.ID, "employeeNavControl_ddFilter")
        dropdown = Select(dropdown_element)
        dropdown.select_by_value("1555")

        search = '//*[@id="employeeNavControl_txtSearchNav"]'
        element = driver.find_element(By.XPATH, search)
        element.click()
        element.clear()
        element.send_keys(paylocity_id)
        time.sleep(5)
       

        enter = '//*[@id="employeeNavControl_btnSearchNavGo"]'
        element = driver.find_element(By.XPATH, enter)
        element.click()
        time.sleep(2)
        print('search done')


        
        
        # xpath = '/html/body/form/div[5]/div[3]/div[1]/table/tbody/tr/td/div/table/tbody/tr/td/table/tbody/tr/td[10]/table/tbody/tr[1]/td/div/input[2]'
        # element = driver.find_element(By.XPATH, xpath)
        # time.sleep(5)
        # # print(element)

        # s = element.get_attribute("value")
        # match = re.search(r'\[(.*?)\]', s)  # Find text within square brackets

        # extracted_text = match.group(1)
        # print(extracted_text)
        # time.sleep(5)
        
        
        
        
        xpath = "//td[@class='fieldvalue']"
        element = driver.find_element(By.XPATH, xpath)

        # Extract the text content
        text = element.text.strip()  # Remove any extra spaces or non-breaking spaces
        print(text)

        if str(paylocity_id) == str(text):
            print(f"ID Match Found: {paylocity_id}")
            # Proceed with the next steps for this employee
            # (Add your next steps here as needed)
        else:
            print(f"ID Mismatch: Database ID {paylocity_id} != System ID {element}")

            # Log mismatch to CSV
            with open('mismatch.csv', mode="a", newline="") as file:
                writer = csv.writer(file)
                writer.writerow([paylocity_id,folder_name, "ID mismatch"])
            return -1       

        xpath = "//a[text()='Check History']"
        element = driver.find_element(By.XPATH, xpath)

        # Extract the text content (if needed)
        element.click()





        year_dropdown_xpath = '//*[@id="checkYearsList"]'
        years_xpath = '//*[@id="checkYearsList"]/option'

        # Specify the target years
        target_years = ["2024", "2023"]
        present_years = []  # List to store years that are found in the system
        missing_years = []  # List to store years that are not found in the system

        # Refresh the dropdown element and get the updated list of options
        year_dropdown = driver.find_element(By.XPATH, year_dropdown_xpath)
        available_years = driver.find_elements(By.XPATH, years_xpath)

        # Loop through the target years to check if they are in the dropdown
        for target_year in target_years:
            year_found = False  # Flag to check if the year is available

            # Check the available years
            for year_option in available_years:
                year_text = year_option.text.strip()
                if year_text == target_year:
                    year_found = True
                    present_years.append(target_year)  # Add the year to present list
                    print(f"Year {target_year} found.")
                    break

            if not year_found:
                missing_years.append(target_year)  # Add the year to missing list
                print(f"Year {target_year} not found.")

        # Write the present and missing years to a CSV file

        file_path = "year_report_before_downloads.csv"
        file_exists = os.path.exists(file_path)
        with open(file_path, mode='a', newline='') as file:
            writer = csv.writer(file)
            if not file_exists or os.stat(file_path).st_size == 0:
                writer.writerow(["Id","Year", "Status"])
            
            for year in present_years:
                writer.writerow([paylocity_id,folder_name,year, "Present"])
            
            for year in missing_years:
                writer.writerow([paylocity_id,folder_name, "Missing"])

        print("Year report has been saved to 'year_report.csv'.")



        # XPaths for year dropdown, submit button, and print summary report button
        year_dropdown_xpath = '//*[@id="checkYearsList"]'
        years_xpath = '//*[@id="checkYearsList"]/option'
        submit_button_xpath = '//*[@id="submitButton"]'
        print_summary_report_xpath = '//*[@id="printSummaryButton"]'

        # Specify the target years
        target_years = ["2024", "2023"]
        present_years = []  # List to store years that are found in the system
        missing_years = []


        # Loop through the target years
        for target_year in target_years:
            year_found = False  # Flag to check if the year is available
            
            # Refresh the dropdown element and get the updated list of options
            year_dropdown = driver.find_element(By.XPATH, year_dropdown_xpath)
            available_years = driver.find_elements(By.XPATH, years_xpath)
            
            for year_option in available_years:
                year_text = year_option.text.strip()
                if year_text == target_year:
                    year_found = True
                    print(f"Year {target_year} found. Processing...")
                    # Ensure the year is selected
                    if not year_option.is_selected():
                        year_option.click()
                        print(f"Clicked on year: {target_year}")
                    
                    # Click the submit button
                    submit_button = driver.find_element(By.XPATH, submit_button_xpath)
                    submit_button.click()
                    print("Submit button clicked")
                    
                    # Wait for results to load
                    time.sleep(3)
                    
                    # Click the print summary report button
                    print_summary_button = driver.find_element(By.XPATH, print_summary_report_xpath)
                    print_summary_button.click()
                    print("Print Summary Report button clicked")
                    present_years.append(target_year)
                    time.sleep(5)
                    createFolder("Z:\Six Robblees\Pay"+folder_name)
                    downloads_folder = r"Z:\Six Robblees\Scripts\Six\temp_downloads"
                    target_folder = "Z:\Six Robblees\Pay"+folder_name+"/"
                    new_name = f"Paystubs_{target_year}.pdf"
                    # Loop through the Downloads folder to find the file to rename
                    for file_name in os.listdir(downloads_folder):
                        if file_name.lower().endswith(".pdf"):
                            old_file_path = os.path.join(downloads_folder, file_name)
                            # Check if "Paystubs_2024.pdf" already exists
                            new_file_path = os.path.join(downloads_folder, new_name)
                            counter = 1
                            while os.path.exists(new_file_path):
                                # Add suffix if file exists
                                new_file_path = os.path.join(downloads_folder, f"Paystubs_{target_year}_{counter}.pdf")
                                counter += 1
                            # Rename the file
                            os.rename(old_file_path, new_file_path)
                            print(f"Renamed '{file_name}' to '{os.path.basename(new_file_path)}'")

                            if not os.path.exists(target_folder):
                                os.makedirs(target_folder)
                            
                            final_file_path = os.path.join(target_folder, os.path.basename(new_file_path))
                            shutil.move(new_file_path, final_file_path)
                            print(f"Moved '{os.path.basename(new_file_path)}' to '{target_folder}'")
                            time.sleep(5)

                            file_path = "downloaded_files.csv"
                            file_exists = os.path.exists(file_path)
                            with open(file_path, mode='a', newline='') as file:
                                writer = csv.writer(file)
                                if not file_exists or os.stat(file_path).st_size == 0:
                                    writer.writerow(["Id", "Employeename", "Filename"])
                                
                                # writer.writerow(["Id", "Employeename", "Filename"])
                                writer.writerow([paylocity_id, emp_name, new_name])
                    # Wait for the process to complete before moving to the next year
                    time.sleep(3)
                    break  # Break out of the inner loop to proceed with the next year
            
            if not year_found:
                print(f"Year {target_year} not found in the dropdown")
                missing_years.append(target_year)

        file_path = "year_report.csv"
        file_exists = os.path.exists(file_path)
        with open(file_path, mode='a', newline='') as file:
            writer = csv.writer(file)
            if not file_exists or os.stat(file_path).st_size == 0:
                writer.writerow(["Id", "Year", "Status"])
            
            for year in present_years:
                writer.writerow([paylocity_id, year, "Present"])
            
            for year in missing_years:
                writer.writerow([paylocity_id, year, "Missing"])

        print("Year report has been saved to 'year_report.csv'.")

        return 0
    except Exception as e :
        print('employee name error', e)
        f= open('Employees_with_no_files_Paystubs.csv',"a")
        f.write(emp_name+'('+paylocity_id+')')
        f.write('\n')
        f.close()
        return -1


def main():
    global SQLconnection
    global err_f
    #read_config_file()
    connectSQL()
    if connectSQL() == 0:
        print("Connection established.")
        try:
            cursor = SQLconnection.cursor()
            # Your database queries go here
            print("Cursor created successfully.")
        except Exception as ex:
            print("Error using SQL connection:", ex)
    else:
        print("Unable to connect to the database.")
    setup_firefox_driver()
    login()
 
    select_statement = "SELECT [Employee_Id],[Last_Name],[First_Name],[IsDownloaded_1] "
    select_statement = select_statement + " FROM [SIX ROBBLEES].[dbo].[Employee list];"
    cursor = SQLconnection.cursor()
    cursor.execute(select_statement)
    employee_details = cursor.fetchall()
    cursor.close()
    print("count is ")
    print(len(employee_details))

    for i in range(0, 20):
        each = employee_details[i]
        paylocity_id = each[0]
        paylocity_firstname = each[2].strip()
        paylocity_lastname = each[1].strip()
        # status = each[3]
        emp_name =paylocity_lastname+" "+paylocity_firstname
        is_downloaded = each[3]

        #path = './'
        path = 'Z:\Six Robblees\Pay'
        directory_contents = os.listdir(path)
        fld_name=emp_name+" ("+str(paylocity_id)+")"
    
        # if fld_name not in directory_contents:
        if not is_downloaded:
            print(fld_name)
            err_f = 1

            count_flag = 1
            while(err_f):
                res = -1
                res = searchandDownload(paylocity_id,fld_name,paylocity_lastname,paylocity_firstname,emp_name)
                print("res=",res)
                if res == 0:
                    cursor = SQLconnection.cursor()
                    cursor.execute("UPDATE [SIX ROBBLEES].[dbo].[Employee list] SET [IsDownloaded_1] = ?, [DownloadedDateTime] = ? WHERE [Employee_Id] = ? ",
                    (1, time.strftime('%Y-%m-%d %H:%M:%S'), paylocity_id))
                    SQLconnection.commit()
                    print("updated in Employee list")
                    err_f = 0
                elif res<0:
                    count_flag +=1
                    if count_flag>5:
                        err_f = 0
                    continue


main()


