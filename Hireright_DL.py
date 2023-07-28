from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import shutil
import os
import time
import fnmatch
import pickle

# URL of the website
url = 'https://ows01.hireright.com/login/'

# Define the webdriver options
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
"download.default_directory": os.getcwd(),  # Define default directory
"download.prompt_for_download": False,  # To auto download the file
"download.directory_upgrade": True,
"plugins.always_open_pdf_externally": True  # To download PDF files
})

driver = webdriver.Chrome(options=options)

# Go to website
driver.get(url)

# load the cookies file
#cookies = pickle.load(open("cookies_HR.pkl", "rb"))
#for cookie in cookies:
#    driver.add_cookie(cookie)
#driver.refresh()

# Load the excel data
df = pd.read_excel('HireRight_List.xlsx')

input("HireRight web page should open in a new window. Please login within the browser and press enter after to continue...")

for index, row in df.iterrows():
    # Check if the BG has been downloaded
    if pd.isna(row['BG Downloaded']) or row['BG Downloaded'] != 'x':
        if pd.isna(row['Middle Name']):
            # If 'Middle Name' is NaN, only use 'First Name' and 'Last Name'
            name = f"{row['Last Name']}, {row['First Name']}"
            print(name)
            profile_request = f"{row['Request #']}"
            print(profile_request)
            df.at[index, 'BG Downloaded'] = 'x'
        else:
            # If 'Middle Name' is not NaN, use 'First Name', 'Middle Name', and 'Last Name'
            name = f"{row['Last Name']}, {row['First Name']} {row['Middle Name']}"
            print(name)
            profile_request = f"{row['Request #']}"
            print(profile_request)
            df.at[index, 'BG Downloaded'] = 'x'
    
        # Wait for search box to be selectable
        search_box = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="_jsx_0_q"]')))
        search_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="_jsx_0_r"]/div/img')))
                                                     
        # Enter the profile into search box
        search_box.send_keys(profile_request)
        search_icon.click()

        # Make profile line active to activate DL button
        # select_active_profile = WebDriverWait(driver, 10).until(EC.element_to_be_selected((By.XPATH, '//*[@id="_jsx_0_cvjsx_0"]/tbody')))
        time.sleep(5)
        select_active_profile = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[4]/div[2]/div/div[2]/div/table/tbody/tr/td[6]/div')
        driver.execute_script("arguments[0].click();", select_active_profile)

        # Get last 4 of ssn
        ssn_element = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div/div/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[4]/div[2]/div/div[2]/div/table/tbody/tr/td[6]/div')
        ssn_save = ssn_element.text
        last_four = ssn_save[-5:].strip()
        print(last_four)

        # Select the download icon
        download_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div/div/div/div[2]/div[2]/div/div[2]/div/div[2]/div[2]/div/div[1]/div[1]/div/div/div/div[2]/span[2]/span[1]')))
        download_icon.click()
        

        # Click continue to download pdf
        continue_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div/span/div/div[2]/div/div[1]/div[2]/span[1]')))
        time.sleep(1)
        continue_button.click()

        # Give the pdf time to download
        time.sleep(3)

        # X out of window
        close_icon = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div/span/div/div[1]/div[2]/span/div/img')))
        close_icon.click()

        close_profile = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/div/div/div/div/div[2]/div[2]/div/div[2]/div/div[1]/span[2]/span[2]/span[2]/span')))
        close_profile.click()

        # Move it to the correct folder
        cwd = os.getcwd()
        src = os.path.join(cwd, 'report.pdf')

        directories = next(os.walk(cwd))[1]
        dst_dir = None
        for directory in directories:
            if fnmatch.fnmatch(directory, name + '*'):
                dst_dir = os.path.join(cwd, directory)
                new_dst_dir = dst_dir + ' ' + last_four
                os.rename(dst_dir, new_dst_dir)
                dst_dir = new_dst_dir
                break
        if not dst_dir:
            dst_dir = os.path.join(cwd, name + ' NOT_FOUND ' + last_four)
            os.makedirs(dst_dir, exist_ok=True)

        # Rename the file after moving
        dst_file = shutil.move(src, dst_dir)
        new_name = 'HireRight.pdf'
        new_file_path = os.path.join(dst_dir, new_name)
        os.rename(dst_file, new_file_path)

# Save the excel file
df.to_excel('HireRight_List.xlsx', index=False)

# Quit Chrome
driver.quit()