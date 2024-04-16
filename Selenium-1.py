from selenium import webdriver
import shutil
import os
import time
import openpyxl
import pandas as pd

# Set up paths
excel_path = r"C:\Users\negir\OneDrive\Desktop\Selenium\input.xlsx"
source_dir = r"C:\Users\negir\Downloads"
destination = r"C:\Users\negir\OneDrive\Desktop\Selenium\input.xlsx"

# Initialize Chrome WebDriver
driver = webdriver.Chrome()

# Open the webpage
driver.get("https://rpachallenge.com/")

# Find the link element using XPath with contains()
link_element = driver.find_element(by='xpath', value="//a[contains(@href,'.xlsx')]")

# Click on the link
link_element.click()
time.sleep(1)  # Wait for the download to complete

# Close the browser


# Move the downloaded Excel file to the destination folder
for filename in os.listdir(source_dir):
    if filename.endswith('.xlsx') and 'challenge' in filename.lower():
        shutil.move(src=os.path.join(source_dir, filename), dst=destination)
        break

# Check if the file was copied successfully
if os.path.exists(excel_path):
    

    # Load the Excel file into a DataFrame
    df = pd.read_excel(excel_path)
    start_element =driver.find_element(by='xpath', value="//button[contains(text(),'Start')]")
    # Click on the link
    start_element.click()
    
    # Loop through each cell in the DataFrame
    for index, row in df.iterrows():
        for column, value in row.items():
            #print(f"Row: {index}, Column: {column}, Value: {value}")
            if column=='First Name':
                firstnamelink_element = driver.find_element(by='xpath', value="//input[@ng-reflect-name='labelFirstName']")
                firstnamelink_element.send_keys(value)
            elif column=='Last Name ':
                lastnamelink_element = driver.find_element(by='xpath', value="//input[@ng-reflect-name='labelLastName']")
                lastnamelink_element.send_keys(value)
            elif column=='Company Name':
                companylink_element = driver.find_element(by='xpath', value="//input[@ng-reflect-name='labelCompanyName']")
                companylink_element.send_keys(value)
            elif column=='Role in Company':
                companylink_element = driver.find_element(by='xpath', value="//input[@ng-reflect-name='labelRole']")
                companylink_element.send_keys(value)
            elif column=='Address':
                companylink_element = driver.find_element(by='xpath', value="//input[@ng-reflect-name='labelAddress']")
                companylink_element.send_keys(value)
            elif column=='Email':
                emaillink_element = driver.find_element(by='xpath', value="//input[@ng-reflect-name='labelEmail']")
                emaillink_element.send_keys(value)
            elif column=='Phone Number':
                companylink_element = driver.find_element(by='xpath', value="//input[@ng-reflect-name='labelPhone']")
                companylink_element.send_keys(value)
            
        
        subimtlink_element = driver.find_element(by='xpath', value="//input[@class='btn uiColorButton']")
        
        subimtlink_element.click()
        
else:
    print('File not found')
#driver.quit()
