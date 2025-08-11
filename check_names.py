import os
script_dir = os.path.dirname(os.path.abspath(__file__))

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time

import warnings
warnings.filterwarnings("ignore", category=FutureWarning)

EXCEL_FILE = os.path.join(script_dir, "names.xlsx")
SHEET_NAME = 'Sheet1'
SEARCH_URL = 'https://or.justice.cz/ias/ui/rejstrik-$firma'
CHROME_DRIVER_PATH = 'chromedriver'

# load excel file
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
if 'MATCH_COUNT' not in df.columns:
    df['MATCH_COUNT'] = ''

# webscrape setup
print('Starting Web Browser')
options = Options()
options.add_argument("--headless")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.get(SEARCH_URL)
time.sleep(2.0) # Wait for page load

# checking logic
print('Checking names:')
for index, row in df.iterrows():

    #for each name in file
    name = row['NAMES']
    print(name)
    #find input box and enter name
    input_box = driver.find_element(By.ID, 'id2')
    input_box.clear()
    input_box.send_keys(name)
    input_box.send_keys(Keys.ENTER)
    
    time.sleep(0.5) #increase with worse internet speed, 
    #below 0.3 justice.cz might not respond in time, irrespective of response time
    
    #read number of name matches
    try:
        result_count = driver.find_element(By.CSS_SELECTOR,'#SearchResults h2.fl span')
        match_count = int(result_count.text)
    except Exception: #if errors give 0
        match_count = 0

    #update dataframe
    if match_count == 0:
        df.at[index, 'MATCH_COUNT'] = 'free'
    else:
        df.at[index, 'MATCH_COUNT'] = match_count

#output
print('Search Complete')
free_df = df[df['MATCH_COUNT'] == 'free']
others_df = df[df['MATCH_COUNT'] != 'free']
df_final = pd.concat([free_df, others_df], ignore_index=True) # reorder such that free names are at top of output file

# save back to excel sheet
with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_final.to_excel(writer, sheet_name=SHEET_NAME, index=False)

print('File Saved')

driver.quit()
print("Done.")

time.sleep(3)
