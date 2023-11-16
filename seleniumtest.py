from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from termcolor import colored

import pandas as pd

from utils.PowerBIRestHandler import PowerBIRestHandler

import argparse
import yaml
import os


os.system('taskkill /f /im  "msedge.exe"')
options = webdriver.EdgeOptions()

profile_name = "Profile 4"

# add profile options if profile is specified
if profile_name:
    profilePath = r"C:\Users\domin\AppData\Local\Microsoft\Edge\User Data"
    
    # Here you set the path of the profile ending with User Data not the profile folder
    options.add_argument(f"--user-data-dir={profilePath}");
    # Here you specify the actual profile folder
    options.add_argument(f"--profile-directory={profile_name}");

driver = webdriver.Edge(options=options)

url = "https://app.powerbi.com/groups/6f03024f-41a2-4c42-9029-dc8214b96d32/reports/cee527bb-c766-4472-a567-8da775d2cc97/ReportSectione665a1d5a408a69b78bd?experience=power-bi"
driver.get(url)


# Assuming the buttons are within a mat-action-list
mat_action_list = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, "//mat-action-list[@data-testid='pages-navigation-list']"))
)

# Locate the buttons within the mat-action-list (adjust the locator as per your HTML structure)
buttons = mat_action_list.find_elements(By.TAG_NAME, "button")

for button in buttons:
    button.click()
    print(driver.current_url)
    # You can use WebDriverWait for this purpose
    #WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.ID, "loading-spinner-id")))
