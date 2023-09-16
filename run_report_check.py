from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from termcolor import colored

import pandas as pd

from utils.PowerBIRestHandler import PowerBIRestHandler

HAS_REPORT_ERRORS = False
RESULTS = []

class PowerBIReportProbe:
    def __init__(self, profile_name=None):
        self.profile_name=profile_name
        self.results = [["url_report", "url_page", "page_nr", "has_error"]]
        self.has_found_any_errors = False

    def init_selenium_driver_edge(self):
        options = webdriver.EdgeOptions()
        
        # add profile options if profile is specified
        if self.profile_name:
            profilePath = r"C:\Users\domin\AppData\Local\Microsoft\Edge\User Data"
            
            # Here you set the path of the profile ending with User Data not the profile folder
            options.add_argument(f"--user-data-dir={profilePath}");
            # Here you specify the actual profile folder
            options.add_argument(f"--profile-directory={self.profile_name}");

        self.driver = webdriver.Edge(options=options)

    def get_report_page_url(self, report_base_url, page_number = None):
        if not page_number or page_number == 1 or page_number ==0:
            # add Report section to make sure we land on the first page of a report
            url = report_base_url + "/ReportSection"
        else:
            url = report_base_url + f"/ReportSection{page_number}"
        return url

    def get_report_page_id(self, url) -> str:
        # returns the id of the power bi report page based on the url

        # check if url contains section information
        if "ReportSection" in url:
            # getting the id of the power bi report page
            sectionId = url.split("ReportSection")[1].split("?")[0]
        
        if not sectionId:
            print("currently on page 1, section id is empty")
        
        return sectionId

    def load_report_page_by_url(self, url):
        # function to handle calls to power bi url, including waits
        print(f"loading url: {url}")
        self.driver.get(url)
        #driver.find_elements(By.TAG_NAME, "pbi-overlay-container")

        element = WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "pbi-overlay-container"))
        )
        if element:
            print("loaded Page")

    def has_report_page_error_visuals(self) -> bool:
        # function to check for visuals which have errors in them
        print(f"Checking for visuals with errors in {self.driver.current_url}")
        try: 
            element = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.TAG_NAME, "canvas-visual-error-overlay"))
            )

            self.has_found_any_errors = True
            return True
        except:
            print("no errors in visuals found")
            return False

    def get_report_all_pages(self, report_base_url):
        # function to loop through all pages within a report
        current_page_id = "" # arbitrary ids
        next_page_id = "init" # arbitrary ids
        next_page_number = 0 # start with zero but directly add 1 in the first loop to properly count pages
        report_page_numbers = 0
        # check if the ids are the same, if we exceed the numbers of pages, the last two ids are the same.
        while current_page_id != next_page_id:
            next_page_number = next_page_number + 1 # start the loop with the next page
            current_page_id = next_page_id
            url = self.get_report_page_url(report_base_url=report_base_url, page_number=next_page_number)
            self.load_report_page_by_url(url=url)
            next_page_id = self.get_report_page_id(self.driver.current_url)

            # ensure we are not checking same page twice
            if current_page_id != next_page_id:
                has_report_page_errors = self.has_report_page_error_visuals()
                
            # Make sure global switch is set that errors have been found
            if has_report_page_errors:
                self.has_found_any_errors = True
            
            report_page_numbers = next_page_number
            
            self.log_results(
                report_base_url=report_base_url,
                url=url,
                report_page_number=report_page_numbers,
                has_report_page_errors=has_report_page_errors
                )
        
            print(f"current page id: {current_page_id}")
            print(f"next page id: {next_page_id}")
            print(f"page numbers in report: {report_page_numbers}")
            print(f"has report page errors: {has_report_page_errors}")

    def log_results(self, report_base_url, url, report_page_number, has_report_page_errors):
        # make sure results are stored somewhere
        self.results.append([report_base_url, url, report_page_number, has_report_page_errors])
    
    def show_results(self):
        df = pd.DataFrame.from_records(self.results)
        print(df)




if __name__ == "__main__":
    client = PowerBIRestHandler()
    workspaceId = client.get_workspace_by_name("Blog - Region")
    print(workspaceId)
    report_urls = client.get_reports_in_workspace(workspaceId=workspaceId)
    print(report_urls)

    input("Make sure thePress Enter to continue...")

    # setup
    probe = PowerBIReportProbe(profile_name="Profile 4")
    probe.init_selenium_driver_edge()

    # main logic
    for report_base_url in report_urls:
        #report_base_url = "https://app.powerbi.com/groups/6f03024f-41a2-4c42-9029-dc8214b96d32/reports/a9f997f8-bf25-4719-a408-21fca5c5ffee"
        probe.get_report_all_pages(report_base_url=report_base_url)

    # results
    if probe.has_found_any_errors:
        print(colored("----------there are errors in the report---------------", "red"))
    else:
        print(colored("------------there are no errors in the report-------------", "green"))

    probe.show_results()

    # close browser
    probe.driver.quit()