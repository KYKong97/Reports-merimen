from playwright.sync_api import sync_playwright
import re
import logging
import pandas as pd
from bs4 import BeautifulSoup
import time
from src.lib import preprocess_df


class MerimenController:
    def __init__(self, merimen_username:str, merimen_password:str, headless=False,
                 executable_path=r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",slow_mo=None,
                 url = "https://report.merimen.com.my/claims/index.cfm") -> None:
        
        playwright = sync_playwright().start()
        self.browser = playwright.chromium.launch(headless=headless, executable_path=executable_path, slow_mo=slow_mo)
        self.merimen_page = self.browser.new_page()
        self.merimen_page.goto(url)
        self.merimen_page.locator("#sleUserName").fill(merimen_username)
        self.merimen_page.locator("#slePassword").fill(merimen_password)
        self.merimen_page.get_by_role("button", name="Login").click()
        self.merimen_page.wait_for_load_state("networkidle")
    
    def filter_claim_type(self, claim_type:str="TP BI"):

        try:
            self.merimen_page.get_by_role("button",name="ClmTypes").click(timeout=1000000)
            self.merimen_page.locator("//span[contains(text(), '[Select All]')]").click(timeout=1000000)
            self.merimen_page.locator("#MTRCTContextMenu").get_by_role("cell", name=claim_type).locator("#MTRCTL").check(timeout=1000000)
            self.merimen_page.get_by_role("button", name="Done").click(timeout=1000000)
            self.merimen_page.wait_for_load_state("networkidle")
            return 1
        except Exception as e:
            logging.error(f"Err in filter claim type {e}")

            return -1


    def filter_report_date(self, from_date:str,to_date:str):
        try:
            from_date_field = self.merimen_page.locator("#DrFrom")
            to_date_field = self.merimen_page.locator("#DrTo")
            from_date_field.clear(timeout=1000000)
            to_date_field.clear(timeout=1000000)

            from_date_field.fill(from_date)
            to_date_field.fill(to_date)
            return 1
        
        except Exception as e:
            logging.error(f"Err in filter report date. {e}")
            return -1

    def generate_report(self):
        try:
            
            # self.merimen_page.locator("//a[@id='ClaimMenu$2']").click(timeout=1000000)
            self.merimen_page.get_by_role("link", name="Process").click()
            self.merimen_page.wait_for_load_state("networkidle")
            return 1
        except Exception as e:
            logging.error(f"Err in generate_report. {e}")
            return -1

    def check_opinion_report(self):
        try:
            self.merimen_page.get_by_label("Latest Solicitor Opinion Report Submitted Date").check(timeout=1000000)
            return 1
        
        except Exception as e:
            logging.error(f"Err in check opinion report")
            return -1

    def read_report_table(self):
        try:
            table = self.merimen_page.locator("//table[@border='1']")
            html = table.inner_html()

            list_claim = []

            soup = BeautifulSoup(html, 'html.parser')
            rows = soup.find_all('tr')
            for row in rows[1:]:
                tds = row.find_all('td')
                no_value = tds[0].get_text().strip()
                claim_no = tds[1].get_text().strip()
                liable_amount = tds[22].get_text().strip()
                panel_solicitor = tds[8].get_text().strip()
                pic = tds[28].get_text().strip()
                panel_solicitor_assigned_date = tds[12].get_text().strip()
                latest_solicitor_opinion_report_date = tds[18].get_text().strip()

                temp_dict = {
                    "No":no_value,
                    "Claim No":claim_no,
                    "Solicitor Worksheet Liable Amount":liable_amount,
                    "Panel Solicitor":panel_solicitor,
                    "PIC":pic,
                    "Panel Solicitor Assigned Date":panel_solicitor_assigned_date,
                    "Latest Solicitor Opinion Report Submitted Date":latest_solicitor_opinion_report_date
                }
                list_claim.append(temp_dict)
                
            
            df = pd.DataFrame(list_claim)
            
            return df
        except Exception as e:
            logging.error(f"Error in read report. Error in {e}")
            return -1



    