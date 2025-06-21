from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException
import pandas as pd
import time
import os
from datetime import datetime


LOGIN_URL = "https://workfreaks.app/ERP/index.php"
EMPLOYEE_DASHBOARD_URL_PART = "/ERP/edit_employee.php" 
LEADS_URL_PART = "/ERP/leads.php" 
ADD_LEAD_URL_PART = "/ERP/add_leads.php" 

YOUR_MOBILE_NUMBER = "Your Number" # e.g., "9876543210"
YOUR_PASSWORD = "Your Password" # e.g., "MySecureP@ss1"


EXCEL_FILE_PATH = "leads.xlsx" 


DEFAULT_COUNTRY_CODE = "India (91)" 
DEFAULT_LEAD_TYPE = "Admission"
DEFAULT_HIGHEST_QUALIFICATION = "Graduate"
DEFAULT_MONTHS_OF_EXP = "0" # Text input, as specified
DEFAULT_PRESENTLY_IN_JOB = "No" # Dropdown, as specified (confirm exact text "Yes" or "yes")
DEFAULT_PREFERRED_COURSE_JOB = "MBA / MCA" # Renamed variable and Text input
DEFAULT_FOLLOW_UP_STATUS = "Yet to receive resume" # Text input
DEFAULT_LEAD_STATUS = "Warm" # Dropdown, as specified ("cold or warm" - picked Warm as default)
DEFAULT_FINAL_LEAD_STATUS = None # Skipping this if no default needed
DEFAULT_NOTES = None # Skipping this if no default needed
DEFAULT_STATUS = "Active" # Dropdown (Final Status, not Lead Status)



def login_to_portal(driver, mobile_no, password):
    """
    Automates the login process.
    """
    print(f"Navigating to login page: {LOGIN_URL}")
    driver.get(LOGIN_URL)

    wait = WebDriverWait(driver, 20)

    print("Waiting for login form fields...")
    try:
        mobile_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='mobileno'][id='mobileno']")))
        password_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='login[password]'][id='password']")))
        login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button#btnSubmit")))
    except Exception as e:
        print(f"Error finding login fields: {e}")
        raise 

    mobile_input.send_keys(mobile_no)
    print("Entered mobile number.")
    password_input.send_keys(password)
    print("Entered password.")
    
    login_button.click()
    print("Clicked Login button.")

    print(f"Waiting for successful login (URL containing '{EMPLOYEE_DASHBOARD_URL_PART}')...")
    wait.until(EC.url_contains(EMPLOYEE_DASHBOARD_URL_PART))
    print("Successfully logged in and reached dashboard.")
    time.sleep(2) 


def set_date_picker_to_today(driver, date_input_locator):
    """
    Attempts to set a date picker input to today's date, prioritizing direct JavaScript input.
    """
    wait = WebDriverWait(driver, 10)
    
    try:
        date_input = wait.until(EC.element_to_be_clickable(date_input_locator))
        today_str = datetime.now().strftime("%Y-%m-%d") 

        
        driver.execute_script(f"arguments[0].value = '{today_str}';", date_input)
        print(f"Set date input value directly to {today_str}.")
        time.sleep(0.2) 

      
        try:
            driver.find_element(By.TAG_NAME, "body").click()
         
        except:
            pass 

        

    except Exception as e:
        print(f"Warning: Could not set Next Follow Up Date to today using direct input. Error: {e}")
        

def add_new_lead(driver, lead_data):
    """
    Automates the process of adding a new lead using the provided data.
    """
    wait = WebDriverWait(driver, 30)

    try:
        
        print("Navigating to Leads page by clicking sidebar link...")
        leads_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'leads.php') and normalize-space(.)='Leads']"))
        )
        leads_link.click()
        print("Clicked 'Leads' sidebar link.")

        
        print(f"Waiting for URL to change to Leads page (containing '{LEADS_URL_PART}')...")
        wait.until(EC.url_contains(LEADS_URL_PART))
        print("Successfully navigated to Leads page (URL confirmed).")
        print(f"Current URL after navigating to Leads page: {driver.current_url}") 

        
        print("Proceeding to click '+' button as 'Search' input field is not a required dependency.")
        time.sleep(1) 

        
        print("Attempting to click the '+' (Add Leads) button using a robust strategy...")
        add_lead_button_locator = (By.CSS_SELECTOR, "button.add-btn.btn-hover")
        
        clicked_plus_button = False
        start_time = time.time()
        timeout_for_plus_click = 10 

        while not clicked_plus_button and (time.time() - start_time) < timeout_for_plus_click:
            try:
                add_lead_button = wait.until(EC.element_to_be_clickable(add_lead_button_locator))
                
                
                add_lead_button.click()
                print("Clicked '+' button using Selenium's click.")
                clicked_plus_button = True
            except ElementClickInterceptedException:
                print("Element click intercepted for '+', attempting JavaScript click.")
                
                driver.execute_script("arguments[0].click();", add_lead_button)
                print("Clicked '+' button using JavaScript.")
                clicked_plus_button = True
            except Exception as e:
                print(f"Attempt to click '+' button failed: {e}. Retrying...")
                time.sleep(0.5) 
            
            
            if ADD_LEAD_URL_PART in driver.current_url:
                print("URL changed after '+' button action, navigation successful.")
                clicked_plus_button = True 
        
        if not clicked_plus_button:
            raise Exception("Failed to click the '+' (Add Leads) button and navigate after multiple attempts.")

        time.sleep(2) 
        print(f"URL immediately after executing '+' button action: {driver.current_url}") 


        
        print(f"Waiting for Add Lead form to load (URL containing '{ADD_LEAD_URL_PART}')...")
        wait.until(EC.url_contains(ADD_LEAD_URL_PART))
        print("Add Lead form URL confirmed.")
        print(f"Current URL after navigating to Add Lead form: {driver.current_url}") 


        print("Waiting for 'Name' field on the add lead form...")
        name_field = wait.until(
            EC.presence_of_element_located((By.ID, "name"))
        )
        print("Add lead form loaded. Filling details.")

        
        name_field.send_keys(lead_data.get('Name', ''))
        print(f"Filled Name: {lead_data.get('Name', '')}")

        mobile_field = driver.find_element(By.ID, "mobile_no")
        mobile_field.send_keys(str(lead_data.get('Mobile Number', '')))
        print(f"Filled Mobile Number: {lead_data.get('Mobile Number', '')}")

        try:
            country_code_element = driver.find_element(By.ID, "country_code")
            if country_code_element.tag_name == "select":
                select_country_code = Select(country_code_element)
                select_country_code.select_by_visible_text(DEFAULT_COUNTRY_CODE)
                print(f"Selected Country Code: {DEFAULT_COUNTRY_CODE}")
            else:
                country_code_element.send_keys(DEFAULT_COUNTRY_CODE)
                print(f"Typed Country Code: {DEFAULT_COUNTRY_CODE}")
        except Exception as e:
            print(f"Warning: Could not set Country Code. Error: {e}")

        try:
            select_lead_type = Select(driver.find_element(By.ID, "lead_type"))
            select_lead_type.select_by_visible_text(DEFAULT_LEAD_TYPE)
            print(f"Selected Lead Type: {DEFAULT_LEAD_TYPE}")
        except Exception as e:
            print(f"Warning: Could not set Lead Type. Error: {e}")
        
        if DEFAULT_LEAD_TYPE == "Job Seeker":
            try:
                highest_qualification_element = wait.until(
                    EC.visibility_of_element_located((By.ID, "highest_qualification"))
                )
                select_qualification = Select(highest_qualification_element)
                select_qualification.select_by_visible_text(DEFAULT_HIGHEST_QUALIFICATION)
                print(f"Selected Highest Qualification: {DEFAULT_HIGHEST_QUALIFICATION}")
            except Exception as e:
                print(f"Warning: Could not set Highest Qualification (may not have appeared). Error: {e}")
        else:
            print(f"Skipping Highest Qualification as Lead Type is not '{DEFAULT_LEAD_TYPE}'.")
        
        try:
            months_exp_field = driver.find_element(By.ID, "months_of_exp")
            months_exp_field.send_keys(DEFAULT_MONTHS_OF_EXP)
            print(f"Filled Months of Exp: {DEFAULT_MONTHS_OF_EXP}")
        except Exception as e:
            print(f"Warning: Could not fill Months of Exp. Error: {e}")

        try:
            select_presently_in_job = Select(driver.find_element(By.ID, "presently_in_job"))
            select_presently_in_job.select_by_visible_text(DEFAULT_PRESENTLY_IN_JOB)
            print(f"Selected Presently In Job: {DEFAULT_PRESENTLY_IN_JOB}")
        except Exception as e:
            print(f"Warning: Could not set Presently In Job. Error: {e}")

        try:
            preferred_field = driver.find_element(By.ID, "preferred_course_job")
            preferred_field.send_keys(DEFAULT_PREFERRED_COURSE_JOB)
            print(f"Filled Preferred Course/Job/Candidates: {DEFAULT_PREFERRED_COURSE_JOB}")
        except Exception as e:
            print(f"Warning: Could not fill Preferred Course/Job/Candidates. Error: {e}")
        
        try:
            follow_up_status_field = driver.find_element(By.ID, "follow_up_status")
            follow_up_status_field.send_keys(DEFAULT_FOLLOW_UP_STATUS)
            print(f"Filled Follow Up Status: {DEFAULT_FOLLOW_UP_STATUS}")
        except Exception as e:
            print(f"Warning: Could not fill Follow Up Status. Error: {e}")

        try:
            set_date_picker_to_today(driver, (By.ID, "next_follow_up_date"))
            print("Set Next Follow Up Date to today.")
        except Exception as e:
            print(f"Warning: Error setting Next Follow Up Date: {e}")

        try:
            select_lead_status = Select(driver.find_element(By.ID, "lead_status"))
            select_lead_status.select_by_visible_text(DEFAULT_LEAD_STATUS)
            print(f"Selected Lead Status: {DEFAULT_LEAD_STATUS}")
        except Exception as e:
            print(f"Warning: Could not set Lead Status. Error: {e}")
        
        if DEFAULT_FINAL_LEAD_STATUS:
             try:
                select_final_lead_status = Select(driver.find_element(By.ID, "final_lead_status"))
                select_final_lead_status.select_by_visible_text(DEFAULT_FINAL_LEAD_STATUS)
                print(f"Selected Final Lead Status: {DEFAULT_FINAL_LEAD_STATUS}")
             except Exception as e:
                print(f"Warning: Could not set Final Lead Status. Error: {e}")
        
        if DEFAULT_NOTES:
            try:
                notes_field = driver.find_element(By.ID, "notes")
                notes_field.send_keys(DEFAULT_NOTES)
                print(f"Filled Notes: {DEFAULT_NOTES}")
            except Exception as e:
                print(f"Warning: Could not fill Notes. Error: {e}")

        try:
            select_status = Select(driver.find_element(By.ID, "status"))
            select_status.select_by_visible_text(DEFAULT_STATUS)
            print(f"Selected Status: {DEFAULT_STATUS}")
        except Exception as e:
            print(f"Warning: Could not set overall Status. Error: {e}")

        
        print("Waiting for Submit button...")
        submit_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-primary"))
        )
        
        
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
        time.sleep(0.5) 

        try:
            submit_button.click()
            print("Clicked the Submit button.")
        except ElementClickInterceptedException:
            print("Element click intercepted, attempting JavaScript click for Submit button.")
            driver.execute_script("arguments[0].click();", submit_button)
            print("Clicked Submit button using JavaScript.")
        except Exception as e: 
            print(f"An unexpected error occurred while clicking Submit button: {e}")
            raise 


        time.sleep(3) 

        
        print("Lead addition attempted. Waiting to return to Leads page...")
        wait.until(EC.url_contains(LEADS_URL_PART))
        print(f"Successfully returned to Leads page after adding {lead_data.get('Name', '')}.")
        return True

    except Exception as e:
        print(f"An error occurred while adding lead '{lead_data.get('Name', 'N/A')}': {e}")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = os.path.join(os.getcwd(), f"error_lead_{lead_data.get('Name', 'N_A').replace(' ', '_')}_{timestamp}.png")
        driver.save_screenshot(screenshot_path)
        print(f"Screenshot saved to: {screenshot_path}")
        return False


def main():
    print("Starting Workfreaks Lead Automation Script (Full Automation Mode)...")

    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"ERROR: Excel file not found at '{EXCEL_FILE_PATH}'. Please create it with 'Name' and 'Mobile Number' columns.")
        return

    try:
        leads_df = pd.read_excel(EXCEL_FILE_PATH)
        if 'Mobile Number' in leads_df.columns:
            leads_df['Mobile Number'] = leads_df['Mobile Number'].astype(str)
        print(f"Successfully loaded {len(leads_df)} leads from '{EXCEL_FILE_PATH}'.")
        if leads_df.empty:
            print("Excel file is empty. No leads to process.")
            return
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        print("Please ensure 'pandas' and 'openpyxl' are installed (`pip install pandas openpyxl`) and the file format is correct.")
        return

    driver = None
    try:
        service = EdgeService(EdgeChromiumDriverManager().install())
        edge_options = EdgeOptions()
        
        
        
        
        driver = webdriver.Edge(service=service, options=edge_options)
        driver.maximize_window() 

        login_to_portal(driver, YOUR_MOBILE_NUMBER, YOUR_PASSWORD)

        successful_leads = 0
        for index, row in leads_df.iterrows():
            print(f"\n--- Processing Lead {index + 1}: {row['Name']} ({row['Mobile Number']}) ---")
            if add_new_lead(driver, row.to_dict()):
                successful_leads += 1
            else:
                print(f"Failed to add lead: {row['Name']}. Moving to next.")
            time.sleep(2) 

        print(f"\n--- Automation Summary ---")
        print(f"Total leads processed: {len(leads_df)}")
        print(f"Successful leads added: {successful_leads}")
        print(f"Failed leads: {len(leads_df) - successful_leads}")

    except Exception as e:
        print(f"An unhandled error occurred during overall automation: {e}")
        if driver:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = os.path.join(os.getcwd(), f"overall_error_screenshot_{timestamp}.png")
            driver.save_screenshot(screenshot_path)
            print(f"Screenshot 'overall_error_screenshot.png' saved to: {screenshot_path}")
    finally:
        if driver:
            print("Closing browser...")
            driver.quit()
        print("Script execution finished.")

if __name__ == "__main__":
    main()