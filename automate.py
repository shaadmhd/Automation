import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
# from webdriver_manager.microsoft import EdgeChromiumDriverManager # COMMENTED OUT: We are using a manual driver now
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException, NoSuchElementException, WebDriverException
import time
import os
from datetime import datetime, timedelta

# --- Configuration Variables ---
LOGIN_URL = "https://workfreaks.app/ERP/index.php"
EMPLOYEE_DASHBOARD_URL_PART = "/ERP/edit_employee.php"
LEADS_URL_PART = "/ERP/leads.php"
ADD_LEAD_URL_PART = "/ERP/add_leads.php"

YOUR_MOBILE_NUMBER = "9600879549" # <--- REPLACE with your actual mobile number
YOUR_PASSWORD = "vasan@naruto" # <--- REPLACE with your actual password

# Excel file path (make sure this file is in the same directory as the script, or provide a full path)
EXCEL_FILE_PATH = "leads.xlsx"

# --- WebDriver Initialization ---
def initialize_driver():
    """Initializes and returns a Selenium WebDriver for Microsoft Edge."""
    try:
        # --- MODIFIED LINE START ---
        # COMMENTED OUT: service = EdgeService(EdgeChromiumDriverManager().install())
        
        # Path to the manually downloaded msedgedriver.exe
        # User specified the driver is located at C:\SeleniumDrivers
        driver_path = r"C:\SeleniumDrivers\msedgedriver.exe" # <--- FIX: Changed driver path to specified location
        service = EdgeService(driver_path)
        # --- MODIFIED LINE END ---

        edge_options = EdgeOptions()
        # Uncomment these lines to run in headless mode (without opening a browser UI)
        # edge_options.add_argument("--headless")
        # edge_options.add_argument("--window-size=1920,1080")
        edge_options.add_argument("--disable-gpu") # Recommended for headless mode on some systems
        edge_options.add_argument("--no-sandbox") # Recommended for Linux environments
        edge_options.add_argument("--disable-dev-shm-usage") # Overcomes limited resource problems
        
        driver = webdriver.Edge(service=service, options=edge_options)
        driver.maximize_window() # Maximize for better element visibility if not headless
        print("WebDriver initialized successfully.")
        return driver
    except Exception as e:
        print(f"Error initializing WebDriver: {e}")
        print("Please ensure Microsoft Edge is installed and up-to-date.")
        # --- MODIFIED ERROR MESSAGE START ---
        print("Also, ensure 'msedgedriver.exe' is downloaded (matching your Edge browser version) and placed in 'C:\\SeleniumDrivers'.")
        # --- MODIFIED ERROR MESSAGE END ---
        return None

# --- ERP Login Function ---
def login_to_portal(driver, mobile_no, password):
    """
    Automates the login process.
    """
    print(f"Navigating to login page: {LOGIN_URL}")
    driver.get(LOGIN_URL)

    wait = WebDriverWait(driver, 30) # Increased timeout for login fields

    print("Waiting for login form fields...")
    try:
        mobile_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='mobileno'][id='mobileno']")))
        password_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='login[password]'][id='password']")))
        login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button#btnSubmit")))
    except TimeoutException:
        print("Timeout: Login fields or button not found within the expected time. Page structure might have changed.")
        raise
    except NoSuchElementException:
        print("NoSuchElement: Login fields or button not found. Check selectors.")
        raise
    except Exception as e:
        print(f"Error finding login fields: {e}")
        raise # Re-raise to fail early if login fields aren't found

    mobile_input.clear()
    mobile_input.send_keys(mobile_no)
    print("Entered mobile number.")
    password_input.clear()
    password_input.send_keys(password)
    print("Entered password.")
    
    login_button.click()
    print("Clicked Login button.")

    print(f"Waiting for successful login (URL containing '{EMPLOYEE_DASHBOARD_URL_PART}')...")
    try:
        wait.until(EC.url_contains(EMPLOYEE_DASHBOARD_URL_PART))
        print("Successfully logged in and reached dashboard.")
        time.sleep(2) # Small pause after successful login for page rendering
        return True
    except TimeoutException:
        print(f"Timeout waiting for dashboard URL ({EMPLOYEE_DASHBOARD_URL_PART}). Login might have failed or page took too long to load.")
        return False
    except Exception as e:
        print(f"An unexpected error occurred after clicking login button: {e}")
        return False

def set_date_picker_to_next_day(driver, date_input_locator):
    """
    Attempts to set a date picker input to the next day's date, prioritizing direct JavaScript input.
    """
    wait = WebDriverWait(driver, 10)
    
    try:
        date_input = wait.until(EC.element_to_be_clickable(date_input_locator))
        
        # Calculate tomorrow's date
        tomorrow = datetime.now() + timedelta(days=1)
        tomorrow_str = tomorrow.strftime("%Y-%m-%d") # Format for input type="date" (e.g., 2025-07-24)

        driver.execute_script(f"arguments[0].value = '{tomorrow_str}';", date_input)
        print(f"Set date input value directly to {tomorrow_str} (next day).")
        time.sleep(0.2) # Small pause to ensure value registers

        # Attempt to dismiss any potential date picker pop-up by clicking body, if it appeared briefly
        try:
            driver.find_element(By.TAG_NAME, "body").click()
        except:
            pass # Ignore if body click fails

    except Exception as e:
        print(f"Warning: Could not set Next Follow Up Date to next day using direct input. Error: {e}")


def read_excel_data(file_path):
    """
    Reads lead data from an Excel file, expecting all necessary form fields.
    Ensures empty cells are treated as empty strings.
    """
    try:
        df = pd.read_excel(file_path)
        
        # Define all columns the script expects, matching the ERP form fields
        # These are the headers that MUST be present in your Excel file
        required_columns = [
            "Name",
            "Mobile Number",
            "Lead Type",
            "Highest Qualification",
            "Presently in Job",
            "Lead Status",
            "Final Lead Status",
            "Status", # Overall Active/Inactive status
            "Notes", # This now directly matches the ERP field name "Notes"
            "Email",
            "Location",
            "Company Name",
            "Company Email",
            "Preferred Course/Job/Candidates",
            "Follow Up Status",
            "Next Follow Up Date", # This column should exist, though value is ignored by script
            "No. of Months of Exp",
            "Country Code"
        ]
        
        # Check for missing columns and add them to the DataFrame with empty strings
        for col in required_columns:
            if col not in df.columns:
                print(f"Warning: Column '{col}' not found in Excel file. "
                      f"Adding it with empty values. Please ensure all required column headers match exactly.")
                # For 'Country Code', if missing, initialize with '91'
                if col == "Country Code":
                    df[col] = '91' 
                else:
                    df[col] = '' # Add missing column with empty string values

        # Ensure 'Mobile Number' is treated as string to prevent scientific notation (e.g., 9.87E+09)
        # Also remove any decimal parts if they exist after conversion to string
        if 'Mobile Number' in df.columns:
            df['Mobile Number'] = df['Mobile Number'].astype(str).apply(lambda x: x.split('.')[0] if '.' in x else x)

        # Convert all values to string, handling NaN (empty cells) and None specifically for send_keys
        leads_data_list = []
        for index, row in df.iterrows():
            lead_dict = {}
            for col in df.columns:
                value = row[col]
                # Check if value is NaN (from empty Excel cell) or None, convert to empty string
                if pd.isna(value) or value is None: 
                    # If Country Code is empty/NaN in Excel, set to '91' here
                    if col == "Country Code":
                        lead_dict[col] = '91'
                    else:
                        lead_dict[col] = '' 
                else:
                    lead_dict[col] = str(value).strip() # Convert to string and strip whitespace
            leads_data_list.append(lead_dict)

        return leads_data_list

    except FileNotFoundError:
        print(f"Error: Excel file not found at '{file_path}'. Please check the path and filename.")
        return None
    except pd.errors.EmptyDataError:
        print(f"Error: Excel file '{file_path}' is empty or has no data. No leads to process.")
        return None
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        print("Please ensure 'pandas' and 'openpyxl' are installed (`pip install pandas openpyxl`) and the file format is correct.")
        return None

def add_new_lead(driver, lead_data):
    """
    Automates the process of adding a new lead using the provided data from Excel.
    """
    wait = WebDriverWait(driver, 30) # Increased timeout for robustness

    try:
        print(f"\n--- Processing Lead: {lead_data.get('Name', 'N/A')} (Mobile: {lead_data.get('Mobile Number', 'N/A')}) ---")

        # --- Navigate to Leads page by clicking the sidebar link ---
        print("Navigating to Leads page by clicking sidebar link...")
        leads_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'leads.php') and normalize-space(.)='Leads']"))
        )
        leads_link.click()
        print("Clicked 'Leads' sidebar link.")

        # --- Wait for Leads page to load ---
        print(f"Waiting for URL to change to Leads page (containing '{LEADS_URL_PART}')...")
        wait.until(EC.url_contains(LEADS_URL_PART))
        print("Successfully navigated to Leads page (URL confirmed).")
        time.sleep(1) # Small buffer after page loads

        print("Proceeding to click '+' button to add a new lead.")

        # --- Click the "Add Leads" (+) Button (More Robust Strategy) ---
        add_lead_button_locator = (By.CSS_SELECTOR, "button.add-btn.btn-hover")
        
        clicked_plus_button = False
        start_time = time.time()
        timeout_for_plus_click = 10 # seconds to try clicking the plus button

        while not clicked_plus_button and (time.time() - start_time) < timeout_for_plus_click:
            try:
                add_lead_button = wait.until(EC.element_to_be_clickable(add_lead_button_locator))
                
                # Try standard Selenium click first
                add_lead_button.click()
                print("Clicked '+' button using Selenium's click.")
                clicked_plus_button = True
            except ElementClickInterceptedException:
                print("Element click intercepted for '+', attempting JavaScript click.")
                # Fallback to JavaScript click if intercepted
                driver.execute_script("arguments[0].click();", add_lead_button)
                print("Clicked '+' button using JavaScript.")
                clicked_plus_button = True
            except Exception as e:
                print(f"Attempt to click '+' button failed: {e}. Retrying...")
                time.sleep(0.5) # Wait a bit before retrying
            
            # After an attempted click, check if the URL has changed, which means success
            if ADD_LEAD_URL_PART in driver.current_url:
                print("URL changed after '+' button action, navigation successful.")
                clicked_plus_button = True # Confirm navigation happened
        
        if not clicked_plus_button:
            raise Exception("Failed to click the '+' (Add Leads) button and navigate after multiple attempts.")

        time.sleep(2) # Give a bit more time for navigation to fully render

        # --- Wait for the Add Lead Form to Load ---
        print(f"Waiting for Add Lead form to load (URL containing '{ADD_LEAD_URL_PART}')...")
        wait.until(EC.url_contains(ADD_LEAD_URL_PART))
        print("Add Lead form URL confirmed. Filling details.")

        # --- Fill in the form details using lead_data from Excel ---
        # Name
        name_field = wait.until(EC.element_to_be_clickable((By.ID, "name")))
        name_field.clear()
        name_field.send_keys(lead_data.get('Name', ''))
        print(f"Filled Name: {lead_data.get('Name', '')}")

        # Mobile Number
        mobile_field = driver.find_element(By.ID, "mobile_no")
        mobile_field.clear()
        mobile_field.send_keys(lead_data.get('Mobile Number', ''))
        print(f"Filled Mobile Number: {lead_data.get('Mobile Number', '')}")

        # Country Code - Using JavaScript for robustness
        try:
            country_code_input = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "country_code")))
            # The default '91' is now handled in read_excel_data, so we just use the value from lead_data
            country_code_value = lead_data.get('Country Code', '') 
            
            # Use JavaScript to clear and set the value for higher reliability
            driver.execute_script("arguments[0].value = '';", country_code_input) # Clear using JS
            driver.execute_script("arguments[0].value = arguments[1];", country_code_input, country_code_value) # Set value using JS
            print(f"Filled Country Code: {country_code_value}")
        except TimeoutException: print("Timeout: Country Code input field not found. Skipping.")
        except NoSuchElementException: print("NoSuchElement: Country Code input field not found. Skipping.")
        except Exception as e: print(f"Error filling Country Code: {e}. Skipping.")


        # Lead Type (Crucial for dynamic forms - select this first if possible)
        excel_lead_type = lead_data.get("Lead Type", "").strip()
        try:
            lead_type_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "lead_type")))
            Select(lead_type_element).select_by_visible_text(excel_lead_type)
            print(f"Selected Lead Type: {excel_lead_type}")
            time.sleep(3) # Increased CRITICAL wait time for form to re-render after selection
        except TimeoutException: 
            print("Timeout: Lead Type dropdown not found or not clickable. Cannot proceed with lead details for conditional fields.")
            return False 
        except NoSuchElementException: 
            print("NoSuchElement: Lead Type dropdown not found. Cannot proceed with lead details for conditional fields.")
            return False 
        except Exception as e: 
            print(f"Error selecting Lead Type '{excel_lead_type}': {e}. Cannot proceed with lead details for conditional fields.")
            return False 
        
        # --- Conditional Fields based on Lead Type ---
        if excel_lead_type in ["Job Provider", "Client - Other Services"]:
            try:
                company_name_field = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "company_name")))
                company_name_field.clear()
                company_name_field.send_keys(lead_data.get("Company Name", ""))
                print(f"Filled Company Name: {lead_data.get('Company Name', '')}")
            except TimeoutException: print("Timeout: Company Name field not found (expected for Job Provider/Client). Skipping.")
            except NoSuchElementException: print("NoSuchElement: Company Name field not found (expected for Job Provider/Client). Skipping.")
            except Exception as e: print(f"Error filling Company Name: {e}. Skipping.")

            try:
                company_email_field = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "company_email")))
                company_email_field.clear()
                company_email_field.send_keys(lead_data.get("Company Email", ""))
                print(f"Filled Company Email: {lead_data.get('Company Email', '')}")
            except TimeoutException: print("Timeout: Company Email field not found (expected for Job Provider/Client). Skipping.")
            except NoSuchElementException: print("NoSuchElement: Company Email field not found (expected for Job Provider/Client). Skipping.")
            except Exception as e: print(f"Error filling Company Email: {e}. Skipping.")
        else:
            print(f"Skipping Company Name/Email as Lead Type is not 'Job Provider' or 'Client - Other Services'.")


        # Highest Qualification 
        excel_highest_qual = lead_data.get("Highest Qualification", "").strip()
        try:
            highest_qual_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "highest_qualification")))
            Select(highest_qual_element).select_by_visible_text(excel_highest_qual)
            print(f"Selected Highest Qualification: {excel_highest_qual}")
        except TimeoutException: print("Timeout: Highest Qualification field not found. Skipping (might be conditional or incorrect value).")
        except NoSuchElementException: print("NoSuchElement: Highest Qualification field not found. Skipping (might be conditional or incorrect value).")
        except Exception as e: print(f"Error selecting Highest Qualification '{excel_highest_qual}': {e}. Skipping.")


        # No. of Months of Exp field 
        excel_months_of_exp = lead_data.get("No. of Months of Exp", "").strip()
        try:
            months_exp_field = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "months_of_exp"))) 
            months_exp_field.clear()
            months_exp_field.send_keys(excel_months_of_exp)
            print(f"Entered No. of Months of Exp: {excel_months_of_exp}")
        except TimeoutException: print("Timeout: No. of Months of Exp field not found. Skipping (might be conditional).")
        except NoSuchElementException: print("NoSuchElement: No. of Months of Exp field not found. Skipping (might be conditional).")
        except Exception as e: print(f"Error filling No. of Months of Exp: {e}. Skipping.")


        # --- Common Fields ---
        try:
            presently_in_job_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "presently_in_job")))
            Select(presently_in_job_element).select_by_visible_text(lead_data.get("Presently in Job", "No").strip())
            print(f"Selected Presently in Job: {lead_data.get('Presently in Job', '').strip()}")
        except TimeoutException: print("Timeout: Presently in Job dropdown not found. Skipping.")
        except NoSuchElementException: print("NoSuchElement: Presently in Job dropdown not found. Skipping.")
        except Exception as e: print(f"Error selecting Presently in Job: {e}. Skipping.")

        try:
            preferred_field = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "preferred_course_job"))) 
            preferred_field.clear()
            preferred_field.send_keys(lead_data.get("Preferred Course/Job/Candidates", ""))
            print(f"Filled Preferred Course/Job/Candidates: {lead_data.get('Preferred Course/Job/Candidates', '')}")
        except TimeoutException: print("Timeout: Preferred Course/Job/Candidates field not found. Skipping.")
        except NoSuchElementException: print("NoSuchElement: Preferred Course/Job/Candidates field not found. Skipping.")
        except Exception as e: print(f"Error filling Preferred Course/Job/Candidates: {e}. Skipping.")
        
        try:
            follow_up_status_field = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "follow_up_status")))
            follow_up_status_field.clear()
            follow_up_status_field.send_keys(lead_data.get("Follow Up Status", ""))
            print(f"Filled Follow Up Status: {lead_data.get('Follow Up Status', '')}")
        except TimeoutException: print("Timeout: Follow Up Status field not found. Skipping.")
        except NoSuchElementException: print("NoSuchElement: Follow Up Status field not found. Skipping.")
        except Exception as e: print(f"Error filling Follow Up Status: {e}. Skipping.")

        # Next Follow Up Date - ALWAYS SET TO NEXT DAY'S DATE
        try:
            set_date_picker_to_next_day(driver, (By.ID, "next_follow_up_date")) # This function sets next day's date
        except Exception as e:
            print(f"Warning: Error setting Next Follow Up Date: {e}")

        try:
            lead_status_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "lead_status")))
            Select(lead_status_element).select_by_visible_text(lead_data.get("Lead Status", "Cold").strip())
            print(f"Selected Lead Status: {lead_data.get('Lead Status', '').strip()}")
        except TimeoutException: print("Timeout: Lead Status dropdown not found. Skipping.")
        except NoSuchElementException: print("NoSuchElement: Lead Status dropdown not found. Skipping.")
        except Exception as e: print(f"Error selecting Lead Status: {e}. Skipping.")

        # Final Lead Status (dependent on Lead Type) - Improved Waiting
        excel_final_lead_status = lead_data.get("Final Lead Status", "").strip()
        try:
            # Wait specifically for the final_lead_status dropdown to be clickable after Lead Type selection
            final_lead_status_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "final_lead_status"))) # Increased timeout here
            valid_statuses = {
                "Admission": ["Sl Scheduled", "Sl Reject / Dropped", "Application Filled", "Application Fees Paid", "Admission Fees Pending", "Admission Fees Collected"],
                "Job Seeker": ["YTR In HRP", "Details Pending", "Registered", "Job Not Assigned", "Job Suggested", "Job Applied", "Processed (Sent To AM)", "Approved (Sent To Client)", "Scheduled (Cl Sch)", "Client Feedback Awaited", "Shortlisted", "Rejected", "Selected", "Offered", "Joined", "Dropped", "Billed"],
                "Job Provider": ["Proposal Yet to Share", "Proposal Shared", "Agreement Shared", "Agreement Signed", "Job Posted", "Client Billed", "Client Payment Collected"],
                "Client - Other Services": ["Business In Progress", "Business Dropped", "Business Billed"],
                "FLR": ["Completed", "Pending"]
            }
            if excel_lead_type in valid_statuses and excel_final_lead_status in valid_statuses[excel_lead_type]:
                Select(final_lead_status_element).select_by_visible_text(excel_final_lead_status)
                print(f"Selected Final Lead Status: {excel_final_lead_status}")
            else:
                print(f"Warning: Invalid or missing Final Lead Status '{excel_final_lead_status}' for Lead Type '{excel_lead_type}'. Skipping selection.")
        except TimeoutException: 
            print(f"Timeout: Final Lead Status dropdown not found or not clickable within 10 seconds. It might not have loaded options for Lead Type '{excel_lead_type}'. Skipping.")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = os.path.join(os.getcwd(), f"debug_final_lead_status_timeout_{lead_data.get('Name', 'N_A').replace(' ', '_')}_{timestamp}.png")
            try:
                driver.save_screenshot(screenshot_path)
                print(f"Debug Screenshot saved to: {screenshot_path}")
            except Exception as se:
                print(f"Could not save debug screenshot: {se}")
        except NoSuchElementException: 
            print("NoSuchElement: Final Lead Status dropdown not found. Skipping.")
        except Exception as e: 
            print(f"Error selecting Final Lead Status: {e}. Skipping.")
        
        # Notes field (now explicitly "Notes" in Excel) - Using JavaScript for robustness
        try:
            notes_value = lead_data.get("Notes", "") # Use "Notes" directly from Excel
            
            # Find the element directly. If not found, NoSuchElementException is raised quickly.
            notes_element = driver.find_element(By.ID, "notes")
            
            # Use JavaScript to clear and set the value for higher reliability
            driver.execute_script("arguments[0].value = '';", notes_element) # Clear using JS
            driver.execute_script("arguments[0].value = arguments[1];", notes_element, notes_value) # Set value using JS
            print(f"Filled Notes: {notes_value}") # Adjusted print statement
        except NoSuchElementException:
            print("Notes field not found. Skipping this field.")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = os.path.join(os.getcwd(), f"debug_notes_not_found_{lead_data.get('Name', 'N_A').replace(' ', '_')}_{timestamp}.png")
            try:
                driver.save_screenshot(screenshot_path)
                print(f"Debug Screenshot saved to: {screenshot_path}")
            except Exception as se:
                print(f"Could not save debug screenshot: {se}")
        except Exception as e:
            print(f"Error filling Notes field: {e}. Skipping.")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = os.path.join(os.getcwd(), f"debug_notes_error_{lead_data.get('Name', 'N_A').replace(' ', '_')}_{timestamp}.png")
            try:
                driver.save_screenshot(screenshot_path)
                print(f"Debug Screenshot saved to: {screenshot_path}")
            except Exception as se:
                print(f"Could not save debug screenshot: {se}")
        
        # Email field
        try:
            email_field = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "email"))) 
            email_field.clear()
            email_field.send_keys(lead_data.get("Email", ""))
            print(f"Filled Email: {lead_data.get("Email", "")}")
        except TimeoutException: print("Timeout: Email field not found. Skipping. (Please verify its ID)")
        except NoSuchElementException: print("NoSuchElement: Email field not found. Skipping. (Please verify its ID)")
        except Exception as e: print(f"Error filling Email: {e}. Skipping.")

        # Location field
        try:
            location_field = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "location"))) 
            location_field.clear()
            location_field.send_keys(lead_data.get("Location", ""))
            print(f"Filled Location: {lead_data.get("Location", "")}")
        except TimeoutException: print("Timeout: Location field not found. Skipping. (Please verify its ID)")
        except NoSuchElementException: print("NoSuchElement: Location field not found. Skipping. (Please verify its ID)")
        except Exception as e: print(f"Error filling Location: {e}. Skipping.")

        try:
            status_element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "status")))
            Select(status_element).select_by_visible_text(lead_data.get("Status", "Active").strip())
            print(f"Selected Status: {lead_data.get('Status', '').strip()}")
        except TimeoutException: print("Timeout: Status dropdown not found. Skipping.")
        except NoSuchElementException: print("NoSuchElement: Status dropdown not found. Skipping.")
        except Exception as e: print(f"Error selecting Status: {e}. Skipping.")


        # --- Find and click the Submit button ---
        print("Waiting for Submit button...")
        submit_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-primary"))
        )
        
        # Scroll the button into view before clicking
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
        time.sleep(0.5) # Small pause after scrolling

        try:
            submit_button.click()
            print("Clicked the Submit button.")
        except ElementClickInterceptedException:
            print("Element click intercepted, attempting JavaScript click for Submit button.")
            driver.execute_script("arguments[0].click();", submit_button)
            print("Clicked Submit button using JavaScript.")
        except Exception as e: # Catch any other unexpected errors during click
            print(f"An unexpected error occurred while clicking Submit button: {e}")
            raise # Re-raise to be caught by the outer try-except for screenshot.

        time.sleep(3) # Give some time for the submission to process and page to reload

        # After submission, it should likely return to the leads page.
        print("Lead addition attempted. Waiting to return to Leads page...")
        wait.until(EC.url_contains(LEADS_URL_PART))
        print(f"Successfully returned to Leads page after adding {lead_data.get('Name', '')}.")
        return True

    except Exception as e:
        print(f"An error occurred while adding lead '{lead_data.get('Name', 'N/A')}': {e}")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = os.path.join(os.getcwd(), f"error_lead_{lead_data.get('Name', 'N_A').replace(' ', '_')}_{timestamp}.png")
        try:
            driver.save_screenshot(screenshot_path)
            print(f"Screenshot saved to: {screenshot_path}")
        except Exception as se:
            print(f"Could not save screenshot: {se}")
        return False


def main():
    print("Starting Workfreaks Lead Automation Script...")

    leads_to_process = read_excel_data(EXCEL_FILE_PATH)

    if leads_to_process is None: # read_excel_data returns None on critical error
        print("Exiting script due to error reading Excel file.")
        return
    elif not leads_to_process: # No leads found or empty file
        print("No leads to process found in the Excel file. Exiting script.")
        return

    driver = None
    try:
        driver = initialize_driver()

        if driver:
            if login_to_portal(driver, YOUR_MOBILE_NUMBER, YOUR_PASSWORD):
                print("Successfully logged in.")
                print(f"Current URL after login: {driver.current_url}")

                successful_leads = 0
                for i, lead in enumerate(leads_to_process):
                    print(f"\n--- Initiating processing for Lead {i+1}/{len(leads_to_process)} ---")
                    if add_new_lead(driver, lead):
                        successful_leads += 1
                        print(f"Lead '{lead.get('Name')}' processed successfully.")
                    else:
                        print(f"Failed to process lead '{lead.get('Name')}'. Please check the error messages and screenshot (if any). Skipping to next lead.")
                    time.sleep(2) # Pause between leads to prevent overwhelming the server

                print(f"\n--- Automation Summary ---")
                print(f"Total leads attempted: {len(leads_to_process)}")
                print(f"Successful leads added: {successful_leads}")
                print(f"Failed leads: {len(leads_to_process) - successful_leads}")

            else:
                print("Login failed. Please check your credentials. Exiting script.")
        else:
            print("WebDriver failed to initialize. Cannot proceed with automation. Exiting script.")
    except WebDriverException as e:
        print(f"A critical WebDriver error occurred during overall script execution: {e}")
        print("This usually means the browser crashed or disconnected unexpectedly, or EdgeDriver issues.")
        if driver:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = os.path.join(os.getcwd(), f"critical_error_screenshot_{timestamp}.png")
            try:
                driver.save_screenshot(screenshot_path)
                print(f"Screenshot saved to: {screenshot_path}")
            except Exception as se:
                print(f"Could not save screenshot during critical error: {se}")
    except Exception as e:
        print(f"An unhandled error occurred during overall script execution: {e}")
        if driver:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = os.path.join(os.getcwd(), f"unhandled_error_screenshot_{timestamp}.png")
            try:
                driver.save_screenshot(screenshot_path)
                print(f"Screenshot saved to: {screenshot_path}")
            except Exception as se:
                print(f"Could not save screenshot during unhandled error: {se}")
    finally:
        if driver:
            print("\nAutomation script finished. Quitting WebDriver.")
            driver.quit()

if __name__ == "__main__":
    main()
