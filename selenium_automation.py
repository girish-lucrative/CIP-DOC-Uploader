"""
Selenium automation for CIP-Signal portal login with process-specific navigation, IEC selection, BRC type selection, and file upload
"""

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
import time
import sys
import os
import glob
from pathlib import Path
from selenium import webdriver


# Mapping of our process types to CIP portal card names
PROCESS_TO_CARD_MAP = {
    'dbk_disbursement': 'DBK_SCROLL',
    'dbk_pendency': 'DBK_PENDING', 
    'brc': 'BRC',
    'igst_scroll': 'IGST_SCROLL',
    'rodtep_scroll': 'RODTEP_SCROLL',
    'rodtep_scrip': 'RODTEP_SCRIP'
    # Add other mappings as needed
}

def select_brc_type(driver, wait, brc_type):
    """Select BRC type (FOB or INV) in the portal before IEC selection"""
    try:
        print(f"\nAttempting to select BRC type: {brc_type}")
        time.sleep(2)  # Wait for page to load completely

        # Map UI brc_type to portal options (FOB or INV)
        brc_type_upper = brc_type.upper() if brc_type else 'FOB'
        
        # Wait for the BRC type selector to be present
        print("Looking for BRC type selector...")
        
        # Strategy 1: Look for the first ant-select (which is BRC Type)
        try:
            # Find the card-body containing BRC
            card_body = wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "card-body"))
            )
            
            # Find all ant-select elements within card-body
            ant_selects = card_body.find_elements(By.CLASS_NAME, "ant-select")
            
            if len(ant_selects) >= 2:
                # First ant-select should be BRC Type (width: 150px)
                brc_type_selector = ant_selects[0]
                
                # Verify it's the BRC type selector by placeholder
                try:
                    placeholder = brc_type_selector.find_element(
                        By.CLASS_NAME, "ant-select-selection-placeholder"
                    )
                    if "Select Type" in placeholder.text:
                        print(f"✓ Found BRC type selector: '{placeholder.text}'")
                        
                        # Click to open dropdown
                        print("Clicking BRC type selector...")
                        brc_type_selector.click()
                        time.sleep(1)
                        
                        # Now find and click the option
                        # Options are in a dropdown list with class "ant-select-item-option"
                        dropdown_options = wait.until(
                            EC.presence_of_all_elements_located((By.CLASS_NAME, "ant-select-item-option"))
                        )
                        
                        print(f"Found {len(dropdown_options)} dropdown options")
                        
                        # Look for option matching our brc_type
                        for option in dropdown_options:
                            option_text = option.text.strip().upper()
                            print(f"Option: {option_text}")
                            if brc_type_upper in option_text or option_text == brc_type_upper:
                                print(f"Found matching option: '{option.text}', clicking...")
                                option.click()
                                time.sleep(1)
                                print(f"✓ BRC type {brc_type_upper} selected")
                                return True
                        
                        # If exact match not found, click first option
                        if dropdown_options:
                            print(f"No exact match, clicking first option: '{dropdown_options[0].text}'")
                            dropdown_options[0].click()
                            time.sleep(1)
                            print("✓ Clicked first dropdown option")
                            return True
                        
                    else:
                        print(f"Placeholder text is: '{placeholder.text}'")
                except Exception as e:
                    print(f"Error checking placeholder: {e}")
            else:
                print(f"Found only {len(ant_selects)} ant-select elements, expected at least 2")
                
        except Exception as e:
            print(f"Error in Strategy 1: {e}")
        
        # Strategy 2: Direct approach using the search input
        try:
            print("\nTrying Strategy 2: Direct search input approach...")
            
            # Find the first ant-select-search-input (for BRC Type)
            search_inputs = driver.find_elements(By.CLASS_NAME, "ant-select-selection-search-input")
            
            if len(search_inputs) >= 1:
                brc_search_input = search_inputs[0]
                
                print("Clicking BRC type search input...")
                brc_search_input.click()
                time.sleep(0.5)
                
                # Type the BRC type
                print(f"Typing BRC type: {brc_type_upper}")
                brc_search_input.send_keys(brc_type_upper)
                time.sleep(1)
                
                # Press Enter
                brc_search_input.send_keys(Keys.RETURN)
                time.sleep(1)
                print(f"✓ BRC type {brc_type_upper} entered")
                return True
                
        except Exception as e:
            print(f"Error in Strategy 2: {e}")
        
        # Strategy 3: Look by style attribute (width: 150px)
        try:
            print("\nTrying Strategy 3: Style attribute approach...")
            
            brc_type_select = wait.until(
                EC.element_to_be_clickable((By.XPATH, 
                    "//div[contains(@class, 'ant-select') and contains(@style, 'width: 150px')]"
                ))
            )
            
            print("Found BRC type selector by style (width: 150px)")
            brc_type_select.click()
            time.sleep(1)
            
            # Now press down arrow and Enter to select
            actions = ActionChains(driver)
            actions.send_keys(brc_type_upper)
            time.sleep(0.5)
            actions.send_keys(Keys.RETURN)
            actions.perform()
            time.sleep(1)
            
            print(f"✓ BRC type {brc_type_upper} selected using keyboard")
            return True
            
        except Exception as e:
            print(f"Error in Strategy 3: {e}")
        
        print("⚠ Could not select BRC type after trying all strategies")
        return False
    
    except Exception as e:
        print(f"⚠ Could not select BRC type: {e}")
        import traceback
        traceback.print_exc()
        return False

def login_and_navigate(username, password, process_type, iec_number=None, file_to_upload=None, brc_type=None):
    """
    Automate login to CIP-Signal portal and navigate to specific process dashboard
    Args:
        username (str): Login username/email
        password (str): Login password
        process_type (str): The process type selected in our app
        iec_number (str, optional): IEC number to select in the portal
        file_to_upload (str, optional): Path to the file to upload
        brc_type (str, optional): BRC type (FOB or INV) for BRC process
    
    Returns:
        dict: Result with status and message
    """
    result = {
        'success': False,
        'message': ''
    }
    
    driver = None
    try:
        print("Starting CIP-Signal automation...")
        
        # Configure Chrome options - NO HEADLESS (visible browser)
        chrome_options = Options()
        
        # Add options for better performance and compatibility
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--window-size=1920,1080')
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument("--disable-save-password-bubble")
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_experimental_option("prefs", {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False
        })
        # NO HEADLESS - Browser will be visible
        # chrome_options.add_argument('--headless')  # REMOVED
        
        # Initialize Chrome driver
        print("Initializing Chrome driver...")
        driver = webdriver.Chrome(options=chrome_options)
        
        # Navigate to CIP-Signal portal
        print("Navigating to https://www.cip-lucrative.com...")
        driver.get("https://www.cip-lucrative.com")
        
        # Wait for page to load
        wait = WebDriverWait(driver, 10)
        time.sleep(2)
        
        # Find login elements - SIMPLE APPROACH
        print("Looking for login form...")
        
        # Find email field
        email_field = wait.until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Enter email' or @type='email']"))
        )
        print("Found email field")
        
        # Find password field
        password_field = driver.find_element(By.XPATH, "//input[@placeholder='Enter password' or @type='password']")
        print("Found password field")
        
        # Find submit button
        try:
            submit_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Sign In') or contains(text(), 'Login') or @type='submit']")
            print("Found submit button")
        except:
            submit_button = None
        
        # Enter credentials
        print(f"Entering username: {username}")
        email_field.clear()
        email_field.send_keys(username)
        time.sleep(0.5)
        
        print("Entering password...")
        password_field.clear()
        password_field.send_keys(password)
        time.sleep(0.5)
        
        # Submit form
        if submit_button:
            print("Clicking Sign In button...")
            submit_button.click()
        else:
            print("Pressing Enter in password field...")
            password_field.send_keys(Keys.RETURN)
        
        # Wait for login to complete
        print("Waiting for login to complete...")
        time.sleep(3)
        
        # Check if login was successful by looking for dashboard elements
        print("Checking for successful login...")
        
        # Look for dashboard indicators
        try:
            # Wait for dashboard or navigation elements
            wait.until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Dashboard') or contains(text(), 'Upload') or contains(text(), 'dashboard')]"))
            )
            print("✓ Login successful - Dashboard detected")
            result['success'] = True
            result['message'] = "Successfully logged in"
            
            # Now navigate to specific process dashboard based on process_type
            if process_type in PROCESS_TO_CARD_MAP:
                card_name = PROCESS_TO_CARD_MAP[process_type]
                print(f"Navigating to {card_name} dashboard...")
                
                # Construct the URL for the specific card
                card_url = f"https://www.cip-lucrative.com/dashboard/upload/upload-files?__card__={card_name}"
                
                # Navigate to the specific card URL
                print(f"Opening URL: {card_url}")
                driver.get(card_url)
                time.sleep(3)
                
                # Check if we're on the right page
                current_url = driver.current_url
                if card_name in current_url:
                    print(f"✓ Successfully navigated to {card_name} dashboard")
                    result['message'] = f"Successfully logged in and navigated to {card_name} dashboard"
                else:
                    print(f"⚠ Might not be on {card_name} dashboard, but login was successful")
                
                # FOR BRC PROCESS: Select BRC type first (FOB or INV)
                if process_type == 'brc' and brc_type:
                    select_brc_type_success = select_brc_type(driver, wait, brc_type)
                    if select_brc_type_success:
                        result['message'] += f", selected BRC type: {brc_type.upper()}"
                    else:
                        print("⚠ Could not select BRC type, continuing...")
                
                # Now select IEC number if provided
                if iec_number and iec_number.strip():
                    select_iec_success = select_iec_number(driver, wait, iec_number)
                    if select_iec_success:
                        result['message'] += f" and selected IEC: {iec_number}"
                    else:
                        print("⚠ Could not select IEC number, continuing...")
                else:
                    print("No IEC number provided, skipping IEC selection")
                
                # Upload file if provided
                if file_to_upload and os.path.exists(file_to_upload):
                    upload_success = upload_file_to_portal(driver, wait, file_to_upload)
                    if upload_success:
                        result['message'] += f" and uploaded file: {os.path.basename(file_to_upload)}"
                        result['success'] = True
                    else:
                        print("⚠ File upload failed")
                else:
                    print(f"No file to upload or file doesn't exist: {file_to_upload}")
                    
            else:
                print(f"⚠ No specific card mapping for process: {process_type}")
                print("Staying on main dashboard")
                
        except Exception as e:
            print(f"⚠ Could not confirm dashboard: {e}")
            # Check if we're still on login page
            current_url = driver.current_url
            if "login" in current_url.lower() or "signin" in current_url.lower():
                result['message'] = "Login failed - Still on login page"
                print("✗ Login failed")
            else:
                result['success'] = True
                result['message'] = "Login likely successful (not on login page)"
                print("✓ Login likely successful")
    
    except TimeoutException as e:
        result['message'] = f"Timeout while waiting for page elements: {str(e)}"
        print(f"✗ Timeout error: {e}")
    except NoSuchElementException as e:
        result['message'] = f"Required element not found: {str(e)}"
        print(f"✗ Element not found: {e}")
    except Exception as e:
        result['message'] = f"Error during automation: {str(e)}"
        print(f"✗ Error: {e}")
        import traceback
        traceback.print_exc()
    
    # Keep browser open for a while if successful
    if driver:
        if result['success']:
            print("\n✓ Automation completed successfully!")
            print("Browser will remain open for 30 seconds for manual inspection...")
            print("Press Ctrl+C in terminal to stop the application")
            
            # Keep browser open for longer
            try:
                time.sleep(30)  # Keep browser open for 30 seconds
            except KeyboardInterrupt:
                print("\nClosing browser...")
        
        # Close the browser
        try:
            driver.quit()
            print("Browser closed.")
        except:
            pass
    
    return result

def select_iec_number(driver, wait, iec_number):
    """Select IEC number in the portal - should be called AFTER BRC type selection for BRC process"""
    try:
        print(f"\nAttempting to select IEC number: {iec_number}")
        time.sleep(2)  # Wait for previous actions to complete
        
        # Strategy 1: Find the second ant-select in card-body (IEC selector)
        try:
            # Find the card-body containing BRC
            card_body = wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "card-body"))
            )
            
            # Find all ant-select elements within card-body
            ant_selects = card_body.find_elements(By.CLASS_NAME, "ant-select")
            
            if len(ant_selects) >= 2:
                # Second ant-select should be IEC (width: 300px)
                iec_selector = ant_selects[1]
                
                # Verify it's the IEC selector by placeholder
                try:
                    placeholder = iec_selector.find_element(
                        By.CLASS_NAME, "ant-select-selection-placeholder"
                    )
                    if "Select IEC" in placeholder.text:
                        print(f"✓ Found IEC selector: '{placeholder.text}'")
                        
                        # Click to open dropdown
                        print("Clicking IEC selector...")
                        iec_selector.click()
                        time.sleep(1)
                        
                        # Find the search input within this selector
                        search_input = iec_selector.find_element(
                            By.CLASS_NAME, "ant-select-selection-search-input"
                        )
                        
                        # Type the IEC number
                        print(f"Typing IEC number: {iec_number}")
                        search_input.click()
                        time.sleep(0.5)
                        search_input.send_keys(Keys.CONTROL + "a")
                        search_input.send_keys(Keys.DELETE)
                        time.sleep(0.5)
                        search_input.send_keys(iec_number)
                        time.sleep(1)
                        
                        # Press Enter
                        search_input.send_keys(Keys.RETURN)
                        time.sleep(1)
                        print(f"✓ IEC number {iec_number} entered")
                        
                        # Wait for dropdown to show options
                        time.sleep(1)
                        
                        # Look for dropdown options and click first one
                        dropdown_options = driver.find_elements(By.CLASS_NAME, "ant-select-item-option")
                        if dropdown_options:
                            print(f"Found {len(dropdown_options)} IEC options")
                            # Click the first option (should be our IEC)
                            dropdown_options[0].click()
                            print("Clicked first IEC option")
                            time.sleep(1)
                        
                        return True
                        
                    else:
                        print(f"IEC placeholder text is: '{placeholder.text}'")
                except Exception as e:
                    print(f"Error checking IEC placeholder: {e}")
            else:
                print(f"Found only {len(ant_selects)} ant-select elements, expected at least 2")
                
        except Exception as e:
            print(f"Error in IEC Strategy 1: {e}")
        
        # Strategy 2: Direct approach using the second search input
        try:
            print("\nTrying Strategy 2: Direct IEC search input approach...")
            
            # Find all search inputs
            search_inputs = driver.find_elements(By.CLASS_NAME, "ant-select-selection-search-input")
            
            if len(search_inputs) >= 2:
                iec_search_input = search_inputs[1]  # Second one is for IEC
                
                print("Clicking IEC search input...")
                iec_search_input.click()
                time.sleep(0.5)
                
                # Type the IEC number
                print(f"Typing IEC number: {iec_number}")
                iec_search_input.send_keys(Keys.CONTROL + "a")
                iec_search_input.send_keys(Keys.DELETE)
                time.sleep(0.5)
                iec_search_input.send_keys(iec_number)
                time.sleep(1)
                
                # Press Enter
                iec_search_input.send_keys(Keys.RETURN)
                time.sleep(1)
                print(f"✓ IEC number {iec_number} entered")
                return True
                
        except Exception as e:
            print(f"Error in IEC Strategy 2: {e}")
        
        # Strategy 3: Look by style attribute (width: 300px)
        try:
            print("\nTrying Strategy 3: Style attribute approach for IEC...")
            
            iec_select = wait.until(
                EC.element_to_be_clickable((By.XPATH, 
                    "//div[contains(@class, 'ant-select') and contains(@style, 'width: 300px')]"
                ))
            )
            
            print("Found IEC selector by style (width: 300px)")
            iec_select.click()
            time.sleep(1)
            
            # Find the search input
            search_input = iec_select.find_element(
                By.CLASS_NAME, "ant-select-selection-search-input"
            )
            
            # Type the IEC number
            search_input.click()
            time.sleep(0.5)
            search_input.send_keys(Keys.CONTROL + "a")
            search_input.send_keys(Keys.DELETE)
            time.sleep(0.5)
            search_input.send_keys(iec_number)
            time.sleep(1)
            search_input.send_keys(Keys.RETURN)
            time.sleep(1)
            
            print(f"✓ IEC number {iec_number} selected")
            return True
            
        except Exception as e:
            print(f"Error in IEC Strategy 3: {e}")
        
        print("⚠ Could not select IEC number after trying all strategies")
        return False
    
    except Exception as e:
        print(f"⚠ Could not select IEC number: {e}")
        import traceback
        traceback.print_exc()
        return False

def upload_file_to_portal(driver, wait, file_path):
    """Upload file to the portal after BRC type and IEC selection"""
    try:
        print(f"\nAttempting to upload file: {file_path}")
        time.sleep(2)  # Wait for previous selections to complete
        
        # Look for file input in the card-body
        print("Looking for file input...")
        
        # Strategy 1: Look in card-body
        try:
            card_body = wait.until(
                EC.presence_of_element_located((By.CLASS_NAME, "card-body"))
            )
            
            # Find file input in card-body
            file_input = card_body.find_element(By.CSS_SELECTOR, "input[type='file']")
            print("Found file input in card-body")
            
            # Send the file path
            print(f"Sending file path: {file_path}")
            file_input.send_keys(os.path.abspath(file_path))
            time.sleep(2)
            
            # Look for upload button in card-body
            upload_button = card_body.find_element(By.CLASS_NAME, "upload-btn")
            print("Found upload button")
            
            # Click upload button
            print("Clicking upload button...")
            upload_button.click()
            time.sleep(3)
            
            print("✓ File upload initiated")
            return True
            
        except Exception as e:
            print(f"Error in upload Strategy 1: {e}")
        
        # Strategy 2: Direct file input search
        try:
            file_input = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input.file-input[type='file']"))
            )
            
            print("Found file input by class 'file-input'")
            file_input.send_keys(os.path.abspath(file_path))
            time.sleep(2)
            
            # Find upload button by class
            upload_button = wait.until(
                EC.element_to_be_clickable((By.CLASS_NAME, "upload-btn"))
            )
            
            upload_button.click()
            time.sleep(3)
            print("✓ File uploaded")
            return True
            
        except Exception as e:
            print(f"Error in upload Strategy 2: {e}")
        
        print("⚠ Could not upload file")
        return False
        
    except Exception as e:
        print(f"⚠ Error during file upload: {e}")
        import traceback
        traceback.print_exc()
        return False

def find_latest_downloaded_file(download_dir=None, pattern="*.xlsx"):
    """
    Find the most recently downloaded file in the download directory
    Args:
        download_dir: Directory to search for files (default: user's Downloads folder)
        pattern: File pattern to match (default: *.xlsx)
    Returns:
        str: Path to the latest file, or None if not found
    """
    if download_dir is None:
        # Try to get user's Downloads folder
        download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    
    if not os.path.exists(download_dir):
        print(f"Download directory doesn't exist: {download_dir}")
        return None
    
    # Find all files matching the pattern
    files = glob.glob(os.path.join(download_dir, pattern))
    
    if not files:
        print(f"No files found matching pattern: {pattern} in {download_dir}")
        return None
    
    # Get the most recently modified file
    latest_file = max(files, key=os.path.getmtime)
    print(f"Found latest downloaded file: {latest_file}")
    print(f"  Modified: {time.ctime(os.path.getmtime(latest_file))}")
    
    return latest_file

if __name__ == "__main__":
    # Test the automation
    if len(sys.argv) >= 4:
        username = sys.argv[1]
        password = sys.argv[2]
        process_type = sys.argv[3]
        iec_number = sys.argv[4] if len(sys.argv) > 4 else None
        brc_type = sys.argv[5] if len(sys.argv) > 5 else None
        file_to_upload = sys.argv[6] if len(sys.argv) > 6 else None
        
        print(f"Testing CIP-Signal login with:")
        print(f"Username: {username}")
        print(f"Process: {process_type}")
        if iec_number:
            print(f"IEC Number: {iec_number}")
        if brc_type and process_type == 'brc':
            print(f"BRC Type: {brc_type}")
        if file_to_upload:
            print(f"File to upload: {file_to_upload}")
        
        # If no file specified, try to find the latest downloaded file
        if not file_to_upload:
            file_to_upload = find_latest_downloaded_file()
            if file_to_upload:
                print(f"Using latest downloaded file: {file_to_upload}")
        
        result = login_and_navigate(username, password, process_type, iec_number, file_to_upload, brc_type)
        
        print(f"\nResult: {result['success']}")
        print(f"Message: {result['message']}")
    else:
        # Use fixed credentials for testing
        fixed_username = "asdf@12331"
        fixed_password = "1234"
        fixed_process = "brc"  # Test BRC process
        fixed_brc_type = "FOB"  # Test BRC type
        fixed_iec = "ALFA12345"  # Test IEC number
        
        print(f"Testing CIP-Signal login with fixed credentials")
        print(f"Username: {fixed_username}")
        print(f"Process: {fixed_process}")
        print(f"BRC Type: {fixed_brc_type}")
        print(f"IEC Number: {fixed_iec}")
        
        # Try to find a file to upload
        test_file = find_latest_downloaded_file()
        if test_file:
            print(f"Test file to upload: {test_file}")
        else:
            print("No test file found for upload")
        
        result = login_and_navigate(fixed_username, fixed_password, fixed_process, fixed_iec, test_file, fixed_brc_type)
        
        print(f"\nResult: {result['success']}")
        print(f"Message: {result['message']}")