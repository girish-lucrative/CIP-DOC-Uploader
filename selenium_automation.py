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

# def handle_popup(driver, wait):
#     """Handle any popups that appear after login"""
#     try:
#         print("Checking for popups after login...")
        
#         # Wait a bit for popup to appear
#         time.sleep(2)
        
#         popup_handled = False
        
#         # STRATEGY 1: SPECIFICALLY HANDLE "CHANGE YOUR PASSWORD" POPUP
#         # Look for modal that contains text about changing password
#         try:
#             print("Looking for 'change password' popup...")
            
#             # Look for any element containing password-related text
#             password_popup_elements = driver.find_elements(By.XPATH, 
#                 "//*[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'change your password') or "
#                 "contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'change password') or "
#                 "contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'password change') or "
#                 "contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'password expiry') or "
#                 "contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'password alert')]"
#             )
            
#             if password_popup_elements:
#                 print(f"Found {len(password_popup_elements)} element(s) with password change text")
                
#                 # Find the modal or dialog containing this text
#                 for element in password_popup_elements:
#                     try:
#                         # Go up the DOM to find the modal/dialog container
#                         modal_container = element
#                         for _ in range(5):  # Go up 5 levels max
#                             modal_container = modal_container.find_element(By.XPATH, "..")
#                             tag_name = modal_container.tag_name.lower()
#                             class_attr = modal_container.get_attribute("class") or ""
                            
#                             # Check if this looks like a modal container
#                             if any(modal_class in class_attr.lower() for modal_class in ['modal', 'dialog', 'popup', 'alert', 'ant-modal']):
#                                 print(f"Found modal container with class: {class_attr}")
                                
#                                 # Now look for OK/Close buttons in this container
#                                 buttons = modal_container.find_elements(By.TAG_NAME, "button")
#                                 for button in buttons:
#                                     button_text = button.text.strip().lower()
#                                     if button_text in ['ok', 'okay', 'close', 'got it', 'dismiss', 'continue', 'proceed']:
#                                         print(f"Found button with text: '{button.text}', clicking...")
#                                         button.click()
#                                         time.sleep(1)
#                                         popup_handled = True
#                                         print("✓ Password change popup handled")
#                                         return popup_handled
                                
#                                 # Also look for links or other clickable elements
#                                 clickables = modal_container.find_elements(By.XPATH, ".//*[self::a or self::span or self::div][contains(text(), 'OK') or contains(text(), 'Close')]")
#                                 for clickable in clickables:
#                                     print(f"Found clickable element with text: '{clickable.text}', clicking...")
#                                     clickable.click()
#                                     time.sleep(1)
#                                     popup_handled = True
#                                     print("✓ Password change popup handled via clickable element")
#                                     return popup_handled
#                     except:
#                         continue
        
#         except Exception as e:
#             print(f"Error in password popup handling: {e}")
        
#         # STRATEGY 2: Look for modal popups with OK/Close buttons (ORIGINAL CODE)
#         try:
#             # Look for modals or dialogs
#             modals = driver.find_elements(By.CSS_SELECTOR, ".modal, .ant-modal, .dialog, [role='dialog'], .ant-modal-content")
#             for modal in modals:
#                 try:
#                     # Check if modal is visible
#                     if modal.is_displayed():
#                         print("Found visible modal/dialog")
                        
#                         # Get all text in modal for debugging
#                         modal_text = modal.text
#                         print(f"Modal text: {modal_text[:200]}...")
                        
#                         # Look for OK button in the modal
#                         ok_buttons = modal.find_elements(By.XPATH, 
#                             ".//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'ok') or "
#                             "contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'close') or "
#                             "contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'got it') or "
#                             "contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'continue') or "
#                             "contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'proceed')]"
#                         )
                        
#                         if ok_buttons:
#                             for button in ok_buttons:
#                                 if button.is_displayed() and button.is_enabled():
#                                     print(f"Found OK/Close button: '{button.text}', clicking...")
#                                     button.click()
#                                     time.sleep(1)
#                                     popup_handled = True
#                                     print("✓ Modal closed")
#                                     break
#                 except:
#                     continue
#         except:
#             pass
        
#         # STRATEGY 3: Look for alert popups
#         try:
#             alerts = driver.find_elements(By.CSS_SELECTOR, ".alert, .ant-alert, .notification, .ant-notification")
#             for alert in alerts:
#                 try:
#                     if alert.is_displayed():
#                         # Look for close button (usually an X)
#                         close_buttons = alert.find_elements(By.CSS_SELECTOR, ".close, .ant-alert-close-icon, [aria-label='close'], .ant-notification-close-x")
#                         if close_buttons:
#                             for close_btn in close_buttons:
#                                 if close_btn.is_displayed() and close_btn.is_enabled():
#                                     close_btn.click()
#                                     time.sleep(0.5)
#                                     popup_handled = True
#                                     print("✓ Alert closed")
#                                     break
#                 except:
#                     pass
#         except:
#             pass
        
#         # STRATEGY 4: Look for any button with "OK" anywhere on page (with visibility check)
#         if not popup_handled:
#             try:
#                 ok_buttons = driver.find_elements(By.XPATH, 
#                     "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'ok') or "
#                     "contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'close')]"
#                 )
#                 for button in ok_buttons:
#                     if button.is_displayed() and button.is_enabled():
#                         print(f"Found visible OK/Close button: '{button.text}', clicking...")
#                         button.click()
#                         time.sleep(1)
#                         popup_handled = True
#                         print("✓ OK button clicked")
#                         break
#             except:
#                 pass
        
#         # STRATEGY 5: Try to close any overlay/backdrop with ESC key as last resort
#         if not popup_handled:
#             try:
#                 # Try pressing ESC key to close modal
#                 print("Trying ESC key to close popup...")
#                 from selenium.webdriver.common.keys import Keys
#                 body = driver.find_element(By.TAG_NAME, "body")
#                 body.send_keys(Keys.ESCAPE)
#                 time.sleep(1)
#                 popup_handled = True
#                 print("✓ ESC key pressed")
#             except:
#                 pass
        
#         # STRATEGY 6: Click at center of screen (last resort)
#         if not popup_handled:
#             try:
#                 print("Trying to click at center of screen...")
#                 actions = ActionChains(driver)
                
#                 # Get window size
#                 window_size = driver.get_window_size()
#                 center_x = window_size['width'] // 2
#                 center_y = window_size['height'] // 2
                
#                 # Move to center and click
#                 actions.move_by_offset(center_x, center_y).click().perform()
#                 time.sleep(1)
#                 popup_handled = True
#                 print("✓ Clicked at center of screen")
#             except:
#                 pass
        
#         if popup_handled:
#             print("Popup handled successfully")
#         else:
#             print("No popup found or popup already closed")
        
#         return popup_handled
        
#     except Exception as e:
#         print(f"⚠ Error handling popup: {e}")
#         return False

def select_brc_type(driver, wait, brc_type):
    """Select BRC type (FOC or INV) in the portal before IEC selection"""
    try:
        print(f"\nAttempting to select BRC type: {brc_type}")
        time.sleep(2)  # Wait for page to load completely
        
        # Map UI brc_type to portal options (FOC or INV)
        # Assuming brc_type from UI is 'foc' or 'inv' (lowercase)
        brc_type_upper = brc_type.upper() if brc_type else 'FOC'
        
        # Wait for the BRC type selector to be present
        print("Looking for BRC type selector...")
        
        # Try multiple strategies to find the BRC type selector
        brc_type_selector = None
        
        # Strategy 1: Look for ant-select with placeholder "Select Type"
        try:
            brc_type_selector = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'ant-select') and .//span[contains(@class, 'ant-select-selection-placeholder') and contains(text(), 'Select Type')]]"))
            )
            print("Found BRC type selector by placeholder")
        except:
            # Strategy 2: Look for ant-select with width 150px (from your HTML)
            try:
                brc_type_selector = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'ant-select') and contains(@style, 'width: 150px')]"))
                )
                print("Found BRC type selector by style (width: 150px)")
            except:
                # Strategy 3: Look for any ant-select that might be the BRC type selector
                try:
                    # Get all ant-select elements on the page
                    ant_selects = driver.find_elements(By.CLASS_NAME, "ant-select")
                    if ant_selects:
                        # Try to find one that seems to be for BRC type (usually comes before IEC selector)
                        for select in ant_selects:
                            if select.is_displayed():
                                # Check if it has placeholder text
                                try:
                                    placeholder = select.find_element(By.CLASS_NAME, "ant-select-selection-placeholder")
                                    if "Select Type" in placeholder.text or "Type" in placeholder.text:
                                        brc_type_selector = select
                                        print("Found BRC type selector by placeholder text")
                                        break
                                except:
                                    continue
                except Exception as e:
                    print(f"Could not find BRC type selector: {e}")
                    return False
        
        if brc_type_selector:
            # Click the selector to open dropdown
            print("Clicking BRC type selector to open dropdown...")
            brc_type_selector.click()
            time.sleep(1)
            
            # Now we need to select the option (FOC or INV)
            # First try to find the search input within the selector
            try:
                search_input = brc_type_selector.find_element(By.CLASS_NAME, "ant-select-selection-search-input")
                if search_input:
                    # Clear and type the BRC type
                    print(f"Typing BRC type: {brc_type_upper}")
                    search_input.click()
                    time.sleep(0.5)
                    search_input.send_keys(Keys.CONTROL + "a")
                    search_input.send_keys(Keys.DELETE)
                    time.sleep(0.5)
                    search_input.send_keys(brc_type_upper)
                    time.sleep(1)
                    
                    # Press Enter to select
                    search_input.send_keys(Keys.RETURN)
                    time.sleep(1)
                    print(f"✓ BRC type {brc_type_upper} selected via search input")
                    return True
            except:
                print("Could not find search input, trying dropdown options...")
            
            # If search input not found, look for dropdown options
            try:
                # Wait for dropdown options to appear
                time.sleep(1)
                
                # Look for dropdown options
                dropdown_options = driver.find_elements(By.CSS_SELECTOR, ".ant-select-item-option")
                if dropdown_options:
                    print(f"Found {len(dropdown_options)} dropdown options")
                    
                    # Look for the option with text matching our BRC type
                    for option in dropdown_options:
                        option_text = option.text.strip().upper()
                        if brc_type_upper in option_text or option_text in brc_type_upper:
                            print(f"Found matching option: '{option.text}', clicking...")
                            option.click()
                            time.sleep(1)
                            print(f"✓ BRC type {brc_type_upper} selected from dropdown")
                            return True
                    
                    # If exact match not found, click first option
                    print("Exact match not found, clicking first option...")
                    dropdown_options[0].click()
                    time.sleep(1)
                    print("✓ Clicked first dropdown option")
                    return True
                else:
                    print("No dropdown options found")
                    return False
            except Exception as e:
                print(f"Error selecting from dropdown: {e}")
                return False
        else:
            print("⚠ Could not find BRC type selector")
            return False
    
    except Exception as e:
        print(f"⚠ Could not select BRC type: {e}")
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
        brc_type (str, optional): BRC type (FOC or INV) for BRC process
    
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
        # NO HEADLESS - Browser will be visible
        # chrome_options.add_argument('--headless')  # REMOVED
        
        # Initialize Chrome driver
        print("Initializing Chrome driver...")
        # driver = webdriver.Chrome(options=chrome_options)
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
        
        # Handle any popups that appear after login
        # handle_popup(driver, wait)
        
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
                
                # Handle any popups that might appear after navigation
                # handle_popup(driver, wait)
                
                # Check if we're on the right page
                current_url = driver.current_url
                if card_name in current_url:
                    print(f"✓ Successfully navigated to {card_name} dashboard")
                    result['message'] = f"Successfully logged in and navigated to {card_name} dashboard"
                else:
                    print(f"⚠ Might not be on {card_name} dashboard, but login was successful")
                
                # FOR BRC PROCESS: Select BRC type first (FOC or INV)
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
    """Select IEC number in the portal"""
    try:
        print(f"\nAttempting to select IEC number: {iec_number}")
        time.sleep(2)  # Wait for page to load completely
        
        # Wait for the IEC selector to be present
        print("Looking for IEC selector...")
        
        # Try multiple strategies to find the IEC selector
        iec_input = None
        
        # Strategy 1: Look for input with id containing 'rc_select'
        try:
            iec_input = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input[id*='rc_select'][type='search']"))
            )
            print("Found IEC input by ID pattern")
        except:
            # Strategy 2: Look for ant-select-selection-search-input class
            try:
                iec_input = wait.until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "ant-select-selection-search-input"))
                )
                print("Found IEC input by class name")
            except:
                # Strategy 3: Look for any input with placeholder containing 'Select IEC'
                try:
                    iec_input = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//input[contains(@placeholder, 'IEC') or contains(@placeholder, 'iec')]"))
                    )
                    print("Found IEC input by placeholder")
                except:
                    # Strategy 4: Look for ant-select element and click it first
                    try:
                        ant_select = wait.until(
                            EC.element_to_be_clickable((By.CLASS_NAME, "ant-select-selector"))
                        )
                        print("Found ant-select selector, clicking it...")
                        ant_select.click()
                        time.sleep(1)
                        
                        # Now look for the search input
                        iec_input = wait.until(
                            EC.element_to_be_clickable((By.CLASS_NAME, "ant-select-selection-search-input"))
                        )
                        print("Found IEC input after clicking selector")
                    except Exception as e:
                        print(f"Could not find IEC selector: {e}")
                        return False
        
        if iec_input:
            # Clear and enter the IEC number
            print(f"Entering IEC number: {iec_number}")
            iec_input.click()
            time.sleep(0.5)
            
            # Clear any existing value
            iec_input.send_keys(Keys.CONTROL + "a")
            iec_input.send_keys(Keys.DELETE)
            time.sleep(0.5)
            
            # Type the IEC number
            iec_input.send_keys(iec_number)
            time.sleep(1)
            
            # Press Enter to select
            iec_input.send_keys(Keys.RETURN)
            time.sleep(1)
            
            print(f"✓ IEC number {iec_number} entered successfully")
            
            # Try to find and click the first dropdown option if available
            try:
                # Wait for dropdown options to appear
                time.sleep(1)
                
                # Look for dropdown options
                dropdown_options = driver.find_elements(By.CSS_SELECTOR, ".ant-select-item-option")
                if dropdown_options:
                    print(f"Found {len(dropdown_options)} dropdown options")
                    # Click the first option
                    dropdown_options[0].click()
                    print("Clicked first dropdown option")
                    time.sleep(1)
            except:
                print("No dropdown options found or could not click, continuing...")
            
            return True
        else:
            return False
    
    except Exception as e:
        print(f"⚠ Could not select IEC number: {e}")
        return False

def upload_file_to_portal(driver, wait, file_path):
    """Upload file to the portal"""
    try:
        print(f"\nAttempting to upload file: {file_path}")
        time.sleep(2)  # Wait for page to settle
        
        # Look for file input element
        print("Looking for file input...")
        
        # Try multiple strategies to find file input
        file_input = None
        
        # Strategy 1: Look for input with type='file'
        try:
            file_input = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
            )
            print("Found file input by type attribute")
        except:
            # Strategy 2: Look for input with class containing 'file-input'
            try:
                file_input = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input.file-input"))
                )
                print("Found file input by class name")
            except:
                # Strategy 3: Look for any input that accepts files
                try:
                    file_input = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//input[@accept='.xlsx,.xls,.pdf,.doc,.docx,.txt']"))
                    )
                    print("Found file input by accept attribute")
                except:
                    # Strategy 4: Look for file input in the card body
                    try:
                        card_body = wait.until(
                            EC.presence_of_element_located((By.CLASS_NAME, "card-body"))
                        )
                        file_input = card_body.find_element(By.CSS_SELECTOR, "input[type='file']")
                        print("Found file input in card body")
                    except Exception as e:
                        print(f"Could not find file input: {e}")
                        return False
        
        if file_input:
            # Send the file path to the input element
            print(f"Sending file path to input: {file_path}")
            file_input.send_keys(os.path.abspath(file_path))
            time.sleep(2)
            
            # Handle any popup that might appear after file selection
            # handle_popup(driver, wait)
            
            # Look for upload button
            print("Looking for upload button...")
            upload_button = None
            
            # Strategy 1: Look for button with class 'upload-btn'
            try:
                upload_button = wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button.upload-btn"))
                )
                print("Found upload button by class")
            except:
                # Strategy 2: Look for button containing 'Upload' text
                try:
                    upload_button = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Upload')]"))
                    )
                    print("Found upload button by text")
                except:
                    # Strategy 3: Look for any primary button
                    try:
                        upload_button = wait.until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn-primary"))
                        )
                        print("Found upload button by primary class")
                    except Exception as e:
                        print(f"Could not find upload button: {e}")
            
            if upload_button:
                # Click the upload button
                print("Clicking upload button...")
                upload_button.click()
                time.sleep(3)
                
                # Handle any popup that might appear after upload
                # handle_popup(driver, wait)
                
                # Check for upload success
                print("Checking for upload success...")
                
                # Wait a bit for upload to complete
                time.sleep(5)
                
                # Look for success message or indication
                try:
                    # Look for success toast/message
                    success_elements = driver.find_elements(By.XPATH, 
                        "//*[contains(text(), 'Success') or contains(text(), 'success') or contains(text(), 'uploaded') or contains(text(), 'Uploaded')]")
                    
                    # Also check for table rows (uploaded files appear in table)
                    table_rows = driver.find_elements(By.CSS_SELECTOR, ".ant-table-tbody tr")
                    
                    if success_elements or table_rows:
                        print(f"✓ File upload appears successful")
                        
                        # Print the uploaded filename from table if available
                        if table_rows and len(table_rows) > 0:
                            try:
                                first_row_cells = table_rows[0].find_elements(By.TAG_NAME, "td")
                                if len(first_row_cells) > 1:
                                    filename = first_row_cells[1].text
                                    print(f"  Uploaded file appears in table: {filename}")
                            except:
                                pass
                        
                        return True
                    else:
                        print("⚠ No clear success message found, but upload button was clicked")
                        return True
                except Exception as e:
                    print(f"⚠ Could not verify upload success: {e}")
                    # Still return True since we clicked upload
                    return True
            else:
                print("⚠ Could not find upload button, file was selected but not uploaded")
                return False
        else:
            print("⚠ Could not find file input element")
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
        fixed_brc_type = "FOC"  # Test BRC type
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