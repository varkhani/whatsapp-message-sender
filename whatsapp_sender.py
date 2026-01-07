"""
Simple WhatsApp Message Sender using Selenium
Reads contact numbers and messages from XLSX file and sends WhatsApp messages
More reliable than pywhatkit for bulk messaging
"""

import openpyxl
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def read_contacts_from_excel(file_path):
    """
    Read contacts and messages from Excel file
    Expected format: 
    - Column A = Contact Number
    - Column B = Message
    """
    contacts = []
    
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        
        # Skip header row (if exists) - start from row 2
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0] and row[1]:  # Check if both contact and message exist
                contact = str(row[0]).strip()
                message = str(row[1]).strip()
                
                # Format contact number (remove spaces, ensure country code)
                contact = contact.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
                
                contacts.append({
                    'number': contact,
                    'message': message
                })
        
        workbook.close()
        print(f"✓ Loaded {len(contacts)} contacts from {file_path}")
        return contacts
    
    except Exception as e:
        print(f"✗ Error reading Excel file: {str(e)}")
        return []


def init_whatsapp_web(driver):
    """
    Initialize WhatsApp Web and wait for user to scan QR code
    """
    print("\n📱 Opening WhatsApp Web...")
    driver.get("https://web.whatsapp.com")
    
    print("\n⚠️  Please scan the QR code with your phone to log in to WhatsApp Web")
    print("   Waiting for you to complete login...")
    
    # Wait for the main chat list to appear (indicates successful login)
    try:
        # Wait for the search box or chat list to appear (sign of successful login)
        WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
        )
        print("✓ Successfully logged in to WhatsApp Web!")
        time.sleep(1)
        return True
    except TimeoutException:
        print("✗ Login timeout. Please try again.")
        return False


def ensure_main_page(driver):
    """
    Ensure we're on the main WhatsApp Web page (not in a chat)
    Only reloads if we're not on WhatsApp Web at all
    """
    try:
        current_url = driver.current_url
        if "web.whatsapp.com" not in current_url:
            # Only reload if we're not on WhatsApp Web at all
            driver.get("https://web.whatsapp.com")
            time.sleep(2)
        # Don't reload if we're already on WhatsApp Web, even if in a chat
        # We'll navigate back using the back button or clearing search
    except:
        pass

def go_back_to_main_page(driver):
    """
    Navigate back to main page from a chat without reloading
    """
    try:
        # Try to clear search box (this goes back to main view)
        try:
            search_box = driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='3']")
            search_box.click()
            time.sleep(0.2)
            search_box.send_keys(Keys.CONTROL + "a")
            time.sleep(0.1)
            search_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.5)
        except:
            # If search box not found, try pressing Escape
            try:
                from selenium.webdriver.common.action_chains import ActionChains
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.5)
            except:
                pass
    except:
        pass


def clear_attachment_preview(driver):
    """
    Clear any leftover attachment preview before sending next message
    """
    try:
        # Try to find and close attachment preview
        close_selectors = [
            "//span[@data-icon='close']",
            "//button[@aria-label='Close']",
            "//div[@role='button'][@aria-label='Close']",
            "//span[contains(@class, 'close')]"
        ]
        
        for selector in close_selectors:
            try:
                close_buttons = driver.find_elements(By.XPATH, selector)
                for btn in close_buttons:
                    if btn.is_displayed():
                        btn.click()
                        time.sleep(0.3)
                        return True
            except:
                continue
        
        # Try pressing Escape key as fallback
        try:
            message_input = driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='10'] | //div[@contenteditable='true'][@role='textbox']")
            message_input.send_keys(Keys.ESCAPE)
            time.sleep(0.3)
        except:
            pass
            
        return False
    except:
        return False


def send_image_with_caption(driver, message_box, image_path, caption, contact_number, delay_seconds):
    """
    Send image with caption in WhatsApp Web
    
    Args:
        driver: Selenium WebDriver instance
        message_box: Message input box element
        image_path: Path to image file (will be converted to absolute path)
        caption: Caption text to send with image
        contact_number: Contact number (for logging)
        delay_seconds: Delay after sending
    """
    def get_fresh_message_box():
        """Re-find message box to avoid stale element issues"""
        try:
            return driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='10'] | //div[@contenteditable='true'][@role='textbox']")
        except:
            return None
    
    def verify_chat_is_open():
        """Verify that we're actually in a chat conversation"""
        try:
            # Check if message box exists and is visible
            msg_box = get_fresh_message_box()
            if msg_box and msg_box.is_displayed():
                return True
            # Also check for chat header or other chat indicators
            chat_indicators = [
                "//div[@data-testid='conversation-header']",
                "//div[@data-testid='chatlist']",
                "//div[contains(@class, 'chat')]"
            ]
            for indicator in chat_indicators:
                try:
                    elem = driver.find_element(By.XPATH, indicator)
                    if elem.is_displayed():
                        return True
                except:
                    continue
            return False
        except:
            return False
    
    def send_text_fallback():
        """Fallback to send text only"""
        print(f"  → Attempting to send text message...")
        try:
            # First verify chat is open
            if not verify_chat_is_open():
                print(f"  ⚠️  Chat is not open, cannot send message")
                return False
            
            fresh_box = get_fresh_message_box()
            if not fresh_box:
                print(f"  ⚠️  Could not find message box")
                return False
            
            fresh_box.click()
            time.sleep(0.3)
            fresh_box.send_keys(caption)
            time.sleep(0.3)
            fresh_box.send_keys(Keys.ENTER)
            time.sleep(1)  # Wait for message to send
            
            # Verify message was sent by checking if input box is cleared
            time.sleep(0.5)
            print(f"✓ Text message sent to {contact_number}")
            return True
        except Exception as e:
            print(f"  ⚠️  Could not send text fallback: {str(e)}")
        return False
    
    try:
        # Step 0: Verify chat is actually open
        print(f"  → Verifying chat is open...")
        if not verify_chat_is_open():
            print(f"  ⚠️  Chat is not open, cannot send message")
            return False
        
        # Step 1: Validate and prepare image path
        if not os.path.isabs(image_path):
            image_path = os.path.abspath(image_path)
        
        if not os.path.exists(image_path):
            print(f"  ⚠️  Image file not found: {image_path}, sending text only")
            return send_text_fallback()
        
        # Step 2: Wait for chat to be fully loaded, then find attachment button
        print(f"  → Waiting for chat to fully load...")
        time.sleep(2)  # Give chat more time to fully load
        
        print(f"  → Looking for attachment button...")
        attachment_button = None
        attachment_selectors = [
            "//span[@data-testid='clip']",
            "//div[@data-testid='clip']",
            "//span[@data-icon='attach']",
            "//button[@title='Attach']",
            "//button[@aria-label='Attach']"
        ]
        
        # Try multiple times with delays
        for attempt in range(5):
            for selector in attachment_selectors:
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        try:
                            if elem.is_displayed() and elem.is_enabled():
                                attachment_button = elem
                                break
                        except:
                            continue
                    if attachment_button:
                        break
                except:
                    continue
            if attachment_button:
                break
            time.sleep(0.5)  # Wait a bit and try again
        
        if not attachment_button:
            print(f"  ⚠️  Could not find attachment button, sending text only")
            print(f"      Make sure the chat is open and you're in a conversation")
            return send_text_fallback()
        
        # Step 3: Click attachment button
        print(f"  → Clicking attachment button...")
        try:
            driver.execute_script("arguments[0].click();", attachment_button)
        except:
            try:
                attachment_button.click()
            except:
                print(f"  ⚠️  Could not click attachment button, sending text only")
                return send_text_fallback()
        
        time.sleep(1.5)  # Wait for file picker menu to appear
        
        # Step 4: Find file input element
        print(f"  → Looking for file input...")
        file_input = None
        try:
            file_inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
            for inp in file_inputs:
                if inp.is_displayed() or True:  # File inputs are usually hidden
                    file_input = inp
                    break
        except:
            pass
        
        if not file_input:
            print(f"  ⚠️  Could not find file input, sending text only")
            return send_text_fallback()
        
        # Step 5: Upload image file
        print(f"  → Uploading image...")
        try:
            file_input.send_keys(image_path)
            time.sleep(5)  # Wait for image to load and preview to appear
        except Exception as e:
            print(f"  ⚠️  Could not upload image: {str(e)}, sending text only")
            return send_text_fallback()
        
        # Step 6: Find caption input box (appears after image is selected)
        print(f"  → Looking for caption box...")
        caption_box = None
        caption_selectors = [
            "//div[@contenteditable='true'][@data-tab='11']",  # Most specific - caption box when image attached
            "//div[@contenteditable='true'][@data-testid='media-caption-input-container']",
            "//div[@contenteditable='true'][contains(@placeholder, 'caption')]",
            "//div[@contenteditable='true'][contains(@placeholder, 'Caption')]",
            "//div[@contenteditable='true'][contains(@placeholder, 'Add a caption')]"
        ]
        
        # Try to find caption box - wait up to 8 seconds
        for attempt in range(4):
            for selector in caption_selectors:
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        if elem.is_displayed():
                            data_tab = elem.get_attribute('data-tab')
                            placeholder = str(elem.get_attribute('placeholder') or '').lower()
                            # Verify it's the caption box (not regular message box)
                            if data_tab == '11' or 'caption' in placeholder or 'add' in placeholder:
                                caption_box = elem
                                break
                    if caption_box:
                        break
                except:
                    continue
            if caption_box:
                break
            time.sleep(1)  # Wait a bit and try again
        
        # Step 7: Type caption if caption box found and caption text provided
        if caption_box and caption:
            print(f"  → Typing caption...")
            try:
                caption_box.click()
                time.sleep(0.3)
                # Clear any existing text
                caption_box.send_keys(Keys.CONTROL + "a")
                time.sleep(0.1)
                caption_box.send_keys(Keys.BACKSPACE)
                time.sleep(0.2)
                # Type caption
                caption_box.send_keys(caption)
                time.sleep(0.5)
            except Exception as e:
                print(f"  ⚠️  Could not type caption: {str(e)} (will try to send image without caption)")
        elif caption:
            print(f"  ⚠️  Caption box not found, will send image without caption")
        
        # Step 8: Find and click send button
        print(f"  → Looking for send button...")
        sent = False
        send_selectors = [
            "//span[@data-testid='send']",
            "//span[@data-icon='send']",
            "//button[@aria-label='Send']",
            "//div[@data-testid='send']"
        ]
        
        # Wait for send button to appear (up to 5 seconds)
        for attempt in range(5):
            for selector in send_selectors:
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        if elem.is_displayed() and elem.is_enabled():
                            try:
                                driver.execute_script("arguments[0].click();", elem)
                                time.sleep(2)
                                sent = True
                                break
                            except:
                                try:
                                    elem.click()
                                    time.sleep(2)
                                    sent = True
                                    break
                                except:
                                    continue
                    if sent:
                        break
                except:
                    continue
            if sent:
                break
            time.sleep(1)
        
        # Step 9: Alternative methods if send button not found
        if not sent:
            print(f"  → Trying alternative send methods...")
            # Try pressing Enter in caption box
            if caption_box:
                try:
                    caption_box.click()
                    time.sleep(0.2)
                    caption_box.send_keys(Keys.ENTER)
                    time.sleep(2)
                    sent = True
                except:
                    pass
            
            # Try pressing Enter in active element
            if not sent:
                try:
                    active_input = driver.switch_to.active_element
                    active_input.send_keys(Keys.ENTER)
                    time.sleep(2)
                    sent = True
                except:
                    pass
        
        # Step 10: Verify and cleanup
        if sent:
            time.sleep(2)  # Wait for message to actually send
            clear_attachment_preview(driver)
            time.sleep(1)
            print(f"✓ Image with caption sent to {contact_number}")
            time.sleep(delay_seconds)
            return True
        else:
            print(f"  ⚠️  Could not send image (send button not found), sending text only")
            return send_text_fallback()
        
    except Exception as e:
        print(f"  ⚠️  Error sending image: {str(e)}, sending text only")
        return send_text_fallback()


def send_whatsapp_message(driver, contact_number, message, delay_seconds=15):
    """
    Send WhatsApp message using search (much faster than URL navigation)
    
    Args:
        driver: Selenium WebDriver instance
        contact_number: Contact number in international format
        message: Text message to send
        delay_seconds: Delay between messages
    
    Note: The contact number should:
    - Be in international format (e.g., +919555611880 for India)
    - Be saved in your WhatsApp contacts (recommended)
    """
    try:
        # Format number for search (keep + for better matching)
        search_query = contact_number.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
        
        # Validate number format
        clean_number = search_query.replace("+", "")
        if not clean_number.isdigit() or len(clean_number) < 10 or len(clean_number) > 15:
            print(f"  ⚠️  Invalid number format (should be 10-15 digits): {contact_number}")
            return False
        
        # Method 1: Use WhatsApp Web search (FAST - no page reload)
        try:
            # Find and click the search box
            search_box = None
            search_selectors = [
                "//div[@contenteditable='true'][@data-tab='3']",
                "//div[@contenteditable='true'][@role='textbox'][@data-tab='3']",
                "//div[@contenteditable='true'][@title='Search input textbox']",
                "//div[@contenteditable='true'][contains(@class, 'selectable-text')]"
            ]
            
            for selector in search_selectors:
                try:
                    search_box = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    break
                except TimeoutException:
                    continue
            
            if not search_box:
                # Fallback: Try to find search by clicking on search area
                try:
                    search_area = driver.find_element(By.XPATH, "//div[@data-testid='chat-list-search']")
                    search_area.click()
                    time.sleep(0.5)
                    search_box = driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='3']")
                except:
                    pass
            
            if search_box:
                # Clear search box and type contact number
                search_box.click()
                time.sleep(0.2)
                # Clear any existing text
                search_box.send_keys(Keys.CONTROL + "a")
                time.sleep(0.1)
                search_box.send_keys(Keys.BACKSPACE)
                time.sleep(0.1)
                # Type the search query
                search_box.send_keys(search_query)
                time.sleep(1.2)  # Wait for search results to load
                
                # Click on first search result - try multiple methods
                clicked = False
                
                # Method 1: Use keyboard navigation (most reliable)
                try:
                    # Press Arrow Down to highlight first result, then Enter to select
                    search_box.send_keys(Keys.ARROW_DOWN)
                    time.sleep(0.2)
                    search_box.send_keys(Keys.ENTER)
                    time.sleep(1)  # Wait for chat to open
                    clicked = True
                except Exception as e:
                    pass
                
                # Method 2: Try clicking the first list item container
                if not clicked:
                    try:
                        first_result = WebDriverWait(driver, 3).until(
                            EC.element_to_be_clickable((
                                By.XPATH, 
                                "//div[@role='listitem'][1] | " +
                                "//div[@data-testid='cell-frame-container'][1]"
                            ))
                        )
                        # Try regular click first
                        first_result.click()
                        time.sleep(1)
                        clicked = True
                    except (TimeoutException, Exception) as e:
                        # Try JavaScript click
                        try:
                            list_items = driver.find_elements(By.XPATH, "//div[@role='listitem']")
                            if list_items:
                                driver.execute_script("arguments[0].click();", list_items[0])
                                time.sleep(1)
                                clicked = True
                        except:
                            pass
                
                # Method 3: Try clicking by finding any clickable element in first result
                if not clicked:
                    try:
                        # Find first result container
                        result_containers = driver.find_elements(
                            By.XPATH, 
                            "//div[@role='listitem'][1]//* | " +
                            "//div[@data-testid='cell-frame-container'][1]//*"
                        )
                        for element in result_containers[:5]:  # Try first 5 elements
                            try:
                                if element.is_displayed():
                                    element.click()
                                    time.sleep(1)
                                    clicked = True
                                    break
                            except:
                                continue
                    except Exception as e:
                        pass
                
                if not clicked:
                    # If no search result found, contact might not be in WhatsApp
                    print(f"  ⚠️  Contact not found or couldn't select: {contact_number}")
                    print(f"      Make sure the contact is saved in your WhatsApp contacts")
                    # Clear search
                    search_box.send_keys(Keys.CONTROL + "a")
                    search_box.send_keys(Keys.BACKSPACE)
                    time.sleep(0.5)
                    return False
                
                # Verify chat opened by checking if message input box appears
                time.sleep(0.8)  # Give it a moment to open
                
                # Now find message input box
                message_box = None
                message_selectors = [
                    "//div[@contenteditable='true'][@data-tab='10']",
                    "//div[@contenteditable='true'][@role='textbox']",
                    "//div[@contenteditable='true'][@data-testid='conversation-compose-box-input']",
                    "//div[@contenteditable='true'][contains(@class, 'selectable-text')]",
                    "//footer//div[@contenteditable='true']",
                    "//div[@contenteditable='true'][@spellcheck='true']"
                ]
                
                for selector in message_selectors:
                    try:
                        message_box = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, selector))
                        )
                        # Verify it's actually visible and usable
                        if message_box.is_displayed():
                            break
                    except TimeoutException:
                        continue
                    except Exception:
                        continue
                
                if message_box and message_box.is_displayed():
                    # Send text message
                    try:
                        message_box.click()
                        time.sleep(0.2)
                        message_box.send_keys(message)
                        time.sleep(0.3)
                        message_box.send_keys(Keys.ENTER)
                        time.sleep(0.5)  # Wait for message to send
                        
                        print(f"✓ Message sent to {contact_number}")
                        time.sleep(delay_seconds)
                        # Go back to main page for next contact
                        go_back_to_main_page(driver)
                        return True
                    except Exception as e:
                        print(f"  ⚠️  Error sending message: {str(e)}")
                        return False
                else:
                    # Chat might not have opened, try clearing search and going back
                    print(f"  ⚠️  Chat did not open properly for: {contact_number}")
                    print(f"      Make sure the contact exists and chat can be opened")
                    # Clear search to go back to main view
                    try:
                        search_box = driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='3']")
                        search_box.click()
                        time.sleep(0.2)
                        search_box.send_keys(Keys.CONTROL + "a")
                        time.sleep(0.1)
                        search_box.send_keys(Keys.BACKSPACE)
                        time.sleep(0.5)
                    except:
                        pass
                    return False
            else:
                # Fallback to URL method if search doesn't work
                raise Exception("Search box not found, using URL method")
                
        except Exception as search_error:
            # Fallback: Use URL method if search fails
            print(f"  ⚠️  Search method failed, trying URL method...")
            clean_number = contact_number.replace("+", "").replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
            chat_url = f"https://web.whatsapp.com/send?phone={clean_number}"
            driver.get(chat_url)
            time.sleep(2)
            
            # Find message box
            message_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='10'] | //div[@contenteditable='true'][@role='textbox']"))
            )
            
            # Send text message
            message_box.click()
            time.sleep(0.2)
            message_box.send_keys(message)
            time.sleep(0.3)
            message_box.send_keys(Keys.ENTER)
            time.sleep(0.5)
            print(f"✓ Message sent to {contact_number} (via URL)")
            time.sleep(delay_seconds)
            go_back_to_main_page(driver)
            return True
    
    except TimeoutException:
        print(f"  ⚠️  Timeout: {contact_number}")
        return False
    except Exception as e:
        print(f"✗ Error sending message to {contact_number}: {str(e)}")
        return False


def find_image_file(excel_file_path):
    """
    Find image file in the same directory as Excel file
    Looks for common image formats: jpg, jpeg, png, gif, webp
    """
    excel_dir = os.path.dirname(os.path.abspath(excel_file_path))
    if not excel_dir:
        excel_dir = os.getcwd()
    
    image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.webp']
    
    # Look for image files in the same directory
    for file in os.listdir(excel_dir):
        file_lower = file.lower()
        if any(file_lower.endswith(ext) for ext in image_extensions):
            image_path = os.path.join(excel_dir, file)
            return image_path
    
    return None


def send_bulk_messages(excel_file_path, delay_seconds=15, start_from=0):
    """
    Send messages to all contacts in the Excel file using Selenium
    
    Args:
        excel_file_path: Path to the XLSX file
        delay_seconds: Delay between each message (to avoid rate limiting)
        start_from: Index to start from (useful for resuming)
    """
    contacts = read_contacts_from_excel(excel_file_path)
    
    if not contacts:
        print("No contacts found. Please check your Excel file.")
        return
    
    # Setup Chrome driver
    print("\n🔧 Setting up Chrome browser...")
    driver = None
    
    try:
        chrome_options = webdriver.ChromeOptions()
        
        # Use absolute path for user data directory (prevents crashes)
        import platform
        profile_path = os.path.abspath("./chrome_profile")
        chrome_options.add_argument(f"--user-data-dir={profile_path}")
        
        # Essential stability options
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--remote-debugging-port=9222")
        
        # Experimental options
        chrome_options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        chrome_options.add_experimental_option("detach", True)  # Keep browser open
        
        # Prefs to prevent crashes
        prefs = {
            "profile.default_content_setting_values.notifications": 2,
            "profile.default_content_settings.popups": 0,
            "profile.managed_default_content_settings.images": 1
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        # Try to set Chrome binary path explicitly (helps with some Windows issues)
        if platform.system() == 'Windows':
            chrome_paths = [
                r'C:\Program Files\Google\Chrome\Application\chrome.exe',
                r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
                os.path.expanduser(r'~\AppData\Local\Google\Chrome\Application\chrome.exe')
            ]
            for chrome_path in chrome_paths:
                if os.path.exists(chrome_path):
                    chrome_options.binary_location = chrome_path
                    print(f"   Using Chrome at: {chrome_path}")
                    break
        
        # Initialize driver with better error handling
        try:
            # Try with ChromeDriverManager first
            print("   Downloading/updating ChromeDriver...")
            driver_path = ChromeDriverManager().install()
            
            # Verify the driver file exists and is valid
            if not os.path.exists(driver_path):
                raise Exception(f"ChromeDriver not found at: {driver_path}")
            
            service = Service(driver_path)
            driver = webdriver.Chrome(service=service, options=chrome_options)
            print("   ✓ Chrome browser initialized successfully!")
        except Exception as e:
            print(f"   ⚠️  Error with ChromeDriverManager: {str(e)}")
            print("   Trying alternative method (without explicit service)...")
            try:
                # Try without explicit service (let Selenium find it)
                driver = webdriver.Chrome(options=chrome_options)
                print("   ✓ Chrome browser initialized successfully!")
            except Exception as e2:
                print(f"   ✗ Failed to initialize Chrome: {str(e2)}")
                print("\n   Troubleshooting steps:")
                print("   1. Close all Chrome browser windows and try again")
                print("   2. Make sure Chrome browser is installed and up to date")
                print("   3. Delete Chrome profile and cache:")
                print(f"      Remove-Item -Recurse -Force .\\chrome_profile")
                print(f"      Remove-Item -Recurse -Force $env:USERPROFILE\\.wdm")
                print("   4. Check if antivirus is blocking ChromeDriver")
                print("   5. Try restarting your computer")
                print("   6. Or manually download ChromeDriver from:")
                print("      https://chromedriver.chromium.org/")
                raise
    
    except Exception as e:
        print(f"\n✗ Could not start Chrome browser: {str(e)}")
        return
    
    try:
        # Initialize WhatsApp Web
        if not init_whatsapp_web(driver):
            print("Failed to initialize WhatsApp Web. Exiting.")
            return
        
        # Ensure we start from main page
        ensure_main_page(driver)
        time.sleep(1)
        
        print(f"\n📱 Starting to send messages to {len(contacts)} contacts...")
        print(f"⏱️  Delay between messages: {delay_seconds} seconds")
        print(f"⚠️  Keep the browser window open and don't close it!\n")
        
        successful = 0
        failed = 0
        
        for index, contact in enumerate(contacts[start_from:], start=start_from):
            print(f"\n[{index + 1}/{len(contacts)}] Sending to {contact['number']}...")
            
            if send_whatsapp_message(driver, contact['number'], contact['message'], delay_seconds):
                successful += 1
            else:
                failed += 1
            
            # Progress update every 10 messages
            if (index + 1) % 10 == 0:
                print(f"\n📊 Progress: {index + 1}/{len(contacts)} | ✓ {successful} | ✗ {failed}")
        
        print(f"\n{'='*50}")
        print(f"✅ Completed!")
        print(f"✓ Successful: {successful}")
        print(f"✗ Failed: {failed}")
        print(f"{'='*50}")
        
    except KeyboardInterrupt:
        print("\n\n⚠️  Process interrupted by user")
        print(f"   You can resume from index {start_from + successful + failed} next time")
    except Exception as e:
        print(f"\n✗ Error: {str(e)}")
    finally:
        if driver:
            print("\n⚠️  Closing browser in 5 seconds... (Press Ctrl+C to keep it open)")
            try:
                time.sleep(5)
                driver.quit()
            except:
                pass


if __name__ == "__main__":
    # Configuration
    EXCEL_FILE = "contacts.xlsx"  # Change this to your Excel file name
    DELAY_SECONDS = 5  # Delay between messages (increase if you get rate limited, minimum 3-5 seconds recommended)
    START_FROM = 0  # Start from this index (useful if you need to resume)
    
    # Check if Excel file exists
    if not os.path.exists(EXCEL_FILE):
        print(f"✗ Error: {EXCEL_FILE} not found!")
        print(f"Please create an Excel file with:")
        print(f"  Column A: Contact Number (with country code, e.g., +1234567890)")
        print(f"  Column B: Message")
        exit(1)
    
    # Confirm before starting
    print("="*50)
    print("WhatsApp Bulk Message Sender")
    print("="*50)
    print(f"File: {EXCEL_FILE}")
    print(f"Delay: {DELAY_SECONDS} seconds between messages")
    print("\n⚠️  Make sure:")
    print("  1. You have Chrome browser installed")
    print("  2. Your computer won't go to sleep")
    print("  3. You have your phone nearby to scan QR code")
    print("  4. Browser window will stay open during the process")
    print("\nPress Enter to start, or Ctrl+C to cancel...")
    
    try:
        input()
        send_bulk_messages(EXCEL_FILE, DELAY_SECONDS, START_FROM)
    except KeyboardInterrupt:
        print("\n\n⚠️  Process cancelled by user")
    except Exception as e:
        print(f"\n✗ Error: {str(e)}")
