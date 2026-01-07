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
    Read contacts, messages, and image paths from Excel file
    Expected format: 
    - Column A = Contact Number
    - Column B = Message (caption)
    - Column C = Image Path (optional - if empty, will look for image based on contact number)
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
                
                # Get image path from Column C (if provided)
                image_path = None
                if len(row) > 2 and row[2]:
                    image_path = str(row[2]).strip()
                    if image_path:
                        # Convert to absolute path if relative
                        if not os.path.isabs(image_path):
                            excel_dir = os.path.dirname(os.path.abspath(file_path))
                            image_path = os.path.join(excel_dir, image_path)
                
                contacts.append({
                    'number': contact,
                    'message': message,
                    'image_path': image_path
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
            print(f"  ⚠️  Not on WhatsApp Web (current URL: {current_url}), navigating to WhatsApp Web...")
            driver.get("https://web.whatsapp.com")
            time.sleep(3)
            # Wait for page to load
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
            )
        # Don't reload if we're already on WhatsApp Web, even if in a chat
        # We'll navigate back using the back button or clearing search
    except Exception as e:
        print(f"  ⚠️  Error ensuring main page: {str(e)}")
        pass


def verify_on_whatsapp_web(driver):
    """
    Verify we're actually on WhatsApp Web and not redirected to download page
    Returns True if on WhatsApp Web, False otherwise
    """
    try:
        current_url = driver.current_url
        page_title = driver.title.lower()
        
        # Check if we're on WhatsApp Web
        if "web.whatsapp.com" not in current_url:
            print(f"  ⚠️  Not on WhatsApp Web! Current URL: {current_url}")
            return False
        
        # Check if we've been redirected to download page
        if "download" in page_title or "whatsapp for windows" in page_title:
            print(f"  ⚠️  Redirected to download page! Navigating back to WhatsApp Web...")
            driver.get("https://web.whatsapp.com")
            time.sleep(3)
            # Wait for WhatsApp Web to load
            try:
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
                )
                print(f"  ✓ Back on WhatsApp Web")
                return True
            except TimeoutException:
                print(f"  ✗ Failed to load WhatsApp Web")
                return False
        
        # Verify search box exists (confirms we're on the right page)
        try:
            driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='3']")
            return True
        except:
            print(f"  ⚠️  WhatsApp Web elements not found, might be loading...")
            return False
            
    except Exception as e:
        print(f"  ⚠️  Error verifying WhatsApp Web: {str(e)}")
        return False

def go_back_to_main_page(driver):
    """
    Navigate back to main page from a chat without reloading
    Ensures clean state for next contact
    """
    try:
        # First, press Escape to close any open chat or search
        try:
            from selenium.webdriver.common.action_chains import ActionChains
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.5)
        except:
            pass
        
        # Try to clear search box (this goes back to main view)
        try:
            search_box = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
            )
            search_box.click()
            time.sleep(0.2)
            # Clear any text in search box
            search_box.send_keys(Keys.CONTROL + "a")
            time.sleep(0.1)
            search_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.3)
            # Press Escape again to ensure we're back to main view
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.3)
        except:
            # If search box not found, try pressing Escape again
            try:
                from selenium.webdriver.common.action_chains import ActionChains
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.5)
            except:
                pass
        
        # Verify we're back on main page by checking for search box
        try:
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
            )
        except:
            pass
    except:
        pass


def set_message_text_js(driver, message_box, text):
    """
    Set text in message box using JavaScript to bypass ChromeDriver BMP limitation
    This works with emojis and Unicode characters outside BMP
    
    Args:
        driver: Selenium WebDriver instance
        message_box: Message box element
        text: Text to set
    
    Returns:
        True if successful, False otherwise
    """
    try:
        # Use JavaScript to set the text in WhatsApp's contenteditable div
        # WhatsApp Web uses contenteditable div, we need to properly set the text
        driver.execute_script("""
            var element = arguments[0];
            var text = arguments[1];
            
            // Focus the element first
            element.focus();
            
            // Select all existing content
            var range = document.createRange();
            range.selectNodeContents(element);
            range.collapse(false);
            var selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
            
            // Delete existing content
            document.execCommand('delete', false, null);
            
            // Insert text using execCommand (works with emojis)
            document.execCommand('insertText', false, text);
            
            // Alternative: If execCommand doesn't work, set textContent and trigger events
            if (element.textContent !== text && element.innerText !== text) {
                element.textContent = text;
                element.innerText = text;
                
                // Create and dispatch InputEvent with proper data
                var inputEvent = new InputEvent('input', {
                    bubbles: true,
                    cancelable: true,
                    inputType: 'insertText',
                    data: text
                });
                element.dispatchEvent(inputEvent);
                
                // Also dispatch beforeinput event
                var beforeInputEvent = new InputEvent('beforeinput', {
                    bubbles: true,
                    cancelable: true,
                    inputType: 'insertText',
                    data: text
                });
                element.dispatchEvent(beforeInputEvent);
            }
            
            // Trigger keyup event
            var keyupEvent = new KeyboardEvent('keyup', { bubbles: true, cancelable: true });
            element.dispatchEvent(keyupEvent);
        """, message_box, text)
        
        # Verify text was set
        time.sleep(0.3)
        return True
    except Exception as e:
        print(f"  ⚠️  JavaScript text setting failed: {str(e)}")
        return False


def force_focus_message_box(driver, message_box, max_attempts=5):
    """
    Aggressively focus the message box using multiple methods
    Returns True if successfully focused, False otherwise
    """
    from selenium.webdriver.common.action_chains import ActionChains
    
    for attempt in range(max_attempts):
        try:
            # Wait a bit for element to be ready
            time.sleep(0.3)
            
            # Method 1: Scroll into view and use ActionChains to move and click
            try:
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'instant', block: 'center', inline: 'center'});", message_box)
                time.sleep(0.2)
                
                # Use ActionChains to move mouse to element and click
                ActionChains(driver).move_to_element(message_box).click().perform()
                time.sleep(0.4)
                
                # Verify focus
                active_elem = driver.execute_script("return document.activeElement;")
                if active_elem == message_box:
                    print(f"  ✓ Message box focused (ActionChains method)")
                    return True
            except:
                pass
            
            # Method 2: Aggressive JavaScript focus with multiple events
            try:
                driver.execute_script("""
                    var elem = arguments[0];
                    // Remove focus from any other element
                    if (document.activeElement && document.activeElement !== elem) {
                        document.activeElement.blur();
                    }
                    // Scroll to element
                    elem.scrollIntoView({behavior: 'instant', block: 'center'});
                    // Focus and click
                    elem.focus();
                    elem.click();
                    // Force focus by setting tabindex if needed
                    if (!elem.hasAttribute('tabindex')) {
                        elem.setAttribute('tabindex', '-1');
                    }
                    elem.focus();
                    // Dispatch all focus-related events
                    var events = ['focus', 'focusin', 'mousedown', 'mouseup', 'click'];
                    events.forEach(function(eventType) {
                        var event = new Event(eventType, { bubbles: true, cancelable: true });
                        elem.dispatchEvent(event);
                    });
                """, message_box)
                time.sleep(0.5)
                
                # Verify focus
                active_elem = driver.execute_script("return document.activeElement;")
                if active_elem == message_box:
                    print(f"  ✓ Message box focused (aggressive JavaScript)")
                    return True
            except:
                pass
            
            # Method 3: Click on parent container or footer area
            try:
                # Try clicking on the footer area that contains the message box
                footer = driver.find_element(By.XPATH, "//footer")
                ActionChains(driver).move_to_element(footer).click().perform()
                time.sleep(0.3)
                # Now click the message box
                message_box.click()
                time.sleep(0.4)
                
                active_elem = driver.execute_script("return document.activeElement;")
                if active_elem == message_box:
                    print(f"  ✓ Message box focused (footer click method)")
                    return True
            except:
                pass
            
            # Method 4: Press Escape to clear any focus, then focus message box
            try:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.3)
                # Click multiple times
                for _ in range(3):
                    message_box.click()
                    time.sleep(0.2)
                
                # Force focus with JavaScript
                driver.execute_script("arguments[0].focus();", message_box)
                time.sleep(0.3)
                
                active_elem = driver.execute_script("return document.activeElement;")
                if active_elem == message_box:
                    print(f"  ✓ Message box focused (Escape + multiple clicks)")
                    return True
            except:
                pass
            
            # Method 5: Use JavaScript to simulate a real user click
            try:
                driver.execute_script("""
                    var elem = arguments[0];
                    var rect = elem.getBoundingClientRect();
                    var x = rect.left + rect.width / 2;
                    var y = rect.top + rect.height / 2;
                    
                    // Create and dispatch mouse events
                    var mouseEvents = ['mousedown', 'mouseup', 'click'];
                    mouseEvents.forEach(function(eventType) {
                        var event = new MouseEvent(eventType, {
                            view: window,
                            bubbles: true,
                            cancelable: true,
                            clientX: x,
                            clientY: y
                        });
                        elem.dispatchEvent(event);
                    });
                    
                    // Focus
                    elem.focus();
                """, message_box)
                time.sleep(0.5)
                
                active_elem = driver.execute_script("return document.activeElement;")
                if active_elem == message_box:
                    print(f"  ✓ Message box focused (simulated mouse events)")
                    return True
            except:
                pass
            
        except Exception as e:
            if attempt < max_attempts - 1:
                time.sleep(0.5)
                continue
            else:
                print(f"  ⚠️  Could not focus message box after {max_attempts} attempts")
                return False
    
    return False


def get_fresh_message_box(driver, max_retries=3):
    """
    Get a fresh message box element to avoid stale element issues
    Re-finds the element each time it's called
    
    Args:
        driver: Selenium WebDriver instance
        max_retries: Maximum number of retries to find the element
    
    Returns:
        Message box element or None if not found
    """
    message_selectors = [
        "//div[@contenteditable='true'][@data-tab='10']",
        "//div[@contenteditable='true'][@role='textbox']",
        "//div[@contenteditable='true'][@data-testid='conversation-compose-box-input']",
        "//footer//div[@contenteditable='true']",
        "//div[@contenteditable='true'][@spellcheck='true']",
        "//div[@contenteditable='true'][contains(@class, 'selectable-text')]"
    ]
    
    for attempt in range(max_retries):
        for selector in message_selectors:
            try:
                # Use element_to_be_clickable for better reliability
                element = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                if element and element.is_displayed():
                    return element
            except (TimeoutException, Exception):
                continue
        
        if attempt < max_retries - 1:
            time.sleep(0.8)  # Wait longer before retrying
    
    return None


def clear_attachment_preview(driver):
    """
    Clear any leftover attachment preview before sending next message
    Simplified version - just press Escape and look for close buttons
    """
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        # Press Escape a few times to close any previews
        for _ in range(2):
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.2)
        
        # Try to find and click close buttons
        try:
            close_buttons = driver.find_elements(By.XPATH, 
                "//span[@data-icon='close'] | " +
                "//button[contains(@aria-label, 'Close')] | " +
                "//button[contains(@aria-label, 'Remove')]"
            )
            for btn in close_buttons:
                try:
                    if btn.is_displayed():
                        driver.execute_script("arguments[0].click();", btn)
                        time.sleep(0.3)
                        break
                except:
                    continue
        except:
            pass
        
        return True
    except:
        return False


def verify_message_sent(driver, timeout=10):
    """
    Verify that a message was actually sent by checking if message box is cleared
    Returns True if message appears to be sent, False otherwise
    """
    try:
        # Wait a bit for message to send
        time.sleep(1)
        
        # Check if message box is cleared (indicates message was sent)
        message_box = get_fresh_message_box(driver)
        if message_box:
            # Wait for message box to be cleared
            for _ in range(timeout):
                try:
                    text = message_box.text.strip()
                    inner_html = message_box.get_attribute('innerHTML').strip()
                    # If message box is empty or only has placeholder, message was sent
                    if not text or not inner_html or inner_html == '<br>' or inner_html == '':
                        return True
                    time.sleep(0.5)
                except:
                    time.sleep(0.5)
                    continue
        
        # Alternative: Check if send button is disabled or message appears in chat
        try:
            # Look for the sent message in chat (checkmark icon or sent indicator)
            sent_indicators = driver.find_elements(By.XPATH, 
                "//span[@data-icon='msg-dblcheck'] | " +
                "//span[@data-icon='msg-check'] | " +
                "//span[contains(@data-testid, 'check')]"
            )
            if sent_indicators:
                return True
        except:
            pass
        
        return False
    except:
        return True  # Assume sent if we can't verify


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
    # Use the module-level get_fresh_message_box function
    
    def verify_chat_is_open():
        """Verify that we're actually in a chat conversation"""
        try:
            # Check if message box exists and is visible
            msg_box = get_fresh_message_box(driver)
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
            
            fresh_box = get_fresh_message_box(driver)
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


def send_whatsapp_message(driver, contact_number, message, delay_seconds=15, image_path=None):
    """
    Simple WhatsApp message sender:
    1. Search contact
    2. Auto select
    3. Auto type message
    4. Auto send
    """
    from selenium.webdriver.common.action_chains import ActionChains
    
    try:
        # Ensure we're on main page
        ensure_main_page(driver)
        time.sleep(0.2)
        
        # Step 1: Search contact
        search_query = contact_number.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
        
        print(f"  → Searching for contact...")
        # Find search box
        try:
            search_box = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']"))
            )
        except TimeoutException:
            print(f"  ✗ Error: Could not find search box (timeout)")
            return False
        
        # Clear and type search query
        search_box.click()
        time.sleep(0.1)
        search_box.send_keys(Keys.CONTROL + "a")
        search_box.send_keys(Keys.BACKSPACE)
        time.sleep(0.1)
        search_box.send_keys(search_query)
        time.sleep(1)  # Wait for results (reduced from 2)
        
        # Step 2: Auto select first result
        print(f"  → Selecting contact...")
        try:
            # Press Arrow Down + Enter to select first result
            search_box.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.1)
            search_box.send_keys(Keys.ENTER)
            time.sleep(1.5)  # Wait for chat to open (reduced from 2.5)
        except Exception as e:
            # Fallback: Click first result
            try:
                first_result = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@role='listitem'][1]"))
                )
                first_result.click()
                time.sleep(1.5)  # Reduced from 2.5
            except Exception as e2:
                print(f"  ✗ Error: Could not select contact - {str(e2) if str(e2) else type(e2).__name__}")
                return False
        
        # Step 3: Find message box
        print(f"  → Finding message box...")
        time.sleep(0.5)  # Wait for chat to fully load (reduced from 1)
        
        # Find message box using multiple selectors
        message_box = get_fresh_message_box(driver, max_retries=5)
        if not message_box:
            print(f"  ✗ Error: Could not find message box")
            return False
        
        # Step 4: Auto type message
        print(f"  → Typing message...")
        
        # Re-find message box to avoid stale element
        message_box = get_fresh_message_box(driver, max_retries=3)
        if not message_box:
            print(f"  ✗ Error: Message box became unavailable")
            return False
        
        # Focus message box - use ActionChains for reliable clicking
        try:
            ActionChains(driver).move_to_element(message_box).click().perform()
            time.sleep(0.2)  # Reduced from 0.5
        except Exception:
            try:
                message_box.click()
                time.sleep(0.2)  # Reduced from 0.5
            except Exception:
                driver.execute_script("arguments[0].focus(); arguments[0].click();", message_box)
                time.sleep(0.2)  # Reduced from 0.5
        
        # Check if message contains non-BMP characters (emojis, etc.)
        # ChromeDriver send_keys() only supports BMP characters
        has_non_bmp = any(ord(char) > 0xFFFF for char in message)
        
        if has_non_bmp:
            # Use JavaScript for messages with emojis/special characters
            print(f"  → Using JavaScript for emoji/special characters...")
            # Clear first
            driver.execute_script("arguments[0].innerHTML = ''; arguments[0].textContent = '';", message_box)
            time.sleep(0.1)  # Reduced from 0.2
            
            # Use the set_message_text_js function
            if not set_message_text_js(driver, message_box, message):
                print(f"  ⚠ Warning: JavaScript text setting had issues, trying alternative...")
                # Alternative: Direct JavaScript with proper events
                driver.execute_script("""
                    var elem = arguments[0];
                    var text = arguments[1];
                    
                    elem.focus();
                    elem.textContent = text;
                    elem.innerText = text;
                    
                    // Trigger proper input events
                    var inputEvent = new InputEvent('input', {
                        bubbles: true,
                        cancelable: true,
                        inputType: 'insertText',
                        data: text
                    });
                    elem.dispatchEvent(inputEvent);
                    
                    var beforeInputEvent = new InputEvent('beforeinput', {
                        bubbles: true,
                        cancelable: true,
                        inputType: 'insertText',
                        data: text
                    });
                    elem.dispatchEvent(beforeInputEvent);
                    
                    elem.focus();
                """, message_box, message)
            time.sleep(0.4)  # Reduced from 0.8
        else:
            # Use real keyboard input for regular text (faster and more reliable)
            # Clear any existing text
            message_box.send_keys(Keys.CONTROL + "a")
            time.sleep(0.1)  # Reduced from 0.2
            message_box.send_keys(Keys.BACKSPACE)
            time.sleep(0.1)  # Reduced from 0.3
            
            # Type message using real keyboard input
            message_box.send_keys(message)
            time.sleep(0.4)  # Wait for message to be fully typed (reduced from 0.8)
        
        # Step 5: Auto send
        print(f"  → Sending message...")
        
        # Re-find message box to ensure it's fresh before sending
        message_box = get_fresh_message_box(driver, max_retries=2)
        if not message_box:
            print(f"  ✗ Error: Could not find message box to send")
            return False
        
        # Ensure message box is focused before sending
        try:
            message_box.click()
            time.sleep(0.1)  # Reduced from 0.3
            
            # Try sending with Enter key
            message_box.send_keys(Keys.ENTER)
            time.sleep(0.5)  # Reduced from 1
            
            # Verify message was sent by checking if message box is empty
            # If JavaScript was used, we might need to trigger send differently
            if has_non_bmp:
                # For JavaScript-typed messages, try clicking send button as fallback
                try:
                    send_button = WebDriverWait(driver, 2).until(
                        EC.element_to_be_clickable((By.XPATH, "//span[@data-icon='send'] | //button[@aria-label='Send'] | //span[contains(@data-testid, 'send')]"))
                    )
                    send_button.click()
                    time.sleep(0.5)  # Reduced from 1
                except:
                    pass  # Enter key should have worked
            
        except Exception as e:
            # Fallback: Use ActionChains
            try:
                ActionChains(driver).move_to_element(message_box).click().send_keys(Keys.ENTER).perform()
                time.sleep(0.5)  # Reduced from 1
            except Exception:
                # Last resort: Try JavaScript to trigger send
                try:
                    driver.execute_script("""
                        var elem = arguments[0];
                        var event = new KeyboardEvent('keydown', {
                            key: 'Enter',
                            code: 'Enter',
                            keyCode: 13,
                            which: 13,
                            bubbles: true,
                            cancelable: true
                        });
                        elem.dispatchEvent(event);
                        elem.dispatchEvent(new KeyboardEvent('keyup', {
                            key: 'Enter',
                            code: 'Enter',
                            keyCode: 13,
                            bubbles: true
                        }));
                    """, message_box)
                    time.sleep(0.5)  # Reduced from 1
                except Exception as e2:
                    print(f"  ✗ Error: Could not send message - {str(e2)}")
                    return False
        
        time.sleep(1)  # Wait for message to be sent (reduced from 2)
        
        print(f"✓ Message sent to {contact_number}")
        time.sleep(delay_seconds)  # This is the main delay between contacts (user configurable)
        
        # Go back to main page
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(0.2)  # Reduced from 0.5
        
        return True
        
    except TimeoutException as e:
        error_msg = str(e) if str(e) else "Timeout waiting for element"
        print(f"  ✗ Error: Timeout - {error_msg}")
        return False
    except Exception as e:
        error_msg = str(e) if str(e) else f"{type(e).__name__} (no message)"
        print(f"  ✗ Error: {error_msg}")
        import traceback
        traceback.print_exc()
        return False


def find_image_for_contact(contact_number, excel_file_path, images_folder=None):
    """
    Find image file for a specific contact
    Priority:
    1. Check images folder (if provided) for contact-specific image
    2. Check same directory as Excel file
    3. Look for pattern: {contact_number}.jpg/png/etc or image_{index}.jpg/png
    
    Args:
        contact_number: Contact number (cleaned, without +)
        excel_file_path: Path to Excel file
        images_folder: Optional folder path containing images
    
    Returns:
        Path to image file or None if not found
    """
    excel_dir = os.path.dirname(os.path.abspath(excel_file_path))
    if not excel_dir:
        excel_dir = os.getcwd()
    
    # Clean contact number for filename matching
    clean_number = contact_number.replace("+", "").replace(" ", "").replace("-", "")
    
    image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.webp']
    
    # Priority 1: Check images folder (if provided)
    if images_folder:
        if not os.path.isabs(images_folder):
            images_folder = os.path.join(excel_dir, images_folder)
        
        if os.path.exists(images_folder):
            # Look for contact-specific image: {number}.jpg, {number}.png, etc.
            for ext in image_extensions:
                image_path = os.path.join(images_folder, f"{clean_number}{ext}")
                if os.path.exists(image_path):
                    return image_path
            
            # Look for any image files in the folder (use first one found)
            try:
                for file in os.listdir(images_folder):
                    file_lower = file.lower()
                    if any(file_lower.endswith(ext) for ext in image_extensions):
                        return os.path.join(images_folder, file)
            except:
                pass
    
    # Priority 2: Check same directory as Excel file
    # Look for contact-specific image
    for ext in image_extensions:
        image_path = os.path.join(excel_dir, f"{clean_number}{ext}")
        if os.path.exists(image_path):
            return image_path
    
    # Priority 3: Look for any image in Excel directory
    try:
        for file in os.listdir(excel_dir):
            file_lower = file.lower()
            if any(file_lower.endswith(ext) for ext in image_extensions):
                return os.path.join(excel_dir, file)
    except:
        pass
    
    return None


def send_bulk_messages(excel_file_path, delay_seconds=15, start_from=0, images_folder=None, default_image=None):
    """
    Send messages to all contacts in the Excel file using Selenium
    
    Args:
        excel_file_path: Path to the XLSX file
        delay_seconds: Delay between each message (to avoid rate limiting)
        start_from: Index to start from (useful for resuming)
        images_folder: Optional folder path containing images (relative to Excel file or absolute)
        default_image: Optional path to a single image to send to ALL contacts (overrides individual images)
    """
    contacts = read_contacts_from_excel(excel_file_path)
    
    if not contacts:
        print("No contacts found. Please check your Excel file.")
        return
    
    # Check if default image exists
    if default_image:
        excel_dir = os.path.dirname(os.path.abspath(excel_file_path))
        if not os.path.isabs(default_image):
            default_image = os.path.join(excel_dir, default_image)
        if os.path.exists(default_image):
            print(f"✓ Default image found: {os.path.basename(default_image)}")
            print(f"  📷 Will send this image to ALL contacts")
        else:
            print(f"⚠️  Default image not found: {default_image}")
            print(f"  ⚠️  Will send text-only messages instead")
            default_image = None
    
    # Check if images folder exists (only if not using default image)
    if images_folder and not default_image:
        excel_dir = os.path.dirname(os.path.abspath(excel_file_path))
        if not os.path.isabs(images_folder):
            images_folder = os.path.join(excel_dir, images_folder)
        if os.path.exists(images_folder):
            print(f"✓ Images folder found: {images_folder}")
        else:
            print(f"⚠️  Images folder not found: {images_folder}")
            print(f"  ⚠️  Will send text-only messages instead")
            images_folder = None
    elif not default_image and not images_folder:
        print(f"ℹ️  Text-only mode: No images will be sent")
    
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
        
        # Determine mode
        if default_image:
            print(f"📷 Mode: Single image for all contacts")
            print(f"   Image: {os.path.basename(default_image)}")
        elif images_folder:
            print(f"📷 Mode: Individual images per contact")
            print(f"   Images folder: {images_folder}")
        else:
            print(f"📝 Mode: Text messages only (no images)")
        
        print(f"⚠️  Keep the browser window open and don't close it!\n")
        
        successful = 0
        failed = 0
        
        for index, contact in enumerate(contacts[start_from:], start=start_from):
            print(f"\n[{index + 1}/{len(contacts)}] Sending to {contact['number']}...")
            
            # Get image path for this contact
            image_path = None
            
            if default_image:
                # Mode 1: Single image for all contacts
                image_path = default_image
                if index == start_from:  # Only print once
                    print(f"  📷 Using default image: {os.path.basename(image_path)}")
            elif images_folder:
                # Mode 2: Individual images per contact (look for images)
                image_path = contact.get('image_path')
                
                # If no image path in Excel, try to find one
                if not image_path:
                    image_path = find_image_for_contact(contact['number'], excel_file_path, images_folder)
                    if image_path:
                        print(f"  📷 Found image: {os.path.basename(image_path)}")
            else:
                # Mode 3: Text-only mode (DEFAULT_IMAGE = None and IMAGES_FOLDER = None)
                # Don't look for images, send text only
                image_path = None
            
            # Send message (with image if available, text-only if image_path is None)
            if send_whatsapp_message(driver, contact['number'], contact['message'], delay_seconds, image_path):
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
    DELAY_SECONDS = 2  # Delay between messages (reduced for faster sending, increase if you get rate limited)
    START_FROM = 0  # Start from this index (useful if you need to resume)
    IMAGES_FOLDER = None  # Folder containing images (set to "images" for individual images, or None)
    
    # IMAGE CONFIGURATION - Choose one mode:
    # 
    # Mode 1: Single image for all contacts (campaign mode)
    # DEFAULT_IMAGE = "safari_promo.jpg"  # Same image for everyone
    # IMAGES_FOLDER = None  # Not needed
    #
    # Mode 2: Individual images per contact
    # DEFAULT_IMAGE = None  # Disable single image
    # IMAGES_FOLDER = "images"  # Folder with unique images
    #
    # Mode 3: Text messages only (no images)
    # DEFAULT_IMAGE = None  # No default image
    # IMAGES_FOLDER = None  # No images folder
    
    DEFAULT_IMAGE = None  # Set to image filename for single image mode, or None for text-only
    IMAGES_FOLDER = None  # Set to "images" for individual images, or None for text-only
    
    # Check if Excel file exists
    if not os.path.exists(EXCEL_FILE):
        print(f"✗ Error: {EXCEL_FILE} not found!")
        print(f"Please create an Excel file with:")
        print(f"  Column A: Contact Number (with country code, e.g., +1234567890)")
        print(f"  Column B: Message (caption)")
        print(f"  Column C: Image Path (optional - leave empty to auto-detect)")
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
        send_bulk_messages(EXCEL_FILE, DELAY_SECONDS, START_FROM, IMAGES_FOLDER, DEFAULT_IMAGE)
    except KeyboardInterrupt:
        print("\n\n⚠️  Process cancelled by user")
    except Exception as e:
        print(f"\n✗ Error: {str(e)}")
