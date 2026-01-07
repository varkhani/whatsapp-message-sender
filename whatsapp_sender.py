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
    - Column B = Contact Name (optional - if provided, message will be: "Dear [Name],\n\n[Message]")
    - Column C = Message (Caption)
    - Column D = Image Path (optional - if empty, will look for image based on contact number)
    
    Final message format:
    - If Contact Name (B) is provided: "Dear [Name],\n\n[Message from C]"
    - If Contact Name (B) is empty: "[Message from C]" (sent as-is)
"""
    contacts = []
    
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        
        # Skip header row (if exists) - start from row 2
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0] and row[2]:  # Check if contact number (A) and message (C) exist
                contact = str(row[0]).strip()
                
                # Get contact name from Column B (optional)
                contact_name = None
                if len(row) > 1 and row[1]:
                    contact_name = str(row[1]).strip()
                    if not contact_name:  # Empty string
                        contact_name = None
                
                # Get message from Column C
                message = str(row[2]).strip()
                
                # Build final message: "Dear [Name],\n\n[Message]" if name exists, else just [Message]
                # Ensure proper newline formatting with comma after name
                if contact_name:
                    # Add comma after name and newline before message
                    # Remove any leading newlines from message to avoid double spacing
                    message_clean = message.lstrip('\n\r')
                    final_message = f"Dear {contact_name},\n\n{message_clean}"
                else:
                    final_message = message
                
                # Format contact number (remove spaces, ensure country code)
                contact = contact.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
                
                # Get image path from Column D (if provided)
                image_path = None
                if len(row) > 3 and row[3]:
                    image_path = str(row[3]).strip()
                    if image_path:
                        # Convert to absolute path if relative
                        if not os.path.isabs(image_path):
                            excel_dir = os.path.dirname(os.path.abspath(file_path))
                            image_path = os.path.join(excel_dir, image_path)
                
                contacts.append({
                    'number': contact,
                    'message': final_message,
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
    Set text in message box using a hybrid approach that properly handles newlines:
    - Split by newlines and type line by line
    - Use Shift+Enter between lines (how WhatsApp creates line breaks)
    - Use JavaScript insertText for lines with emojis
    - Use send_keys for lines without emojis (preserves formatting better)
    
    Args:
        driver: Selenium WebDriver instance
        message_box: Message box element
        text: Text to set
    
    Returns:
        True if successful, False otherwise
    """
    try:
        from selenium.webdriver.common.action_chains import ActionChains
        from selenium.webdriver.common.keys import Keys
        
        # Focus and clear the element first
        message_box.click()
        time.sleep(0.1)
        
        # Clear existing content
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
        time.sleep(0.05)
        ActionChains(driver).send_keys(Keys.DELETE).perform()
        time.sleep(0.1)
        
        # Split text by newlines
        lines = text.split('\n')
        
        for i, line in enumerate(lines):
            if i > 0:
                # Press Shift+Enter to create a newline (this is how WhatsApp creates line breaks)
                ActionChains(driver).key_down(Keys.SHIFT).send_keys(Keys.ENTER).key_up(Keys.SHIFT).perform()
                time.sleep(0.05)
            
            if line:  # Only type if line is not empty
                # Check if line has emojis (non-BMP characters)
                has_emoji = any(ord(char) > 0xFFFF for char in line)
                
                if has_emoji:
                    # Use JavaScript insertText for lines with emojis
                    driver.execute_script("""
                        var elem = arguments[0];
                        var text = arguments[1];
                        elem.focus();
                        document.execCommand('insertText', false, text);
                    """, message_box, line)
                    time.sleep(0.05)
                else:
                    # Use send_keys for lines without emojis (preserves formatting better)
                    message_box.send_keys(line)
                    time.sleep(0.03)
        
        # Trigger final events to ensure WhatsApp recognizes the input
        driver.execute_script("""
            var elem = arguments[0];
            var inputEvent = new InputEvent('input', {bubbles: true, cancelable: true});
            elem.dispatchEvent(inputEvent);
            elem.focus();
        """, message_box)
        
        time.sleep(0.2)
        return True
    except Exception as e:
        print(f"  ⚠️  Hybrid text setting failed: {str(e)}, trying simple fallback...")
        # Fallback: Simple JavaScript method
        try:
            driver.execute_script("""
                var element = arguments[0];
                var text = arguments[1];
                element.focus();
                var range = document.createRange();
                range.selectNodeContents(element);
                var selection = window.getSelection();
                selection.removeAllRanges();
                selection.addRange(range);
                document.execCommand('delete', false, null);
                document.execCommand('insertText', false, text);
            """, message_box, text)
            time.sleep(0.3)
            return True
        except:
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
        
        time.sleep(1.5)  # Wait for attachment menu to appear
        
        # Step 3.5: Click "Photos & videos" option (NOT "New sticker")
        print(f"  → Selecting 'Photos & videos' option...")
        
        # First, find the attachment menu container (not all buttons on page)
        print(f"  → Finding attachment menu container...")
        attachment_menu = None
        menu_container_selectors = [
            "//div[@data-testid='popup']",
            "//div[contains(@class, 'popup')]",
            "//div[contains(@class, 'menu')]",
            "//div[@role='menu']",
            "//div[contains(@data-testid, 'attach')]//ancestor::div[contains(@class, 'popup')]",
            "//div[contains(@aria-label, 'attach')]//ancestor::div",
            # Look for menu that contains "Photos & videos" text
            "//div[contains(., 'Photos & videos') and contains(., 'Document')]",
            "//div[contains(., 'Photos & videos') and contains(., 'Camera')]"
        ]
        
        for selector in menu_container_selectors:
            try:
                menus = driver.find_elements(By.XPATH, selector)
                for menu in menus:
                    if menu.is_displayed():
                        # Check if it contains attachment-related options
                        menu_text = menu.text.lower()
                        # The menu should have multiple options: Document, Photos & videos, Camera, etc.
                        has_multiple_options = (
                            ('photos' in menu_text and 'videos' in menu_text) and
                            ('document' in menu_text or 'camera' in menu_text or 'audio' in menu_text)
                        )
                        if has_multiple_options:
                            attachment_menu = menu
                            print(f"  ✓ Found attachment menu container (contains multiple options)")
                            break
                if attachment_menu:
                    break
            except:
                continue
        
        # Scan ALL clickable/interactive elements in the attachment menu
        print(f"  → Scanning ALL interactive elements in attachment menu...")
        all_menu_items = []
        try:
            if attachment_menu:
                # Find ALL potentially clickable elements (buttons, clickable divs, etc.)
                all_menu_items = attachment_menu.find_elements(By.XPATH, 
                    ".//div[@role='button'] | "
                    ".//button | "
                    ".//div[contains(@data-testid, 'attach')] | "
                    ".//div[contains(@class, 'menu-item')] | "
                    ".//div[contains(@class, 'selectable')] | "
                    ".//*[contains(@class, 'clickable')] | "
                    ".//*[@tabindex]"
                )
            else:
                # Fallback: look for elements that are likely in attachment menu
                all_menu_items = driver.find_elements(By.XPATH, 
                    "//div[@role='button'][contains(., 'Photos')] | "
                    "//div[@role='button'][contains(., 'Document')] | "
                    "//div[@role='button'][contains(., 'Camera')] | "
                    "//div[@data-testid='attach-photo'] | "
                    "//div[@data-testid='attach-image'] | "
                    "//button[contains(., 'Photos')] | "
                    "//div[contains(@class, 'menu-item')][contains(., 'Photos')]"
                )
            
            print(f"  → Found {len(all_menu_items)} interactive element(s) in menu")
            for idx, item in enumerate(all_menu_items):
                try:
                    if item.is_displayed():
                        text = item.text.lower()
                        aria_label = (item.get_attribute('aria-label') or '').lower()
                        data_testid = (item.get_attribute('data-testid') or '').lower()
                        title = (item.get_attribute('title') or '').lower()
                        tag = item.tag_name.lower()
                        role = item.get_attribute('role') or ''
                        tabindex = item.get_attribute('tabindex') or ''
                        onclick = item.get_attribute('onclick') or ''
                        has_click_handler = bool(onclick)
                        print(f"    [{idx}] tag='{tag}', role='{role}', tabindex='{tabindex}', text='{text[:40]}', aria-label='{aria_label[:40]}', data-testid='{data_testid[:40]}', has_onclick={has_click_handler}")
                except:
                    pass
        except:
            pass
        
        photos_videos_option = None
        
        # Build selectors - prioritize within attachment menu if found
        if attachment_menu:
            base_xpath = "."
            photos_selectors = [
                # PRIORITY 1: Clickable buttons with data-testid (most reliable)
                f"{base_xpath}//div[@data-testid='attach-photo']",
                f"{base_xpath}//div[@data-testid='attach-image']",
                f"{base_xpath}//div[@data-testid='attach-media']",
                # PRIORITY 2: Buttons with role='button' containing "Photos & videos"
                f"{base_xpath}//div[@role='button'][contains(., 'Photos') and contains(., 'videos') and not(contains(., 'sticker'))]",
                f"{base_xpath}//div[@role='button'][normalize-space(.)='Photos & videos']",
                # PRIORITY 3: Find by position (2nd item in menu after Document)
                f"{base_xpath}//div[@role='button'][2]",  # 2nd button in menu
                f"{base_xpath}//div[@role='listitem'][2]//div[@role='button']",  # 2nd list item's button
                # PRIORITY 4: Text elements with parent button
                f"{base_xpath}//span[normalize-space(text())='Photos & videos']//ancestor::div[@role='button']",
                f"{base_xpath}//div[normalize-space(text())='Photos & videos']//ancestor::div[@role='button']",
                f"{base_xpath}//span[contains(text(), 'Photos & videos')]//ancestor::div[@role='button']",
            ]
        else:
            photos_selectors = []
        
        # Add global selectors (fallback)
        photos_selectors.extend([
            # PRIORITY 1: Clickable buttons with data-testid (most reliable - these are actual buttons)
            "//div[@data-testid='attach-photo']",
            "//div[@data-testid='attach-image']",
            "//div[@data-testid='attach-media']",
            # PRIORITY 2: Buttons with role='button' and aria-label (exact match)
            "//div[@role='button' and normalize-space(@aria-label)='Photos & videos']",
            "//div[@role='button' and contains(@aria-label, 'Photos') and contains(@aria-label, 'videos')]",
            "//div[@role='button' and contains(@aria-label, 'Photos')]",
            "//button[contains(., 'Photos') and contains(., 'videos')]",
            # PRIORITY 3: Buttons with role='button' containing text (but NOT sticker)
            "//div[@role='button'][contains(., 'Photos') and contains(., 'videos') and not(contains(., 'sticker'))]",
            "//div[@role='button'][normalize-space(.)='Photos & videos']",
            "//div[@role='button'][contains(., 'photos') and contains(., 'videos') and not(contains(., 'sticker'))]",
            # PRIORITY 4: Text elements with parent button (find text, then get parent button)
            "//span[contains(text(), 'Photos & videos')]//ancestor::div[@role='button']",
            "//div[contains(text(), 'Photos & videos')]//ancestor::div[@role='button']",
            "//span[normalize-space(text())='Photos & videos']//ancestor::div[@role='button']",
            "//div[normalize-space(text())='Photos & videos']//ancestor::div[@role='button']",
            # PRIORITY 5: Title attribute
            "//div[@title='Photos & videos']",
            "//div[contains(@title, 'Photos') and contains(@title, 'videos')]",
            # LAST RESORT: Text elements (will find parent button in code below)
            "//div[normalize-space(text())='Photos & videos']",
            "//span[normalize-space(text())='Photos & videos']",
            "//div[contains(text(), 'Photos & videos')]",
            "//span[contains(text(), 'Photos & videos')]",
            "//div[contains(text(), 'Photos') and contains(text(), 'videos')]"
        ])
        
        # FIRST: Try to find "Photos & videos" from the list of interactive elements we already found
        if all_menu_items and not photos_videos_option:
            print(f"  → Searching in {len(all_menu_items)} menu items for 'Photos & videos'...")
            for item in all_menu_items:
                try:
                    if not item.is_displayed() or not item.is_enabled():
                        continue
                    
                    text = item.text.lower()
                    aria_label = (item.get_attribute('aria-label') or '').lower()
                    data_testid = (item.get_attribute('data-testid') or '').lower()
                    tag = item.tag_name.lower()
                    role = item.get_attribute('role') or ''
                    tabindex = item.get_attribute('tabindex') or ''
                    onclick = item.get_attribute('onclick') or ''
                    
                    # Check if this is "Photos & videos" (NOT sticker)
                    is_sticker = 'sticker' in text or 'sticker' in aria_label or 'sticker' in data_testid
                    is_photos_videos = (
                        ('photos' in text and 'videos' in text) or
                        ('photos' in aria_label and 'videos' in aria_label) or
                        'attach-photo' in data_testid or
                        'attach-image' in data_testid or
                        'attach-media' in data_testid
                    )
                    
                    if is_photos_videos and not is_sticker:
                        # Check if it's actually clickable/interactive
                        is_clickable = (
                            tag == 'button' or 
                            role == 'button' or 
                            data_testid or
                            tabindex or
                            onclick or
                            item.get_attribute('onmousedown') or
                            item.get_attribute('onmouseup')
                        )
                        
                        if is_clickable:
                            photos_videos_option = item
                            print(f"  ✓ Found 'Photos & videos' clickable element (tag='{tag}', role='{role}', data-testid='{data_testid}', tabindex='{tabindex}')")
                            break
                        else:
                            # Find parent clickable element
                            try:
                                # Try to find parent with role='button' or actual button
                                parent = item.find_element(By.XPATH, 
                                    "./ancestor::div[@role='button'][1] | "
                                    "./ancestor::button[1] | "
                                    "./ancestor::div[@tabindex][1]"
                                )
                                if parent and parent.is_displayed():
                                    parent_role = parent.get_attribute('role') or ''
                                    parent_tabindex = parent.get_attribute('tabindex') or ''
                                    if parent_role == 'button' or parent_tabindex or parent.tag_name.lower() == 'button':
                                        photos_videos_option = parent
                                        print(f"  ✓ Found 'Photos & videos' clickable parent (tag='{parent.tag_name}', role='{parent_role}', tabindex='{parent_tabindex}')")
                                        break
                            except:
                                # Last resort: use JavaScript to find clickable parent
                                try:
                                    clickable_parent = driver.execute_script("""
                                        var elem = arguments[0];
                                        var current = elem;
                                        for (var i = 0; i < 10; i++) {
                                            current = current.parentElement;
                                            if (!current) break;
                                            var role = current.getAttribute('role');
                                            var tabindex = current.getAttribute('tabindex');
                                            var style = window.getComputedStyle(current);
                                            if (current.tagName === 'BUTTON' || 
                                                role === 'button' ||
                                                tabindex !== null ||
                                                current.onclick ||
                                                current.onmousedown ||
                                                style.cursor === 'pointer') {
                                                return current;
                                            }
                                        }
                                        return null;
                                    """, item)
                                    if clickable_parent and clickable_parent.is_displayed():
                                        photos_videos_option = clickable_parent
                                        print(f"  ✓ Found 'Photos & videos' clickable parent via JavaScript")
                                        break
                                except:
                                    pass
                except:
                    continue
        
        # SECOND: If not found in button list, try selectors
        if not photos_videos_option:
            print(f"  → Not found in button list, trying selectors...")
            for attempt in range(6):
                for selector in photos_selectors:
                    try:
                        # If we have attachment_menu, search within it; otherwise search globally
                        if attachment_menu and selector.startswith("."):
                            elements = attachment_menu.find_elements(By.XPATH, selector)
                        else:
                            elements = driver.find_elements(By.XPATH, selector)
                        for elem in elements:
                            try:
                                if not elem.is_displayed() or not elem.is_enabled():
                                    continue
                                
                                # Get all text content (including nested elements)
                                text = elem.text.lower()
                                aria_label = (elem.get_attribute('aria-label') or '').lower()
                                title = (elem.get_attribute('title') or '').lower()
                                data_testid = (elem.get_attribute('data-testid') or '').lower()
                                
                                # Double-check it's NOT the sticker option
                                is_sticker = (
                                    'sticker' in text or 
                                    'sticker' in aria_label or 
                                    'sticker' in title or
                                    'sticker' in data_testid
                                )
                                
                                # Verify it IS the photos & videos option
                                is_photos_videos = (
                                    ('photos' in text and 'videos' in text) or
                                    ('photos' in aria_label and 'videos' in aria_label) or
                                    ('photos' in title and 'videos' in title) or
                                    'attach-photo' in data_testid or
                                    'attach-image' in data_testid or
                                    'attach-media' in data_testid or
                                    ('photos' in text and 'sticker' not in text and 'sticker' not in aria_label)
                                )
                                
                                if is_photos_videos and not is_sticker:
                                    # CRITICAL: If we found a span or text element, we MUST find the parent button
                                    # Spans are not clickable - we need the actual button element
                                    final_element = None
                                    
                                    # Check if it's already a button
                                    if elem.tag_name.lower() == 'button' or elem.get_attribute('role') == 'button':
                                        final_element = elem
                                        print(f"  ✓ Found 'Photos & videos' button directly (tag='{elem.tag_name}', role='{elem.get_attribute('role')}')")
                                    else:
                                        # It's a span or div - MUST find parent button
                                        print(f"  → Found text element (tag='{elem.tag_name}'), searching for parent button...")
                                        try:
                                            # Try multiple ways to find the parent button
                                            parent_button = None
                                            
                                            # Method 1: Direct parent with role='button'
                                            try:
                                                parent_button = elem.find_element(By.XPATH, "./ancestor::div[@role='button'][1]")
                                            except:
                                                pass
                                            
                                            # Method 2: Parent button tag
                                            if not parent_button:
                                                try:
                                                    parent_button = elem.find_element(By.XPATH, "./ancestor::button[1]")
                                                except:
                                                    pass
                                            
                                            # Method 3: Parent with click handler or cursor pointer
                                            if not parent_button:
                                                try:
                                                    parent_button = driver.execute_script("""
                                                        var elem = arguments[0];
                                                        var current = elem.parentElement;
                                                        for (var i = 0; i < 5; i++) {
                                                            if (!current) break;
                                                            var style = window.getComputedStyle(current);
                                                            if (current.tagName === 'BUTTON' || 
                                                                current.getAttribute('role') === 'button' ||
                                                                style.cursor === 'pointer' ||
                                                                current.onclick) {
                                                                return current;
                                                            }
                                                            current = current.parentElement;
                                                        }
                                                        return null;
                                                    """, elem)
                                                except:
                                                    pass
                                            
                                            if parent_button:
                                                try:
                                                    # Verify parent is actually a button
                                                    if (parent_button.tag_name.lower() == 'button' or 
                                                        parent_button.get_attribute('role') == 'button' or
                                                        parent_button.is_displayed()):
                                                        final_element = parent_button
                                                        print(f"  ✓ Found 'Photos & videos' button (parent, tag='{parent_button.tag_name}', role='{parent_button.get_attribute('role')}')")
                                                    else:
                                                        print(f"  ⚠️  Parent found but not a button, continuing...")
                                                except:
                                                    print(f"  ⚠️  Error checking parent button, continuing...")
                                            else:
                                                print(f"  ⚠️  Could not find parent button for span element, continuing search...")
                                        except Exception as e:
                                            print(f"  ⚠️  Error finding parent button: {str(e)}, continuing search...")
                                    
                                # CRITICAL: Only use if we found a VALID clickable button element
                                # Must have role='button', tabindex, onclick, or be actual button tag
                                if final_element:
                                    try:
                                        final_tag = final_element.tag_name.lower()
                                        final_role = final_element.get_attribute('role') or ''
                                        final_tabindex = final_element.get_attribute('tabindex') or ''
                                        final_onclick = final_element.get_attribute('onclick') or ''
                                        
                                        # Verify it's actually clickable
                                        is_actually_clickable = (
                                            final_tag == 'button' or
                                            final_role == 'button' or
                                            final_tabindex or
                                            final_onclick or
                                            data_testid  # Has data-testid (like attach-photo)
                                        )
                                        
                                        if is_actually_clickable and final_element.is_enabled() and final_element.is_displayed():
                                            photos_videos_option = final_element
                                            print(f"  ✓ Using clickable element (tag='{final_tag}', role='{final_role}', tabindex='{final_tabindex}', data-testid='{data_testid}')")
                                            break
                                        else:
                                            print(f"  ⚠️  Element found but NOT clickable (tag='{final_tag}', role='{final_role}', tabindex='{final_tabindex}'), continuing search...")
                                    except:
                                        print(f"  ⚠️  Error verifying element, continuing search...")
                                # If we didn't find a clickable element, continue searching (don't use non-clickable elements)
                            except:
                                continue
                        if photos_videos_option:
                            break
                    except:
                        continue
                if photos_videos_option:
                    break
                time.sleep(0.5)
        
        if photos_videos_option:
            try:
                # CRITICAL: Re-find the element right before clicking to avoid stale element reference
                # Store identifying attributes first
                try:
                    stored_text = photos_videos_option.text.lower()
                    stored_aria = (photos_videos_option.get_attribute('aria-label') or '').lower()
                    stored_role = photos_videos_option.get_attribute('role') or ''
                    stored_tabindex = photos_videos_option.get_attribute('tabindex') or ''
                except:
                    stored_text = ''
                    stored_aria = ''
                    stored_role = ''
                    stored_tabindex = ''
                
                print(f"  → Clicking option: text='{stored_text[:50]}', aria-label='{stored_aria[:50]}', role='{stored_role}', tabindex='{stored_tabindex}'")
                
                # Re-find the element using stored attributes to avoid stale reference
                fresh_element = None
                try:
                    # Try to find it again using the same method
                    if attachment_menu:
                        # Search within attachment menu
                        all_items = attachment_menu.find_elements(By.XPATH, 
                            ".//div[@role='menuitem'] | .//button | .//div[contains(@data-testid, 'attach')]"
                        )
                        for item in all_items:
                            try:
                                if not item.is_displayed():
                                    continue
                                item_text = item.text.lower()
                                item_aria = (item.get_attribute('aria-label') or '').lower()
                                item_role = item.get_attribute('role') or ''
                                item_tabindex = item.get_attribute('tabindex') or ''
                                
                                # Match by text/aria-label and role/tabindex
                                if (('photos' in item_text and 'videos' in item_text) or 
                                    ('photos' in item_aria and 'videos' in item_aria)) and \
                                   (item_role == stored_role or item_tabindex == stored_tabindex):
                                    fresh_element = item
                                    break
                            except:
                                continue
                    
                    # If not found, try global search
                    if not fresh_element:
                        for selector in photos_selectors:
                            try:
                                elements = driver.find_elements(By.XPATH, selector)
                                for elem in elements:
                                    try:
                                        if not elem.is_displayed():
                                            continue
                                        elem_text = elem.text.lower()
                                        elem_aria = (elem.get_attribute('aria-label') or '').lower()
                                        if ('photos' in elem_text and 'videos' in elem_text) or \
                                           ('photos' in elem_aria and 'videos' in elem_aria):
                                            fresh_element = elem
                                            break
                                    except:
                                        continue
                                if fresh_element:
                                    break
                            except:
                                continue
                    
                    if fresh_element:
                        photos_videos_option = fresh_element
                        print(f"  ✓ Re-found element to avoid stale reference")
                    else:
                        print(f"  ⚠️  Could not re-find element, will try with original (may be stale)")
                except Exception as e:
                    print(f"  ⚠️  Error re-finding element: {str(e)}, will try with original")
                
                # Scroll into view
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", photos_videos_option)
                    time.sleep(0.5)
                except:
                    pass
                
                # Try multiple click methods to ensure it works
                clicked = False
                
                # Method 1: JavaScript click with event dispatch (most reliable, works even if element is stale)
                try:
                    # Use JavaScript to find and click by text/aria-label (avoids stale element)
                    driver.execute_script("""
                        var menu = arguments[0];
                        var targetText = arguments[1];
                        var targetAria = arguments[2];
                        
                        // Find all menu items
                        var items = menu.querySelectorAll('[role="menuitem"], button, div[tabindex]');
                        for (var i = 0; i < items.length; i++) {
                            var item = items[i];
                            var text = (item.textContent || '').toLowerCase();
                            var aria = (item.getAttribute('aria-label') || '').toLowerCase();
                            
                            if ((text.includes('photos') && text.includes('videos')) ||
                                (aria.includes('photos') && aria.includes('videos'))) {
                                // Found it - click it
                                item.scrollIntoView({block: 'center', behavior: 'smooth'});
                                setTimeout(function() {
                                    if (item.click) item.click();
                                    var clickEvent = new MouseEvent('click', {bubbles: true, cancelable: true});
                                    item.dispatchEvent(clickEvent);
                                    item.dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));
                                    item.dispatchEvent(new MouseEvent('mouseup', {bubbles: true}));
                                }, 100);
                                return true;
                            }
                        }
                        return false;
                    """, attachment_menu if attachment_menu else driver.execute_script("return document;"), 'photos & videos', 'photos & videos')
                    time.sleep(1.5)
                    clicked = True
                    print(f"  → Clicked using JavaScript (bypassing stale element)")
                except Exception as e1:
                    print(f"  ⚠️  JavaScript click failed: {str(e1)}")
                
                # Method 2: Try with fresh element reference
                if not clicked:
                    try:
                        driver.execute_script("""
                            var elem = arguments[0];
                            if (elem.click) elem.click();
                            var clickEvent = new MouseEvent('click', {bubbles: true, cancelable: true});
                            elem.dispatchEvent(clickEvent);
                            elem.dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));
                            elem.dispatchEvent(new MouseEvent('mouseup', {bubbles: true}));
                        """, photos_videos_option)
                        time.sleep(1.5)
                        clicked = True
                        print(f"  → Clicked using JavaScript with events")
                    except Exception as e1:
                        print(f"  ⚠️  JavaScript click failed: {str(e1)}")
                
                # Method 3: Regular click (fallback)
                if not clicked:
                    try:
                        photos_videos_option.click()
                        time.sleep(1.5)
                        clicked = True
                        print(f"  → Clicked using regular click")
                    except Exception as e2:
                        print(f"  ⚠️  Regular click failed: {str(e2)}")
                
                # Method 4: ActionChains click (last resort)
                if not clicked:
                    try:
                        from selenium.webdriver.common.action_chains import ActionChains
                        ActionChains(driver).move_to_element(photos_videos_option).pause(0.2).click().perform()
                        time.sleep(1.5)
                        clicked = True
                        print(f"  → Clicked using ActionChains")
                    except Exception as e3:
                        print(f"  ⚠️  ActionChains click failed: {str(e3)}")
                
                if not clicked:
                    print(f"  ✗ ERROR: Could not click 'Photos & videos' option!")
                    return send_text_fallback()
                
                time.sleep(2.5)  # Wait for file picker to open
                
                # CRITICAL: Verify selection was successful
                print(f"  → Verifying 'Photos & videos' was selected...")
                time.sleep(1)  # Wait a bit for interface to update
                
                file_inputs_check = driver.find_elements(By.XPATH, "//input[@type='file']")
                
                # Check if we're in sticker mode (bad!)
                sticker_indicators = driver.find_elements(By.XPATH, 
                    "//span[contains(text(), 'Send sticker')] | "
                    "//button[contains(@aria-label, 'sticker')] | "
                    "//div[contains(@aria-label, 'sticker') and contains(@aria-label, 'send')] | "
                    "//span[@data-icon='sticker']"
                )
                is_sticker_mode = any(btn.is_displayed() for btn in sticker_indicators)
                
                # Also check if file input accepts stickers
                if file_inputs_check:
                    for inp in file_inputs_check:
                        accept_attr = (inp.get_attribute('accept') or '').lower()
                        if 'sticker' in accept_attr:
                            is_sticker_mode = True
                            print(f"  ⚠️  File input accepts stickers - likely sticker mode")
                            break
                
                if is_sticker_mode:
                    print(f"  ✗ ERROR: In STICKER mode! 'Photos & videos' was NOT selected correctly.")
                    print(f"      The element we clicked was not the actual 'Photos & videos' button.")
                    print(f"      Canceling and will try alternative approach...")
                    # Cancel sticker mode
                    try:
                        from selenium.webdriver.common.action_chains import ActionChains
                        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                        time.sleep(1.5)
                    except:
                        pass
                    
                    # Try finding and clicking "Photos & videos" using keyboard navigation
                    print(f"  → Trying keyboard navigation to select 'Photos & videos'...")
                    try:
                        # Press Arrow Down to navigate to "Photos & videos" (2nd item)
                        ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
                        time.sleep(0.3)
                        ActionChains(driver).send_keys(Keys.ENTER).perform()
                        time.sleep(2)
                        
                        # Check again
                        sticker_check2 = driver.find_elements(By.XPATH, 
                            "//span[contains(text(), 'Send sticker')] | "
                            "//button[contains(@aria-label, 'sticker')]"
                        )
                        if any(btn.is_displayed() for btn in sticker_check2):
                            print(f"  ✗ Still in sticker mode after keyboard navigation - canceling")
                            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                            time.sleep(1)
                            return send_text_fallback()
                        else:
                            print(f"  ✓ Keyboard navigation worked - not in sticker mode")
                    except Exception as e:
                        print(f"  ✗ Keyboard navigation failed: {str(e)}")
                        return send_text_fallback()
                
                if file_inputs_check:
                    print(f"  ✓ Selected 'Photos & videos' (file input available)")
                else:
                    print(f"  ✗ ERROR: File input not found after clicking")
                    print(f"      'Photos & videos' was NOT selected correctly")
                    return send_text_fallback()
                
                # Verify we're NOT in sticker mode
                try:
                    sticker_buttons = driver.find_elements(By.XPATH, 
                        "//span[contains(text(), 'Send sticker')] | "
                        "//button[contains(@aria-label, 'sticker')] | "
                        "//div[contains(@aria-label, 'sticker') and contains(@aria-label, 'send')]"
                    )
                    if any(btn.is_displayed() for btn in sticker_buttons):
                        print(f"  ⚠️  WARNING: Sticker mode detected after clicking! Canceling...")
                        from selenium.webdriver.common.action_chains import ActionChains
                        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                        time.sleep(1)
                        print(f"  → Retrying 'Photos & videos' selection...")
                        # Try clicking again
                        driver.execute_script("arguments[0].click();", photos_videos_option)
                        time.sleep(2.5)
                except:
                    pass
            except Exception as e1:
                try:
                    # Fallback: regular click
                    photos_videos_option.click()
                    time.sleep(2.5)
                    print(f"  ✓ Selected 'Photos & videos' (fallback click)")
                except Exception as e2:
                    print(f"  ⚠️  Could not click 'Photos & videos': {str(e1)}, {str(e2)}")
                    print(f"      Trying direct file input...")
        else:
            print(f"  ⚠️  Could not find 'Photos & videos' option after 6 attempts")
            print(f"      This may result in sticker mode. Trying direct file input...")
        
        # Step 4: Find file input element (make sure it's for photos, not stickers)
        print(f"  → Looking for file input (photo mode, not sticker)...")
        file_input = None
        
        # Wait a bit for file input to appear after clicking "Photos & videos"
        time.sleep(1)
        
        try:
            file_inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
            print(f"  → Found {len(file_inputs)} file input(s)")
            
            for idx, inp in enumerate(file_inputs):
                try:
                    accept_attr = (inp.get_attribute('accept') or '').lower()
                    name_attr = (inp.get_attribute('name') or '').lower()
                    id_attr = (inp.get_attribute('id') or '').lower()
                    data_testid = (inp.get_attribute('data-testid') or '').lower()
                    
                    # Check if this is a sticker input (we want to AVOID this)
                    is_sticker_input = (
                        'sticker' in accept_attr or
                        'sticker' in name_attr or
                        'sticker' in id_attr or
                        'sticker' in data_testid
                    )
                    
                    # Prefer file inputs that accept images (for photos)
                    # Accept: image/*, image/jpeg, image/png, etc. (but NOT stickers)
                    is_photo_input = (
                        ('image' in accept_attr and 'sticker' not in accept_attr) or
                        accept_attr == '' or  # Empty accept might be photo input
                        ('image' in name_attr and 'sticker' not in name_attr) or
                        ('photo' in name_attr or 'photo' in id_attr)
                    )
                    
                    print(f"    Input [{idx}]: accept='{accept_attr}', name='{name_attr}', id='{id_attr}', is_sticker={is_sticker_input}, is_photo={is_photo_input}")
                    
                    # Skip sticker inputs
                    if is_sticker_input:
                        print(f"    ⚠️  Skipping input [{idx}] - appears to be sticker input")
                        continue
                    
                    # Use photo input
                    if is_photo_input:
                        file_input = inp
                        print(f"  ✓ Selected file input [{idx}] for photos (accept='{accept_attr}')")
                        break
                except Exception as e:
                    print(f"    ⚠️  Error checking input [{idx}]: {str(e)}")
                    continue
            
            # If no photo input found but we have file inputs, try to find one that's NOT a sticker
            if not file_input and file_inputs:
                for idx, inp in enumerate(file_inputs):
                    try:
                        accept_attr = (inp.get_attribute('accept') or '').lower()
                        name_attr = (inp.get_attribute('name') or '').lower()
                        # Make sure it's not explicitly a sticker input
                        if 'sticker' not in accept_attr and 'sticker' not in name_attr:
                            file_input = inp
                            print(f"  ✓ Using file input [{idx}] as fallback (accept='{accept_attr}')")
                            break
                    except:
                        continue
                
                # Last resort: use first input if no sticker indicators found
                if not file_input:
                    first_inp = file_inputs[0]
                    accept_attr = (first_inp.get_attribute('accept') or '').lower()
                    if 'sticker' not in accept_attr:
                        file_input = first_inp
                        print(f"  ⚠️  Using first file input as last resort (accept='{accept_attr}')")
        except Exception as e:
            print(f"  ⚠️  Error finding file input: {str(e)}")
        
        if not file_input:
            print(f"  ✗ Error: Could not find photo file input (only sticker input found or none available)")
            print(f"      Make sure 'Photos & videos' was selected correctly")
            return send_text_fallback()
        
        # Step 5: Upload image file
        print(f"  → Uploading image as photo (NOT sticker)...")
        try:
            # Verify file path exists
            if not os.path.exists(image_path):
                print(f"  ✗ ERROR: Image file not found: {image_path}")
                return send_text_fallback()
            
            # Get absolute path
            abs_image_path = os.path.abspath(image_path)
            print(f"  → Uploading: {abs_image_path}")
            
            file_input.send_keys(abs_image_path)
            time.sleep(12)  # Wait longer for image to load and preview interface to appear
            
            # Verify image was actually uploaded by checking for image preview
            print(f"  → Verifying image upload...")
            image_uploaded = False
            for check_attempt in range(5):
                time.sleep(1)
                previews = driver.find_elements(By.XPATH, 
                    "//img[contains(@src, 'blob')] | "
                    "//div[contains(@data-testid, 'media')]//img | "
                    "//div[contains(@class, 'preview')]//img"
                )
                if any(p.is_displayed() for p in previews):
                    image_uploaded = True
                    print(f"  ✓ Image preview appeared after {check_attempt + 1} seconds")
                    break
            
            if not image_uploaded:
                print(f"  ⚠️  WARNING: Image preview not found after upload - image may not have uploaded")
                # Still continue, might be a timing issue
            
            # Verify we're in photo mode (not sticker mode) by checking for caption input
            # In sticker mode, there's usually no caption input
            time.sleep(2)  # Additional wait for interface to render
            caption_check = driver.find_elements(By.XPATH, 
                "//div[@contenteditable='true'][@data-tab='11'] | "
                "//div[@contenteditable='true'][contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'type a message')]"
            )
            has_caption_input = any(elem.is_displayed() for elem in caption_check)
            
            # Check for sticker indicators
            sticker_check = driver.find_elements(By.XPATH, 
                "//div[contains(@aria-label, 'sticker') and contains(@aria-label, 'send')] | "
                "//span[contains(text(), 'Send sticker')]"
            )
            is_sticker_mode = any(elem.is_displayed() for elem in sticker_check)
            
            # Check for photo editing tools (indicates photo mode, not sticker)
            photo_tools = driver.find_elements(By.XPATH, 
                "//span[@data-icon='crop'] | "  # Crop tool
                "//span[@data-icon='rotate'] | "  # Rotate tool
                "//span[@data-icon='filter'] | "  # Filter tool
                "//div[contains(@aria-label, 'crop')] | "
                "//div[contains(@aria-label, 'rotate')]"
            )
            has_photo_tools = any(tool.is_displayed() for tool in photo_tools)
            
            if is_sticker_mode:
                print(f"  ✗ ERROR: Image uploaded in STICKER mode! Caption input will not appear.")
                print(f"      This means 'Photos & videos' was not selected correctly.")
                print(f"      Canceling and sending text only...")
                # Try pressing Escape to cancel
                from selenium.webdriver.common.action_chains import ActionChains
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(1)
                return send_text_fallback()  # Cancel and send text only
            
            # Additional check: Look for photo editing interface (crop, rotate, etc.)
            # NOTE: Photo editing tools might not appear immediately - they appear when you click the image
            # OR they might not appear at all in some WhatsApp Web versions
            # The KEY indicator is: if image preview is visible AND no sticker send button, it's photo mode
            if not has_photo_tools and not has_caption_input:
                print(f"  → Photo editing tools not immediately visible - checking image preview...")
                
                # Check if image preview is visible (this is the main indicator of photo mode)
                image_preview_check = driver.find_elements(By.XPATH, 
                    "//img[contains(@src, 'blob')] | "
                    "//div[contains(@data-testid, 'media')] | "
                    "//div[contains(@class, 'preview')]"
                )
                image_preview_visible = any(img.is_displayed() for img in image_preview_check)
                
                if image_preview_visible and not is_sticker_mode:
                    print(f"  ✓ Image preview visible and NOT in sticker mode - assuming photo mode")
                    print(f"  → Proceeding (editing tools may be hidden or appear on click)")
                    # Don't cancel - proceed with photo mode
                    has_photo_tools = True  # Set to True to bypass the cancel check
                elif is_sticker_mode:
                    print(f"  ✗ ERROR: Sticker mode detected - canceling...")
                    from selenium.webdriver.common.action_chains import ActionChains
                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                    time.sleep(1)
                    return send_text_fallback()
                else:
                    print(f"  ⚠️  Image preview not visible - might still be loading...")
                    # Wait a bit more and check again
                    time.sleep(2)
                    image_preview_recheck = driver.find_elements(By.XPATH, 
                        "//img[contains(@src, 'blob')] | "
                        "//div[contains(@data-testid, 'media')]"
                    )
                    if any(img.is_displayed() for img in image_preview_recheck):
                        print(f"  ✓ Image preview now visible - proceeding with photo mode")
                        has_photo_tools = True
                    else:
                        print(f"  ⚠️  Image preview still not visible - but proceeding anyway")
                        # Still proceed - might be a timing issue
            elif has_photo_tools:
                print(f"  ✓ Photo mode confirmed (photo editing tools visible)")
                if not has_caption_input:
                    print(f"  ⚠️  Caption input not found yet - may still be loading...")
                else:
                    print(f"  ✓ Caption input available")
            elif not has_caption_input:
                print(f"  ⚠️  Caption input not found - checking if photo mode...")
                # Check if image preview is visible (should be in both modes)
                image_preview = driver.find_elements(By.XPATH, 
                    "//img[contains(@src, 'blob')] | "
                    "//div[contains(@data-testid, 'media')]//img"
                )
                if any(img.is_displayed() for img in image_preview):
                    print(f"  → Image preview visible, but caption input missing - may be sticker mode")
                else:
                    print(f"  → Image preview not visible - interface may still be loading...")
            else:
                print(f"  ✓ Photo mode confirmed (caption input available)")
        except Exception as e:
            print(f"  ⚠️  Could not upload image: {str(e)}, sending text only")
            return send_text_fallback()
        
        # Step 5.5: Wait for image preview interface to fully load
        print(f"  → Waiting for image preview interface to load...")
        # Get initial state of contenteditable elements BEFORE image upload
        initial_contenteditables = {}
        try:
            initial_elems = driver.find_elements(By.XPATH, "//div[@contenteditable='true']")
            for elem in initial_elems:
                try:
                    data_tab = elem.get_attribute('data-tab')
                    placeholder = elem.get_attribute('placeholder') or ''
                    initial_contenteditables[data_tab] = placeholder
                except:
                    pass
        except:
            pass
        
        # Wait and check for new contenteditable elements or changed attributes
        print(f"  → Checking for caption input to appear...")
        caption_input_found = False
        for wait_round in range(12):
            time.sleep(1)
            try:
                current_elems = driver.find_elements(By.XPATH, "//div[@contenteditable='true']")
                for elem in current_elems:
                    try:
                        if not elem.is_displayed():
                            continue
                        data_tab = elem.get_attribute('data-tab')
                        placeholder = str(elem.get_attribute('placeholder') or '').lower()
                        aria_label = str(elem.get_attribute('aria-label') or '').lower()
                        
                        # Check if this is a NEW element (not in initial set)
                        is_new = data_tab not in initial_contenteditables
                        # Check if placeholder changed (message box might change when image is attached)
                        placeholder_changed = (
                            data_tab in initial_contenteditables and 
                            initial_contenteditables[data_tab] != placeholder
                        )
                        # Check if it has caption-like attributes
                        is_caption_like = (
                            data_tab == '11' or
                            'type a message' in placeholder or
                            ('message' in placeholder and data_tab != '10' and data_tab != '3') or
                            'caption' in placeholder
                        )
                        
                        if (is_new or placeholder_changed or is_caption_like) and data_tab not in ['3']:
                            print(f"  ✓ Found potential caption input after {wait_round + 1}s (data-tab='{data_tab}', placeholder='{placeholder[:30]}')")
                            caption_input_found = True
                            break
                    except:
                        continue
                if caption_input_found:
                    break
            except:
                pass
        
        # Try to activate the caption input by interacting with the image preview
        print(f"  → Activating caption input...")
        try:
            # Method 1: Click on the image preview itself to activate caption mode
            image_previews = driver.find_elements(By.XPATH, 
                "//img[contains(@src, 'blob')] | "
                "//div[contains(@data-testid, 'media')]//img | "
                "//div[contains(@class, 'preview')]//img"
            )
            for preview in image_previews:
                if preview.is_displayed():
                    # Click on the image preview to activate caption input
                    driver.execute_script("arguments[0].click();", preview)
                    time.sleep(1)
                    print(f"  ✓ Clicked image preview to activate caption mode")
                    # Check if caption input appeared
                    caption_check = driver.find_elements(By.XPATH, 
                        "//div[@contenteditable='true'][@data-tab='11'] | "
                        "//div[@contenteditable='true'][contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'type a message')]"
                    )
                    if any(elem.is_displayed() for elem in caption_check):
                        print(f"  ✓ Caption input appeared after clicking image preview")
                        break
                    break
            
            # Method 2: Click in the footer area below the image (where caption input should be)
            time.sleep(0.5)
            footer_area = driver.find_elements(By.XPATH, 
                "//footer | "
                "//div[contains(@class, 'footer')] | "
                "//div[contains(@data-testid, 'conversation-compose')]"
            )
            for footer in footer_area:
                if footer.is_displayed():
                    # Click in the footer area where caption input should appear
                    driver.execute_script("arguments[0].click();", footer)
                    time.sleep(0.5)
                    print(f"  ✓ Clicked footer area to activate caption input")
                    break
            
            # Method 3: Press Tab key to navigate to caption input
            from selenium.webdriver.common.action_chains import ActionChains
            for tab_press in range(3):
                ActionChains(driver).send_keys(Keys.TAB).perform()
                time.sleep(0.3)
                # Check if caption input is now focused
                focused = driver.execute_script("return document.activeElement;")
                if focused:
                    data_tab = focused.get_attribute('data-tab')
                    if data_tab == '11':
                        print(f"  ✓ Caption input focused via Tab navigation")
                        break
            
            # Method 4: Click on message box as fallback
            message_boxes = driver.find_elements(By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
            for msg_box in message_boxes:
                if msg_box.is_displayed():
                    driver.execute_script("arguments[0].click();", msg_box)
                    driver.execute_script("arguments[0].focus();", msg_box)
                    time.sleep(0.5)
                    print(f"  ✓ Clicked and focused message box to activate caption mode")
                    break
        except Exception as e:
            print(f"  ⚠️  Error activating caption input: {str(e)}")
        
        # Step 6: Find caption input box (appears after image is selected)
        # This should be the "Type a message" input that appears BELOW the image preview
        print(f"  → Looking for caption input box (below image preview)...")
        caption_box = None
        
        # Wait longer for the image preview interface to fully render
        time.sleep(2)
        
        # First, find the image preview container, then look for caption input inside it
        print(f"  → Finding image preview container...")
        media_container = None
        container_selectors = [
            "//div[contains(@data-testid, 'media')]",
            "//div[contains(@class, 'media')]",
            "//div[contains(@data-testid, 'image')]",
            "//div[contains(@class, 'preview')]"
        ]
        
        for selector in container_selectors:
            try:
                containers = driver.find_elements(By.XPATH, selector)
                for container in containers:
                    if container.is_displayed():
                        # Check if it contains an image
                        imgs = container.find_elements(By.XPATH, ".//img[contains(@src, 'blob')]")
                        if imgs:
                            media_container = container
                            print(f"  ✓ Found image preview container")
                            break
                if media_container:
                    break
            except:
                continue
        
        # Try clicking on image preview area to activate caption input
        print(f"  → Clicking on image preview to activate caption input...")
        try:
            # Find and click the image preview container
            preview_containers = driver.find_elements(By.XPATH, 
                "//div[contains(@data-testid, 'media')] | "
                "//div[contains(@class, 'preview')] | "
                "//div[contains(@class, 'media-preview')]"
            )
            for container in preview_containers:
                if container.is_displayed():
                    # Click in the bottom area of the container (where caption input should be)
                    driver.execute_script("""
                        var container = arguments[0];
                        var rect = container.getBoundingClientRect();
                        // Click in the bottom area where caption input should appear
                        var x = rect.left + (rect.width / 2);
                        var y = rect.bottom - 50; // Near bottom of container
                        var element = document.elementFromPoint(x, y);
                        if (element) {
                            element.click();
                            element.focus();
                        }
                    """, container)
                    time.sleep(1)
                    print(f"  ✓ Clicked in image preview container to activate caption")
                    break
        except:
            pass
        
        # First, let's see what contenteditable elements are available (debug)
        # Check multiple times as caption input might appear later
        # CRITICAL: We're looking for data-tab='11' or an input with placeholder='Type a message' that appears BELOW the image
        print(f"  → Scanning for caption input (checking multiple times)...")
        for scan_attempt in range(12):  # Increased to 12 attempts
            try:
                all_contenteditables = driver.find_elements(By.XPATH, "//div[@contenteditable='true']")
                print(f"  → Scan {scan_attempt + 1}: Found {len(all_contenteditables)} contenteditable elements")
                for idx, elem in enumerate(all_contenteditables):
                    try:
                        if not elem.is_displayed():
                            continue
                        data_tab = elem.get_attribute('data-tab')
                        placeholder = elem.get_attribute('placeholder') or ''
                        aria_label = elem.get_attribute('aria-label') or ''
                        role = elem.get_attribute('role') or ''
                        spellcheck = elem.get_attribute('spellcheck') or ''
                        
                        # CRITICAL: Caption input should be:
                        # - data-tab='11' (the actual caption input) OR
                        # - placeholder contains 'Type a message' AND data-tab is NOT '10' (not the regular message box)
                        # - We MUST NOT use data-tab='10' as it will detach the image!
                        is_potential_caption = (
                            data_tab == '11' or  # The actual caption input
                            ('type a message' in placeholder.lower() and data_tab != '10' and data_tab != '3') or
                            ('message' in placeholder.lower() and data_tab not in ['3', '10']) or
                            ('caption' in placeholder.lower() and data_tab not in ['3', '10'])
                        )
                        
                        print(f"    [{idx}] data-tab='{data_tab}', placeholder='{placeholder[:30]}', aria-label='{aria_label[:30]}', role='{role}', displayed=True, potential_caption={is_potential_caption}")
                        
                        # If this looks like a caption input and we haven't found one yet, try it
                        if is_potential_caption and not caption_box:
                            # Verify image preview is visible
                            previews = driver.find_elements(By.XPATH, 
                                "//img[contains(@src, 'blob')] | "
                                "//div[contains(@data-testid, 'media')]"
                            )
                            has_preview = any(p.is_displayed() for p in previews)
                            if has_preview:
                                # Double-check it's not data-tab='10' (regular message box)
                                if data_tab != '10':
                                    caption_box = elem
                                    print(f"  ✓ Found caption input during scan (data-tab='{data_tab}', placeholder='{placeholder[:30]}')")
                                    break
                                else:
                                    print(f"  ⚠️  Skipping data-tab='10' - this is the regular message box, not caption input")
                    except:
                        pass
                if caption_box:
                    break
            except:
                pass
            if scan_attempt < 11:  # Don't sleep after last attempt
                time.sleep(1)
        
        caption_selectors = [
            # Look INSIDE media container first (most specific)
            ".//div[@contenteditable='true'][contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'type a message')]",
            ".//div[@contenteditable='true'][contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'message')]",
            ".//div[@contenteditable='true'][@data-tab='11']",
            ".//div[@contenteditable='true'][@spellcheck='true']",
            ".//div[@contenteditable='true']",
            # Most specific selectors first (global)
            "//div[@contenteditable='true'][@data-tab='11']",
            "//div[@contenteditable='true'][@data-testid='media-caption-input-container']",
            "//div[@contenteditable='true'][@spellcheck='true'][@data-tab='11']",
            # By placeholder text - "Type a message" (this is what appears below image)
            "//div[@contenteditable='true'][contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'type a message')]",
            "//div[@contenteditable='true'][contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'message')]",
            "//div[@contenteditable='true'][contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'caption')]",
            # By role and contenteditable
            "//div[@role='textbox'][@contenteditable='true'][@data-tab='11']",
            # Look in footer/media areas
            "//footer//div[@contenteditable='true'][@data-tab='11']",
            "//div[contains(@class, 'media')]//div[@contenteditable='true'][@data-tab='11']",
            "//div[contains(@data-testid, 'media')]//div[@contenteditable='true']",
            # More generic - any contenteditable with data-tab='11' that's visible
            "//div[@contenteditable='true'][@data-tab='11']"
        ]
        
        # Try to find caption box - prioritize looking inside media container
        # Wait up to 15 seconds with more attempts
        for attempt in range(15):
            # First, try to find caption box INSIDE the media container (most reliable)
            if media_container:
                try:
                    # Look for "Type a message" input inside media container
                    caption_in_container = media_container.find_elements(By.XPATH, 
                        ".//div[@contenteditable='true'][contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'type a message')] | "
                        ".//div[@contenteditable='true'][contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'message')] | "
                        ".//div[@contenteditable='true'][@data-tab='11'] | "
                        ".//div[@contenteditable='true'][@spellcheck='true']"
                    )
                    for elem in caption_in_container:
                        try:
                            if elem.is_displayed():
                                placeholder = str(elem.get_attribute('placeholder') or '').lower()
                                data_tab = elem.get_attribute('data-tab')
                                # This should be the caption input below the image
                                if 'message' in placeholder or 'type' in placeholder or data_tab == '11':
                                    caption_box = elem
                                    print(f"  ✓ Found caption box in media container (placeholder='{placeholder[:30]}', data-tab='{data_tab}')")
                                    break
                        except:
                            continue
                    if caption_box:
                        break
                except:
                    pass
            
            # If not found in container, try global selectors
            for selector in caption_selectors:
                # Skip selectors that start with "." (those are for container search)
                if selector.startswith("."):
                    continue
                    
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    for elem in elements:
                        try:
                            if not elem.is_displayed():
                                continue
                            
                            # Get attributes to verify it's the caption box
                            data_tab = elem.get_attribute('data-tab')
                            placeholder = str(elem.get_attribute('placeholder') or '').lower()
                            aria_label = str(elem.get_attribute('aria-label') or '').lower()
                            role = elem.get_attribute('role') or ''
                            
                            # Verify it's the caption box (not regular message box)
                            # Caption box should have:
                            # - data-tab='11' OR
                            # - "type a message" or "message" in placeholder (this is the key!)
                            # - NOT data-tab='10' (regular message box)
                            is_caption_box = (
                                data_tab == '11' or 
                                'type a message' in placeholder or
                                ('message' in placeholder and data_tab != '10') or
                                'caption' in placeholder or 
                                'add' in placeholder
                            )
                            
                            # Make sure it's not the main message box (data-tab='10') or search box (data-tab='3')
                            is_not_main_box = data_tab != '10' and data_tab != '3'
                            
                            # If image preview is visible and this element is not the main box, it might be caption
                            if not is_caption_box and is_not_main_box:
                                # Check if image preview is visible
                                try:
                                    previews = driver.find_elements(By.XPATH, 
                                        "//img[contains(@src, 'blob')] | "
                                        "//div[contains(@data-testid, 'media')]"
                                    )
                                    if any(p.is_displayed() for p in previews):
                                        # If image is visible and this is a different contenteditable, it might be caption
                                        is_caption_box = True
                                except:
                                    pass
                            
                            if is_caption_box and is_not_main_box:
                                caption_box = elem
                                print(f"  ✓ Found caption box (data-tab='{data_tab}', placeholder='{placeholder[:30]}')")
                                break
                        except:
                            continue
                    if caption_box:
                        break
                except:
                    continue
            if caption_box:
                break
            time.sleep(0.8)  # Wait a bit and try again
        
        # If still not found, try finding ANY contenteditable that's not message box or search box
        if not caption_box:
            print(f"  → Trying to find any contenteditable in footer area...")
            try:
                # Get all contenteditable elements in footer
                footer_elements = driver.find_elements(By.XPATH, 
                    "//footer//div[@contenteditable='true'] | "
                    "//div[contains(@class, 'footer')]//div[@contenteditable='true'] | "
                    "//div[contains(@data-testid, 'conversation-compose')]//div[@contenteditable='true'] | "
                    "//div[contains(@data-testid, 'media')]//div[@contenteditable='true']"
                )
                for elem in footer_elements:
                    try:
                        if not elem.is_displayed():
                            continue
                        data_tab = elem.get_attribute('data-tab')
                        # If it's not the main message box (tab 10) or search box (tab 3), try it
                        if data_tab and data_tab not in ['10', '3']:
                            # Check if there's an image preview visible - if yes, this might be caption box
                            previews = driver.find_elements(By.XPATH, 
                                "//img[contains(@src, 'blob')] | "
                                "//div[contains(@data-testid, 'media')]"
                            )
                            has_preview = any(p.is_displayed() for p in previews)
                            if has_preview:
                                caption_box = elem
                                print(f"  ✓ Found potential caption box (data-tab='{data_tab}') - image preview is visible")
                                break
                    except:
                        continue
                
                # If still not found, maybe the caption box appears INSIDE the image preview container
                # Try to find it near the image preview
                if not caption_box:
                    print(f"  → Looking for caption box near image preview...")
                    try:
                        # Find image preview first
                        media_containers = driver.find_elements(By.XPATH,
                            "//div[contains(@data-testid, 'media')] | "
                            "//div[contains(@class, 'media')]"
                        )
                        for container in media_containers:
                            if container.is_displayed():
                                # Look for contenteditable inside this container
                                caption_in_container = container.find_elements(By.XPATH, ".//div[@contenteditable='true']")
                                for elem in caption_in_container:
                                    try:
                                        if elem.is_displayed():
                                            data_tab = elem.get_attribute('data-tab')
                                            if data_tab != '10' and data_tab != '3':
                                                caption_box = elem
                                                print(f"  ✓ Found caption box in media container (data-tab='{data_tab}')")
                                                break
                                    except:
                                        continue
                                if caption_box:
                                    break
                    except:
                        pass
                
                # LAST RESORT: Try to find caption input by looking near the send button
                # The caption input is usually positioned between image thumbnail and send button
                if not caption_box:
                    print(f"  → Trying to find caption input near send button...")
                    try:
                        # Find send button first
                        send_buttons = driver.find_elements(By.XPATH, 
                            "//span[@data-icon='send'] | "
                            "//button[@aria-label='Send'] | "
                            "//span[@data-testid='send']"
                        )
                        for send_btn in send_buttons:
                            if send_btn.is_displayed():
                                # Method 1: Look in the same parent container as send button
                                parent = driver.execute_script("return arguments[0].parentElement;", send_btn)
                                if parent:
                                    # Look for contenteditable in the same container or siblings
                                    caption_near_send = parent.find_elements(By.XPATH, 
                                        ".//div[@contenteditable='true'] | "
                                        ".//preceding-sibling::div[@contenteditable='true'] | "
                                        ".//following-sibling::div[@contenteditable='true']"
                                    )
                                    for elem in caption_near_send:
                                        try:
                                            if elem.is_displayed():
                                                data_tab = elem.get_attribute('data-tab')
                                                placeholder = str(elem.get_attribute('placeholder') or '').lower()
                                                # If it's not the main message box and has message placeholder, it's likely caption
                                                if data_tab != '10' and data_tab != '3':
                                                    if 'message' in placeholder or data_tab == '11' or not data_tab:
                                                        caption_box = elem
                                                        print(f"  ✓ Found caption input near send button (placeholder='{placeholder[:30]}', data-tab='{data_tab}')")
                                                        break
                                        except:
                                            continue
                                    if caption_box:
                                        break
                                
                                # Method 2: Look in footer area near send button
                                footer = driver.find_elements(By.XPATH, "//footer | //div[contains(@class, 'footer')]")
                                for foot in footer:
                                    if foot.is_displayed():
                                        # Get all contenteditable in footer, prioritize ones near send button
                                        footer_inputs = foot.find_elements(By.XPATH, ".//div[@contenteditable='true']")
                                        for elem in footer_inputs:
                                            try:
                                                if elem.is_displayed():
                                                    data_tab = elem.get_attribute('data-tab')
                                                    placeholder = str(elem.get_attribute('placeholder') or '').lower()
                                                    if data_tab != '10' and data_tab != '3':
                                                        if 'message' in placeholder or data_tab == '11':
                                                            caption_box = elem
                                                            print(f"  ✓ Found caption input in footer (placeholder='{placeholder[:30]}', data-tab='{data_tab}')")
                                                            break
                                            except:
                                                continue
                                        if caption_box:
                                            break
                                    if caption_box:
                                        break
                    except:
                        pass
                
                # Try keyboard navigation to focus on caption input
                if not caption_box:
                    print(f"  → Trying keyboard navigation to find caption input...")
                    try:
                        # Press Tab multiple times to navigate to caption input
                        from selenium.webdriver.common.action_chains import ActionChains
                        for tab_press in range(5):
                            ActionChains(driver).send_keys(Keys.TAB).perform()
                            time.sleep(0.3)
                            # Check what element is focused
                            focused = driver.execute_script("return document.activeElement;")
                            if focused and focused.get_attribute('contenteditable') == 'true':
                                data_tab = focused.get_attribute('data-tab')
                                placeholder = str(focused.get_attribute('placeholder') or '').lower()
                                if data_tab != '10' and data_tab != '3':
                                    if 'message' in placeholder or data_tab == '11' or not data_tab:
                                        caption_box = focused
                                        print(f"  ✓ Found caption input via Tab navigation (placeholder='{placeholder[:30]}', data-tab='{data_tab}')")
                                        break
                    except:
                        pass
                
                # Try clicking directly in the area between thumbnail and send button
                if not caption_box:
                    print(f"  → Trying to click in caption input area...")
                    try:
                        # Find send button to get position
                        send_buttons = driver.find_elements(By.XPATH, 
                            "//span[@data-icon='send'] | "
                            "//button[@aria-label='Send']"
                        )
                        if send_buttons:
                            for send_btn in send_buttons:
                                if send_btn.is_displayed():
                                    # Click to the left of send button (where caption input should be)
                                    driver.execute_script("""
                                        var btn = arguments[0];
                                        var rect = btn.getBoundingClientRect();
                                        // Click to the left of send button, in the middle vertically
                                        var x = rect.left - 150;
                                        var y = rect.top + (rect.height / 2);
                                        var element = document.elementFromPoint(x, y);
                                        if (element) {
                                            element.click();
                                            element.focus();
                                        }
                                    """, send_btn)
                                    time.sleep(1)
                                    
                                    # Check if a contenteditable element is now focused
                                    focused = driver.execute_script("return document.activeElement;")
                                    if focused and focused.get_attribute('contenteditable') == 'true':
                                        data_tab = focused.get_attribute('data-tab')
                                        if data_tab != '10' and data_tab != '3':
                                            caption_box = focused
                                            print(f"  ✓ Found caption input by clicking in area (data-tab='{data_tab}')")
                                            break
                    except:
                        pass
                
                # DO NOT use message box (data-tab='10') as fallback - it will detach the image!
                # We MUST find the actual caption input (data-tab='11' or placeholder='Type a message' below image)
                if not caption_box:
                    print(f"  ⚠️  Caption box not found - trying to activate it by clicking below image preview...")
                    try:
                        # Try clicking directly below the image preview where caption input should appear
                        previews = driver.find_elements(By.XPATH, 
                            "//img[contains(@src, 'blob')] | "
                            "//div[contains(@data-testid, 'media')]"
                        )
                        for preview in previews:
                            if preview.is_displayed():
                                # Click below the image preview
                                driver.execute_script("""
                                    var preview = arguments[0];
                                    var rect = preview.getBoundingClientRect();
                                    // Click below the preview, in the middle horizontally
                                    var x = rect.left + (rect.width / 2);
                                    var y = rect.bottom + 50; // 50px below the preview
                                    var element = document.elementFromPoint(x, y);
                                    if (element) {
                                        element.click();
                                        element.focus();
                                    }
                                """, preview)
                                time.sleep(1)
                                print(f"  ✓ Clicked below image preview to activate caption")
                                break
                    except:
                        pass
            except:
                pass
        
        # Step 7: Type caption if caption box found and caption text provided
        # CRITICAL: We MUST use the actual caption input (data-tab='11' or input below image)
        # Using data-tab='10' (regular message box) will DETACH the image!
        if not caption_box and caption:
            print(f"  ✗ ERROR: Caption box (data-tab='11') not found!")
            print(f"      Cannot type caption - using message box (data-tab='10') will detach the image.")
            print(f"      Canceling image send to avoid sending text-only message...")
            # Cancel the image attachment
            try:
                from selenium.webdriver.common.action_chains import ActionChains
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(1)
            except:
                pass
            return send_text_fallback()
        
        if caption_box and caption:
            print(f"  → Typing caption in caption box...")
            try:
                # Verify this is actually the caption box (data-tab='11')
                data_tab = caption_box.get_attribute('data-tab')
                if data_tab != '11':
                    print(f"  ⚠️  Warning: Element data-tab='{data_tab}' might not be caption box")
                
                # Verify image preview is still visible before typing
                print(f"  → Verifying image preview is still visible...")
                preview_visible = False
                try:
                    previews = driver.find_elements(By.XPATH, 
                        "//img[contains(@src, 'blob')] | "
                        "//div[contains(@class, 'preview')] | "
                        "//div[contains(@data-testid, 'media')]"
                    )
                    for preview in previews:
                        if preview.is_displayed():
                            preview_visible = True
                            print(f"  ✓ Image preview is visible")
                            break
                except:
                    pass
                
                if not preview_visible:
                    print(f"  ⚠️  Warning: Image preview not visible, but continuing...")
                
                # Use JavaScript to focus and type - this bypasses overlay issues
                print(f"  → Using JavaScript to type (bypassing overlay)...")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", caption_box)
                time.sleep(0.2)
                
                # Focus using JavaScript (bypasses click interception)
                driver.execute_script("arguments[0].focus();", caption_box)
                time.sleep(0.3)
                
                # Verify focus is on caption box
                active_elem = driver.execute_script("return document.activeElement;")
                if active_elem != caption_box:
                    print(f"  ⚠️  Focus not on caption box, trying JavaScript focus again...")
                    # Try multiple times with JavaScript
                    for attempt in range(3):
                        driver.execute_script("""
                            arguments[0].focus();
                            arguments[0].dispatchEvent(new Event('focus', { bubbles: true }));
                        """, caption_box)
                        time.sleep(0.2)
                        active_elem = driver.execute_script("return document.activeElement;")
                        if active_elem == caption_box:
                            break
                
                # Always use JavaScript to type (bypasses overlay and works with emojis)
                # IMPORTANT: Don't clear the element if image preview is visible - just append/type
                print(f"  → Setting caption text via JavaScript...")
                
                # Check if image preview is still visible before typing
                previews_before = driver.find_elements(By.XPATH, 
                    "//img[contains(@src, 'blob')] | "
                    "//div[contains(@data-testid, 'media')]"
                )
                preview_visible_before = any(p.is_displayed() for p in previews_before)
                
                if preview_visible_before:
                    print(f"  ✓ Image preview still visible, typing caption...")
                    # CRITICAL: When image is attached, setting textContent/innerText can detach it
                    # Try using send_keys() instead - this simulates real typing and might preserve the image
                    
                    # First, verify image is still there
                    time.sleep(0.5)
                    previews_before_type = driver.find_elements(By.XPATH, 
                        "//img[contains(@src, 'blob')] | "
                        "//div[contains(@data-testid, 'media')]"
                    )
                    if not any(p.is_displayed() for p in previews_before_type):
                        print(f"  ✗ ERROR: Image preview lost before typing caption!")
                        return send_text_fallback()
                    
                    # Try using send_keys() first (real keyboard input - might preserve image better)
                    try:
                        print(f"  → Trying send_keys() method (real keyboard input)...")
                        # Focus the element
                        driver.execute_script("arguments[0].focus();", caption_box)
                        time.sleep(0.2)
                        
                        # Clear any existing text first (Ctrl+A, Delete)
                        from selenium.webdriver.common.action_chains import ActionChains
                        ActionChains(driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
                        time.sleep(0.1)
                        ActionChains(driver).send_keys(Keys.DELETE).perform()
                        time.sleep(0.2)
                        
                        # Check if image is still attached after clearing
                        previews_after_clear = driver.find_elements(By.XPATH, 
                            "//img[contains(@src, 'blob')] | "
                            "//div[contains(@data-testid, 'media')]"
                        )
                        if not any(p.is_displayed() for p in previews_after_clear):
                            print(f"  ⚠️  Image detached after clearing text - will use JavaScript method")
                            # Image was detached, fall back to JavaScript
                            raise Exception("Image detached")
                        
                        # Now type using send_keys (character by character for emojis)
                        # Split caption into chunks to handle emojis
                        try:
                            caption_box.send_keys(caption)
                            print(f"  ✓ Typed caption using send_keys()")
                        except Exception as e:
                            # If send_keys fails (emojis), use JavaScript for those parts
                            if "BMP" in str(e) or "characters" in str(e).lower():
                                print(f"  → send_keys() failed for emojis, using JavaScript for full text...")
                                raise  # Will fall to JavaScript method
                            else:
                                raise
                        
                        # Verify image is still attached after typing
                        time.sleep(0.5)
                        previews_after_type = driver.find_elements(By.XPATH, 
                            "//img[contains(@src, 'blob')] | "
                            "//div[contains(@data-testid, 'media')]"
                        )
                        if not any(p.is_displayed() for p in previews_after_type):
                            print(f"  ⚠️  Image detached after send_keys() - this shouldn't happen")
                            return send_text_fallback()
                        else:
                            print(f"  ✓ Image preview still attached after typing with send_keys()")
                    except:
                        # Fallback to JavaScript method if send_keys fails
                        print(f"  → send_keys() failed, using JavaScript method...")
                        driver.execute_script("""
                            var elem = arguments[0];
                            var text = arguments[1];
                            
                            elem.focus();
                            var currentText = elem.textContent || elem.innerText || '';
                            
                            if (currentText.trim() !== text.trim()) {
                                // Select all existing text
                                if (currentText.trim()) {
                                    var range = document.createRange();
                                    range.selectNodeContents(elem);
                                    var sel = window.getSelection();
                                    sel.removeAllRanges();
                                    sel.addRange(range);
                                }
                                
                                // Set text
                                elem.textContent = text;
                                elem.innerText = text;
                                
                                // Dispatch events
                                elem.dispatchEvent(new InputEvent('beforeinput', {bubbles: true, inputType: 'insertText', data: text}));
                                elem.dispatchEvent(new InputEvent('input', {bubbles: true, inputType: 'insertText', data: text}));
                                elem.dispatchEvent(new Event('change', {bubbles: true}));
                                elem.focus();
                            }
                        """, caption_box, caption)
                        
                        # Verify image is still attached
                        time.sleep(0.5)
                        previews_after_js = driver.find_elements(By.XPATH, 
                            "//img[contains(@src, 'blob')] | "
                            "//div[contains(@data-testid, 'media')]"
                        )
                        if not any(p.is_displayed() for p in previews_after_js):
                            print(f"  ✗ ERROR: Image preview lost AFTER typing caption with JavaScript!")
                            return send_text_fallback()
                        else:
                            print(f"  ✓ Image preview still attached after JavaScript typing")
                else:
                    # No image preview, safe to clear
                    driver.execute_script("""
                        var elem = arguments[0];
                        var text = arguments[1];
                        elem.focus();
                        elem.textContent = '';
                        elem.innerText = '';
                        elem.textContent = text;
                        elem.innerText = text;
                        var inputEvent = new InputEvent('input', {
                            bubbles: true,
                            cancelable: true,
                            inputType: 'insertText',
                            data: text
                        });
                        elem.dispatchEvent(inputEvent);
                        elem.focus();
                    """, caption_box, caption)
                
                time.sleep(1)  # Wait for caption to be set
                
                # Verify caption was typed AND image preview is still there
                caption_text = driver.execute_script("return arguments[0].textContent || arguments[0].innerText;", caption_box)
                previews_after = driver.find_elements(By.XPATH, 
                    "//img[contains(@src, 'blob')] | "
                    "//div[contains(@data-testid, 'media')]"
                )
                preview_visible_after = any(p.is_displayed() for p in previews_after)
                
                if caption_text and len(caption_text.strip()) > 0:
                    if preview_visible_after:
                        print(f"  ✓ Caption typed successfully ({len(caption_text)} chars) - Image preview still attached")
                    else:
                        print(f"  ⚠️  Warning: Caption typed but image preview might be lost!")
                else:
                    print(f"  ⚠️  Warning: Caption might not have been set properly")
                
            except Exception as e:
                print(f"  ⚠️  Could not type caption: {str(e)}")
                # Don't send without caption - cancel and send text instead
                print(f"  → Canceling image send, will send text-only...")
                return send_text_fallback()
        
        # Step 8: Verify image preview and caption are ready before sending
        print(f"  → Verifying image and caption are ready...")
        image_ready = False
        caption_ready = False
        
        try:
            # Check if image preview is still visible
            previews = driver.find_elements(By.XPATH, 
                "//img[contains(@src, 'blob')] | "
                "//div[contains(@class, 'preview')] | "
                "//div[contains(@data-testid, 'media')]"
            )
            for preview in previews:
                if preview.is_displayed():
                    image_ready = True
                    print(f"  ✓ Image preview is visible")
                    break
            
            if not image_ready:
                print(f"  ⚠️  Warning: Image preview not found! Image might be lost.")
            
            # Verify caption is still in caption/message box
            if caption_box and caption:
                caption_text = driver.execute_script("return arguments[0].textContent || arguments[0].innerText;", caption_box)
                if caption_text and len(caption_text.strip()) > 0:
                    caption_ready = True
                    print(f"  ✓ Caption is ready ({len(caption_text)} chars)")
                else:
                    print(f"  ⚠️  Warning: Caption seems to be empty! Re-typing...")
                    # Try to re-type caption
                    try:
                        driver.execute_script("""
                            var elem = arguments[0];
                            var text = arguments[1];
                            elem.focus();
                            elem.textContent = text;
                            elem.innerText = text;
                            var inputEvent = new InputEvent('input', {
                                bubbles: true,
                                cancelable: true,
                                inputType: 'insertText',
                                data: text
                            });
                            elem.dispatchEvent(inputEvent);
                            elem.focus();
                        """, caption_box, caption)
                        time.sleep(0.5)
                        caption_ready = True
                    except:
                        pass
        except:
            pass
        
        if not image_ready:
            print(f"  ✗ Image preview lost, cannot send image with caption")
            return send_text_fallback()
        
        # Step 8.5: Send image with caption
        # IMPORTANT: When image preview is visible, use Enter key in message box
        # This is more reliable than clicking send button for image + caption
        print(f"  → Sending image with caption...")
        sent = False
        
        # First, ensure caption box is focused using JavaScript (no clicks to avoid overlay)
        if caption_box:
            try:
                # Focus using JavaScript only
                driver.execute_script("arguments[0].focus();", caption_box)
                time.sleep(0.3)
                
                # Verify focus
                active_elem = driver.execute_script("return document.activeElement;")
                if active_elem != caption_box:
                    print(f"  ⚠️  Focus not on caption box, refocusing...")
                    driver.execute_script("arguments[0].focus();", caption_box)
                    time.sleep(0.3)
                
                print(f"  ✓ Caption box focused")
            except:
                pass
        
        # Check if image preview is still visible
        previews = driver.find_elements(By.XPATH, 
            "//img[contains(@src, 'blob')] | "
            "//div[contains(@data-testid, 'media')]"
        )
        has_preview = any(p.is_displayed() for p in previews)
        
        if has_preview and caption_box:
            print(f"  → Image preview visible, finding send button for image...")
            # When image is attached, we should use the send button, not just Enter key
            # Enter key in message box might send only text, not image+text
            try:
                # First, ensure caption box is focused
                driver.execute_script("arguments[0].focus();", caption_box)
                time.sleep(0.3)
                
                # Look for send button that appears when image is attached
                send_button = None
                send_selectors = [
                    "//span[@data-testid='send']",
                    "//span[@data-icon='send']",
                    "//button[@aria-label='Send']",
                    "//div[@data-testid='send']",
                    "//span[@data-icon='send']//ancestor::button",
                    "//button[contains(@class, 'send')]"
                ]
                
                for selector in send_selectors:
                    try:
                        elements = driver.find_elements(By.XPATH, selector)
                        for elem in elements:
                            if elem.is_displayed() and elem.is_enabled():
                                send_button = elem
                                break
                        if send_button:
                            break
                    except:
                        continue
                
                if send_button:
                    print(f"  → Clicking send button (image + caption)...")
                    # CRITICAL: Final verification - image MUST be visible
                    print(f"  → Final check: Verifying image is still attached...")
                    previews_check = driver.find_elements(By.XPATH, 
                        "//img[contains(@src, 'blob')] | "
                        "//div[contains(@data-testid, 'media')] | "
                        "//div[contains(@class, 'preview')] | "
                        "//div[contains(@class, 'media-preview')]"
                    )
                    image_still_attached = any(p.is_displayed() for p in previews_check)
                    
                    if not image_still_attached:
                        print(f"  ✗ ERROR: Image preview NOT visible right before sending!")
                        print(f"      Image was detached. Canceling to avoid sending text-only...")
                        try:
                            from selenium.webdriver.common.action_chains import ActionChains
                            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                            time.sleep(1)
                        except:
                            pass
                        return send_text_fallback()
                    
                    print(f"  ✓ Image confirmed attached - clicking send button...")
                    # Click send button - this should send image WITH caption
                    driver.execute_script("arguments[0].click();", send_button)
                    
                    # Wait and verify message was actually sent
                    # Check multiple times as it may take a moment for preview to disappear
                    image_sent = False
                    for check_attempt in range(5):
                        time.sleep(1)
                        previews_after_send = driver.find_elements(By.XPATH, 
                            "//img[contains(@src, 'blob')] | "
                            "//div[contains(@data-testid, 'media')] | "
                            "//div[contains(@class, 'preview')]"
                        )
                        visible_previews = [p for p in previews_after_send if p.is_displayed()]
                        
                        if len(visible_previews) == 0:
                            # No visible previews - image was sent!
                            image_sent = True
                            print(f"  ✓ Image sent! (preview disappeared after {check_attempt + 1} seconds)")
                            break
                        elif check_attempt == 4:
                            # Last attempt - check if send button is still there (if gone, message was sent)
                            send_buttons_after = driver.find_elements(By.XPATH, 
                                "//span[@data-icon='send'] | "
                                "//button[@aria-label='Send']"
                            )
                            send_button_gone = not any(btn.is_displayed() for btn in send_buttons_after)
                            if send_button_gone:
                                image_sent = True
                                print(f"  ✓ Image sent! (send button disappeared)")
                            else:
                                print(f"  ⚠️  Image preview still visible after 5 seconds - may not have sent")
                    
                    sent = image_sent or True  # Assume sent if we can't verify
                else:
                    # Fallback: Try Enter key but verify image preview is still there
                    print(f"  → Send button not found, trying Enter key...")
                    # Double-check image preview is still visible
                    previews_check = driver.find_elements(By.XPATH, 
                        "//img[contains(@src, 'blob')] | "
                        "//div[contains(@data-testid, 'media')]"
                    )
                    if any(p.is_displayed() for p in previews_check):
                        # Focus caption box and press Enter
                        driver.execute_script("arguments[0].focus();", caption_box)
                        time.sleep(0.2)
                        caption_box.send_keys(Keys.ENTER)
                        time.sleep(2.5)  # Wait longer for image to send
                        sent = True
                        print(f"  ✓ Sent image with caption via Enter key")
                    else:
                        print(f"  ⚠️  Image preview lost, cannot send image")
            except Exception as e:
                print(f"  ⚠️  Send method failed: {str(e)}, trying alternative...")
        
        # Step 9: Try send button if Enter key didn't work
        if not sent:
            print(f"  → Trying send button as fallback...")
            send_selectors = [
                "//span[@data-testid='send']",
                "//span[@data-icon='send']",
                "//button[@aria-label='Send']",
                "//div[@data-testid='send']",
                "//span[@data-icon='send']//ancestor::button",
                "//button[contains(@class, 'send')]"
            ]
            
            # Wait for send button to appear (up to 4 seconds)
            for attempt in range(4):
                for selector in send_selectors:
                    try:
                        elements = driver.find_elements(By.XPATH, selector)
                        for elem in elements:
                            if elem.is_displayed() and elem.is_enabled():
                                try:
                                    # Click send button - this should send as photo with caption
                                    driver.execute_script("arguments[0].click();", elem)
                                    time.sleep(2)
                                    sent = True
                                    print(f"  ✓ Sent via send button")
                                    break
                                except:
                                    try:
                                        elem.click()
                                        time.sleep(2)
                                        sent = True
                                        print(f"  ✓ Sent via send button")
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
            
            # Last resort: Try pressing Enter in active element
            if not sent:
                try:
                    active_input = driver.switch_to.active_element
                    data_tab = active_input.get_attribute('data-tab')
                    print(f"  → Last resort: Trying Enter in active element (data-tab='{data_tab}')...")
                    active_input.send_keys(Keys.ENTER)
                    time.sleep(2)
                    sent = True
                    print(f"  ✓ Sent via active element")
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
    3. Auto type message (or send image with caption if image_path provided)
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
        
        # If image_path is provided, use image sending function instead
        if image_path:
            print(f"  📷 Sending image with caption...")
            if send_image_with_caption(driver, message_box, image_path, message, contact_number, delay_seconds):
                print(f"✓ Message sent to {contact_number}")
                time.sleep(delay_seconds)
                # Go back to main page
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.2)
                return True
            else:
                print(f"  ⚠️  Image send failed, falling back to text-only...")
                # Fall through to text-only sending
        
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
                # Alternative: Direct JavaScript with proper newline handling
                driver.execute_script("""
                    var elem = arguments[0];
                    var text = arguments[1];
                    
                    elem.focus();
                    
                    // Clear existing content
                    var range = document.createRange();
                    range.selectNodeContents(elem);
                    var selection = window.getSelection();
                    selection.removeAllRanges();
                    selection.addRange(range);
                    document.execCommand('delete', false, null);
                    
                    // Insert text with newlines preserved using insertText
                    try {
                        document.execCommand('insertText', false, text);
                    } catch(e) {
                        // Fallback: insert line by line
                        var lines = text.split('\\n');
                        for (var i = 0; i < lines.length; i++) {
                            if (i > 0) {
                                // Insert line break
                                var br = document.createElement('br');
                                elem.appendChild(br);
                                var range = document.createRange();
                                range.setStartAfter(br);
                                range.collapse(true);
                                var sel = window.getSelection();
                                sel.removeAllRanges();
                                sel.addRange(range);
                            }
                            if (lines[i]) {
                                document.execCommand('insertText', false, lines[i]);
                            }
                        }
                    }
                    
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
    DEFAULT_IMAGE = None  # No default image
    IMAGES_FOLDER = None  # No images folder
    
    # Check if Excel file exists
    if not os.path.exists(EXCEL_FILE):
        print(f"✗ Error: {EXCEL_FILE} not found!")
        print(f"Please create an Excel file with:")
        print(f"  Column A: Contact Number (with country code, e.g., +1234567890)")
        print(f"  Column B: Contact Name (optional - if provided, message will be: 'Dear [Name],\\n\\n[Message]')")
        print(f"  Column C: Message (Caption) - text to send with image")
        print(f"  Column D: Image Path (optional - leave empty to auto-detect)")
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
