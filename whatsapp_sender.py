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
        """
        IMPORTANT (Mode 2): We do NOT send a separate text message if image sending fails.
        Returning False here prevents the "caption sent separately, image still attached" behavior.
        """
        print(f"  ⚠️  Image flow failed - text-only fallback is disabled (Mode 2).")
        try:
            # Best-effort: close any open media composer / attachment UI
            from selenium.webdriver.common.action_chains import ActionChains
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.8)
        except:
            pass
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
        
        # Step 3.5: STRICT "Photos & videos" selection (avoid Sticker Maker)
        # Key rule: DO NOT use the first <input type="file"> and DO NOT click random divs.
        # Prefer WhatsApp's attach button by data-testid; fallback to 2nd menu item but click its clickable ancestor.
        print(f"  → Selecting 'Photos & videos' option (strict)...")
        time.sleep(0.8)  # menu animation settle

        def _click(el):
            try:
                driver.execute_script("arguments[0].click();", el)
                return True
            except:
                try:
                    el.click()
                    return True
                except:
                    return False

        def _clickable_ancestor(el):
            try:
                return el.find_element(By.XPATH, "./ancestor::button[1] | ./ancestor::div[@role='button'][1]")
            except:
                return None

        # Try official data-testid selectors first (language independent)
        selected = False
        for sel in ["//*[@data-testid='attach-photo']",
                    "//*[@data-testid='attach-image']",
                    "//*[@data-testid='attach-media']"]:
            try:
                candidates = driver.find_elements(By.XPATH, sel)
                for c in candidates:
                    if c.is_displayed() and c.is_enabled():
                        if _click(c):
                            selected = True
                            print(f"  ✓ Selected Photos & videos via data-testid='{c.get_attribute('data-testid')}'")
                            break
                if selected:
                    break
            except:
                pass

        # Fallback: click the 2nd visible menu item (Document is 1st, Photos & videos is 2nd)
        if not selected:
            try:
                menu_items = driver.find_elements(By.XPATH, "//div[@role='menuitem']")
                visible_items = [m for m in menu_items if m.is_displayed()]
                if len(visible_items) >= 2:
                    target = visible_items[1]
                    click_target = _clickable_ancestor(target) or target
                    if _click(click_target):
                        selected = True
                        print(f"  ✓ Selected Photos & videos via menu item #2 (clicked ancestor button)")
            except:
                pass

        if not selected:
            print(f"  ✗ ERROR: Could not select Photos & videos")
            return send_text_fallback()

        # Step 4: Pick the correct media <input type='file'> (reject accept='' and webp)
        time.sleep(0.8)
        file_input = None
        try:
            inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
            scored = []
            for idx, inp in enumerate(inputs):
                try:
                    accept_attr = (inp.get_attribute('accept') or '').lower()
                    multiple = inp.get_attribute('multiple')
                    data_testid = (inp.get_attribute('data-testid') or '').lower()
                    name_attr = (inp.get_attribute('name') or '').lower()

                    score = 0
                    # Best: media picker supports videos + multiple selection
                    if 'video' in accept_attr:
                        score += 100
                    if multiple is not None:
                        score += 50
                    if 'image' in accept_attr:
                        score += 20

                    # Hard rejects / penalties
                    if accept_attr == '':
                        score -= 200  # very often sticker/new-sticker input
                    if 'webp' in accept_attr:
                        score -= 300
                    if 'sticker' in accept_attr or 'sticker' in data_testid or 'sticker' in name_attr:
                        score -= 500

                    scored.append((score, idx, inp, accept_attr, bool(multiple)))
                except:
                    continue

            if scored:
                scored.sort(key=lambda x: x[0], reverse=True)
                best = scored[0]
                file_input = best[2]
                print(f"  → File inputs found: {len(scored)} (best score={best[0]}, accept='{best[3]}', multiple={best[4]})")
        except Exception as e:
            print(f"  ✗ ERROR finding file input: {str(e)}")
            return send_text_fallback()

        if not file_input:
            print(f"  ✗ ERROR: Could not find a valid media file input for Photos & videos")
            return send_text_fallback()
        
        # Step 5: PRE-UPLOAD STICKER CHECK (detect sticker mode BEFORE uploading)
        print(f"  → Verifying photo mode (not sticker)...")
        try:
            sticker_ui = driver.find_elements(By.XPATH,
                "//span[contains(text(),'Send sticker')] | "
                "//button[contains(@aria-label,'sticker')] | "
                "//div[contains(@aria-label, 'sticker') and contains(@aria-label, 'send')] | "
                "//span[@data-icon='sticker']"
            )
            if any(x.is_displayed() for x in sticker_ui):
                print(f"  ✗ ERROR: STICKER MODE detected before upload!")
                print(f"      Keyboard navigation failed. Canceling...")
                from selenium.webdriver.common.action_chains import ActionChains
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(1)
                return send_text_fallback()
            else:
                print(f"  ✓ Photo mode confirmed (no sticker UI detected)")
        except Exception as e:
            print(f"  ⚠️  Could not verify mode: {str(e)}, proceeding anyway...")
        
        # Step 6: Upload image (no sanitization; use original file)
        print(f"  → Preparing image for upload...")
        try:
            # Verify file path exists
            if not os.path.exists(image_path):
                print(f"  ✗ ERROR: Image file not found: {image_path}")
                return send_text_fallback()

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
        
        # Caption input in the media composer is in the footer ("Type a message").
        # Use an explicit wait for the correct box to become clickable.
        print(f"  → Finding caption/message box in media composer footer...")
        try:
            caption_box = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((
                    By.XPATH,
                    "//footer//div[@contenteditable='true' and (@data-tab='10' or @data-tab='11')]"
                ))
            )
            try:
                dt = caption_box.get_attribute('data-tab')
                al = (caption_box.get_attribute('aria-label') or '')[:40]
                ap = (caption_box.get_attribute('aria-placeholder') or '')[:40]
                print(f"  ✓ Using footer caption box (data-tab='{dt}', aria-label='{al}', aria-placeholder='{ap}')")
            except:
                print(f"  ✓ Using footer caption box")
        except Exception:
            # Fallback to previous heuristic scan
            try:
                for scan_attempt in range(10):
                    time.sleep(0.5)
                    candidates = driver.find_elements(By.XPATH, "//footer//div[@contenteditable='true']")
                    best = None
                    best_score = -10_000
                    for elem in candidates:
                        try:
                            if not elem.is_displayed():
                                continue
                            data_tab = (elem.get_attribute('data-tab') or '')
                            aria_label = (elem.get_attribute('aria-label') or '')
                            aria_placeholder = (elem.get_attribute('aria-placeholder') or '')
                            placeholder = (elem.get_attribute('placeholder') or '')
                            score = 0
                            text_hint = (aria_label + " " + aria_placeholder + " " + placeholder).lower()
                            if "type a message" in text_hint:
                                score += 1000
                            if data_tab == '10':
                                score += 200
                            if data_tab == '3' or "search" in text_hint:
                                score -= 1000
                            if score > best_score:
                                best_score = score
                                best = elem
                        except:
                            continue
                    if best:
                        caption_box = best
                        print(f"  ✓ Using footer caption box (fallback scan)")
                        break
            except:
                pass
        
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
                            
                            # Check if image preview is visible first
                            has_image_preview = False
                            try:
                                previews = driver.find_elements(By.XPATH, 
                                    "//img[contains(@src, 'blob')] | "
                                    "//div[contains(@data-testid, 'media')]"
                                )
                                has_image_preview = any(p.is_displayed() for p in previews)
                            except:
                                pass
                            
                            # Caption box logic:
                            # - If image attached: data-tab='10' IS the caption box
                            # - If no image: data-tab='11' or other caption-specific elements
                            is_caption_box = (
                                (data_tab == '10' and has_image_preview) or  # Message box becomes caption when image attached
                                data_tab == '11' or 
                                'caption' in placeholder or 
                                ('message' in placeholder and data_tab not in ['3', '10'])
                            )
                            
                            # Make sure it's not the search box
                            is_not_search_box = data_tab != '3'
                            
                            # Use it if it's a valid caption box and not the search box
                            if is_caption_box and is_not_search_box:
                                caption_box = elem
                                print(f"  ✓ Found caption box (data-tab='{data_tab}', has_image={has_image_preview}, placeholder='{placeholder[:30]}')")
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
                
                # If caption box still not found, we'll try data-tab='10' as fallback (it's the caption input when image is attached)
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
        
        # Step 7: If caption box not found yet, try using data-tab='10' (when image is attached, it becomes the caption input)
        if not caption_box and caption:
            print(f"  → Caption box not found in scans, looking for data-tab='10' with image attached...")
            # When image is attached, data-tab='10' IS the caption input
            try:
                # Check if image is still attached
                previews = driver.find_elements(By.XPATH, "//img[contains(@src, 'blob')] | //div[contains(@data-testid, 'media')]")
                has_preview = any(p.is_displayed() for p in previews)
                
                if has_preview:
                    # Find data-tab='10' - it's the caption input when image is attached
                    message_boxes = driver.find_elements(By.XPATH, "//footer//div[@contenteditable='true'][@data-tab='10']")
                    for box in message_boxes:
                        if box.is_displayed():
                            caption_box = box
                            print(f"  ✓ Found caption input: data-tab='10' (message box becomes caption when image attached)")
                            break
                
                if not caption_box:
                    print(f"  ✗ ERROR: Could not find caption input!")
                    print(f"      Canceling image send...")
                    # Cancel the image attachment
                    try:
                        from selenium.webdriver.common.action_chains import ActionChains
                        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                        time.sleep(1)
                    except:
                        pass
                    return send_text_fallback()
            except Exception as e:
                print(f"  ✗ ERROR finding caption input: {str(e)}")
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
                    print(f"  ✓ Image preview still visible, typing caption in footer message box...")
                    # Use the same robust hybrid typing used for normal messages:
                    # - Shift+Enter between lines
                    # - JS insertText for emoji lines
                    # This avoids "stuck" caption box behavior.
                    if not set_message_text_js(driver, caption_box, caption):
                        print(f"  ⚠️  Could not type caption with hybrid method, trying JS fallback...")
                        driver.execute_script("""
                            var elem = arguments[0];
                            var text = arguments[1] || '';
                            elem.scrollIntoView({block:'center'});
                            elem.focus();
                            // Clear existing
                            try {
                                document.execCommand('selectAll', false, null);
                                document.execCommand('delete', false, null);
                            } catch(e) {}
                            try { document.execCommand('insertText', false, text); } catch(e) {}
                            try { elem.dispatchEvent(new InputEvent('input', {bubbles:true, cancelable:true})); } catch(e) {}
                            elem.focus();
                        """, caption_box, caption)
                    time.sleep(0.6)
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
                print(f"  → Canceling image send (Mode 2) - NOT sending text-only...")
                try:
                    from selenium.webdriver.common.action_chains import ActionChains
                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                    time.sleep(0.8)
                except:
                    pass
                return False
        
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
                    # Caption might be in WhatsApp's internal format (wrapped in spans, etc.)
                    # Assume it's already typed successfully if we got here
                    print(f"  → Caption verification skipped (already typed successfully)")
                    caption_ready = True
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
            print(f"  → Image preview visible, sending via media composer...")
            # IMPORTANT: Never press Enter as primary send here; it can send text-only and leave media attached.
            try:
                driver.execute_script("arguments[0].focus();", caption_box)
                time.sleep(0.2)

                def _find_media_send_button():
                    """
                    Find the *media composer* send button (not the normal chat send button).
                    Key detail: search relative to the current caption box's footer to avoid wrong matches.
                    """
                    try:
                        footer = caption_box.find_element(By.XPATH, "./ancestor::footer[1]")
                    except Exception:
                        footer = None

                    roots = [footer] if footer is not None else [driver]

                    rel_selectors = [
                        ".//button[@aria-label='Send']",
                        ".//span[@data-icon='send']/ancestor::button[1]",
                        ".//span[@data-testid='send']/ancestor::button[1]",
                        ".//span[@data-icon='send']/ancestor::*[@role='button'][1]",
                        ".//span[@data-testid='send']/ancestor::*[@role='button'][1]",
                        ".//*[@data-testid='send']/ancestor::button[1]",
                        ".//*[@data-testid='send']/ancestor::*[@role='button'][1]",
                    ]

                    for root in roots:
                        for sel in rel_selectors:
                            try:
                                for el in root.find_elements(By.XPATH, sel):
                                    if el.is_displayed() and el.is_enabled():
                                        return el
                            except Exception:
                                continue
                    return None

                def _composer_open():
                    """
                    Detect whether the media composer overlay is still open.
                    DO NOT use the chat's 'Add file' button (it exists even when composer is closed).
                    """
                    # 1) If caption_box is still attached & visible, composer is open
                    try:
                        if caption_box and caption_box.is_displayed():
                            return True
                    except Exception:
                        pass

                    # 2) Look for a dialog-like media overlay with a blob preview
                    try:
                        overlay_previews = driver.find_elements(
                            By.XPATH,
                            "//div[@role='dialog']//img[contains(@src,'blob')] | "
                            "//div[@role='dialog']//div[contains(@data-testid,'media')]"
                        )
                        if any(p.is_displayed() for p in overlay_previews):
                            return True
                    except Exception:
                        pass

                    return False

                # Try sending up to 3 times, verify by composer closing
                for send_try in range(3):
                    send_button = _find_media_send_button()
                    if not send_button:
                        print(f"  ⚠️  Media send button not found (try {send_try+1}/3)")
                        time.sleep(0.6)
                        continue

                    print(f"  → Clicking media send button (try {send_try+1}/3)...")
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block:'center',inline:'center'});", send_button)
                    except Exception:
                        pass
                    try:
                        driver.execute_script("arguments[0].click();", send_button)
                    except Exception:
                        try:
                            send_button.click()
                        except Exception:
                            pass

                    # The send button may morph into a spinner/disabled state after click.
                    # So: wait longer for composer to close, and only re-click if still open.
                    closed = False
                    for _ in range(120):  # up to 60s
                        time.sleep(0.5)
                        if not _composer_open():
                            closed = True
                            break
                    if closed:
                        sent = True
                        print(f"  ✓ Image sent! (media composer closed)")
                        break
                    else:
                        print(f"  ⚠️  Still in media composer after click (may still be uploading) - retrying send...")

                # If still not sent, cancel attachment BEFORE falling back to text-only
                if not sent and _composer_open():
                    try:
                        from selenium.webdriver.common.action_chains import ActionChains
                        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                        time.sleep(0.8)
                        print(f"  → Canceled attachment (Escape) before fallback")
                    except:
                        pass
            except Exception as e:
                print(f"  ⚠️  Send method failed: {str(e)}")
        
        # Step 9: No global send fallbacks.
        # IMPORTANT: Clicking generic send buttons (or pressing Enter) can send text-only and leave media attached.
        # If we couldn't send via the media composer send button + close verification, treat as failure.
        
        # Step 10: Verify and cleanup
        if sent:
            time.sleep(3)  # Wait for WhatsApp to fully process the image
            clear_attachment_preview(driver)
            time.sleep(1)
            
            print(f"✓ Image with caption sent to {contact_number}")
            time.sleep(delay_seconds)
            return True
        else:
            print(f"  ✗ Could not send image (send button not found)")
            # Close media composer if still open, but do NOT send text-only
            try:
                from selenium.webdriver.common.action_chains import ActionChains
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.8)
            except:
                pass
            return False
        
    except Exception as e:
        print(f"  ⚠️  Error sending image: {str(e)}")
        try:
            from selenium.webdriver.common.action_chains import ActionChains
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(0.8)
        except:
            pass
        return False


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
                # IMPORTANT (Mode 2): Never send caption as a separate text message if image send failed.
                # This avoids "text sent, image still attached" behavior.
                print(f"  ✗ Image send failed - NOT sending text-only fallback (Mode 2).")
                # Try to close any open media composer to avoid leaving attachments stuck
                try:
                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                    time.sleep(0.5)
                except:
                    pass
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


def send_bulk_messages(excel_file_path, delay_seconds=15, start_from=0, default_image=None):
    """
    Send messages to all contacts in the Excel file using Selenium
    
    Args:
        excel_file_path: Path to the XLSX file
        delay_seconds: Delay between each message (to avoid rate limiting)
        start_from: Index to start from (useful for resuming)
        default_image: Optional path to a single image to send to ALL contacts (Mode 2)
                      If None, only text messages are sent (Mode 1)
    """
    contacts = read_contacts_from_excel(excel_file_path)
    
    if not contacts:
        print("No contacts found. Please check your Excel file.")
        return
    
    # Check if default image exists (Mode 2)
    if default_image:
        excel_dir = os.path.dirname(os.path.abspath(excel_file_path))
        if not os.path.isabs(default_image):
            default_image = os.path.join(excel_dir, default_image)
        if os.path.exists(default_image):
            print(f"✓ Image found: {os.path.basename(default_image)}")
            print(f"  📷 Mode 2: Will send image with caption to ALL contacts")
        else:
            print(f"⚠️  Image not found: {default_image}")
            print(f"  ⚠️  Switching to Mode 1 (text only)")
            default_image = None
    else:
        print(f"  📝 Mode 1: Text messages only (no images)")
    
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
            print(f"📷 Active Mode: Mode 2 (Photo with caption)")
            print(f"   Image: {os.path.basename(default_image)}")
        else:
            print(f"📝 Active Mode: Mode 1 (Text messages only)")
        
        print(f"⚠️  Keep the browser window open and don't close it!\n")
        
        successful = 0
        failed = 0
        
        for index, contact in enumerate(contacts[start_from:], start=start_from):
            print(f"\n[{index + 1}/{len(contacts)}] Sending to {contact['number']}...")
            
            # Determine image path based on mode
            if default_image:
                # Mode 2: Photo with caption (same photo for all)
                image_path = default_image
                if index == start_from:  # Only print once
                    print(f"  📷 Using image: {os.path.basename(image_path)}")
            else:
                # Mode 1: Text messages only
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
    
    # IMAGE CONFIGURATION - Interactive Mode Selection
    print("\n" + "="*50)
    print("SELECT MODE:")
    print("="*50)
    print("Mode 1: Text messages only (no images)")
    print("Mode 2: Photo with caption (same photo for all contacts)")
    print("="*50)
    print()
    
    # Get mode selection from user
    while True:
        mode_input = input("Enter mode (1 or 2): ").strip()
        if mode_input == "1":
            # Mode 1: Text messages only
            DEFAULT_IMAGE = None
            print("\n✓ Mode 1 selected: Text messages only")
            break
        elif mode_input == "2":
            # Mode 2: Photo with caption using safari_promo.jpg
            image_file = "pexels-karola-g-4016579.jpg"
            
            # Check if image exists
            if os.path.exists(image_file):
                DEFAULT_IMAGE = image_file
                print(f"\n✓ Mode 2 selected: Photo with caption")
                print(f"✓ Image found: {image_file}")
                break
            else:
                print(f"\n⚠️  Error: Image file '{image_file}' not found!")
                print(f"    Switching to Mode 1 (text only)")
                DEFAULT_IMAGE = None
                break
        else:
            print("Invalid input. Please enter 1 or 2.")
    
    # Display active mode
    if DEFAULT_IMAGE:
        print(f"\n{'='*50}")
        print(f"📷 Active Mode: Mode 2 (Photo with caption)")
        print(f"   Image: {DEFAULT_IMAGE}")
        print(f"{'='*50}")
    else:
        print(f"\n{'='*50}")
        print(f"📝 Active Mode: Mode 1 (Text messages only)")
        print(f"{'='*50}")
    
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
        send_bulk_messages(EXCEL_FILE, DELAY_SECONDS, START_FROM, DEFAULT_IMAGE)
    except KeyboardInterrupt:
        print("\n\n⚠️  Process cancelled by user")
    except Exception as e:
        print(f"\n✗ Error: {str(e)}")
