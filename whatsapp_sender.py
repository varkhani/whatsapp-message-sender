"""
Simple WhatsApp Message Sender using Selenium
Reads contact numbers and messages from XLSX file and sends WhatsApp messages
More reliable than pywhatkit for bulk messaging

🔒 SAFETY FEATURES FOR WHATSAPP BUSINESS:
- Daily limit: 200 messages (adjustable)
- Hourly limit: 40 messages (adjustable)
- Smart delays: 15-45 seconds between messages
- Auto-breaks: Every 50 messages, 5-10 minute break
- Time restrictions: 9 AM - 9 PM only
- Progress tracking: Resume from where you left off
"""

import openpyxl
import time
import os
import json
import random
import pyautogui
from datetime import datetime, timedelta
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

# ============================================================
# SAFETY CONFIGURATION (Optimized for WhatsApp Business)
# ============================================================
MAX_MESSAGES_PER_DAY = 300      # WhatsApp Business allows more messages (increased from 200)
MAX_MESSAGES_PER_HOUR = 50      # Higher hourly limit for Business (increased from 40)
MIN_DELAY_BETWEEN_MESSAGES = 10  # Minimum seconds between messages (reduced from 15 for faster sending)
MAX_DELAY_BETWEEN_MESSAGES = 25  # Maximum seconds between messages (reduced from 45 for faster sending)
MESSAGES_BEFORE_BREAK = 60      # Take a break after every 60 messages (increased from 50)
BREAK_DURATION_MIN = 300        # Minimum break duration (5 minutes)
BREAK_DURATION_MAX = 600        # Maximum break duration (10 minutes)
ACTIVE_HOURS_START = 6          # Start sending from 6 AM
ACTIVE_HOURS_END = 22           # Stop sending after 10 PM
PROGRESS_FILE = "whatsapp_progress.json"  # File to track progress

# ============================================================
# DEBUG CONFIGURATION
# ============================================================
DEBUG_MODE = False  # Set to True for detailed debugging logs, False for clean user output

def debug_print(message):
    """Print message only if DEBUG_MODE is enabled"""
    if DEBUG_MODE:
        print(message)


# ============================================================
# SAFETY MANAGER CLASS
# ============================================================
class SafetyManager:
    """
    Manages message sending safety to prevent WhatsApp bans.
    Tracks message counts, enforces delays, and manages breaks.
    """
    
    def __init__(self, progress_file=PROGRESS_FILE):
        self.progress_file = progress_file
        self.progress = self.load_progress()
    
    def load_progress(self):
        """Load progress from JSON file"""
        if os.path.exists(self.progress_file):
            try:
                with open(self.progress_file, 'r') as f:
                    data = json.load(f)
                    # Reset daily count if it's a new day
                    last_date = data.get('last_message_date', '')
                    if last_date != datetime.now().strftime('%Y-%m-%d'):
                        data['messages_today'] = 0
                        data['last_message_date'] = datetime.now().strftime('%Y-%m-%d')
                    return data
            except Exception as e:
                print(f"⚠️  Could not load progress file: {e}")
        
        return {
            'messages_today': 0,
            'messages_this_hour': 0,
            'last_message_time': None,
            'last_message_date': datetime.now().strftime('%Y-%m-%d'),
            'last_hour': datetime.now().hour,
            'total_sent': 0,
            'last_contact_index': 0
        }
    
    def save_progress(self):
        """Save progress to JSON file"""
        try:
            with open(self.progress_file, 'w') as f:
                json.dump(self.progress, f, indent=2)
        except Exception as e:
            print(f"⚠️  Could not save progress: {e}")
    
    def can_send_message(self):
        """Check if it's safe to send a message"""
        current_hour = datetime.now().hour
        
        # Check time restrictions
        if current_hour < ACTIVE_HOURS_START or current_hour >= ACTIVE_HOURS_END:
            return False, f"Outside active hours ({ACTIVE_HOURS_START}:00 - {ACTIVE_HOURS_END}:00)"
        
        # Reset hourly count if new hour
        if current_hour != self.progress['last_hour']:
            self.progress['messages_this_hour'] = 0
            self.progress['last_hour'] = current_hour
        
        # Check daily limit
        if self.progress['messages_today'] >= MAX_MESSAGES_PER_DAY:
            return False, f"Daily limit reached ({MAX_MESSAGES_PER_DAY} messages/day)"
        
        # Check hourly limit
        if self.progress['messages_this_hour'] >= MAX_MESSAGES_PER_HOUR:
            return False, f"Hourly limit reached ({MAX_MESSAGES_PER_HOUR} messages/hour)"
        
        return True, "OK"
    
    def get_next_delay(self):
        """Calculate the next delay with randomization"""
        base_delay = random.uniform(MIN_DELAY_BETWEEN_MESSAGES, MAX_DELAY_BETWEEN_MESSAGES)
        
        # Add extra delay if approaching limits
        if self.progress['messages_this_hour'] >= MAX_MESSAGES_PER_HOUR * 0.8:
            base_delay *= 1.5  # 50% longer delays when near hourly limit
        
        return base_delay
    
    def needs_break(self):
        """Check if a break is needed"""
        return self.progress['total_sent'] > 0 and self.progress['total_sent'] % MESSAGES_BEFORE_BREAK == 0
    
    def take_break(self):
        """Take a random break"""
        break_duration = random.uniform(BREAK_DURATION_MIN, BREAK_DURATION_MAX)
        break_minutes = break_duration / 60
        
        print(f"\n{'='*60}")
        print(f"⏸️  TAKING A BREAK")
        print(f"{'='*60}")
        print(f"📊 Progress so far:")
        print(f"   • Messages sent: {self.progress['total_sent']}")
        print(f"   • Today: {self.progress['messages_today']}/{MAX_MESSAGES_PER_DAY}")
        print(f"   • This hour: {self.progress['messages_this_hour']}/{MAX_MESSAGES_PER_HOUR}")
        print(f"\n⏳ Break duration: {break_minutes:.1f} minutes")
        print(f"   Resume time: {(datetime.now() + timedelta(seconds=break_duration)).strftime('%I:%M %p')}")
        print(f"{'='*60}\n")
        
        time.sleep(break_duration)
        print("✅ Break complete! Resuming...\n")
    
    def record_message_sent(self, contact_index):
        """Record that a message was sent"""
        self.progress['messages_today'] += 1
        self.progress['messages_this_hour'] += 1
        self.progress['total_sent'] += 1
        self.progress['last_message_time'] = datetime.now().isoformat()
        self.progress['last_contact_index'] = contact_index
        self.save_progress()
    
    def wait_until_active_hours(self):
        """Wait if outside active hours"""
        current_hour = datetime.now().hour
        
        if current_hour < ACTIVE_HOURS_START:
            wait_minutes = (ACTIVE_HOURS_START - current_hour) * 60
            resume_time = (datetime.now() + timedelta(minutes=wait_minutes)).strftime('%I:%M %p')
            print(f"\n⏰ Too early! Waiting until {ACTIVE_HOURS_START}:00 AM")
            print(f"   Resume time: {resume_time}")
            time.sleep(wait_minutes * 60)
        
        elif current_hour >= ACTIVE_HOURS_END:
            # Wait until next day's active hours
            hours_until_start = (24 - current_hour + ACTIVE_HOURS_START)
            wait_minutes = hours_until_start * 60
            resume_time = (datetime.now() + timedelta(minutes=wait_minutes)).strftime('%I:%M %p')
            print(f"\n⏰ Too late! Waiting until {ACTIVE_HOURS_START}:00 AM tomorrow")
            print(f"   Resume time: {resume_time}")
            time.sleep(wait_minutes * 60)
    
    def print_stats(self):
        """Print current statistics"""
        print(f"\n{'='*60}")
        print(f"📊 WHATSAPP SENDING STATISTICS")
        print(f"{'='*60}")
        print(f"Today's date: {self.progress['last_message_date']}")
        print(f"Messages sent today: {self.progress['messages_today']}/{MAX_MESSAGES_PER_DAY}")
        print(f"Messages this hour: {self.progress['messages_this_hour']}/{MAX_MESSAGES_PER_HOUR}")
        print(f"Total messages sent: {self.progress['total_sent']}")
        print(f"Last contact index: {self.progress['last_contact_index']}")
        if self.progress['last_message_time']:
            last_time = datetime.fromisoformat(self.progress['last_message_time'])
            print(f"Last message sent: {last_time.strftime('%Y-%m-%d %I:%M:%S %p')}")
        print(f"{'='*60}\n")
    
    def reset_progress(self):
        """Reset all progress (use with caution!)"""
        self.progress = {
            'messages_today': 0,
            'messages_this_hour': 0,
            'last_message_time': None,
            'last_message_date': datetime.now().strftime('%Y-%m-%d'),
            'last_hour': datetime.now().hour,
            'total_sent': 0,
            'last_contact_index': 0
        }
        self.save_progress()
        print("✅ Progress reset successfully!")


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
                debug_print(f"  ✓ Back on WhatsApp Web")
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
                    debug_print(f"  ✓ Message box focused (ActionChains method)")
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
                    debug_print(f"  ✓ Message box focused (aggressive JavaScript)")
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
                    debug_print(f"  ✓ Message box focused (footer click method)")
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
                    debug_print(f"  ✓ Message box focused (Escape + multiple clicks)")
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
                    debug_print(f"  ✓ Message box focused (simulated mouse events)")
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
        debug_print(f"  → Verifying chat is open...")
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
        debug_print(f"  → Waiting for chat to fully load...")
        
        # Check if this is an unsaved contact (longer wait needed)
        is_unsaved_contact = False
        try:
            # Look for "Add to Contacts" button or unsaved contact indicators
            unsaved_indicators = driver.find_elements(By.XPATH, 
                "//*[contains(text(), 'Add to Contacts')] | "
                "//*[contains(text(), 'Add to contacts')] | "
                "//button[@aria-label='Add to Contacts'] | "
                "//button[@aria-label='Add contact']"
            )
            if unsaved_indicators:
                is_unsaved_contact = True
                debug_print(f"  → Detected UNSAVED contact - will wait longer for interface...")
        except:
            pass
        
        # Wait longer for unsaved contacts (interface takes more time to initialize)
        if is_unsaved_contact:
            time.sleep(2.5)  # Unsaved contacts need more time
        else:
            time.sleep(1)  # Saved contacts load faster
        
        # Try to focus the chat by clicking the message box (ensures interface is active)
        try:
            debug_print(f"  → Focusing chat interface...")
            msg_box = get_fresh_message_box(driver)
            if msg_box:
                driver.execute_script("arguments[0].focus();", msg_box)
                driver.execute_script("arguments[0].click();", msg_box)
                time.sleep(0.3)
                debug_print(f"  ✓ Chat focused")
        except Exception as e:
            debug_print(f"  → Could not focus chat (non-critical): {e}")
        
        debug_print(f"  → Looking for attachment button...")
        attachment_button = None
        attachment_selectors = [
            "//span[@data-testid='clip']",
            "//div[@data-testid='clip']",
            "//span[@data-icon='attach']",
            "//button[@title='Attach']",
            "//button[@aria-label='Attach']"
        ]
        
        # Try multiple times with delays (more attempts for unsaved contacts)
        max_attempts = 10 if is_unsaved_contact else 5
        for attempt in range(max_attempts):
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
            
            # Longer retry delay for unsaved contacts
            if is_unsaved_contact:
                time.sleep(1)  # Unsaved: 1 second between attempts
            else:
                time.sleep(0.5)  # Saved: 0.5 second between attempts
        
        if not attachment_button:
            print(f"  ⚠️  Could not find attachment button, sending text only")
            print(f"      Make sure the chat is open and you're in a conversation")
            return send_text_fallback()
        
        # Step 3: Click attachment button
        debug_print(f"  → Clicking attachment button...")
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
        debug_print(f"  → Selecting 'Photos & videos' option (strict)...")
        time.sleep(1.2)  # menu animation settle - increased wait time

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

        def _find_menu_container():
            """Find the attachment menu container to scope searches"""
            try:
                # Look for the menu popup container
                menu_containers = [
                    "//div[@role='menu']",
                    "//div[contains(@class, 'menu') and @role='menu']",
                    "//div[@data-testid='menu']",
                    "//ul[@role='menu']"
                ]
                for sel in menu_containers:
                    try:
                        containers = driver.find_elements(By.XPATH, sel)
                        for container in containers:
                            if container.is_displayed():
                                # Verify it's the attachment menu by checking for menu items
                                items = container.find_elements(By.XPATH, ".//div[@role='menuitem']")
                                if len(items) >= 2:  # Should have at least Document and Photos & videos
                                    return container
                    except:
                        continue
                return None
            except:
                return None

        def _has_good_file_input():
            """Check if there's a good file input (for photos/videos, not stickers)"""
            try:
                file_inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
                for inp in file_inputs:
                    try:
                        if not inp.is_displayed():
                            continue
                        accept_attr = (inp.get_attribute('accept') or '').lower()
                        data_testid = (inp.get_attribute('data-testid') or '').lower()
                        name_attr = (inp.get_attribute('name') or '').lower()
                        
                        # Reject sticker inputs
                        if 'sticker' in accept_attr or 'sticker' in data_testid or 'sticker' in name_attr:
                            continue
                        # Reject empty accept (often sticker)
                        if accept_attr == '':
                            continue
                        # Reject webp-only
                        if accept_attr == 'image/webp' or accept_attr == '.webp':
                            continue
                        # Accept if has video or image (and not webp-only)
                        if 'video' in accept_attr:
                            return True
                        if 'image' in accept_attr and 'webp' not in accept_attr:
                            return True
                    except:
                        continue
                return False
            except:
                return False

        def _is_sticker_mode():
            """Verify if we're in sticker mode - return True ONLY if sticker mode is CONFIRMED"""
            try:
                time.sleep(0.8)  # Wait longer for UI to update
                # STRICT: Only return True if we have POSITIVE evidence of sticker mode
                # Check for explicit sticker UI elements
                sticker_indicators = driver.find_elements(By.XPATH,
                    "//button[contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'send sticker')] | "
                    "//button[contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'create sticker')] | "
                    "//span[@data-icon='sticker'] | "
                    "//div[contains(@data-testid, 'sticker')]"
                )
                if any(x.is_displayed() for x in sticker_indicators):
                    return True
                
                # Check file inputs - if ALL visible inputs are sticker-related, it's sticker mode
                file_inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
                visible_inputs = [inp for inp in file_inputs if inp.is_displayed()]
                if visible_inputs:
                    all_sticker = True
                    for inp in visible_inputs:
                        try:
                            accept_attr = (inp.get_attribute('accept') or '').lower()
                            data_testid = (inp.get_attribute('data-testid') or '').lower()
                            # If it's NOT sticker-related, then we're not in sticker mode
                            if 'sticker' not in accept_attr and 'sticker' not in data_testid:
                                if accept_attr and accept_attr != 'image/webp' and accept_attr != '.webp':
                                    all_sticker = False
                                    break
                        except:
                            continue
                    if all_sticker and len(visible_inputs) > 0:
                        return True
                
                # If we can't confirm sticker mode, return False (assume it's not sticker)
                return False
            except:
                return False

        def _is_photo_mode():
            """Verify if we're in photo mode - return True if photo mode confirmed or likely"""
            try:
                time.sleep(0.8)  # Wait longer for UI to update
                # Primary check: do we have a good file input?
                if _has_good_file_input():
                    return True
                # Secondary check: is media composer visible with caption?
                caption_inputs = driver.find_elements(By.XPATH,
                    "//div[contains(@data-testid, 'media')]//div[@contenteditable='true'][@data-tab='11'] | "
                    "//div[contains(@data-testid, 'media')]//div[@contenteditable='true'][contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'caption')]"
                )
                if any(elem.is_displayed() for elem in caption_inputs):
                    return True
                
                # Tertiary check: menu closed after clicking (good sign it worked)
                # If menu closed, assume it worked unless we have positive evidence it's sticker mode
                menu_container = _find_menu_container()
                menu_closed = menu_container is None or not menu_container.is_displayed()
                if menu_closed:
                    # Menu closed is a good sign - check if we have file inputs
                    file_inputs = driver.find_elements(By.XPATH, "//input[@type='file']")
                    visible_inputs = [inp for inp in file_inputs if inp.is_displayed()]
                    if visible_inputs:
                        # Check if any are NOT sticker-related
                        for inp in visible_inputs:
                            try:
                                accept_attr = (inp.get_attribute('accept') or '').lower()
                                data_testid = (inp.get_attribute('data-testid') or '').lower()
                                if 'sticker' not in accept_attr and 'sticker' not in data_testid:
                                    # If it has any accept attribute and not sticker, assume photo mode
                                    if accept_attr or data_testid:
                                        return True
                            except:
                                continue
                        # If we have inputs but can't confirm, and menu closed, assume it worked
                        # (file picker might be native OS dialog which we can't detect)
                        return True
                    # Menu closed but no file inputs yet - might be opening file picker
                    # Give it the benefit of the doubt if menu closed
                    return True
                
                return False
            except:
                return False

        selected = False
        
        # Method 1: Try official data-testid selectors first (language independent)
        for sel in ["//*[@data-testid='attach-photo']",
                    "//*[@data-testid='attach-image']",
                    "//*[@data-testid='attach-media']",
                    "//*[@data-testid='attach-photos-videos']",
                    "//button[@data-testid='attach-photo']",
                    "//div[@data-testid='attach-photo']",
                    "//span[@data-testid='attach-photo']"]:
            try:
                candidates = driver.find_elements(By.XPATH, sel)
                for c in candidates:
                    try:
                        # STRICT: Exclude any sticker-related data-testid
                        testid = (c.get_attribute('data-testid') or "").lower()
                        if "sticker" in testid:
                            continue
                        if c.is_displayed() and c.is_enabled():
                            if _click(c):
                                time.sleep(1.0)  # Wait for menu to close and file picker/file inputs to appear
                                # VERIFY: In photo mode (has good file input)
                                if _is_photo_mode():
                                    selected = True
                                    debug_print(f"  ✓ Selected Photos & videos via data-testid='{c.get_attribute('data-testid')}' (verified photo mode)")
                                    break
                                # VERIFY: Not in sticker mode
                                elif _is_sticker_mode():
                                    print(f"  ⚠️  Rejected: data-testid='{testid}' led to sticker mode")
                                    # Press Escape to close
                                    from selenium.webdriver.common.action_chains import ActionChains
                                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                    time.sleep(0.5)
                                    continue
                                else:
                                    print(f"  ⚠️  Rejected: data-testid='{testid}' - photo mode not confirmed")
                                    from selenium.webdriver.common.action_chains import ActionChains
                                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                    time.sleep(0.5)
                    except:
                        continue
                if selected:
                    break
            except:
                pass

        # Method 2: Text-based selection (STRICT: require BOTH "photo" AND "video" text, scoped to menu)
        if not selected:
            try:
                from selenium.webdriver.common.action_chains import ActionChains
                # Find the menu container first to scope the search
                menu_container = _find_menu_container()
                if menu_container:
                    # STRICT: Only match items in the menu that contain BOTH "photo" AND "video"
                    text_selectors = [
                        ".//div[@role='menuitem' and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'photo') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'video')]",
                        ".//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'photo') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'video')]"
                    ]
                    for sel in text_selectors:
                        try:
                            candidates = menu_container.find_elements(By.XPATH, sel)
                            for c in candidates:
                                try:
                                    # STRICT: Exclude any sticker-related items
                                    text = (c.text or "").strip()
                                    text_lower = text.lower()
                                    aria_label = (c.get_attribute('aria-label') or "").lower()
                                    # VERY STRICT: Menu items are short, single-line, contain both words
                                    if len(text) > 30 or '\n' in text or '\r' in text:  # Menu items are max ~25 chars, single line
                                        continue
                                    if "sticker" in text_lower or "sticker" in aria_label or "new sticker" in text_lower:
                                        continue
                                    # Must contain both photo and video
                                    if "photo" not in text_lower or "video" not in text_lower:
                                        continue
                                    if c.is_displayed() and c.is_enabled():
                                        click_target = _clickable_ancestor(c) or c
                                        if _click(click_target):
                                            time.sleep(1.0)  # Wait for menu to close and file picker to appear
                                            # VERIFY: In photo mode (has good file input)
                                            if _is_photo_mode():
                                                selected = True
                                                debug_print(f"  ✓ Selected Photos & videos via text match: '{text[:30]}' (verified photo mode)")
                                                break
                                            # VERIFY: Not in sticker mode
                                            elif _is_sticker_mode():
                                                print(f"  ⚠️  Rejected: text='{text[:30]}' led to sticker mode")
                                                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                                time.sleep(0.5)
                                                continue
                                            else:
                                                print(f"  ⚠️  Rejected: text='{text[:30]}' - photo mode not confirmed")
                                                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                                time.sleep(0.5)
                                except:
                                    continue
                            if selected:
                                break
                        except:
                            pass
                else:
                    # Fallback: search globally but with VERY strict filters
                    # Only match if it's clearly a menu item (very short text, no newlines, in menu context)
                    text_selectors = [
                        "//div[@role='menuitem' and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'photo') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'video')]"
                    ]
                    for sel in text_selectors:
                        try:
                            candidates = driver.find_elements(By.XPATH, sel)
                            for c in candidates:
                                try:
                                    text = (c.text or "").strip()
                                    text_lower = text.lower()
                                    # VERY STRICT: Menu items are short, single-line, contain both words
                                    # Skip if: too long, has newlines, contains sticker, missing words
                                    if len(text) > 30:  # Menu items are max ~25 chars
                                        continue
                                    if '\n' in text or '\r' in text:  # Menu items are single line
                                        continue
                                    if "sticker" in text_lower:
                                        continue
                                    if "photo" not in text_lower or "video" not in text_lower:
                                        continue
                                    # Check if parent is actually a menu (not chat)
                                    try:
                                        parent = c.find_element(By.XPATH, "./ancestor::div[@role='menu'][1]")
                                        if not parent or not parent.is_displayed():
                                            continue
                                    except:
                                        # If no menu parent found, skip (likely not a menu item)
                                        continue
                                    
                                    if c.is_displayed() and c.is_enabled():
                                        click_target = _clickable_ancestor(c) or c
                                        if _click(click_target):
                                            time.sleep(1.0)
                                            if _is_photo_mode():
                                                selected = True
                                                debug_print(f"  ✓ Selected Photos & videos via text match: '{text[:30]}' (verified photo mode)")
                                                break
                                            elif _is_sticker_mode():
                                                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                                time.sleep(0.5)
                                                continue
                                except:
                                    continue
                            if selected:
                                break
                        except:
                            pass
            except:
                pass

        # Method 3: Keyboard navigation (Arrow Down + Enter) - with verification
        if not selected:
            try:
                from selenium.webdriver.common.action_chains import ActionChains
                # Focus on the menu and navigate (Document is 1st, Photos & videos is 2nd)
                ActionChains(driver).send_keys(Keys.ARROW_DOWN).send_keys(Keys.ARROW_DOWN).send_keys(Keys.ENTER).perform()
                time.sleep(1.0)  # Wait for menu to close and file picker to appear
                # VERIFY: In photo mode (has good file input)
                if _is_photo_mode():
                    selected = True
                    debug_print(f"  ✓ Selected Photos & videos via keyboard navigation (verified photo mode)")
                    
                    # FORCE CLOSE attachment menu by clicking the attachment button again (toggle)
                    try:
                        debug_print(f"  → Force closing attachment menu...")
                        time.sleep(0.3)
                        
                        # Method 1: Click the attachment button again to toggle it closed
                        attachment_btns = driver.find_elements(By.XPATH,
                            "//span[@data-testid='clip'] | "
                            "//div[@data-testid='clip'] | "
                            "//span[@data-icon='attach']"
                        )
                        clicked = False
                        for btn in attachment_btns:
                            try:
                                if btn.is_displayed() and btn.is_enabled():
                                    btn.click()
                                    clicked = True
                                    debug_print(f"  ✓ Clicked attachment button to close menu")
                                    time.sleep(0.2)
                                    break
                            except:
                                continue
                        
                        # Method 2: If clicking didn't work, try pressing Escape
                        if not clicked:
                            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                            time.sleep(0.2)
                            
                    except Exception as e:
                        print(f"  ⚠️  Could not force close menu: {str(e)}")
                
                # VERIFY: Not in sticker mode
                elif _is_sticker_mode():
                    print(f"  ⚠️  Rejected: Keyboard navigation led to sticker mode")
                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                    time.sleep(0.5)
                else:
                    print(f"  ⚠️  Rejected: Keyboard navigation - photo mode not confirmed")
                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                    time.sleep(0.5)
            except:
                pass

        # Method 4: Click the 2nd visible menu item (Document is 1st, Photos & videos is 2nd) - with verification
        if not selected:
            try:
                from selenium.webdriver.common.action_chains import ActionChains
                # Try to scope to menu container first
                menu_container = _find_menu_container()
                if menu_container:
                    menu_items = menu_container.find_elements(By.XPATH, ".//div[@role='menuitem']")
                else:
                    menu_items = driver.find_elements(By.XPATH, "//div[@role='menuitem']")
                
                visible_items = [m for m in menu_items if m.is_displayed()]
                # Filter out items with very long text (likely not menu items)
                visible_items = [m for m in visible_items if len((m.text or "").strip()) < 50]
                
                if len(visible_items) >= 2:
                    target = visible_items[1]
                    # STRICT: Check text of target to exclude stickers
                    target_text = (target.text or "").lower().strip()
                    target_aria = (target.get_attribute('aria-label') or "").lower()
                    if "sticker" in target_text or "sticker" in target_aria:
                        print(f"  ⚠️  Rejected: Menu item #2 is sticker-related: '{target_text[:30]}'")
                    elif "photo" not in target_text or "video" not in target_text:
                        # If it doesn't contain both photo and video, might not be the right item
                        print(f"  ⚠️  Menu item #2 text doesn't match: '{target_text[:30]}'")
                    else:
                        click_target = _clickable_ancestor(target) or target
                        if _click(click_target):
                            time.sleep(1.0)  # Wait for menu to close and file picker to appear
                            # VERIFY: In photo mode (has good file input)
                            if _is_photo_mode():
                                selected = True
                                debug_print(f"  ✓ Selected Photos & videos via menu item #2 (verified photo mode)")
                            # VERIFY: Not in sticker mode
                            elif _is_sticker_mode():
                                print(f"  ⚠️  Rejected: Menu item #2 led to sticker mode")
                                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                time.sleep(0.5)
                            else:
                                print(f"  ⚠️  Rejected: Menu item #2 - photo mode not confirmed")
                                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                time.sleep(0.5)
            except:
                pass

        # Method 5: Try finding by aria-label (STRICT: require BOTH photo AND video)
        if not selected:
            try:
                from selenium.webdriver.common.action_chains import ActionChains
                # STRICT: Only match aria-labels that contain BOTH "photo" AND "video"
                aria_selectors = [
                    "//button[contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'photo') and contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'video')]",
                    "//div[contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'photo') and contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'video')]"
                ]
                for sel in aria_selectors:
                    try:
                        candidates = driver.find_elements(By.XPATH, sel)
                        for c in candidates:
                            try:
                                aria_label = (c.get_attribute('aria-label') or "").lower()
                                # STRICT: Exclude any sticker-related
                                if "sticker" in aria_label:
                                    continue
                                if c.is_displayed() and c.is_enabled():
                                    if _click(c):
                                        time.sleep(1.0)  # Wait for menu to close and file picker to appear
                                        # VERIFY: In photo mode (has good file input)
                                        if _is_photo_mode():
                                            selected = True
                                            debug_print(f"  ✓ Selected Photos & videos via aria-label (verified photo mode)")
                                            break
                                        # VERIFY: Not in sticker mode
                                        elif _is_sticker_mode():
                                            print(f"  ⚠️  Rejected: aria-label='{aria_label[:30]}' led to sticker mode")
                                            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                            time.sleep(0.5)
                                            continue
                                        else:
                                            print(f"  ⚠️  Rejected: aria-label='{aria_label[:30]}' - photo mode not confirmed")
                                            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                                            time.sleep(0.5)
                            except:
                                continue
                        if selected:
                            break
                    except:
                        pass
            except:
                pass

        if not selected:
            # Debug: Print available menu items for troubleshooting
            try:
                print(f"  ⚠️  Debug: Available menu items:")
                menu_items = driver.find_elements(By.XPATH, "//div[@role='menuitem'] | //button[contains(@class, 'menu')]")
                for idx, item in enumerate(menu_items[:5]):  # Show first 5
                    try:
                        if item.is_displayed():
                            text = item.text[:50] if item.text else ""
                            testid = item.get_attribute('data-testid') or ""
                            aria = item.get_attribute('aria-label') or ""
                            print(f"      [{idx}] text='{text}' testid='{testid}' aria='{aria}'")
                    except:
                        pass
            except:
                pass
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
                debug_print(f"  → File inputs found: {len(scored)} (best score={best[0]}, accept='{best[3]}', multiple={best[4]})")
        except Exception as e:
            print(f"  ✗ ERROR finding file input: {str(e)}")
            return send_text_fallback()

        if not file_input:
            print(f"  ✗ ERROR: Could not find a valid media file input for Photos & videos")
            return send_text_fallback()
        
        # Step 5: PRE-UPLOAD STICKER CHECK (detect sticker mode BEFORE uploading)
        debug_print(f"  → Verifying photo mode (not sticker)...")
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
                debug_print(f"  ✓ Photo mode confirmed (no sticker UI detected)")
        except Exception as e:
            print(f"  ⚠️  Could not verify mode: {str(e)}, proceeding anyway...")
        
        # Step 6: Upload image (no sanitization; use original file)
        debug_print(f"  → Preparing image for upload...")
        try:
            # Verify file path exists
            if not os.path.exists(image_path):
                print(f"  ✗ ERROR: Image file not found: {image_path}")
                return send_text_fallback()

            abs_image_path = os.path.abspath(image_path)
            debug_print(f"  → Uploading: {abs_image_path}")
            
            file_input.send_keys(abs_image_path)
            
            # Wait for image to start loading first (don't close anything yet!)
            time.sleep(1.5)  # Reduced from 2.5s - image loads faster now
            
            # Verify image was actually uploaded by checking for image preview
            debug_print(f"  → Verifying image upload...")
            image_uploaded = False
            for check_attempt in range(6):
                time.sleep(0.5)
                previews = driver.find_elements(By.XPATH, 
                    "//img[contains(@src, 'blob')] | "
                    "//div[contains(@data-testid, 'media')]//img | "
                    "//div[contains(@class, 'preview')]//img"
                )
                if any(p.is_displayed() for p in previews):
                    image_uploaded = True
                    debug_print(f"  ✓ Image preview appeared after {check_attempt + 1} seconds")
                    
                    # Close Windows file picker (native OS dialog) - ALWAYS needed on Windows
                    debug_print(f"  → Closing Windows file picker...")
                    try:
                        # Press Escape to close the Windows file picker dialog
                        # This is a NATIVE Windows dialog, separate from the browser
                        pyautogui.press('esc')
                        time.sleep(0.3)
                        
                        # Verify media composer is still visible (not accidentally closed)
                        driver.execute_script("window.focus();")
                        time.sleep(0.2)
                        
                        media_check = driver.find_elements(By.XPATH, 
                            "//img[contains(@src, 'blob')] | "
                            "//div[contains(@data-testid, 'media-composer')]"
                        )
                        
                        if any(elem.is_displayed() for elem in media_check):
                            debug_print(f"  ✓ File picker closed - media composer still active")
                        else:
                            # Oops - we might have closed the media composer too
                            debug_print(f"  ⚠️  Warning: Media composer might have closed - will retry")
                        
                    except Exception as e:
                        print(f"  ⚠️  Cleanup warning: {str(e)}")
                    
                    break
            
            if not image_uploaded:
                print(f"  ⚠️  WARNING: Image preview not found after upload - image may not have uploaded")
                # Still continue, might be a timing issue
            
            # Wait for interface to stabilize after cleanup
            time.sleep(0.8)
            
            # Verify we're in photo mode (not sticker mode) by checking for caption input
            # In sticker mode, there's usually no caption input
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
                debug_print(f"  → Photo editing tools not immediately visible - checking image preview...")
                
                # Check if image preview is visible (this is the main indicator of photo mode)
                image_preview_check = driver.find_elements(By.XPATH, 
                    "//img[contains(@src, 'blob')] | "
                    "//div[contains(@data-testid, 'media')] | "
                    "//div[contains(@class, 'preview')]"
                )
                image_preview_visible = any(img.is_displayed() for img in image_preview_check)
                
                if image_preview_visible and not is_sticker_mode:
                    debug_print(f"  ✓ Image preview visible and NOT in sticker mode - assuming photo mode")
                    debug_print(f"  → Proceeding (editing tools may be hidden or appear on click)")
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
                        debug_print(f"  ✓ Image preview now visible - proceeding with photo mode")
                        has_photo_tools = True
                    else:
                        print(f"  ⚠️  Image preview still not visible - but proceeding anyway")
                        # Still proceed - might be a timing issue
            elif has_photo_tools:
                debug_print(f"  ✓ Photo mode confirmed (photo editing tools visible)")
                if not has_caption_input:
                    print(f"  ⚠️  Caption input not found yet - may still be loading...")
                else:
                    debug_print(f"  ✓ Caption input available")
            elif not has_caption_input:
                print(f"  ⚠️  Caption input not found - checking if photo mode...")
                # Check if image preview is visible (should be in both modes)
                image_preview = driver.find_elements(By.XPATH, 
                    "//img[contains(@src, 'blob')] | "
                    "//div[contains(@data-testid, 'media')]//img"
                )
                if any(img.is_displayed() for img in image_preview):
                    debug_print(f"  → Image preview visible, but caption input missing - may be sticker mode")
                else:
                    debug_print(f"  → Image preview not visible - interface may still be loading...")
            else:
                debug_print(f"  ✓ Photo mode confirmed (caption input available)")
        except Exception as e:
            print(f"  ⚠️  Could not upload image: {str(e)}, sending text only")
            return send_text_fallback()
        
        # Step 5.5: Wait for image preview interface to fully load
        debug_print(f"  → Waiting for image preview interface to load...")
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
        debug_print(f"  → Checking for caption input to appear...")
        caption_input_found = False
        for wait_round in range(5):  # Reduced from 12 to 5 seconds
            time.sleep(0.5)  # Reduced from 1 second to 0.5 seconds
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
                            debug_print(f"  ✓ Found potential caption input after {wait_round + 1}s (data-tab='{data_tab}', placeholder='{placeholder[:30]}')")
                            caption_input_found = True
                            break
                    except:
                        continue
                if caption_input_found:
                    break
            except:
                pass
        
        # Try to activate the caption input by interacting with the image preview
        debug_print(f"  → Activating caption input...")
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
                    debug_print(f"  ✓ Clicked image preview to activate caption mode")
                    # Check if caption input appeared
                    caption_check = driver.find_elements(By.XPATH, 
                        "//div[@contenteditable='true'][@data-tab='11'] | "
                        "//div[@contenteditable='true'][contains(translate(@placeholder, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'type a message')]"
                    )
                    if any(elem.is_displayed() for elem in caption_check):
                        debug_print(f"  ✓ Caption input appeared after clicking image preview")
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
                    debug_print(f"  ✓ Clicked footer area to activate caption input")
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
                        debug_print(f"  ✓ Caption input focused via Tab navigation")
                        break
            
            # Method 4: Click on message box as fallback
            message_boxes = driver.find_elements(By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
            for msg_box in message_boxes:
                if msg_box.is_displayed():
                    driver.execute_script("arguments[0].click();", msg_box)
                    driver.execute_script("arguments[0].focus();", msg_box)
                    time.sleep(0.5)
                    debug_print(f"  ✓ Clicked and focused message box to activate caption mode")
                    break
        except Exception as e:
            print(f"  ⚠️  Error activating caption input: {str(e)}")
        
        # Step 6: Find caption input box (appears after image is selected)
        # This should be the "Type a message" input that appears BELOW the image preview
        debug_print(f"  → Looking for caption input box (below image preview)...")
        caption_box = None
        
        # Wait longer for the image preview interface to fully render
        time.sleep(1)  # Reduced from 2 seconds - interface renders faster
        
        # First, find the image preview container, then look for caption input inside it
        debug_print(f"  → Finding image preview container...")
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
                            debug_print(f"  ✓ Found image preview container")
                            break
                if media_container:
                    break
            except:
                continue
        
        # Try clicking on image preview area to activate caption input
        debug_print(f"  → Clicking on image preview to activate caption input...")
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
                    debug_print(f"  ✓ Clicked in image preview container to activate caption")
                    break
        except:
            pass
        
        # Caption input in the media composer is in the footer ("Type a message").
        # CRITICAL: Must search ONLY inside the media overlay that contains the blob preview,
        # NOT the normal chat footer (which also has a message box but is in the background).
        debug_print(f"  → Finding caption/message box INSIDE media composer overlay...")
        
        def _find_caption_in_media_overlay():
            """Find caption box that's inside the same container as the blob preview."""
            try:
                # Step 1: Find ALL elements that contain blob previews
                preview_containers = driver.find_elements(
                    By.XPATH,
                    "//*[.//img[contains(@src,'blob')]]"
                )
                
                # Step 2: For each preview container, look for a caption box inside it
                for container in preview_containers:
                    try:
                        if not container.is_displayed():
                            continue
                        
                        # Try to find caption box inside this container
                        caption_boxes = container.find_elements(
                            By.XPATH,
                            ".//footer//div[@contenteditable='true'][@data-tab='10'] | "
                            ".//footer//div[@contenteditable='true'][@data-tab='11'] | "
                            ".//div[@contenteditable='true'][contains(@aria-placeholder,'message')] | "
                            ".//div[@contenteditable='true'][contains(@placeholder,'message')]"
                        )
                        
                        for box in caption_boxes:
                            if box.is_displayed():
                                return box
                    except:
                        continue
                
                return None
            except:
                return None
        
        try:
            caption_box = WebDriverWait(driver, 20).until(lambda d: _find_caption_in_media_overlay())
            try:
                dt = caption_box.get_attribute('data-tab')
                al = (caption_box.get_attribute('aria-label') or '')[:40]
                ap = (caption_box.get_attribute('aria-placeholder') or '')[:40]
                debug_print(f"  ✓ Found caption box INSIDE media overlay (data-tab='{dt}', aria-label='{al}', aria-placeholder='{ap}')")
            except:
                debug_print(f"  ✓ Found caption box INSIDE media overlay")
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
                        debug_print(f"  ✓ Using footer caption box (fallback scan)")
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
                                    debug_print(f"  ✓ Found caption box in media container (placeholder='{placeholder[:30]}', data-tab='{data_tab}')")
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
                                debug_print(f"  ✓ Found caption box (data-tab='{data_tab}', has_image={has_image_preview}, placeholder='{placeholder[:30]}')")
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
            debug_print(f"  → Trying to find any contenteditable in footer area...")
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
                                debug_print(f"  ✓ Found potential caption box (data-tab='{data_tab}') - image preview is visible")
                                break
                    except:
                        continue
                
                # If still not found, maybe the caption box appears INSIDE the image preview container
                # Try to find it near the image preview
                if not caption_box:
                    debug_print(f"  → Looking for caption box near image preview...")
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
                                                debug_print(f"  ✓ Found caption box in media container (data-tab='{data_tab}')")
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
                    debug_print(f"  → Trying to find caption input near send button...")
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
                                                        debug_print(f"  ✓ Found caption input near send button (placeholder='{placeholder[:30]}', data-tab='{data_tab}')")
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
                                                            debug_print(f"  ✓ Found caption input in footer (placeholder='{placeholder[:30]}', data-tab='{data_tab}')")
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
                    debug_print(f"  → Trying keyboard navigation to find caption input...")
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
                                        debug_print(f"  ✓ Found caption input via Tab navigation (placeholder='{placeholder[:30]}', data-tab='{data_tab}')")
                                        break
                    except:
                        pass
                
                # Try clicking directly in the area between thumbnail and send button
                if not caption_box:
                    debug_print(f"  → Trying to click in caption input area...")
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
                                            debug_print(f"  ✓ Found caption input by clicking in area (data-tab='{data_tab}')")
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
                                debug_print(f"  ✓ Clicked below image preview to activate caption")
                                break
                    except:
                        pass
            except:
                pass
        
        # Step 7: If caption box not found yet, try using data-tab='10' (when image is attached, it becomes the caption input)
        if not caption_box and caption:
            debug_print(f"  → Caption box not found in scans, looking for data-tab='10' with image attached...")
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
                            debug_print(f"  ✓ Found caption input: data-tab='10' (message box becomes caption when image attached)")
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
            debug_print(f"  → Typing caption in caption box...")
            try:
                # Verify this is actually the caption box (data-tab='11')
                data_tab = caption_box.get_attribute('data-tab')
                if data_tab != '11':
                    debug_print(f"  ⚠️  Warning: Element data-tab='{data_tab}' might not be caption box")
                
                # Verify image preview is still visible before typing
                debug_print(f"  → Verifying image preview is still visible...")
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
                            debug_print(f"  ✓ Image preview is visible")
                            break
                except:
                    pass
                
                if not preview_visible:
                    debug_print(f"  ⚠️  Warning: Image preview not visible, but continuing...")
                
                # Use JavaScript to focus and type - this bypasses overlay issues
                debug_print(f"  → Using JavaScript to type (bypassing overlay)...")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", caption_box)
                time.sleep(0.2)
                
                # Focus using JavaScript (bypasses click interception) - single attempt is enough
                driver.execute_script("""
                    arguments[0].focus();
                    arguments[0].dispatchEvent(new Event('focus', { bubbles: true }));
                """, caption_box)
                time.sleep(0.15)
                
                # Always use JavaScript to type (bypasses overlay and works with emojis)
                # IMPORTANT: Don't clear the element if image preview is visible - just append/type
                debug_print(f"  → Setting caption text via JavaScript...")
                
                # Check if image preview is still visible before typing
                previews_before = driver.find_elements(By.XPATH, 
                    "//img[contains(@src, 'blob')] | "
                    "//div[contains(@data-testid, 'media')]"
                )
                preview_visible_before = any(p.is_displayed() for p in previews_before)
                
                if preview_visible_before:
                    debug_print(f"  ✓ Image preview still visible, typing caption in footer message box...")
                    # IMPORTANT: Use pure JS (no .click()) to avoid "Add file" button interception
                    # Check if caption box already has content - if yes, clear it FIRST
                    existing_text = driver.execute_script("return arguments[0].textContent || arguments[0].innerText || '';", caption_box)
                    if existing_text and len(existing_text.strip()) > 0:
                        print(f"  ⚠️  Caption box already has {len(existing_text)} chars - clearing before typing...")
                    
                    driver.execute_script("""
                        var elem = arguments[0];
                        var text = arguments[1] || '';
                        
                        // Scroll into view and focus (NO physical click)
                        elem.scrollIntoView({block:'center', inline:'center'});
                        elem.focus();
                        
                        // Step 1: CLEAR completely - try multiple methods
                        try {
                            // Clear via innerHTML (most aggressive for Lexical editor)
                            elem.innerHTML = '';
                        } catch(e) {}
                        
                        try {
                            // Also clear text properties
                            elem.textContent = '';
                            elem.innerText = '';
                        } catch(e) {}
                        
                        try {
                            // Select all and delete using execCommand
                            var range = document.createRange();
                            range.selectNodeContents(elem);
                            var selection = window.getSelection();
                            selection.removeAllRanges();
                            selection.addRange(range);
                            document.execCommand('delete', false, null);
                        } catch(e) {}
                        
                        // Refocus after clearing
                        elem.focus();
                        
                        // Step 2: Insert text with proper newline handling
                        // Split by newlines and insert each line, using Shift+Enter for line breaks
                        var lines = text.split('\\n');
                        for (var i = 0; i < lines.length; i++) {
                            if (i > 0) {
                                // Insert newline using Shift+Enter (WhatsApp's way of creating line breaks)
                                var shiftEnterEvent = new KeyboardEvent('keydown', {
                                    key: 'Enter',
                                    code: 'Enter',
                                    keyCode: 13,
                                    which: 13,
                                    shiftKey: true,
                                    bubbles: true,
                                    cancelable: true
                                });
                                elem.dispatchEvent(shiftEnterEvent);
                                
                                // Also dispatch keyup
                                var keyupEvent = new KeyboardEvent('keyup', {
                                    key: 'Enter',
                                    code: 'Enter',
                                    keyCode: 13,
                                    which: 13,
                                    shiftKey: true,
                                    bubbles: true
                                });
                                elem.dispatchEvent(keyupEvent);
                            }
                            
                            // Insert the line text (if not empty)
                            if (lines[i]) {
                                try {
                                    document.execCommand('insertText', false, lines[i]);
                                } catch(e) {}
                            }
                        }
                        
                        elem.focus();
                    """, caption_box, caption)
                    time.sleep(0.4)  # Reduced from 0.8
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
                
                time.sleep(0.5)  # Reduced from 1 - Wait for caption to be set
                
                # Verify caption was typed AND image preview is still there
                caption_text = driver.execute_script("return arguments[0].textContent || arguments[0].innerText;", caption_box)
                previews_after = driver.find_elements(By.XPATH, 
                    "//img[contains(@src, 'blob')] | "
                    "//div[contains(@data-testid, 'media')]"
                )
                preview_visible_after = any(p.is_displayed() for p in previews_after)
                
                if caption_text and len(caption_text.strip()) > 0:
                    if preview_visible_after:
                        debug_print(f"  ✓ Caption typed successfully ({len(caption_text)} chars) - Image preview still attached")
                    else:
                        debug_print(f"  ⚠️  Warning: Caption typed but image preview might be lost!")
                else:
                    debug_print(f"  ⚠️  Warning: Caption might not have been set properly")
                
            except Exception as e:
                print(f"  ⚠️  Could not type caption: {str(e)}")
                # Don't send without caption - cancel and send text instead
                debug_print(f"  → Canceling image send (Mode 2) - NOT sending text-only...")
                try:
                    from selenium.webdriver.common.action_chains import ActionChains
                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                    time.sleep(0.8)
                except:
                    pass
                return False
        
        # Step 8: Verify image preview and caption are ready before sending
        debug_print(f"  → Verifying image and caption are ready...")
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
                    debug_print(f"  ✓ Image preview is visible")
                    break
            
            if not image_ready:
                debug_print(f"  ⚠️  Warning: Image preview not found! Image might be lost.")
            
            # Verify caption is still in caption/message box
            if caption_box and caption:
                caption_text = driver.execute_script("return arguments[0].textContent || arguments[0].innerText;", caption_box)
                if caption_text and len(caption_text.strip()) > 0:
                    caption_ready = True
                    debug_print(f"  ✓ Caption is ready ({len(caption_text)} chars)")
                else:
                    # Caption might be in WhatsApp's internal format (wrapped in spans, etc.)
                    # Assume it's already typed successfully if we got here
                    debug_print(f"  → Caption verification skipped (already typed successfully)")
                    caption_ready = True
        except:
            pass
        
        if not image_ready:
            print(f"  ✗ Image preview lost, cannot send image with caption")
            return send_text_fallback()
        
        # Step 8.5: Send image with caption
        # IMPORTANT: When image preview is visible, use Enter key in message box
        # This is more reliable than clicking send button for image + caption
        debug_print(f"  → Sending image with caption...")
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
                
                debug_print(f"  ✓ Caption box focused")
            except:
                pass
        
        # Check if image preview is still visible
        previews = driver.find_elements(By.XPATH, 
            "//img[contains(@src, 'blob')] | "
            "//div[contains(@data-testid, 'media')]"
        )
        has_preview = any(p.is_displayed() for p in previews)
        
        if has_preview and caption_box:
            debug_print(f"  → Image preview visible, sending via media composer...")
            # IMPORTANT: Never press Enter as primary send here; it can send text-only and leave media attached.
            try:
                driver.execute_script("arguments[0].focus();", caption_box)
                time.sleep(0.2)

                def _find_media_send_button():
                    """
                    Find the *media composer* send button (not the normal chat send button).
                    Search INSIDE the same media overlay container that has the blob preview.
                    """
                    try:
                        # Find containers with blob previews (same strategy as caption box finding)
                        preview_containers = driver.find_elements(
                            By.XPATH,
                            "//*[.//img[contains(@src,'blob')]]"
                        )
                        
                        # Look for send button inside each preview container
                        for container in preview_containers:
                            try:
                                if not container.is_displayed():
                                    continue
                                
                                # Search for send button inside this media overlay
                                send_selectors = [
                                    ".//button[@aria-label='Send']",
                                    ".//span[@data-icon='send']/ancestor::button[1]",
                                    ".//span[@data-testid='send']/ancestor::button[1]",
                                    ".//div[@role='button'][.//span[@data-icon='send']]",
                                    ".//*[@data-testid='send']/ancestor::button[1]",
                                    ".//*[@aria-label='Send']",
                                ]
                                
                                for sel in send_selectors:
                                    try:
                                        buttons = container.find_elements(By.XPATH, sel)
                                        for btn in buttons:
                                            if btn.is_displayed() and btn.is_enabled():
                                                return btn
                                    except:
                                        continue
                            except:
                                continue
                        
                        return None
                    except:
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

                    debug_print(f"  → Clicking media send button (try {send_try+1}/3)...")
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
                        debug_print(f"  ✓ Image sent! (media composer closed)")
                        break
                    else:
                        print(f"  ⚠️  Still in media composer after click (may still be uploading) - retrying send...")

                # If still not sent, cancel attachment BEFORE falling back to text-only
                if not sent and _composer_open():
                    try:
                        from selenium.webdriver.common.action_chains import ActionChains
                        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                        time.sleep(0.8)
                        debug_print(f"  → Canceled attachment (Escape) before fallback")
                    except:
                        pass
            except Exception as e:
                print(f"  ⚠️  Send method failed: {str(e)}")
        
        # Step 9: No global send fallbacks.
        # IMPORTANT: Clicking generic send buttons (or pressing Enter) can send text-only and leave media attached.
        # If we couldn't send via the media composer send button + close verification, treat as failure.
        
        # Step 10: Verify and cleanup
        if sent:
            time.sleep(1)  # Reduced from 2 seconds - WhatsApp processes faster
            
            # Close any remaining popups/menus (attachment menu, media composer, etc.)
            debug_print(f"  → Closing attachment menu...")
            try:
                from selenium.webdriver.common.action_chains import ActionChains
                # Press Escape multiple times to close all overlays
                for i in range(3):
                    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                    time.sleep(0.3)
                
                # Click on the main chat area to ensure focus is back on chat
                try:
                    # Click on the message box area to close any menus
                    msg_box = get_fresh_message_box(driver)
                    if msg_box:
                        msg_box.click()
                        time.sleep(0.2)
                except:
                    pass
                    
            except Exception as e:
                print(f"  ⚠️  Could not close attachment menu: {str(e)}")
            
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
    Simple WhatsApp message sender using direct URL method (more reliable):
    1. Navigate directly to chat via URL
    2. Wait for chat to load
    3. Send message (or image with caption if image_path provided)
    4. Auto send
    
    This method works for:
    - Saved contacts
    - Unsaved contacts with chat history
    - Completely NEW contacts (not saved, no chat history)
    """
    from selenium.webdriver.common.action_chains import ActionChains
    
    try:
        # Step 1: Navigate directly to chat using WhatsApp's send URL
        # This is more reliable than searching, especially for new contacts
        clean_number = contact_number.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
        
        # Remove leading + if present (will add it back)
        if clean_number.startswith("+"):
            clean_number = clean_number[1:]
        
        # Build WhatsApp send URL
        whatsapp_url = f"https://web.whatsapp.com/send?phone={clean_number}"
        
        debug_print(f"  → Opening chat via direct URL...")
        driver.get(whatsapp_url)
        
        # Step 2: Wait for chat to load
        # Different wait times based on whether this is a new contact
        debug_print(f"  → Waiting for chat to initialize...")
        
        # Wait for page to load first
        time.sleep(2)
        
        # Check if this is a new/unsaved contact
        is_new_or_unsaved = False
        try:
            # Look for indicators of new/unsaved contact
            indicators = driver.find_elements(By.XPATH,
                "//*[contains(text(), 'Add to Contacts')] | "
                "//*[contains(text(), 'Add to contacts')] | "
                "//*[contains(text(), 'Start new conversation')] | "
                "//button[@aria-label='Add to Contacts'] | "
                "//button[@aria-label='Add contact']"
            )
            if indicators:
                is_new_or_unsaved = True
                debug_print(f"  → Detected NEW/UNSAVED contact - waiting longer...")
                time.sleep(3)  # Extra wait for new contacts
        except:
            pass
        
        # Step 3: Wait for chat interface to be ready
        debug_print(f"  → Waiting for chat interface...")
        time.sleep(0.5)
        
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
                
                # Close ALL menus, dialogs, and overlays (attachment menu, file pickers, etc.)
                debug_print(f"  → Final cleanup: Closing all dialogs...")
                try:
                    # First: Use pyautogui to close any Windows dialogs that might still be open
                    try:
                        for i in range(2):
                            pyautogui.press('esc')
                            time.sleep(0.1)
                        debug_print(f"  ✓ Sent Escape to close any OS dialogs")
                    except:
                        pass
                    
                    # Second: Press Escape in browser to close WhatsApp menus
                    for i in range(4):
                        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                        time.sleep(0.12)
                    
                    # Third: Ensure browser window has focus
                    try:
                        driver.execute_script("window.focus();")
                        time.sleep(0.15)
                    except:
                        pass
                    
                    # Fourth: Toggle attachment button to ensure menu is closed
                    try:
                        attachment_btns = driver.find_elements(By.XPATH, "//span[@data-testid='clip']")
                        for btn in attachment_btns:
                            try:
                                if btn.is_displayed():
                                    driver.execute_script("arguments[0].click();", btn)
                                    time.sleep(0.12)
                                    driver.execute_script("arguments[0].click();", btn)
                                    time.sleep(0.12)
                                    debug_print(f"  ✓ Toggled attachment button")
                                    break
                            except:
                                continue
                    except:
                        pass
                    
                    # Fifth: Click on chat area to ensure focus is back
                    try:
                        driver.execute_script("""
                            var mainPane = document.querySelector('[data-testid="conversation-panel-wrapper"]') 
                                || document.querySelector('#main');
                            if (mainPane) {
                                mainPane.click();
                            }
                        """)
                        time.sleep(0.15)
                    except:
                        pass
                    
                    debug_print(f"  ✓ All dialogs closed successfully")
                        
                except Exception as e:
                    print(f"  ⚠️  Could not close all menus: {str(e)}")
                
                time.sleep(delay_seconds)
                return True
            else:
                # IMPORTANT (Mode 2): Never send caption as a separate text message if image send failed.
                # This avoids "text sent, image still attached" behavior.
                print(f"  ✗ Image send failed - NOT sending text-only fallback (Mode 2).")
                # Try to close any open media composer and attachment menu
                debug_print(f"  → Closing all menus after failure...")
                try:
                    # Press Escape multiple times to ensure all menus close
                    for i in range(4):
                        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                        time.sleep(0.2)
                except:
                    pass
                return False
        
        # Step 4: Auto type message
        debug_print(f"  → Typing message...")
        
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
            debug_print(f"  → Using JavaScript for emoji/special characters...")
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
        debug_print(f"  → Sending message...")
        
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
        
        # Initialize Safety Manager
        safety = SafetyManager()
        
        print(f"\n📱 Starting to send messages to {len(contacts)} contacts...")
        print(f"🔒 Safety features enabled:")
        print(f"   • Daily limit: {MAX_MESSAGES_PER_DAY} messages")
        print(f"   • Hourly limit: {MAX_MESSAGES_PER_HOUR} messages")
        print(f"   • Smart delays: {MIN_DELAY_BETWEEN_MESSAGES}-{MAX_DELAY_BETWEEN_MESSAGES}s")
        print(f"   • Auto-breaks: Every {MESSAGES_BEFORE_BREAK} messages")
        print(f"   • Active hours: {ACTIVE_HOURS_START}:00 - {ACTIVE_HOURS_END}:00")
        
        # Determine mode
        if default_image:
            print(f"\n📷 Active Mode: Mode 2 (Photo with caption)")
            print(f"   Image: {os.path.basename(default_image)}")
        else:
            print(f"\n📝 Active Mode: Mode 1 (Text messages only)")
        
        print(f"\n⚠️  Keep the browser window open and don't close it!")
        
        # Show current stats
        print(f"\n📊 Current Status:")
        print(f"   • Today: {safety.progress['messages_today']}/{MAX_MESSAGES_PER_DAY}")
        print(f"   • This hour: {safety.progress['messages_this_hour']}/{MAX_MESSAGES_PER_HOUR}")
        if safety.progress['last_contact_index'] > 0:
            print(f"   • Last contact: #{safety.progress['last_contact_index']}")
        print()
        
        successful = 0
        failed = 0
        failed_contacts = []  # Track failed contacts for retry file
        
        for index, contact in enumerate(contacts[start_from:], start=start_from):
            # Check if we can send (time restrictions, limits, etc.)
            can_send, reason = safety.can_send_message()
            if not can_send:
                print(f"\n⚠️  Cannot send: {reason}")
                if "hour" in reason.lower():
                    # Wait for next hour
                    print(f"⏳ Waiting for next hour...")
                    time.sleep(3600 - (datetime.now().minute * 60))
                    continue
                elif "Outside active hours" in reason:
                    safety.wait_until_active_hours()
                    continue
                else:
                    # Daily limit reached
                    print(f"✅ Daily limit reached. Will resume tomorrow.")
                    break
            
            # Check if we need a break
            if safety.needs_break():
                safety.take_break()
            
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
            
            # Get smart delay (randomized)
            smart_delay = safety.get_next_delay()
            
            # Send message (with image if available, text-only if image_path is None)
            if send_whatsapp_message(driver, contact['number'], contact['message'], smart_delay, image_path):
                successful += 1
                safety.record_message_sent(index)
                print(f"  ⏱️  Next delay: {smart_delay:.1f}s")
            else:
                failed += 1
                failed_contacts.append(contact)  # Track failed contact for retry
            
            # Progress update every 10 messages
            if (index + 1) % 10 == 0:
                print(f"\n📊 Progress: {index + 1}/{len(contacts)} | ✓ {successful} | ✗ {failed}")
                print(f"   Today: {safety.progress['messages_today']}/{MAX_MESSAGES_PER_DAY} | Hour: {safety.progress['messages_this_hour']}/{MAX_MESSAGES_PER_HOUR}")
        
        print(f"\n{'='*50}")
        print(f"✅ Completed!")
        print(f"✓ Successful: {successful}")
        print(f"✗ Failed: {failed}")
        print(f"{'='*50}")
        
        # Save failed contacts to retry file
        if failed_contacts:
            try:
                failed_file = "failed_contacts.xlsx"
                print(f"\n📝 Saving {len(failed_contacts)} failed contacts to {failed_file}...")
                
                # Create new workbook for failed contacts
                from openpyxl import Workbook
                wb = Workbook()
                ws = wb.active
                ws.title = "Failed Contacts"
                
                # Add header
                ws.append(["Number", "Name", "Message"])
                
                # Add failed contacts
                for contact in failed_contacts:
                    ws.append([contact['number'], contact.get('name', ''), contact['message']])
                
                # Save file
                wb.save(failed_file)
                print(f"✅ Failed contacts saved to {failed_file}")
                print(f"   Run the script again to automatically retry these contacts.")
                
            except Exception as e:
                print(f"⚠️  Could not save failed contacts: {str(e)}")
        elif os.path.exists("failed_contacts.xlsx") and failed == 0:
            # All contacts succeeded, delete the retry file
            try:
                os.remove("failed_contacts.xlsx")
                print(f"\n🎉 All retry contacts sent successfully!")
                print(f"✅ Deleted failed_contacts.xlsx (no longer needed)")
            except:
                pass
        
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
    import sys
    
    # Command line arguments for safety management
    if len(sys.argv) > 1:
        if sys.argv[1] == "stats":
            # Show statistics
            safety = SafetyManager()
            safety.print_stats()
            sys.exit(0)
        elif sys.argv[1] == "reset":
            # Reset progress
            print("\n⚠️  WARNING: This will reset all progress!")
            print("   You will lose:")
            print("   - Today's message count")
            print("   - Last contact index")
            print("   - All statistics")
            confirm = input("\nType 'YES' to confirm reset: ")
            if confirm == "YES":
                safety = SafetyManager()
                safety.reset_progress()
            else:
                print("Reset cancelled.")
            sys.exit(0)
    
    # Configuration
    DELAY_SECONDS = 2  # This will be overridden by SafetyManager's smart delays
    START_FROM = 0  # Start from this index (useful if you need to resume)
    
    # Check for failed contacts retry file
    FAILED_FILE = "failed_contacts.xlsx"
    if os.path.exists(FAILED_FILE):
        print("\n" + "="*60)
        print("⚠️  RETRY FILE DETECTED!")
        print("="*60)
        print(f"Found: {FAILED_FILE}")
        print(f"This file contains contacts that failed in the previous run.")
        print(f"\nOptions:")
        print(f"  1. Retry failed contacts (use {FAILED_FILE})")
        print(f"  2. Start fresh (use contacts.xlsx)")
        print("="*60)
        
        while True:
            retry_choice = input("\nRetry failed contacts? (1=Yes, 2=No): ").strip()
            if retry_choice == "1":
                EXCEL_FILE = FAILED_FILE
                print(f"\n✓ Using {FAILED_FILE} for retry")
                break
            elif retry_choice == "2":
                EXCEL_FILE = "contacts.xlsx"
                print(f"\n✓ Using contacts.xlsx (fresh start)")
                print(f"⚠️  Note: {FAILED_FILE} will remain on disk")
                break
            else:
                print("Invalid input. Please enter 1 or 2.")
    else:
        EXCEL_FILE = "contacts.xlsx"  # Default file
    
    # IMAGE CONFIGURATION - Interactive Mode Selectione.
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
            image_file = "image.jpg"
            
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
