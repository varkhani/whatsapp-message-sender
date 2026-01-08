# MODE 1: Single Image for All Contacts - Complete Step-by-Step Flow

## Overview
**MODE 1** is designed for **campaign messaging** where you want to send the **same image** to all contacts with **personalized captions** from your Excel file.

## Configuration
In `whatsapp_sender.py`:
```python
DEFAULT_IMAGE = "safari_promo.jpg"  # Same image for everyone
IMAGES_FOLDER = None  # Not needed for Mode 1
```

---

## Complete Execution Flow

### **PHASE 1: Initialization & Setup**

#### Step 1: Script Starts
- Location: `if __name__ == "__main__"` (line 3526)
- Actions:
  - Reads configuration variables:
    - `EXCEL_FILE = "contacts.xlsx"`
    - `DELAY_SECONDS = 2` (delay between messages)
    - `START_FROM = 0` (resume from index)
    - `DEFAULT_IMAGE = "safari_promo.jpg"` ⭐ **MODE 1 KEY**
    - `IMAGES_FOLDER = None` ⭐ **MODE 1 KEY**

#### Step 2: Excel File Validation
- Location: Lines 3547-3555
- Actions:
  - Checks if `contacts.xlsx` exists
  - If not found, displays error and exits
  - Shows expected Excel format:
    - Column A: Contact Number
    - Column B: Contact Name (optional)
    - Column C: Message/Caption
    - Column D: Image Path (ignored in Mode 1)

#### Step 3: User Confirmation
- Location: Lines 3557-3568
- Actions:
  - Displays configuration summary
  - Shows warnings (Chrome, phone, QR code)
  - Waits for user to press Enter to start

---

### **PHASE 2: Excel Reading & Contact Loading**

#### Step 4: Read Contacts from Excel
- Location: `read_contacts_from_excel()` function (line 23)
- Actions:
  - Opens `contacts.xlsx` using `openpyxl`
  - Skips header row (starts from row 2)
  - For each row:
    - **Column A**: Contact Number (cleaned: removes spaces, dashes, parentheses)
    - **Column B**: Contact Name (optional)
      - If provided: Message format = `"Dear {Name},\n\n{Message}"`
      - If empty: Message format = `"{Message}"` (as-is)
    - **Column C**: Message/Caption text
    - **Column D**: Image Path (read but **ignored in Mode 1**)
  - Returns list of contact dictionaries:
    ```python
    {
        'number': '+919555611880',
        'message': 'Dear John,\n\nYour personalized message here',
        'image_path': None  # Ignored in Mode 1
    }
    ```

#### Step 5: Validate Default Image
- Location: `send_bulk_messages()` function, lines 3332-3342
- Actions:
  - Checks if `DEFAULT_IMAGE` file exists
  - Converts relative path to absolute path (if needed)
  - If image found:
    - ✅ Prints: `"✓ Default image found: safari_promo.jpg"`
    - ✅ Prints: `"📷 Will send this image to ALL contacts"`
  - If image NOT found:
    - ⚠️ Prints warning
    - Sets `default_image = None` (falls back to text-only)

#### Step 6: Skip Images Folder Check
- Location: Lines 3344-3354
- Actions:
  - Since `IMAGES_FOLDER = None` and `default_image` exists, this check is **skipped**
  - Mode 1 doesn't use individual images folder

---

### **PHASE 3: Chrome Browser Setup**

#### Step 7: Initialize Chrome Driver
- Location: Lines 3358-3440
- Actions:
  - Creates Chrome options with:
    - User data directory: `./chrome_profile` (persistent session)
    - Stability flags: `--no-sandbox`, `--disable-dev-shm-usage`
    - Anti-detection: `--disable-blink-features=AutomationControlled`
    - Remote debugging port: `9222`
  - Downloads/updates ChromeDriver using `ChromeDriverManager`
  - Starts Chrome browser with custom profile
  - Prints: `"✓ Chrome browser initialized successfully!"`

#### Step 8: Initialize WhatsApp Web
- Location: `init_whatsapp_web()` function (line 95)
- Actions:
  - Navigates to `https://web.whatsapp.com`
  - Waits for QR code to appear
  - Displays: `"⚠️ Please scan the QR code with your phone"`
  - Waits up to 300 seconds for user to scan QR code
  - Detects successful login by finding search box element
  - Prints: `"✓ Successfully logged in to WhatsApp Web!"`

#### Step 9: Ensure Main Page
- Location: `ensure_main_page()` function (line 119)
- Actions:
  - Verifies we're on WhatsApp Web (not redirected)
  - Ensures main chat list is visible
  - Ready to start sending messages

---

### **PHASE 4: Message Sending Loop**

#### Step 10: Display Mode Information
- Location: Lines 3455-3465
- Actions:
  - Determines mode based on `default_image`:
    - ✅ **Mode 1 detected**: `default_image` is set
  - Prints:
    ```
    📷 Mode: Single image for all contacts
       Image: safari_promo.jpg
    ```
  - Shows total contacts and delay settings

#### Step 11: Loop Through Each Contact
- Location: Lines 3470-3503
- For each contact in Excel file:

##### **11a. Get Image Path for Contact**
- Location: Lines 3473-3493
- Actions:
  - Since `default_image` is set (Mode 1):
    - ✅ Sets `image_path = default_image` (same image for all)
    - Prints (only once): `"📷 Using default image: safari_promo.jpg"`

##### **11b. Send WhatsApp Message**
- Location: `send_whatsapp_message()` function (line 2977)
- Called with: `(driver, contact_number, message, delay_seconds, image_path)`

---

### **PHASE 5: Individual Message Sending Process**

For each contact, the following steps execute:

#### Step 12: Ensure Main Page
- Location: `send_whatsapp_message()`, line 2989
- Actions:
  - Calls `ensure_main_page(driver)` to ensure we're on main chat list
  - Not in any specific chat conversation

#### Step 13: Search for Contact
- Location: Lines 2992-3012
- Actions:
  - Finds search box: `//div[@contenteditable='true'][@data-tab='3']`
  - Cleans contact number (removes spaces, dashes, etc.)
  - Types contact number into search box
  - Waits 1 second for search results to appear
  - Prints: `"→ Searching for contact..."`

#### Step 14: Select Contact
- Location: Lines 3014-3032
- Actions:
  - Presses `Arrow Down` key to highlight first search result
  - Presses `Enter` key to open chat
  - Fallback: Clicks first result if keyboard navigation fails
  - Waits 1.5 seconds for chat to open
  - Prints: `"→ Selecting contact..."`

#### Step 15: Find Message Box
- Location: Lines 3034-3042
- Actions:
  - Calls `get_fresh_message_box(driver)` to find message input box
  - Uses multiple XPath selectors to locate message box
  - Waits up to 5 retries if not found immediately
  - Prints: `"→ Finding message box..."`

#### Step 16: Check for Image Path (Mode 1 Detection)
- Location: Lines 3044-3056
- Actions:
  - Since `image_path` is provided (default_image):
    - ✅ Detects image mode
    - Prints: `"📷 Sending image with caption..."`
    - Calls `send_image_with_caption()` function
    - If successful: Returns `True`, waits delay, goes back to main page
    - If failed: Falls back to text-only sending

---

### **PHASE 6: Image Upload Process (Mode 1 Specific)**

#### Step 17: Verify Chat is Open
- Location: `send_image_with_caption()`, lines 652-656
- Actions:
  - Verifies we're actually in a chat conversation
  - Checks if message box is visible
  - Ensures chat header exists

#### Step 18: Validate Image File
- Location: Lines 658-664
- Actions:
  - Converts image path to absolute path
  - Checks if `safari_promo.jpg` exists on disk
  - If not found: Falls back to text-only message
  - If found: Continues with image upload

#### Step 19: Wait for Chat to Load
- Location: Line 668
- Actions:
  - Waits 2 seconds for chat interface to fully load
  - Ensures all UI elements are ready

#### Step 20: Find Attachment Button
- Location: Lines 670-703
- Actions:
  - Searches for attachment/clip button using multiple selectors:
    - `//span[@data-testid='clip']`
    - `//div[@data-testid='clip']`
    - `//span[@data-icon='attach']`
  - Tries up to 5 times with 0.5s delays
  - Clicks attachment button when found
  - Prints: `"→ Looking for attachment button..."` → `"→ Clicking attachment button..."`

#### Step 21: Wait for Attachment Menu
- Location: Line 716
- Actions:
  - Waits 1.5 seconds for attachment menu popup to appear
  - Menu shows options: Document, Photos & videos, Camera, etc.

#### Step 22: Select "Photos & videos" Option
- Location: Lines 718-1100+ (complex logic)
- Actions:
  - **PRIMARY METHOD**: Keyboard Navigation (most reliable)
    - Focuses on attachment menu
    - Presses `Arrow Down` to navigate to "Photos & videos" (2nd option)
    - Presses `Enter` to select
    - Verifies NOT in sticker mode
    - Verifies file input is available
  - **FALLBACK METHOD**: Element Clicking
    - Finds "Photos & videos" button in menu
    - Explicitly excludes "New sticker" option
    - Clicks the correct option
  - Waits 3.5 seconds for file picker to appear
  - Prints: `"→ Selecting 'Photos & videos' option..."`

#### Step 23: Upload Image File
- Location: Lines 1100-1200+
- Actions:
  - Finds file input element: `//input[@type='file']`
  - Verifies it accepts images (not stickers)
  - Uses `send_keys(image_path)` to upload file
  - Waits for image preview to appear
  - Prints: `"→ Uploading image file..."`

#### Step 24: Add Caption to Image
- Location: Lines 1200-1400+
- Actions:
  - Waits for image preview to fully load
  - Finds caption input box (usually appears below image preview)
  - Types the personalized message from Excel (Column C)
  - Handles emojis and special characters using JavaScript
  - Handles newlines properly (Shift+Enter)
  - Prints: `"→ Adding caption..."`

#### Step 25: Send Image with Caption
- Location: Lines 1400-1500+
- Actions:
  - Finds send button or presses `Enter` key
  - Verifies message was sent (checks if message box is cleared)
  - Waits for confirmation
  - Prints: `"✓ Message sent to {contact_number}"`

#### Step 26: Return to Main Page
- Location: Lines 3051-3052
- Actions:
  - Presses `Escape` key to close chat
  - Returns to main chat list
  - Ready for next contact

---

### **PHASE 7: Delay & Next Contact**

#### Step 27: Wait Delay Period
- Location: Line 3049
- Actions:
  - Waits for `DELAY_SECONDS` (default: 2 seconds)
  - Prevents rate limiting by WhatsApp
  - Allows time for message to be processed

#### Step 28: Progress Update
- Location: Lines 3501-3503
- Actions:
  - Every 10 messages, prints progress:
    ```
    📊 Progress: 10/50 | ✓ 10 | ✗ 0
    ```

#### Step 29: Repeat for Next Contact
- Location: Line 3470 (loop continues)
- Actions:
  - Goes back to Step 11 for next contact
  - Uses **same image** (`safari_promo.jpg`) for all contacts
  - Uses **different caption** from Excel Column C for each contact

---

### **PHASE 8: Completion**

#### Step 30: Final Summary
- Location: Lines 3505-3509
- Actions:
  - After all contacts processed:
  - Prints completion summary:
    ```
    ==================================================
    ✅ Completed!
    ✓ Successful: 45
    ✗ Failed: 5
    ==================================================
    ```

#### Step 31: Cleanup
- Location: Lines 3516-3523
- Actions:
  - Waits 5 seconds
  - Closes Chrome browser
  - Exits script

---

## Key Characteristics of MODE 1

### ✅ **Advantages:**
1. **Simple Setup**: Only need one image file
2. **Fast**: No need to search for images per contact
3. **Consistent**: Same image for all contacts (campaign mode)
4. **Flexible**: Each contact gets personalized caption from Excel

### 📋 **Excel File Requirements:**
- **Column A**: Contact Number (required)
- **Column B**: Contact Name (optional - adds "Dear {Name}," prefix)
- **Column C**: Message/Caption (required - personalized text)
- **Column D**: Image Path (ignored in Mode 1 - can be empty)

### 🖼️ **Image Requirements:**
- Image file must exist in project folder (or provide full path)
- Supported formats: `.jpg`, `.jpeg`, `.png`, `.gif`, `.webp`
- Recommended size: Under 5MB for faster upload

### ⚙️ **Configuration:**
```python
DEFAULT_IMAGE = "safari_promo.jpg"  # Same image for all
IMAGES_FOLDER = None  # Must be None for Mode 1
```

---

## Error Handling

### If Image Not Found:
- Script detects missing image at Step 5
- Falls back to **text-only mode**
- Sends caption as regular message (no image)

### If Contact Not Found:
- Search returns no results
- Skips contact, marks as failed
- Continues with next contact

### If WhatsApp Web Disconnects:
- Script may timeout
- User can resume from last successful index using `START_FROM`

---

## Summary Flow Diagram

```
START
  ↓
Read Excel File
  ↓
Validate DEFAULT_IMAGE exists
  ↓
Initialize Chrome Browser
  ↓
Open WhatsApp Web & Scan QR
  ↓
FOR EACH CONTACT:
  ├─ Search Contact
  ├─ Open Chat
  ├─ Click Attachment Button
  ├─ Select "Photos & videos"
  ├─ Upload DEFAULT_IMAGE (same for all)
  ├─ Type Caption (from Excel Column C)
  ├─ Send Message
  ├─ Wait DELAY_SECONDS
  └─ Return to Main Page
  ↓
Print Final Summary
  ↓
Close Browser
  ↓
END
```

---

## Example Execution

**Input:**
- `DEFAULT_IMAGE = "safari_promo.jpg"`
- Excel with 3 contacts:
  - Contact 1: `+919555611880`, Message: `"Check out this offer!"`
  - Contact 2: `+919355611880`, Message: `"Special discount for you!"`
  - Contact 3: `+919455611880`, Message: `"Limited time offer!"`

**Output:**
- All 3 contacts receive **same image**: `safari_promo.jpg`
- Each contact receives **different caption** from their Excel row
- Total: 3 messages sent (1 image + 3 unique captions)

---

*This document covers the complete execution flow for MODE 1: Single Image for All Contacts.*
