# Step-by-Step Guide: Checking and Running the Code

This guide will walk you through checking your setup and running the WhatsApp message sender.

## Prerequisites Check

### Step 1: Verify Python Installation

**Windows PowerShell:**
```powershell
python --version
```

**Expected Output:**
```
Python 3.10.0
```
(or any version 3.7 or higher)

**If Python is not found:**
- Install Python from [python.org](https://www.python.org/downloads/)
- Make sure to check "Add Python to PATH" during installation

---

### Step 2: Verify pip is Installed

**Windows PowerShell:**
```powershell
pip --version
```

**Expected Output:**
```
pip 23.0.1 from ...
```

**If pip is not found:**
```powershell
python -m ensurepip --upgrade
```

---

### Step 3: Install Dependencies

Navigate to the project folder:
```powershell
cd C:\SalesIntel\repo\my_project\message-sender
```

Install required packages:
```powershell
pip install -r requirements.txt
```

**Expected Output:**
```
Successfully installed openpyxl-3.1.2 selenium-4.15.2 webdriver-manager-4.0.1 python-dotenv-1.0.0
```

**If you get errors:**
- Try: `pip install --user -r requirements.txt`
- Or: `python -m pip install -r requirements.txt`

---

## Code Validation Steps

### Step 4: Check Python Syntax

Verify the main script has no syntax errors:

**Windows PowerShell:**
```powershell
python -m py_compile whatsapp_sender.py
```

**Expected Output:**
- No output = Success (no errors)
- If errors appear, fix them before proceeding

**Check utility scripts:**
```powershell
python -m py_compile utils/check_code.py
python -m py_compile utils/create_template.py
```

---

### Step 5: Run Code Validator

Run the built-in code checker to validate your entire setup:

**Windows PowerShell:**
```powershell
python utils/check_code.py
```

**Expected Output:**
```
============================================================
1. Checking Python Version...
============================================================
   Python Version: 3.10.5
   ✓ Python version is compatible (3.7+)

============================================================
2. Checking Dependencies...
============================================================
   ✓ openpyxl is installed
   ✓ selenium is installed
   ✓ webdriver-manager is installed
   ✓ python-dotenv is installed

============================================================
3. Checking Excel File...
============================================================
   ✓ Excel file found: contacts.xlsx
   ✓ Excel file format is correct
   ✓ Loaded 2 contacts from contacts.xlsx

============================================================
4. Checking Chrome Browser...
============================================================
   ✓ Chrome browser is available

============================================================
✅ All checks passed! Your setup is ready.
```

**If any check fails:**
- Follow the error messages to fix issues
- Re-run the checker after fixing

---

### Step 6: Create/Verify Excel Template

Create a sample Excel file to test:

**Windows PowerShell:**
```powershell
python utils/create_template.py
```

**Expected Output:**
```
✓ File 'contacts.xlsx' created successfully!

Format:
  Column A: Contact Number (with country code, e.g., +1234567890)
  Column B: Message/Caption (text to send with image)
  Column C: Image Path (optional - leave empty to auto-detect)
```

**Verify the file was created:**
```powershell
Test-Path contacts.xlsx
```

**Expected Output:**
```
True
```

---

## Configuration Steps

### Step 7: Configure the Script

Open `whatsapp_sender.py` and check/edit these settings (around line 1001-1023):

```python
# Configuration
EXCEL_FILE = "contacts.xlsx"
DELAY_SECONDS = 5
START_FROM = 0

# Image Configuration - Choose ONE mode:
# Mode 1: Single image for all
DEFAULT_IMAGE = "safari_promo.jpg"  # or None
IMAGES_FOLDER = None

# Mode 2: Individual images
# DEFAULT_IMAGE = None
# IMAGES_FOLDER = "images"

# Mode 3: Text-only (no images)
# DEFAULT_IMAGE = None
# IMAGES_FOLDER = None
```

**For Text-Only Mode (Recommended for first test):**
```python
DEFAULT_IMAGE = None
IMAGES_FOLDER = None
```

---

### Step 8: Prepare Your Excel File

**Option A: Edit the template:**
1. Open `contacts.xlsx` in Excel
2. Add your contact numbers in Column A
3. Add your messages in Column B
4. Leave Column C empty (for text-only) or add image paths

**Option B: Create manually:**
- Column A: Contact Number (e.g., +919555611880)
- Column B: Message text
- Column C: Image Path (optional)

**Example:**
```
Contact Number    | Message                    | Image Path
+919555611880     | Hello! This is a test.     | 
+919355611880     | Hi there!                  |
```

---

## Running the Application

### Step 9: Run the Main Script

**Windows PowerShell:**
```powershell
python whatsapp_sender.py
```

**What happens:**
1. Script checks if `contacts.xlsx` exists
2. Shows configuration summary
3. Prompts: "Press Enter to start, or Ctrl+C to cancel..."
4. Opens Chrome browser
5. Navigates to WhatsApp Web
6. **You need to scan QR code with your phone**
7. Once logged in, starts sending messages

**Expected Output:**
```
==================================================
WhatsApp Bulk Message Sender
==================================================
File: contacts.xlsx
Delay: 5 seconds between messages

⚠️  Make sure:
  1. You have Chrome browser installed
  2. Your computer won't go to sleep
  3. You have your phone nearby to scan QR code
  4. Browser window will stay open during the process

Press Enter to start, or Ctrl+C to cancel...

✓ Loaded 2 contacts from contacts.xlsx
📝 Mode: Text messages only (no images)

🔧 Setting up Chrome browser...
   ✓ Chrome browser initialized successfully!

📱 Opening WhatsApp Web...

⚠️  Please scan the QR code with your phone to log in to WhatsApp Web
   Waiting for you to complete login...
✓ Successfully logged in to WhatsApp Web!

📱 Starting to send messages to 2 contacts...
⏱️  Delay between messages: 5 seconds
📝 Mode: Text messages only (no images)
⚠️  Keep the browser window open and don't close it!


[1/2] Sending to +919555611880...
✓ Message sent to +919555611880

[2/2] Sending to +919355611880...
✓ Message sent to +919355611880

==================================================
✅ Completed!
✓ Successful: 2
✗ Failed: 0
==================================================
```

---

## Testing Checklist

Before running on all contacts, test with 2-3 contacts:

- [ ] Python is installed and working
- [ ] All dependencies are installed
- [ ] Code validator passes all checks
- [ ] Excel file is created and formatted correctly
- [ ] Configuration is set correctly
- [ ] Chrome browser is installed
- [ ] Test with 2-3 contacts first
- [ ] Verify messages are sent successfully
- [ ] Then run with full contact list

---

## Troubleshooting Steps

### If "Python not found":
```powershell
# Check if Python is in PATH
where.exe python

# If not found, reinstall Python with "Add to PATH" checked
```

### If "Module not found":
```powershell
# Reinstall dependencies
pip install --upgrade -r requirements.txt
```

### If "Excel file not found":
```powershell
# Check current directory
Get-Location

# Verify file exists
Test-Path contacts.xlsx

# Create template if missing
python utils/create_template.py
```

### If "ChromeDriver error":
```powershell
# Run the fix script
.\scripts\fix_chromedriver.ps1

# Or manually clear cache
Remove-Item -Recurse -Force $env:USERPROFILE\.wdm
```

### If "Contact not found":
- Make sure contact numbers are saved in your WhatsApp contacts
- Verify number format includes country code (e.g., +919555611880)
- Check that numbers are correct in Excel file

---

## Quick Test Sequence

Run these commands in order for a complete test:

```powershell
# 1. Check Python
python --version

# 2. Install dependencies
pip install -r requirements.txt

# 3. Validate setup
python utils/check_code.py

# 4. Create test Excel file
python utils/create_template.py

# 5. Edit contacts.xlsx with 2-3 test contacts

# 6. Run the script
python whatsapp_sender.py
```

---

## Common Issues and Solutions

### Issue: "pip: command not found"
**Solution:**
```powershell
python -m pip install -r requirements.txt
```

### Issue: Permission denied
**Solution:**
```powershell
pip install --user -r requirements.txt
```

### Issue: ChromeDriver error [WinError 193]
**Solution:**
```powershell
.\scripts\fix_chromedriver.ps1
```

### Issue: Messages not sending
**Solutions:**
- Verify WhatsApp Web is logged in (QR code scanned)
- Check contact numbers are in your WhatsApp contacts
- Ensure internet connection is stable
- Increase DELAY_SECONDS if rate limited

### Issue: Browser closes immediately
**Solution:**
- Don't close the browser window manually
- Let the script handle browser closing
- Check Chrome is up to date

---

## Next Steps After Testing

Once testing is successful:

1. **Add all your contacts** to `contacts.xlsx`
2. **Configure images** (if needed):
   - Single image: Set `DEFAULT_IMAGE`
   - Individual images: Set `IMAGES_FOLDER = "images"`
3. **Adjust delay** if needed (5-10 seconds recommended)
4. **Run for full batch**
5. **Monitor progress** in console output

---

## Support

- **Code Issues**: Run `python utils/check_code.py`
- **Image Setup**: See `docs/IMAGE_GUIDE.md`
- **PowerShell Help**: See `docs/POWERSHELL_GUIDE.md`
- **Git Setup**: See `docs/GIT_SETUP.md`
