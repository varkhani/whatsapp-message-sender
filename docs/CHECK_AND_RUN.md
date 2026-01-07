# How to Check and Run the Code - Visual Guide

## 📋 Complete Step-by-Step Process

### Phase 1: Initial Setup (One-Time)

#### ✅ Step 1: Verify Python
```powershell
python --version
```
**Expected:** `Python 3.x.x` (3.7 or higher)

#### ✅ Step 2: Install Dependencies
```powershell
cd C:\SalesIntel\repo\my_project\message-sender
pip install -r requirements.txt
```
**Expected:** `Successfully installed...`

#### ✅ Step 3: Verify Installation
```powershell
pip list | Select-String "selenium|openpyxl|webdriver"
```
**Expected:** Shows installed packages

---

### Phase 2: Code Validation

#### ✅ Step 4: Check Python Syntax
```powershell
python -m py_compile whatsapp_sender.py
```
**Expected:** No output (means success)

#### ✅ Step 5: Run Complete Validator
```powershell
python utils/check_code.py
```
**Expected Output:**
```
============================================================
1. Checking Python Version...
============================================================
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
   ✓ Loaded X contacts from contacts.xlsx

============================================================
4. Checking Chrome Browser...
============================================================
   ✓ Chrome browser is available

============================================================
✅ All checks passed! Your setup is ready.
```

**If any check fails:**
- Fix the issue shown in the error message
- Re-run the validator

---

### Phase 3: Prepare Data

#### ✅ Step 6: Create Excel Template
```powershell
python utils/create_template.py
```
**Expected:** `✓ File 'contacts.xlsx' created successfully!`

#### ✅ Step 7: Edit Excel File
1. Open `contacts.xlsx` in Excel
2. Add your data:
   - **Column A**: Contact Numbers (e.g., +919555611880)
   - **Column B**: Messages
   - **Column C**: Image Path (optional, leave empty for text-only)
3. Save the file

**Example:**
```
Contact Number    | Message              | Image Path
+919555611880     | Hello! Test message   | 
+919355611880     | Hi there!            |
```

---

### Phase 4: Configuration

#### ✅ Step 8: Configure Script
Open `whatsapp_sender.py` and check these settings (around line 1001-1023):

**For Text-Only Mode (Recommended for first test):**
```python
DEFAULT_IMAGE = None
IMAGES_FOLDER = None
```

**For Single Image Mode:**
```python
DEFAULT_IMAGE = "safari_promo.jpg"
IMAGES_FOLDER = None
```

**For Individual Images:**
```python
DEFAULT_IMAGE = None
IMAGES_FOLDER = "images"
```

---

### Phase 5: Run the Application

#### ✅ Step 9: Execute the Script
```powershell
python whatsapp_sender.py
```

**What You'll See:**
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
```

#### ✅ Step 10: Start the Process
1. **Press Enter** to start
2. Chrome browser will open automatically
3. **Scan QR code** with your WhatsApp phone app
4. Wait for login confirmation
5. Script will start sending messages automatically

**During Execution:**
```
✓ Loaded 2 contacts from contacts.xlsx
📝 Mode: Text messages only (no images)

🔧 Setting up Chrome browser...
   ✓ Chrome browser initialized successfully!

📱 Opening WhatsApp Web...
✓ Successfully logged in to WhatsApp Web!

📱 Starting to send messages to 2 contacts...

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

## 🔄 Daily Workflow

Once setup is complete, your daily workflow is:

1. **Update Excel file** with new contacts/messages
2. **Run the script:**
   ```powershell
   python whatsapp_sender.py
   ```
3. **Press Enter** and scan QR code
4. **Wait for completion**

---

## 🛠️ Troubleshooting Commands

### Check Everything is Working:
```powershell
# 1. Python version
python --version

# 2. Dependencies
pip list | Select-String "selenium|openpyxl"

# 3. Code validation
python utils/check_code.py

# 4. Excel file
Test-Path contacts.xlsx

# 5. Syntax check
python -m py_compile whatsapp_sender.py
```

### Fix Common Issues:
```powershell
# Reinstall dependencies
pip install --upgrade -r requirements.txt

# Fix ChromeDriver
.\scripts\fix_chromedriver.ps1

# Clear ChromeDriver cache
Remove-Item -Recurse -Force $env:USERPROFILE\.wdm
```

---

## 📊 Verification Checklist

Before running on all contacts:

- [ ] Python installed and working
- [ ] All packages installed (`pip install -r requirements.txt`)
- [ ] Code validator passes (`python utils/check_code.py`)
- [ ] Excel file exists and has correct format
- [ ] Tested with 2-3 contacts first
- [ ] Configuration is correct in `whatsapp_sender.py`
- [ ] Chrome browser is installed
- [ ] WhatsApp account is ready
- [ ] Phone is nearby for QR code scanning

---

## 🎯 Quick Test Sequence

Copy and paste this entire sequence:

```powershell
# Navigate to project
cd C:\SalesIntel\repo\my_project\message-sender

# Check Python
python --version

# Install dependencies
pip install -r requirements.txt

# Validate setup
python utils/check_code.py

# Create template
python utils/create_template.py

# Edit contacts.xlsx with 2-3 test contacts

# Run the script
python whatsapp_sender.py
```

---

## 📚 Additional Resources

- **Detailed Guide**: See `docs/RUN_STEPS.md` for complete instructions
- **Quick Reference**: See `docs/QUICK_START.md` for fast setup
- **Image Setup**: See `docs/IMAGE_GUIDE.md` for image configuration
- **PowerShell Help**: See `docs/POWERSHELL_GUIDE.md` for command reference

---

## ⚠️ Important Notes

1. **First Time**: Always test with 2-3 contacts before bulk sending
2. **QR Code**: You need to scan QR code every time (unless using saved Chrome profile)
3. **Browser**: Keep Chrome window open during the process
4. **Internet**: Keep your phone connected to internet
5. **Rate Limits**: Use appropriate delays (5+ seconds) to avoid restrictions

---

## 🆘 Need Help?

1. **Run the validator**: `python utils/check_code.py`
2. **Check error messages** in console output
3. **See troubleshooting section** in `docs/RUN_STEPS.md`
4. **Verify each step** in the checklist above
