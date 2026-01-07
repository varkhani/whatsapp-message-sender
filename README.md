# WhatsApp Bulk Message Sender

A simple Python application to send WhatsApp messages (with images and captions) to multiple contacts from an Excel file.

## 🚀 Quick Start

1. **Install dependencies:**
   ```powershell
   pip install -r requirements.txt
   ```

2. **Check your setup:**
   ```powershell
   python utils/check_code.py
   ```

3. **Create Excel file:**
   ```powershell
   python utils/create_template.py
   ```

4. **Configure images** (in `whatsapp_sender.py`):
   ```python
   DEFAULT_IMAGE = None      # Text-only mode
   IMAGES_FOLDER = None      # No images
   ```

5. **Run:**
   ```powershell
   python whatsapp_sender.py
   ```

📖 **Documentation:**
- **Step-by-Step Guide**: `docs/RUN_STEPS.md` - Complete setup and run instructions
- **Quick Reference**: `docs/QUICK_START.md` - Fast setup guide
- **Image Setup**: `docs/IMAGE_GUIDE.md` - Image sending guide
- **Single Image**: `docs/SINGLE_IMAGE_GUIDE.md` - Single image mode

## Features

- 📊 Reads contact numbers and messages from Excel (XLSX) files
- 📱 Sends WhatsApp messages automatically
- 🖼️ **Send images with captions** - Support for personalized images per contact
- 📷 **Single image mode** - Send one image to all contacts with different captions
- ⏱️ Configurable delay between messages to avoid rate limiting
- 📈 Progress tracking and error handling
- 🔄 Resume capability (can start from a specific index)
- 🌐 Supports Hindi/English text with emojis

## Prerequisites

- Python 3.7 or higher
- Google Chrome browser (required for Selenium automation)
- WhatsApp account (you'll scan QR code to log in)

## Installing Python (If Not Already Installed)

If you don't have Python installed on your computer, follow these steps:

### Check if Python is Already Installed

First, check if Python is already installed:

**Windows PowerShell:**
1. Open PowerShell (Press `Win + X`, then select "Windows PowerShell" or "Terminal")
2. Type: `python --version`
3. If you see a version number (e.g., "Python 3.11.5"), Python is installed!

**Windows Command Prompt:**
1. Open Command Prompt (Press `Win + R`, type `cmd`, press Enter)
2. Type: `python --version`
3. If you see a version number (e.g., "Python 3.11.5"), Python is installed!

**Mac/Linux:**
1. Open Terminal
2. Type: `python3 --version`
3. If you see a version number, Python is installed!

### Install Python

#### **Windows:**

1. **Download Python:**
   - Go to [python.org/downloads](https://www.python.org/downloads/)
   - Click the big yellow "Download Python" button (downloads the latest version)
   - Or go to [python.org/downloads/windows](https://www.python.org/downloads/windows/) for specific versions

2. **Run the Installer:**
   - Double-click the downloaded `.exe` file
   - **IMPORTANT:** Check the box "Add Python to PATH" at the bottom of the installer
   - Click "Install Now"
   - Wait for installation to complete

3. **Verify Installation:**
   - Open PowerShell (Press `Win + X`, then select "Windows PowerShell")
   - Type: `python --version`
   - You should see something like "Python 3.11.5"

#### **Mac:**

1. **Option 1 - Using Official Installer:**
   - Go to [python.org/downloads](https://www.python.org/downloads/)
   - Download the macOS installer
   - Run the `.pkg` file and follow the installation wizard

2. **Option 2 - Using Homebrew (Recommended):**
   ```bash
   brew install python3
   ```

3. **Verify Installation:**
   - Open Terminal
   - Type: `python3 --version`
   - You should see a version number

#### **Linux (Ubuntu/Debian):**

```bash
# Update package list
sudo apt update

# Install Python 3
sudo apt install python3 python3-pip

# Verify installation
python3 --version
```

#### **Linux (Fedora/CentOS/RHEL):**

```bash
# Install Python 3
sudo dnf install python3 python3-pip

# Verify installation
python3 --version
```

### Install pip (Python Package Manager)

pip usually comes with Python, but if you get errors, install it:

**Windows:**
```bash
python -m ensurepip --upgrade
```

**Mac/Linux:**
```bash
python3 -m ensurepip --upgrade
```

### Troubleshooting Python Installation

- **"python is not recognized" (Windows):**
  - Python wasn't added to PATH during installation
  - Reinstall Python and make sure to check "Add Python to PATH"
  - Or manually add Python to PATH in System Environment Variables

- **"command not found" (Mac/Linux):**
  - Try using `python3` instead of `python`
  - Make sure Python is installed correctly

- **Permission errors:**
  - On Mac/Linux, you might need to use `sudo` for some commands
  - Or install Python using a package manager

## Installation Steps

### 1. Install Python Dependencies

**Windows PowerShell:**
```powershell
pip install -r requirements.txt
```

**Windows Command Prompt:**
```cmd
pip install -r requirements.txt
```

**Mac/Linux:**
```bash
pip3 install -r requirements.txt
```

> **Note:** 
> - On Windows, use `pip` (not `pip3`)
> - If you get permission errors on Mac/Linux, use `pip3 install --user -r requirements.txt` or `sudo pip3 install -r requirements.txt`
> - **For Windows PowerShell users, see `POWERSHELL_GUIDE.md` for complete PowerShell command reference**

### 2. Create Your Contacts Excel File

**Option 1: Use the template generator (Recommended)**
```powershell
python utils/create_template.py
```
This will create `contacts.xlsx` with sample data that you can edit.

**Option 2: Create manually**
Create an Excel file named `contacts.xlsx` with the following format:

| Contact Number | Message (Caption) | Image Path (Optional) |
|----------------|------------------|----------------------|
| +1234567890 | Your message here | images/agent1.jpg |
| +0987654321 | Another message | |

**Excel File Format:**
- **Column A**: Contact Number (must include country code, e.g., +919555611880 for India)
- **Column B**: Message/Caption text (supports Hindi, English, emojis)
- **Column C**: Image Path (optional - leave empty for auto-detection or text-only)
- First row can be headers (will be skipped automatically)
- The file must be named `contacts.xlsx`

**📷 Image Support - Two Modes:**

#### **Mode 1: Single Image for All Contacts** (Recommended for Campaigns)

Send the same image to all contacts with personalized captions:

1. Place your image file (e.g., `safari_promo.jpg`) in the project folder
2. In `whatsapp_sender.py`, set:
   ```python
   DEFAULT_IMAGE = "safari_promo.jpg"  # Same image for everyone
   ```
3. Excel file only needs 2 columns (A & B) - Column C is ignored

**Example:**
```
message-sender/
├── contacts.xlsx
├── safari_promo.jpg    <- Single image for all
└── whatsapp_sender.py
```

#### **Mode 2: Individual Images per Contact**

Send unique images to each contact:

1. Create an `images/` folder
2. Add images named by contact number: `919555611880.jpg`, `919355611880.jpg`
3. Or specify image path in Excel Column C: `images/agent1.jpg`
4. In `whatsapp_sender.py`, set:
   ```python
   DEFAULT_IMAGE = None  # Disable single image mode
   IMAGES_FOLDER = "images"  # Folder with individual images
   ```

**Example folder structure:**
```
message-sender/
├── contacts.xlsx
├── images/
│   ├── 919555611880.jpg    # Auto-detected for +919555611880
│   ├── agent1.jpg          # Use in Excel: images/agent1.jpg
│   └── agent2.jpg
└── whatsapp_sender.py
```

**Image Detection Priority:**
1. Column C in Excel (if specified)
2. Contact-specific image: `{contact_number}.jpg` in images folder
3. Any image in images folder
4. Default image (if `DEFAULT_IMAGE` is set)

**Supported Image Formats:** `.jpg`, `.jpeg`, `.png`, `.gif`, `.webp`

**⚠️ Important Notes:**
- **Contacts**: For best results, **save the contact numbers in your WhatsApp contacts first**
- **Images**: Keep images under 5MB for faster upload
- If image not found, script will send text-only message as fallback
- Each contact can have a unique caption (Column B) even with same image
- See `docs/IMAGE_GUIDE.md` for detailed image setup instructions

### 3. Run the Application

When you run the script, it will:
1. Open Chrome browser automatically
2. Navigate to WhatsApp Web
3. **You need to scan the QR code** with your phone to log in
4. Once logged in, it will start sending messages automatically

## Checking Your Code & Setup

Before running the application, it's a good idea to verify everything is set up correctly:

### Quick Check (Recommended)

Run the code checker to validate your setup:

**Windows PowerShell:**
```powershell
python utils/check_code.py
```

**Windows Command Prompt:**
```cmd
python utils/check_code.py
```

**Mac/Linux:**
```bash
python3 utils/check_code.py
```

This will check:
- ✓ Python version compatibility
- ✓ All required packages are installed
- ✓ Excel file exists and has correct format
- ✓ Chrome browser is available
- ✓ Code syntax is valid
- ✓ Can read contacts from Excel file

### Manual Code Checks

#### 1. **Check Python Syntax**

**Windows PowerShell:**
```powershell
python -m py_compile whatsapp_sender.py
```

**Mac/Linux:**
```bash
python3 -m py_compile whatsapp_sender.py
```

#### 2. **Test Reading Excel File**
Create a test file with 2-3 contacts and verify it reads correctly:

**Windows PowerShell:**
```powershell
python utils/create_template.py
python utils/check_code.py
```

**Mac/Linux:**
```bash
python3 utils/create_template.py
python3 utils/check_code.py
```

#### 3. **Verify Dependencies**

**Windows PowerShell:**
```powershell
pip list | Select-String "selenium|openpyxl"
```

**Windows Command Prompt:**
```cmd
pip list | findstr "selenium openpyxl"
```

**Mac/Linux:**
```bash
pip3 list | grep -E "selenium|openpyxl"
```

## Usage

### Basic Usage

**Windows PowerShell:**
```powershell
python whatsapp_sender.py
```

**Windows Command Prompt:**
```cmd
python whatsapp_sender.py
```

**Mac/Linux:**
```bash
python3 whatsapp_sender.py
```

> **Note:** 
> - On Windows, use `python` (not `python3`)
> - On Mac and Linux, use `python3`
> - **Windows users: See `POWERSHELL_GUIDE.md` for complete PowerShell command reference**

The script will:
1. Read contacts from `contacts.xlsx`
2. Open WhatsApp Web in your browser
3. Send messages one by one with a delay between each

### Configuration

Edit the configuration in `whatsapp_sender.py`:

```python
EXCEL_FILE = "contacts.xlsx"  # Your Excel file name
DELAY_SECONDS = 5  # Delay between messages (minimum 3-5 seconds recommended, increase if rate limited)
START_FROM = 0  # Start from this index (useful for resuming)

# Image Configuration
DEFAULT_IMAGE = None  # Single image for all contacts (e.g., "safari_promo.jpg" or None)
IMAGES_FOLDER = "images"  # Folder containing individual images (or None to disable)
```

**Image Configuration Examples:**

**Single Image Mode (Campaign):**
```python
DEFAULT_IMAGE = "safari_promo.jpg"  # Same image for all contacts
IMAGES_FOLDER = None  # Not needed
```

**Individual Images Mode:**
```python
DEFAULT_IMAGE = None  # Disable single image
IMAGES_FOLDER = "images"  # Folder with unique images per contact
```

**Text-Only Mode:**
```python
DEFAULT_IMAGE = None
IMAGES_FOLDER = None  # No images, text messages only
```

### For Large Batches (700-2000 users)

For sending to 700-2000 users weekly, consider:

1. **Delay settings**: 
   - Default is 5 seconds (fast but safe)
   - Minimum recommended: 3-5 seconds
   - If you get rate limited, increase to 8-10 seconds
2. **Run in batches**: Split your Excel file into smaller files (e.g., 200 contacts each)
3. **Run overnight**: Start the process when you won't need your computer
4. **Monitor progress**: Check the console output for any errors

**Example timing:**
- 1000 contacts with 5-second delay: ~1.5-2 hours
- 1000 contacts with 8-second delay: ~2.5-3 hours
- Recommended: Run overnight or during off-hours

## Why Selenium Instead of pywhatkit or Twilio?

### ✅ **Selenium (Current Solution)**
- **Free** - No cost
- **Reliable** - More stable than pywhatkit
- **Full Control** - Direct browser automation
- **Good for 700-2000 users** - Handles bulk messaging well
- **Cons**: Requires browser to stay open, needs Chrome

### ❌ **pywhatkit** (Not Recommended)
- **Unreliable** - Breaks frequently with WhatsApp updates
- **Limited control** - Can't handle errors well
- **Not suitable for bulk** - Not designed for 700-2000 messages
- **Cons**: Often fails, poor error handling

### 💰 **Twilio** (For Production/Enterprise)
- **Most Reliable** - Professional API service
- **Scalable** - Handles thousands of messages easily
- **No Browser Needed** - Pure API calls
- **Cons**: **Paid service** (~$0.005-0.01 per message)
- **Cost for 2000 messages/week**: ~$40-80/month

**Recommendation**: Use **Selenium** (current solution) for your needs. It's free, reliable enough for 700-2000 weekly messages, and gives you full control. Consider Twilio only if you need enterprise-level reliability and don't mind the cost.

## Important Notes

⚠️ **Limitations:**
- This uses WhatsApp Web automation, which has rate limits
- WhatsApp may temporarily restrict your account if you send too many messages too quickly
- The browser must stay open during the entire process
- Your computer must not go to sleep
- Chrome browser window will be visible (you can minimize it)

⚠️ **Best Practices:**
- Start with a small test batch (10-20 contacts)
- Use appropriate delays (15-30 seconds) between messages
- Don't send spam or unsolicited messages
- Respect WhatsApp's terms of service
- Keep your phone connected to internet (WhatsApp Web needs active connection)

## Troubleshooting

### "Python not found" or "python: command not found"
- Make sure Python is installed (see "Installing Python" section above)
- On Mac/Linux, try using `python3` instead of `python`
- On Windows, make sure Python was added to PATH during installation
- Restart your terminal/command prompt after installing Python

### "pip: command not found"
- pip should come with Python, but if missing:
  - Windows: `python -m ensurepip --upgrade`
  - Mac/Linux: `python3 -m ensurepip --upgrade`
- Or install pip separately:
  - Windows: Download get-pip.py and run `python get-pip.py`
  - Mac/Linux: `sudo apt install python3-pip` (Ubuntu) or `brew install python3` (Mac)

### "Script not on PATH" Warning (Windows)
If you see: `WARNING: The script dotenv.exe is installed in '...' which is not on PATH`

**This is safe to ignore** - packages still work fine! But if you want to fix it:

**Option 1: Suppress the warning**
```powershell
pip install --no-warn-script-location -r requirements.txt
```

**Option 2: Add to PATH** (see `POWERSHELL_GUIDE.md` for detailed instructions)

**Option 3: Just ignore it** - Your application will work normally

### "Excel file not found"
- Make sure `contacts.xlsx` is in the same folder as the script
- Check the file name matches `EXCEL_FILE` in the script

### "Message not sent"
- Check if the contact number format is correct (must include country code)
- Ensure WhatsApp Web is logged in
- Check your internet connection

### Rate Limiting
- Increase `DELAY_SECONDS` to 30 or more
- Split your contacts into smaller batches
- Wait a few hours before resuming

### Browser Issues

**"[WinError 193]" or ChromeDriver errors:**
- Clear ChromeDriver cache:
  ```powershell
  Remove-Item -Recurse -Force $env:USERPROFILE\.wdm
  ```
- Or run: `.\scripts\fix_chromedriver.ps1`
- Make sure Chrome browser is installed and up to date
- Try running the script again

### Image Sending Issues

**Image not found:**
- Check image path in Excel Column C is correct
- Verify image file exists at the specified location
- Check file permissions
- Ensure image format is supported (.jpg, .png, etc.)

**Image not sending, only text:**
- Check console output for error messages
- Verify chat is open before sending
- Make sure attachment button is visible
- Try with a smaller image file (< 5MB)

**All contacts getting same image:**
- If using individual images, make sure Column C has different paths
- Or name images with contact numbers (e.g., `919555611880.jpg`)
- Check `DEFAULT_IMAGE` is set to `None` for individual image mode

**Other browser issues:**
- Make sure Chrome browser is installed
- Close other Chrome windows/tabs if needed
- If QR code doesn't appear, close and restart the script
- Chrome profile is saved in `./chrome_profile` folder (you may need to delete it if login issues persist)

## File Structure

```
message-sender/
├── whatsapp_sender.py          # Main application
├── requirements.txt            # Python dependencies
├── README.md                   # Main documentation (this file)
├── .gitignore                  # Git ignore rules
│
├── utils/                      # Utility scripts
│   ├── __init__.py
│   ├── check_code.py           # Setup validator
│   └── create_template.py      # Excel template generator
│
├── scripts/                    # Helper scripts
│   └── fix_chromedriver.ps1   # ChromeDriver troubleshooting
│
├── templates/                  # Template files
│   └── contacts_template.xlsx  # Sample Excel template
│
└── docs/                       # All documentation files
    ├── IMAGE_GUIDE.md          # Image sending guide
    ├── SINGLE_IMAGE_GUIDE.md   # Single image mode guide
    ├── IMAGE_FEATURE_SUMMARY.md # Image feature quick reference
    ├── POWERSHELL_GUIDE.md     # PowerShell command reference
    ├── GIT_SETUP.md            # Git setup instructions
    ├── PROJECT_STRUCTURE.md    # Project structure documentation
    └── CHANGES_SUMMARY.md      # Project reorganization summary
```

## Documentation

- **Main Guide**: This README.md
- **Run Steps**: `docs/RUN_STEPS.md` - **Complete step-by-step guide for checking and running the code**
- **Quick Start**: `docs/QUICK_START.md` - Fast setup reference
- **Image Guide**: `docs/IMAGE_GUIDE.md` - Complete guide for sending images with captions
- **Single Image Guide**: `docs/SINGLE_IMAGE_GUIDE.md` - How to send one image to all contacts
- **PowerShell Guide**: `docs/POWERSHELL_GUIDE.md` - Windows PowerShell commands
- **Git Setup**: `docs/GIT_SETUP.md` - Setting up Git repository

## Example Use Cases

### Campaign with Single Image
Perfect for promotional campaigns where you want to send the same image (like a Safari bag promotion) to all contacts with personalized captions:

```python
DEFAULT_IMAGE = "safari_promo.jpg"
```

Excel file:
```
Contact Number    | Message (Caption)
+919555611880     | 👆🏻 आपका फोटो यहाँ आएगा 📸✨...
+919355611880     | 👆🏻 आपका फोटो यहाँ आएगा 📸✨...
```

### Personalized Images per Contact
Send unique images to each agent/contact:

```python
DEFAULT_IMAGE = None
IMAGES_FOLDER = "images"
```

Excel file:
```
Contact Number    | Message          | Image Path
+919555611880     | Your message...  | images/agent1.jpg
+919355611880     | Your message...  | images/agent2.jpg
```

### Text-Only Messages
Send text messages without images:

```python
DEFAULT_IMAGE = None
IMAGES_FOLDER = None
```

## License

This is a simple utility script. Use responsibly and in accordance with WhatsApp's terms of service.
