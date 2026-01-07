# Windows PowerShell Quick Guide

All commands for running this project in **Windows PowerShell**.

## 🚀 Quick Start (PowerShell)

### 1. Check Python Installation

```powershell
python --version
```

If you see an error, install Python from [python.org](https://www.python.org/downloads/)

### 2. Navigate to Project Folder

```powershell
cd C:\SalesIntel\repo\my_project\message-sender
```

Or if you're already in the project folder:
```powershell
pwd
```

### 3. Install Dependencies

```powershell
pip install -r requirements.txt
```

**If you see a PATH warning** (like "script not on PATH"):
- You can safely ignore it - packages still work
- Or suppress it: `pip install --no-warn-script-location -r requirements.txt`

**If you get permission errors:**
```powershell
pip install --user -r requirements.txt
```

### 4. Check Your Setup

```powershell
python check_code.py
```

### 5. Create Sample Excel Template (Optional)

```powershell
python create_template.py
```

### 6. Run the Application

```powershell
python whatsapp_sender.py
```

---

## 📋 All PowerShell Commands Reference

### Installation Commands

```powershell
# Check Python version
python --version

# Install all dependencies
pip install -r requirements.txt

# Install specific package if needed
pip install selenium
pip install openpyxl
pip install webdriver-manager
pip install python-dotenv

# Upgrade pip (if needed)
python -m pip install --upgrade pip
```

### Setup & Validation Commands

```powershell
# Check code and setup
python check_code.py

# Create Excel template
python create_template.py

# Check Python syntax
python -m py_compile whatsapp_sender.py

# List installed packages
pip list

# Check if specific package is installed
pip show selenium
pip show openpyxl
```

### Running the Application

```powershell
# Run main application
python whatsapp_sender.py

# Run with specific file (if you modify the script)
python whatsapp_sender.py
```

### File Operations

```powershell
# List files in current directory
Get-ChildItem
# or
ls
# or
dir

# Check if file exists
Test-Path contacts.xlsx

# View file contents (text files)
Get-Content README.md

# Open current folder in File Explorer
explorer .
```

### Troubleshooting Commands

```powershell
# Check Python path
where.exe python

# Check pip path
where.exe pip

# Check if Chrome is installed
Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe"

# Check environment variables
$env:PATH

# Clear Python cache (if needed)
Get-ChildItem -Recurse -Filter __pycache__ | Remove-Item -Recurse -Force
```

---

## 🔧 Common PowerShell Issues & Solutions

### Issue: "python is not recognized"

**Solution:**
```powershell
# Check if Python is in PATH
$env:PATH -split ';' | Select-String python

# If not found, add Python to PATH manually or reinstall Python
# Make sure to check "Add Python to PATH" during installation
```

### Issue: "pip is not recognized"

**Solution:**
```powershell
# Try using python -m pip instead
python -m pip install -r requirements.txt
```

### Issue: Permission Denied

**Solution:**
```powershell
# Run PowerShell as Administrator
# Right-click PowerShell → Run as Administrator

# Or install packages for current user only
pip install --user -r requirements.txt
```

### Issue: ChromeDriver Error "[WinError 193]"

**Error message:**
```
✗ Error: [WinError 193] %1 is not a valid Win32 application
```

**Solution:**
```powershell
# Option 1: Clear ChromeDriver cache (Recommended)
Remove-Item -Recurse -Force $env:USERPROFILE\.wdm

# Or run the fix script:
.\fix_chromedriver.ps1

# Then try again:
python whatsapp_sender.py
```

**If that doesn't work:**
```powershell
# Make sure Chrome is up to date
# Check Chrome version: chrome://version/ in Chrome browser
# Update Chrome if needed
```

### Issue: "Script not on PATH" Warning

**Warning message:**
```
WARNING: The script dotenv.exe is installed in 'C:\Users\...\Scripts' which is not on PATH.
```

**This is usually safe to ignore**, but if you want to fix it:

**Option 1: Add to PATH (Recommended)**
```powershell
# Get your Python Scripts path
$scriptsPath = "$env:APPDATA\Python\Python310\Scripts"
# Or find it with:
python -m site --user-base
# Then add \Scripts to that path

# Add to PATH for current session
$env:PATH += ";$scriptsPath"

# Add to PATH permanently (run as Administrator)
[Environment]::SetEnvironmentVariable("Path", $env:Path + ";$scriptsPath", "User")
```

**Option 2: Suppress Warning (Easier)**
```powershell
# Install with --no-warn-script-location flag
pip install --no-warn-script-location -r requirements.txt
```

**Option 3: Ignore It (Simplest)**
- This warning is harmless - packages still work fine
- You can safely ignore it and continue using the application

### Issue: Script Execution Policy

**Solution:**
```powershell
# Check current policy
Get-ExecutionPolicy

# If Restricted, change it (run as Administrator)
Set-ExecutionPolicy RemoteSigned

# Or run Python directly (doesn't need policy change)
python whatsapp_sender.py
```

---

## 📝 Step-by-Step First Time Setup (PowerShell)

### Complete Setup Process:

```powershell
# Step 1: Open PowerShell
# Press Win + X, then select "Windows PowerShell" or "Terminal"

# Step 2: Navigate to project folder
cd C:\SalesIntel\repo\my_project\message-sender

# Step 3: Verify Python
python --version
# Should show: Python 3.x.x

# Step 4: Install dependencies
pip install -r requirements.txt

# Step 5: Verify installation
python check_code.py

# Step 6: Create sample template (optional)
python create_template.py

# Step 7: Edit contacts.xlsx with your data
# Open in Excel and add your contacts

# Step 8: Run the application
python whatsapp_sender.py
```

---

## 🎯 Quick Command Cheat Sheet

| Task | PowerShell Command |
|------|-------------------|
| Check Python | `python --version` |
| Install packages | `pip install -r requirements.txt` |
| Check setup | `python check_code.py` |
| Create template | `python create_template.py` |
| Run app | `python whatsapp_sender.py` |
| List files | `Get-ChildItem` or `ls` |
| Check file exists | `Test-Path contacts.xlsx` |
| View file | `Get-Content filename.txt` |
| Open folder | `explorer .` |

---

## 💡 PowerShell Tips

1. **Auto-complete**: Press `Tab` to auto-complete file/folder names
2. **History**: Press `↑` to see previous commands
3. **Clear screen**: Type `cls` or `Clear-Host`
4. **Copy text**: Select text and press `Enter` (copies to clipboard)
5. **Paste**: Right-click or `Shift + Insert`

---

## ⚠️ Important Notes for PowerShell

- Always use `python` (not `python3`) on Windows
- Use `pip` (not `pip3`) on Windows
- File paths can use backslashes: `C:\path\to\file`
- Or forward slashes: `C:/path/to/file` (both work)
- Use quotes for paths with spaces: `"C:\My Folder\file.xlsx"`

---

## 🆘 Need Help?

If you encounter errors:

1. **Run the checker:**
   ```powershell
   python check_code.py
   ```

2. **Check error messages** - they usually tell you what's wrong

3. **Verify Python:**
   ```powershell
   python --version
   where.exe python
   ```

4. **Reinstall packages:**
   ```powershell
   pip uninstall selenium openpyxl webdriver-manager python-dotenv
   pip install -r requirements.txt
   ```
