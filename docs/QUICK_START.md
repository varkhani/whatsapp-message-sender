# Quick Start Guide

## 🚀 Fast Setup (5 Minutes)

### Step 1: Install Dependencies
```powershell
cd C:\SalesIntel\repo\my_project\message-sender
pip install -r requirements.txt
```

### Step 2: Check Setup
```powershell
python utils/check_code.py
```
✅ Should show "All checks passed!"

### Step 3: Create Excel File
```powershell
python utils/create_template.py
```
✅ Creates `contacts.xlsx`

### Step 4: Edit Excel File
- Open `contacts.xlsx`
- Add your contact numbers (Column A)
- Add your messages (Column B)
- Save the file

### Step 5: Configure (Optional)
Edit `whatsapp_sender.py`:
```python
DEFAULT_IMAGE = None      # Text-only mode
IMAGES_FOLDER = None      # No images
```

### Step 6: Run
```powershell
python whatsapp_sender.py
```
- Press Enter when prompted
- Scan QR code with your phone
- Wait for messages to send

---

## 📋 Complete Checklist

Before running:
- [ ] Python installed (`python --version`)
- [ ] Dependencies installed (`pip install -r requirements.txt`)
- [ ] Code validated (`python utils/check_code.py`)
- [ ] Excel file created (`python utils/create_template.py`)
- [ ] Contacts added to Excel
- [ ] Chrome browser installed
- [ ] Configuration checked in `whatsapp_sender.py`

---

## 🔧 Common Commands

```powershell
# Check Python
python --version

# Install packages
pip install -r requirements.txt

# Validate setup
python utils/check_code.py

# Create template
python utils/create_template.py

# Run application
python whatsapp_sender.py

# Fix ChromeDriver issues
.\scripts\fix_chromedriver.ps1
```

---

## ⚙️ Configuration Modes

**Text-Only:**
```python
DEFAULT_IMAGE = None
IMAGES_FOLDER = None
```

**Single Image:**
```python
DEFAULT_IMAGE = "promo.jpg"
IMAGES_FOLDER = None
```

**Individual Images:**
```python
DEFAULT_IMAGE = None
IMAGES_FOLDER = "images"
```

---

## 📖 Full Documentation

- **Complete Guide**: `docs/RUN_STEPS.md`
- **Image Setup**: `docs/IMAGE_GUIDE.md`
- **Single Image**: `docs/SINGLE_IMAGE_GUIDE.md`
