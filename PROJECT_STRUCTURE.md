# Project Structure

```
message-sender/
├── whatsapp_sender.py          # Main application entry point
├── requirements.txt             # Python dependencies
├── README.md                    # Main documentation
├── .gitignore                   # Git ignore rules
│
├── utils/                       # Utility scripts
│   ├── __init__.py
│   ├── check_code.py           # Setup validation script
│   └── create_template.py      # Excel template generator
│
├── scripts/                     # Helper scripts
│   └── fix_chromedriver.ps1    # ChromeDriver troubleshooting
│
├── templates/                   # Template files
│   └── contacts_template.xlsx  # Excel template for contacts
│
└── docs/                        # Documentation
    ├── POWERSHELL_GUIDE.md     # PowerShell usage guide
    └── GIT_SETUP.md            # Git setup instructions
```

## File Descriptions

### Root Files
- **whatsapp_sender.py**: Main script that sends WhatsApp messages
- **requirements.txt**: List of Python packages needed
- **README.md**: Project documentation and usage instructions
- **.gitignore**: Files/folders to exclude from Git

### Utils Directory
- **check_code.py**: Validates Python version, dependencies, and Excel format
- **create_template.py**: Creates a sample contacts.xlsx file

### Scripts Directory
- **fix_chromedriver.ps1**: PowerShell script to fix ChromeDriver issues

### Templates Directory
- **contacts_template.xlsx**: Template Excel file showing the required format

### Docs Directory
- **POWERSHELL_GUIDE.md**: Guide for using PowerShell commands
- **GIT_SETUP.md**: Instructions for setting up Git repository

## Excluded from Git

The following are automatically excluded (see `.gitignore`):
- `chrome_profile/` - Browser profile data
- `contacts.xlsx` - Your actual contact data (use template instead)
- `__pycache__/` - Python cache files
- `.env` - Environment variables
- Test images and media files

## Usage

1. **Main script**: `python whatsapp_sender.py`
2. **Check setup**: `python utils/check_code.py`
3. **Create template**: `python utils/create_template.py`
4. **Fix ChromeDriver**: `.\scripts\fix_chromedriver.ps1`
