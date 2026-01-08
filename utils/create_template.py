"""
Script to create a sample Excel template for contacts
"""

import openpyxl
import sys
import io

# Fix Windows console encoding for Unicode characters
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def create_template():
    """Create a sample Excel template file with image support"""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    # Add headers
    sheet['A1'] = 'Contact Number'
    sheet['B1'] = 'Contact Name'
    sheet['C1'] = 'Message (Caption)'
    sheet['D1'] = 'Image Path (Optional)'
    
    # Add sample data with Hindi caption example
    # Row 2: With contact name
    sheet['A2'] = '+919555611880'
    sheet['B2'] = 'Jinu'  # Contact Name
    sheet['C2'] = '👆🏻 आपका फोटो यहाँ आएगा 📸✨\n\n🎒 Safari बैग के साथ\n🌴✈️ चलो Goa की ओर 🏖️😎\n\nस्मार्ट तरीके से बिक्री करें। तेज़ी से आगे बढ़ें। ⚡📊'
    sheet['D2'] = 'images/agent1.jpg'  # Optional: specific image for this contact
    
    # Row 3: Without contact name (empty B column)
    # sheet['A3'] = '+919355611880'
    # sheet['B3'] = ''  # Empty contact name - message will be sent as-is
    # sheet['C3'] = '👆🏻 आपका फोटो यहाँ आएगा 📸✨\n\n🎒 Safari बैग के साथ\n🌴✈️ चलो Goa की ओर 🏖️😎\n\nस्मार्ट तरीके से बिक्री करें। तेज़ी से आगे बढ़ें। ⚡📊'
    # sheet['D3'] = ''  # Leave empty to auto-detect image
    
    # Save the file as contacts.xlsx (the file used by the main application)
    workbook.save('contacts.xlsx')
    print("✓ File 'contacts.xlsx' created successfully!")
    print("\nFormat:")
    print("  Column A: Contact Number (with country code, e.g., +1234567890)")
    print("  Column B: Contact Name (optional - if provided, message will be: 'Dear [Name],\\n\\n[Message]')")
    print("  Column C: Message (Caption) - text to send with image")
    print("  Column D: Image Path (optional - leave empty to auto-detect)")
    print("\nMessage Format:")
    print("  - If Contact Name (B) is provided: 'Dear [Name],\\n\\n[Message from C]'")
    print("  - If Contact Name (B) is empty: '[Message from C]' (sent as-is)")
    print("\nImage Support:")
    print("  Option 1: Specify image path in Column D (e.g., 'images/agent1.jpg')")
    print("  Option 2: Leave Column D empty - script will auto-detect:")
    print("    - Looks for image named like contact number (e.g., 919555611880.jpg)")
    print("    - Looks in 'images' folder (if configured)")
    print("    - Falls back to any image in same folder")
    print("\nExample folder structure:")
    print("  contacts.xlsx")
    print("  images/")
    print("    ├── 919555611880.jpg  <- Auto-detected for contact +919555611880")
    print("    ├── 919355611880.jpg  <- Auto-detected for contact +919355611880")
    print("    └── agent1.jpg        <- Use in Column D: 'images/agent1.jpg'")

if __name__ == "__main__":
    create_template()
