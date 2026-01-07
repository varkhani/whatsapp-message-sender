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
    """Create a sample Excel template file"""
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    # Add headers
    sheet['A1'] = 'Contact Number'
    sheet['B1'] = 'Message'
    
    # Add sample data
    sheet['A2'] = '+919555611880'
    sheet['B2'] = 'Hello! This is a test message.'
    
    sheet['A3'] = '+919355611880'
    sheet['B3'] = 'Hi there! This is another test message.'
    
    # Save the file as contacts.xlsx (the file used by the main application)
    workbook.save('contacts.xlsx')
    print("✓ File 'contacts.xlsx' created successfully!")
    print("\nFormat:")
    print("  Column A: Contact Number (with country code, e.g., +1234567890)")
    print("  Column B: Message (used as caption if image is found)")
    print("\nImage Support:")
    print("  - Place an image file (jpg, png, gif, webp) in the same folder as contacts.xlsx")
    print("  - The image will be sent to all contacts with the message as caption")
    print("  - If no image is found, only text messages will be sent")
    print("\nExample:")
    print("  Folder structure:")
    print("    contacts.xlsx")
    print("    promo.jpg  <- This image will be sent to all contacts")

if __name__ == "__main__":
    create_template()
