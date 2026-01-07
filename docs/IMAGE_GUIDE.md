# Image with Caption Guide

This guide explains how to send personalized images with captions to WhatsApp contacts.

## Excel File Format

Your `contacts.xlsx` file should have 3 columns:

| Column A | Column B | Column C |
|----------|----------|----------|
| Contact Number | Message (Caption) | Image Path (Optional) |

### Example:

```
Contact Number          | Message (Caption)                                    | Image Path
+919555611880          | 👆🏻 आपका फोटो यहाँ आएगा 📸✨...                    | images/agent1.jpg
+919355611880          | 👆🏻 आपका फोटो यहाँ आएगा 📸✨...                    | 
```

## How Image Detection Works

The script follows this priority order to find images:

### Priority 1: Column C (Excel File)
If you specify an image path in Column C, that exact image will be used.

**Example:**
- `images/agent1.jpg` - Uses image from images folder
- `agent1.jpg` - Uses image from same folder as Excel file
- `C:\Users\YourName\Pictures\agent1.jpg` - Uses absolute path

### Priority 2: Auto-Detection (if Column C is empty)
If Column C is empty, the script will automatically look for:

1. **Contact-specific image**: `{contact_number}.jpg` (e.g., `919555611880.jpg`)
2. **In images folder**: Any image file in the `images/` folder
3. **In Excel folder**: Any image file in the same folder as `contacts.xlsx`

## Setup Instructions

### Step 1: Prepare Your Images

Create an `images` folder in the same directory as your `contacts.xlsx`:

```
message-sender/
├── contacts.xlsx
├── images/
│   ├── 919555611880.jpg    # Image for contact +919555611880
│   ├── 919355611880.jpg    # Image for contact +919355611880
│   └── agent1.jpg          # Generic image (use in Column C)
└── whatsapp_sender.py
```

### Step 2: Update Excel File

**Option A: Specify image path in Column C**
```
Contact Number    | Message                    | Image Path
+919555611880     | Your caption here...       | images/agent1.jpg
```

**Option B: Leave Column C empty (auto-detect)**
```
Contact Number    | Message                    | Image Path
+919555611880     | Your caption here...       | 
```
The script will look for `919555611880.jpg` in the images folder.

### Step 3: Configure Script

In `whatsapp_sender.py`, set the images folder:

```python
IMAGES_FOLDER = "images"  # Folder containing images
```

Or set to `None` to disable auto-detection:
```python
IMAGES_FOLDER = None  # Only use images from Column C
```

### Step 4: Run the Script

```powershell
python whatsapp_sender.py
```

## Example: Safari Bag Campaign

For your Safari bag campaign with personalized images:

### Excel Format:
```
Contact Number    | Message (Caption)                                                                  | Image Path
+919555611880     | 👆🏻 आपका फोटो यहाँ आएगा 📸✨\n\n🎒 Safari बैग के साथ...                        | images/agent1.jpg
+919355611880     | 👆🏻 आपका फोटो यहाँ आएगा 📸✨\n\n🎒 Safari बैग के साथ...                        | images/agent2.jpg
```

### Folder Structure:
```
message-sender/
├── contacts.xlsx
├── images/
│   ├── agent1.jpg    # Personalized image for agent 1
│   ├── agent2.jpg    # Personalized image for agent 2
│   └── ...
└── whatsapp_sender.py
```

## Image Requirements

- **Supported formats**: `.jpg`, `.jpeg`, `.png`, `.gif`, `.webp`
- **Recommended size**: Under 5MB for faster upload
- **Recommended dimensions**: 1080x1080 or similar (square works best for WhatsApp)

## Troubleshooting

### Image not found
- Check that the image path in Column C is correct
- Verify the image file exists
- Check file permissions
- Ensure image format is supported (.jpg, .png, etc.)

### Image not sending, only text
- Check console output for error messages
- Verify chat is open before sending
- Make sure attachment button is visible
- Try with a smaller image file

### All contacts getting same image
- Make sure Column C has different image paths for each contact
- Or name images with contact numbers (e.g., `919555611880.jpg`)

## Tips

1. **Unique images per contact**: Use Column C to specify exact image path
2. **Batch processing**: Name images with contact numbers for auto-detection
3. **Testing**: Test with 1-2 contacts first before bulk sending
4. **Backup**: Keep backup of your images folder
5. **Naming**: Use descriptive names like `agent1.jpg`, `agent2.jpg` for easy management

## Sample Caption Template

```
👆🏻 आपका फोटो यहाँ आएगा 📸✨

🎒 Safari बैग के साथ
🌴✈️ चलो Goa की ओर 🏖️😎

स्मार्ट तरीके से बिक्री करें। तेज़ी से आगे बढ़ें। ⚡📊

Safari बैग जीतें — और चलो Goa की ओर 🌴✈️

🎒 सिर्फ़ 2 पॉलिसी 👉 Safari बैग अनलॉक 🔓✨

📈 ₹10 लाख प्रीमियम 👉 Goa के लिए क्वालिफ़ाई 🏖️🏆
```

Copy this into Column B of your Excel file and customize as needed!
