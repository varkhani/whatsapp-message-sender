# Single Image for All Contacts - Quick Guide

## 🎯 Use Case

When you want to send the **same image** to all contacts (like a campaign poster, promotional image, etc.) with personalized captions.

## ⚙️ Configuration

### Step 1: Place Your Image

Put your image file in the same folder as `contacts.xlsx`:

```
message-sender/
├── contacts.xlsx
├── promo.jpg          ← Your single image
└── whatsapp_sender.py
```

Or in a subfolder:

```
message-sender/
├── contacts.xlsx
├── images/
│   └── campaign.jpg   ← Your single image
└── whatsapp_sender.py
```

### Step 2: Update Configuration

Open `whatsapp_sender.py` and find this section:

```python
# SINGLE IMAGE FOR ALL CONTACTS
DEFAULT_IMAGE = None  # Change this!
```

**Set the image path:**

```python
# Option 1: Image in same folder
DEFAULT_IMAGE = "promo.jpg"

# Option 2: Image in images folder
DEFAULT_IMAGE = "images/campaign.jpg"

# Option 3: Absolute path
DEFAULT_IMAGE = "C:/Users/YourName/Pictures/promo.jpg"

# Option 4: Disable (use individual images)
DEFAULT_IMAGE = None
```

### Step 3: Prepare Excel File

Your Excel file only needs 2 columns (Column C is ignored when using DEFAULT_IMAGE):

| Column A | Column B |
|----------|----------|
| Contact Number | Message (Caption) |

**Example:**
```
Contact Number    | Message (Caption)
+919555611880     | 👆🏻 आपका फोटो यहाँ आएगा 📸✨...
+919355611880     | 👆🏻 आपका फोटो यहाँ आएगा 📸✨...
```

### Step 4: Run

```powershell
python whatsapp_sender.py
```

## 📋 Complete Example

### Folder Structure:
```
message-sender/
├── contacts.xlsx
├── safari_promo.jpg      ← Single image for all
└── whatsapp_sender.py
```

### Configuration in `whatsapp_sender.py`:
```python
DEFAULT_IMAGE = "safari_promo.jpg"  # Same image for everyone
IMAGES_FOLDER = None  # Not needed when using DEFAULT_IMAGE
```

### Excel File (`contacts.xlsx`):
```
Contact Number    | Message
+919555611880     | 👆🏻 आपका फोटो यहाँ आएगा 📸✨
                  | 🎒 Safari बैग के साथ
                  | 🌴✈️ चलो Goa की ओर 🏖️😎
+919355611880     | 👆🏻 आपका फोटो यहाँ आएगा 📸✨
                  | 🎒 Safari बैग के साथ
                  | 🌴✈️ चलो Goa की ओर 🏖️😎
```

## ✅ Benefits

- ✅ **Simple**: One image, one configuration
- ✅ **Fast**: No need to find images per contact
- ✅ **Consistent**: Same image for all contacts
- ✅ **Flexible**: Each contact can have different caption

## 🔄 Switching Between Modes

### Single Image Mode:
```python
DEFAULT_IMAGE = "promo.jpg"
IMAGES_FOLDER = None
```

### Individual Images Mode:
```python
DEFAULT_IMAGE = None
IMAGES_FOLDER = "images"
```

## 💡 Tips

1. **Image Name**: Use descriptive names like `safari_campaign.jpg`, `goa_promo.jpg`
2. **Image Size**: Keep under 5MB for faster upload
3. **Testing**: Test with 1-2 contacts first
4. **Caption**: Each contact can have a unique caption in Column B

## ❓ FAQ

**Q: Can I still use Column C when DEFAULT_IMAGE is set?**  
A: No, DEFAULT_IMAGE overrides Column C. Use one or the other.

**Q: What if DEFAULT_IMAGE file doesn't exist?**  
A: Script will show a warning and send text-only messages.

**Q: Can I use DEFAULT_IMAGE and IMAGES_FOLDER together?**  
A: No, DEFAULT_IMAGE takes priority. Set IMAGES_FOLDER to None when using DEFAULT_IMAGE.

**Q: How do I switch back to individual images?**  
A: Set `DEFAULT_IMAGE = None` and configure `IMAGES_FOLDER`.
