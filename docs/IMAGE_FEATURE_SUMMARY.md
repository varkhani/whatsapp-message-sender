# Image with Caption Feature - Summary

## ✅ What's New

The WhatsApp sender now supports sending **unique images with personalized captions** to each contact!

## 📋 Excel File Format

Your `contacts.xlsx` now supports 3 columns:

| Column | Description | Required |
|--------|-------------|----------|
| A | Contact Number | ✅ Yes |
| B | Message/Caption | ✅ Yes |
| C | Image Path | ❌ Optional |

### Example:
```
Contact Number    | Message (Caption)                    | Image Path
+919555611880     | 👆🏻 आपका फोटो यहाँ आएगा...        | images/agent1.jpg
+919355611880     | 👆🏻 आपका फोटो यहाँ आएगा...        | 
```

## 🖼️ How It Works

### Option 1: Specify Image in Excel (Column C)
- Put the image path in Column C
- Example: `images/agent1.jpg` or `agent1.jpg`
- Each contact can have a unique image

### Option 2: Auto-Detection (Leave Column C Empty)
- Script automatically finds images based on contact number
- Looks for: `{contact_number}.jpg` (e.g., `919555611880.jpg`)
- Searches in `images/` folder first, then Excel folder

## 📁 Folder Structure

```
message-sender/
├── contacts.xlsx
├── images/                    # Create this folder
│   ├── 919555611880.jpg      # Auto-detected for +919555611880
│   ├── agent1.jpg            # Use in Excel: images/agent1.jpg
│   └── agent2.jpg
└── whatsapp_sender.py
```

## ⚙️ Configuration

In `whatsapp_sender.py`:

```python
IMAGES_FOLDER = "images"  # Folder containing images (or None to disable)
```

## 🚀 Quick Start

1. **Create images folder:**
   ```powershell
   mkdir images
   ```

2. **Add your images:**
   - Name them with contact numbers: `919555611880.jpg`
   - Or use descriptive names: `agent1.jpg`, `agent2.jpg`

3. **Update Excel file:**
   - Column A: Contact Number
   - Column B: Your caption (can include emojis and Hindi text)
   - Column C: Image path (optional - leave empty for auto-detect)

4. **Run the script:**
   ```powershell
   python whatsapp_sender.py
   ```

## 📝 Sample Caption

```
👆🏻 आपका फोटो यहाँ आएगा 📸✨

🎒 Safari बैग के साथ
🌴✈️ चलो Goa की ओर 🏖️😎

स्मार्ट तरीके से बिक्री करें। तेज़ी से आगे बढ़ें। ⚡📊

Safari बैग जीतें — और चलो Goa की ओर 🌴✈️

🎒 सिर्फ़ 2 पॉलिसी 👉 Safari बैग अनलॉक 🔓✨

📈 ₹10 लाख प्रीमियम 👉 Goa के लिए क्वालिफ़ाई 🏖️🏆
```

## 📚 Documentation

- **Full Guide**: See `docs/IMAGE_GUIDE.md`
- **Template Creator**: Run `python utils/create_template.py`

## ✨ Features

- ✅ Unique image per contact
- ✅ Personalized captions from Excel
- ✅ Auto-detection of images
- ✅ Fallback to text-only if image not found
- ✅ Supports Hindi/English text with emojis
- ✅ Multiple image formats (.jpg, .png, .gif, .webp)

## 🔧 Technical Details

- Image path can be relative or absolute
- If image not found, sends text message as fallback
- Images are uploaded via WhatsApp Web file input
- Caption is typed in the caption box (data-tab='11')
- Supports all standard image formats

## 💡 Tips

1. **Test first**: Send to 1-2 contacts before bulk sending
2. **Image size**: Keep images under 5MB for faster upload
3. **Naming**: Use contact numbers in filenames for easy auto-detection
4. **Backup**: Always backup your images folder
5. **Format**: Square images (1080x1080) work best for WhatsApp
