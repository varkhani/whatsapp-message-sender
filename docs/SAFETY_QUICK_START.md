# 🚀 Quick Start Guide

## For WhatsApp Business Users

### ✅ Default Settings (Already Optimized!)

The script comes pre-configured for **WhatsApp Business** accounts:

```
✓ 200 messages/day
✓ 40 messages/hour
✓ 15-45 second delays
✓ Auto-breaks every 50 messages
✓ Active hours: 9 AM - 9 PM
```

### 🎯 Just Run It!

```bash
python whatsapp_sender.py
```

That's it! The safety system handles everything automatically.

---

## 📊 Check Your Progress Anytime

```bash
python whatsapp_sender.py stats
```

Shows:
- Messages sent today: X/200
- Messages this hour: X/40
- Last contact: #X
- Last sent: timestamp

---

## ⚙️ Want to Send More Messages?

Edit `whatsapp_sender.py` (top of file):

```python
# For verified/established Business accounts:
MAX_MESSAGES_PER_DAY = 500     # Increase from 200
MAX_MESSAGES_PER_HOUR = 80      # Increase from 40
MIN_DELAY_BETWEEN_MESSAGES = 10 # Decrease from 15
```

---

## 🔄 Reset Progress (If Needed)

```bash
python whatsapp_sender.py reset
```

Type `YES` to confirm. This resets:
- Today's count
- Hour's count
- Last contact index

---

## ⚠️ Using Normal WhatsApp (Not Business)?

**IMPORTANT**: Normal WhatsApp has MUCH stricter limits!

Edit these settings:

```python
MAX_MESSAGES_PER_DAY = 50       # Lower from 200
MAX_MESSAGES_PER_HOUR = 15      # Lower from 40
MIN_DELAY_BETWEEN_MESSAGES = 30 # Increase from 15
MAX_DELAY_BETWEEN_MESSAGES = 90 # Increase from 45
MESSAGES_BEFORE_BREAK = 25      # Lower from 50
```

---

## 📱 What You'll See

### Starting:
```
📱 Starting to send messages to 170 contacts...
🔒 Safety features enabled:
   • Daily limit: 200 messages
   • Hourly limit: 40 messages
   • Smart delays: 15-45s
   • Auto-breaks: Every 50 messages
   • Active hours: 9:00 - 21:00

📊 Current Status:
   • Today: 0/200
   • This hour: 0/40
```

### During breaks:
```
============================================================
⏸️  TAKING A BREAK
============================================================
📊 Progress so far:
   • Messages sent: 50
   • Today: 50/200
   • This hour: 40/40

⏳ Break duration: 7.3 minutes
   Resume time: 02:15 PM
============================================================
```

### Progress updates (every 10 messages):
```
📊 Progress: 30/170 | ✓ 28 | ✗ 2
   Today: 30/200 | Hour: 15/40
```

---

## 🎯 For 170 Contacts

**Expected Timeline** (default Business settings):

- Hour 1: 40 messages + break
- Hour 2: 40 messages + break  
- Hour 3: 40 messages + break
- Hour 4: 40 messages + break
- Hour 5: 10 messages
- **Total**: ~5 hours

**Why so long?**
- Safety delays (15-45s each)
- Auto-breaks (5-10 min each)
- This protects your account!

---

## 💡 Pro Tips

1. **Start in the morning** (9 AM) to finish in one day
2. **Let it run** - don't interrupt unnecessarily
3. **Check stats** during the day
4. **Don't disable** safety features (ban risk!)
5. **Business account** = much safer than normal WhatsApp

---

## 🆘 Common Issues

### "Daily limit reached"
→ Wait until tomorrow or increase limit

### "Outside active hours"  
→ Wait until 9 AM or adjust hours

### Script interrupted
→ Just run again, it resumes automatically!

### Want faster sending
→ Increase limits (but be careful!)

---

## 📞 Commands Cheat Sheet

```bash
# Normal sending
python whatsapp_sender.py

# Check stats
python whatsapp_sender.py stats

# Reset progress
python whatsapp_sender.py reset
```

---

**That's it! The safety system handles everything else automatically.** 🎉
