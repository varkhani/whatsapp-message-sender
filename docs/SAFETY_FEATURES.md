# 🔒 WhatsApp Business Safety Features

## Overview

This WhatsApp message sender includes comprehensive safety features specifically optimized for **WhatsApp Business** accounts to prevent bans and ensure reliable bulk messaging.

---

## 🎯 Safety Limits (WhatsApp Business Optimized)

### Daily & Hourly Limits
- **Daily limit**: 200 messages/day (adjustable)
- **Hourly limit**: 40 messages/hour (adjustable)
- **Auto-reset**: Limits reset automatically each day/hour

### Smart Delays
- **Random delays**: 15-45 seconds between messages
- **Intelligent spacing**: Delays increase near hourly limits
- **Human-like behavior**: Mimics natural messaging patterns

### Auto-Break System
- **Break frequency**: Every 50 messages
- **Break duration**: 5-10 minutes (randomized)
- **Progress display**: Shows stats during breaks

### Time Restrictions
- **Active hours**: 9 AM - 9 PM only
- **Auto-pause**: Automatically waits until active hours
- **Resume capability**: Continues where it left off

### Progress Tracking
- **Auto-save**: Progress saved after each message
- **Resume support**: Continue from where you stopped
- **Statistics**: Track daily/hourly message counts

---

## 📊 Usage

### Normal Usage (with Safety Features)
```bash
python whatsapp_sender.py
```

The script will:
1. Show current safety status
2. Display today's and this hour's message counts
3. Automatically enforce all safety limits
4. Take breaks when needed
5. Save progress continuously

### Check Statistics
```bash
python whatsapp_sender.py stats
```

Shows:
- Messages sent today
- Messages sent this hour
- Total messages sent
- Last contact index
- Last message timestamp

### Reset Progress (Use with Caution!)
```bash
python whatsapp_sender.py reset
```

Resets:
- Today's message count
- Hourly message count
- Last contact index
- All statistics

**⚠️ Warning**: You'll need to type `YES` to confirm.

---

## 🔧 Customizing Safety Settings

Edit the configuration constants at the top of `whatsapp_sender.py`:

```python
# Increase daily limit (for verified Business accounts)
MAX_MESSAGES_PER_DAY = 500  # Default: 200

# Adjust hourly limit
MAX_MESSAGES_PER_HOUR = 60  # Default: 40

# Shorter delays (if your account is well-established)
MIN_DELAY_BETWEEN_MESSAGES = 10  # Default: 15
MAX_DELAY_BETWEEN_MESSAGES = 30  # Default: 45

# More messages before break
MESSAGES_BEFORE_BREAK = 100  # Default: 50

# Shorter breaks (not recommended)
BREAK_DURATION_MIN = 180  # Default: 300 (5 min)
BREAK_DURATION_MAX = 300  # Default: 600 (10 min)

# Extended active hours
ACTIVE_HOURS_START = 8   # Default: 9
ACTIVE_HOURS_END = 22    # Default: 21
```

---

## 📈 Recommended Settings by Account Type

### 1. New WhatsApp Business Account (< 3 months)
```python
MAX_MESSAGES_PER_DAY = 100
MAX_MESSAGES_PER_HOUR = 20
MIN_DELAY_BETWEEN_MESSAGES = 20
MAX_DELAY_BETWEEN_MESSAGES = 60
MESSAGES_BEFORE_BREAK = 25
```

### 2. Established WhatsApp Business (3-12 months)
```python
MAX_MESSAGES_PER_DAY = 200  # Default
MAX_MESSAGES_PER_HOUR = 40   # Default
MIN_DELAY_BETWEEN_MESSAGES = 15
MAX_DELAY_BETWEEN_MESSAGES = 45
MESSAGES_BEFORE_BREAK = 50
```

### 3. Verified WhatsApp Business (12+ months)
```python
MAX_MESSAGES_PER_DAY = 500
MAX_MESSAGES_PER_HOUR = 80
MIN_DELAY_BETWEEN_MESSAGES = 10
MAX_DELAY_BETWEEN_MESSAGES = 30
MESSAGES_BEFORE_BREAK = 100
```

### 4. Normal WhatsApp (Not Business) - CONSERVATIVE
```python
MAX_MESSAGES_PER_DAY = 50    # ⚠️ Much lower!
MAX_MESSAGES_PER_HOUR = 15
MIN_DELAY_BETWEEN_MESSAGES = 30
MAX_DELAY_BETWEEN_MESSAGES = 90
MESSAGES_BEFORE_BREAK = 25
```

---

## ⚠️ Important Notes

### WhatsApp Business vs Normal WhatsApp

**WhatsApp Business** accounts have:
- ✅ Higher message limits
- ✅ Better bulk messaging tolerance
- ✅ Business verification options
- ✅ Official API support

**Normal WhatsApp** accounts:
- ❌ Stricter limits (50 messages/day recommended)
- ❌ Higher ban risk
- ❌ More aggressive spam detection
- ⚠️ Use conservative settings!

### Ban Prevention Tips

1. **Start Slow**: Begin with lower limits, increase gradually
2. **Use Business Account**: Much safer for bulk messaging
3. **Quality Messages**: Personalized messages are safer
4. **Avoid Spam**: Don't send identical messages repeatedly
5. **Monitor Responses**: Stop if you get warnings
6. **Respect Breaks**: Don't skip auto-breaks
7. **Active Hours**: Only send during business hours
8. **Verify Account**: Get WhatsApp Business verified

### What Happens When Limits Are Reached?

- **Hourly limit**: Script waits until next hour automatically
- **Daily limit**: Script stops and will resume tomorrow
- **Outside active hours**: Script waits until active hours
- **Break time**: Script pauses for 5-10 minutes

### Progress File

The script creates `whatsapp_progress.json` to track:
- Messages sent today
- Messages sent this hour
- Last contact index
- Timestamps

**⚠️ Don't delete this file** unless you want to reset progress!

---

## 🚀 Example Workflow

### Sending 170 Messages Safely

**Scenario**: You have 170 contacts to message

**Timeline** (with default Business settings):

```
Hour 1 (9:00 AM):  40 messages → Break (5-10 min)
Hour 2 (10:00 AM): 40 messages → Break
Hour 3 (11:00 AM): 40 messages → Break
Hour 4 (12:00 PM): 40 messages → Break
Hour 5 (1:00 PM):  10 messages → Complete!

Total time: ~5 hours (including breaks and delays)
```

**Average delay per message**: ~1 minute (including smart delays and breaks)

### Interruption & Resume

If interrupted at message #75:

```bash
# Next time you run:
python whatsapp_sender.py

# Output:
📊 Current Status:
   • Today: 75/200
   • This hour: 15/40
   • Last contact: #75

# It will automatically continue from #76!
```

---

## 🛠️ Troubleshooting

### "Daily limit reached"
- Wait until tomorrow, or
- Increase `MAX_MESSAGES_PER_DAY` (risky!)
- Run `python whatsapp_sender.py reset` to reset (⚠️ use carefully!)

### "Outside active hours"
- Wait until active hours (default: 9 AM - 9 PM)
- Adjust `ACTIVE_HOURS_START` and `ACTIVE_HOURS_END`

### "Hourly limit reached"
- Script will auto-wait for next hour
- Or increase `MAX_MESSAGES_PER_HOUR`

### Progress not saving
- Check file permissions
- Ensure `whatsapp_progress.json` is not locked
- Check disk space

---

## 📞 Support

For issues or questions:
1. Check this documentation
2. Review error messages in console
3. Run `python whatsapp_sender.py stats` to check status
4. Adjust safety settings if needed

---

## ⚡ Quick Command Reference

```bash
# Normal sending with safety
python whatsapp_sender.py

# Check statistics
python whatsapp_sender.py stats

# Reset progress (requires confirmation)
python whatsapp_sender.py reset
```

---

**Remember**: These safety features protect YOUR account. Don't disable them unless you know what you're doing!
