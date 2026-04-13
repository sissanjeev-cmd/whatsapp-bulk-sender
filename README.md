# WhatsApp Bulk Message Sender

A Python script that reads contacts from an Excel file, personalizes a message template from a Word document, and sends WhatsApp messages via WhatsApp Web automation.

## Features

- 📋 Reads contacts (Name + Mobile) from an Excel file
- 📝 Reads message template from a Word (`.docx`) file
- 🔤 Personalizes each message with `{{Name}}` placeholder replacement
- 📱 Sends via WhatsApp Web (no API key or paid service required)
- 📊 Logs every send result to a `.txt` log file
- 🔒 Dry-run mode for safe testing before sending

## Requirements

- Python 3.7+
- WhatsApp account (personal or Business)

## Installation

```bash
pip install pandas openpyxl python-docx pywhatkit
```

## File Structure

```
whatsapp-bulk-sender/
├── send_whatsapp.py        # Main script
├── contacts.xlsx           # Excel file — columns: First Name, Mobile
└── message_template.docx  # Word file — contains {{Name}} placeholder
```

## Usage

### 1. Prepare your Excel contacts file

| First Name  | Mobile     |
|-------------|------------|
| Mr. Sharma  | 9876543210 |
| Ms. Verma   | 9123456789 |

### 2. Prepare your Word message template

```
Dear {{Name}},

Your personalized message goes here...

Warm Regards,
Your Name
```

### 3. Update the config at the top of `send_whatsapp.py`

```python
EXCEL_FILE   = "path/to/contacts.xlsx"
WORD_FILE    = "path/to/message_template.docx"
LOG_FILE     = "path/to/send_log.txt"
COUNTRY_CODE = "+91"   # Change for your country
WAIT_TIME    = 20      # Seconds per message (increase if internet is slow)
DRY_RUN      = False   # Set True to test without sending
```

### 4. Log in to WhatsApp Web

Open Chrome → go to [https://web.whatsapp.com](https://web.whatsapp.com) → scan QR code with your phone.

### 5. Run the script

```bash
python send_whatsapp.py
```

## Configuration Options

| Setting       | Default | Description                                      |
|---------------|---------|--------------------------------------------------|
| `COUNTRY_CODE`| `+91`   | Country dialing code (India). Change as needed.  |
| `WAIT_TIME`   | `20`    | Seconds to wait for WhatsApp Web to load per msg |
| `DRY_RUN`     | `False` | `True` = print messages only, don't send         |

## Output

- Console: real-time status of each message
- Log file: timestamped record of sent / failed / skipped messages

```
2026-04-13 10:00:01  INFO     [  1] SEND  Mr. Sharma          | +919876543210
2026-04-13 10:00:22  INFO           ✓ Sent successfully
2026-04-13 10:00:27  INFO     [  2] SEND  Ms. Verma           | +919123456789
2026-04-13 10:00:48  INFO           ✓ Sent successfully
```

## Notes

- Keep your phone connected to the internet while the script runs
- WhatsApp may throttle or block accounts sending too many messages too fast — the script adds a 5-second pause between messages
- For large lists (100+ contacts), consider spreading sends across multiple days

## License

MIT
