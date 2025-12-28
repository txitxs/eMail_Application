# eMail_Application

Small Tkinter helper to send mails with attachments. Idea: open the window, fill sender/recipients/subject/body, add files if needed, hit send. No extra libs.

## What it can do
- Shows all fields without fiddling with window size
- Multiple recipients via comma
- Pick attachments and see them listed
- SMTP host/port adjustable (default: smtp.gmail.com:587)
- Status bar for quick feedback

## What you need
- Python 3.11+ (tested with 3.12)
- Tkinter ships with Windows Python
- SMTP credentials (app password recommended for Gmail/Outlook)

## Install
```bash
python -m venv .venv
.venv\Scripts\activate
```

## Run
```bash
python eMail_Application/eMail_Application.py
```

## Quick guide
1) Enter sender email and password.
2) Separate recipients with commas.
3) Write subject and message.
4) Optional: add attachments via "Anhänge hinzufügen".
5) Check/adjust provider domain (e.g., gmail.com) and port (e.g., 587).
6) Click "E-Mail senden" and watch the status bar.

## Practical notes
- The app builds `smtp.<provider>` from the domain.
- Default uses TLS on port 587; change if your provider needs another port.
- Simple regex for emails and a file-exists check for attachments.

## Files
- Main: [eMail_Application/eMail_Application.py](eMail_Application/eMail_Application.py)
