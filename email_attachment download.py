import imaplib
import email
from email.header import decode_header
import os

# ----------- USER DETAILS -----------
EMAIL = "aishwarya.dindore07@gmail.com"           # 🔁 Your Gmail address
PASSWORD = "pbum brvo cbmu dekh"           # 🔁 App Password from Google
FOLDER = r"C:\Users\Admin\Desktop\tanmay proj\attachments"  # 🔁 Local folder to save files
SUBJECT = "E-account statement for your SBI account(s)"                   # 🔁 Subject to filter emails (or leave as "" for all)
# -------------------------------------

# Create folder if it doesn't exist
os.makedirs(FOLDER, exist_ok=True)

# Connect to Gmail
mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(EMAIL, PASSWORD)
mail.select("inbox")

# Search emails
search_criteria = f'SUBJECT "{SUBJECT}"' if SUBJECT else "ALL"
status, data = mail.search(None, search_criteria)

for num in data[0].split():
    _, msg_data = mail.fetch(num, "(RFC822)")
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    for part in msg.walk():
        if part.get_content_disposition() == "attachment":
            filename = part.get_filename()
            if filename:
                filename = decode_header(filename)[0][0]
                if isinstance(filename, bytes):
                    filename = filename.decode()
                filepath = os.path.join(FOLDER, filename)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print("Downloaded:", filepath)

mail.logout()
