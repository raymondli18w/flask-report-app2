import win32com.client
import pandas as pd
import requests
import os

# ---- Config ----
FLASK_URL = "http://127.0.0.1:5000/append"  # change to your Flask URL
SUBJECT_KEYWORD = "latesttu01"
TMP_FOLDER = "C:/Temp"  # temporary folder to save attachments

os.makedirs(TMP_FOLDER, exist_ok=True)

# ---- Connect to Outlook ----
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox

messages = inbox.Items
messages.Sort("[ReceivedTime]", True)  # newest first

for msg in messages:
    if SUBJECT_KEYWORD in msg.Subject:
        attachments = msg.Attachments
        for att in attachments:
            if att.FileName.endswith(".xlsx"):
                filepath = os.path.join(TMP_FOLDER, att.FileName)
                att.SaveAsFile(filepath)
                print(f"Saved attachment: {filepath}")

                # ---- Send to Flask ----
                with open(filepath, "rb") as f:
                    files = {"file": (att.FileName, f, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
                    response = requests.post(FLASK_URL, files=files)

                print("Flask response:", response.text)
        break  # only process newest matching email
