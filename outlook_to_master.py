import os
import pandas as pd
import win32com.client as win32
from datetime import datetime, timedelta
import logging

# ----------------- Logging -----------------
logging.basicConfig(
    filename="outlook_to_master.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logging.info("Script started")

# ----------------- Config -----------------
MASTER_FILE = "master.xlsx"
MAX_EMAILS = 3                 # Stop after first 3 emails
DAYS_LOOKBACK = 2              # Look back 2 days
KEYWORD = "latesttu01"         # Filter emails containing this in subject
cutoff_datetime = datetime.now() - timedelta(days=DAYS_LOOKBACK)

# ----------------- Connect to Outlook -----------------
try:
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    logging.info("Connected to Outlook Inbox")
except Exception as e:
    logging.error(f"Failed to connect to Outlook: {e}")
    exit()

# ----------------- Load or create master file -----------------
if os.path.exists(MASTER_FILE):
    master_df = pd.read_excel(MASTER_FILE)
    logging.info(f"Loaded master file: {MASTER_FILE} ({len(master_df)} rows)")
else:
    master_df = pd.DataFrame()
    logging.info(f"No existing master file found. Starting new DataFrame.")

# ----------------- Filter messages -----------------
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)  # Sort newest first
# Restrict to last N days
messages = messages.Restrict(f"[ReceivedTime] >= '{cutoff_datetime.strftime('%m/%d/%Y %H:%M %p')}'")

processed_count = 0

for msg in messages:
    try:
        subject = msg.Subject or ""
        if KEYWORD in subject.lower():
            for att in msg.Attachments:
                if att.FileName.endswith(".xlsx"):
                    temp_path = os.path.join(os.getcwd(), att.FileName)
                    att.SaveAsFile(temp_path)
                    logging.info(f"Downloaded attachment: {att.FileName}")

                    # Read and append to master
                    df = pd.read_excel(temp_path)
                    master_df = pd.concat([master_df, df], ignore_index=True).drop_duplicates()

                    logging.info(f"Current master row count: {len(master_df)}")
                    logging.info(f"Attachment preview:\n{df.head()}")

                    # Clean up temporary file
                    os.remove(temp_path)

            processed_count += 1
            if processed_count >= MAX_EMAILS:
                logging.info(f"Processed {MAX_EMAILS} emails, stopping.")
                break
    except Exception as e:
        logging.warning(f"Failed to process message '{subject}': {e}")

# ----------------- Save updated master -----------------
master_df.to_excel(MASTER_FILE, index=False)
logging.info(f"Master rows AFTER append: {len(master_df)}")
logging.info("Master file updated successfully.")
