import os
import pandas as pd
import win32com.client as win32
from datetime import datetime

# Path to your master Excel file
MASTER_FILE = "master.xlsx"
CUTOFF_DATE = datetime(2025, 8, 15)
TARGET_FOLDER_NAME = "Warehouse Excel"  # Name of the folder you want to search

# ----------------------------------------
# Utility: find any Outlook folder by name
# ----------------------------------------
def find_folder(outlook, folder_name):
    for store in outlook.Folders:
        for f in store.Folders:
            if f.Name.lower() == folder_name.lower():
                return f
    return None

# Connect to Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
warehouse_folder = find_folder(outlook, TARGET_FOLDER_NAME)

if warehouse_folder is None:
    print(f"Folder '{TARGET_FOLDER_NAME}' not found!")
    exit()

# Load or create master file
if os.path.exists(MASTER_FILE):
    master_df = pd.read_excel(MASTER_FILE)
else:
    master_df = pd.DataFrame()

print("Master rows BEFORE append:", len(master_df))

# Filter messages
messages = warehouse_folder.Items
messages.Sort("[ReceivedTime]", True)

for msg in messages:
    if msg.ReceivedTime < CUTOFF_DATE:
        continue
    subject = msg.Subject or ""
    if "latesttu01" in subject.lower():
        for att in msg.Attachments:
            if att.FileName.endswith(".xlsx"):
                temp_path = os.path.join(os.getcwd(), att.FileName)
                att.SaveAsFile(temp_path)
                print(f"Downloaded: {att.FileName}")

                # Read attachment
                df = pd.read_excel(temp_path)

                # Merge + skip duplicates
                combined = pd.concat([master_df, df], ignore_index=True).drop_duplicates()
                master_df = combined

                print("Current master row count:", len(master_df))
                print("Preview:")
                print(df.head())

                os.remove(temp_path)  # cleanup temp file

# Save updated master
master_df.to_excel(MASTER_FILE, index=False)
print("Master rows AFTER append:", len(master_df))
print("Master file updated.")
