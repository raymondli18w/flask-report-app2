import os
import pandas as pd
from datetime import datetime
import logging

# ----------------- Logging -----------------
logging.basicConfig(
    filename="folder_to_master.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logging.info("Script started")

# ----------------- Config -----------------
MASTER_FILE = "master.xlsx"
SOURCE_FOLDER = r"C:\Users\RaymondLi\OneDrive - 18wheels.ca\auto 1\auto 22"
MAX_FILES = 3  # stop after first 3 files

# ----------------- Load or create master file -----------------
if os.path.exists(MASTER_FILE):
    master_df = pd.read_excel(MASTER_FILE)
    logging.info(f"Loaded master file: {MASTER_FILE} ({len(master_df)} rows)")
else:
    master_df = pd.DataFrame()
    logging.info("No existing master file found. Starting new DataFrame.")

# ----------------- Find matching files -----------------
all_files = [f for f in os.listdir(SOURCE_FOLDER) 
             if f.startswith("18WHE") and f.endswith(".xlsx")]

# Sort by last modified time (newest first)
all_files.sort(key=lambda x: os.path.getmtime(os.path.join(SOURCE_FOLDER, x)), reverse=True)

processed_count = 0

for file_name in all_files:
    if processed_count >= MAX_FILES:
        break
    try:
        file_path = os.path.join(SOURCE_FOLDER, file_name)
        df = pd.read_excel(file_path)

        master_df = pd.concat([master_df, df], ignore_index=True).drop_duplicates()
        logging.info(f"Processed file: {file_name}")
        logging.info(f"Current master row count: {len(master_df)}")
        logging.info(f"Preview:\n{df.head()}")

        processed_count += 1
    except Exception as e:
        logging.warning(f"Failed to process file '{file_name}': {e}")

# ----------------- Save updated master -----------------
master_df.to_excel(MASTER_FILE, index=False)
logging.info(f"Master rows AFTER append: {len(master_df)}")
logging.info("Master file updated successfully.")
