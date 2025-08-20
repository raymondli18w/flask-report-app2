import os
import pandas as pd
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

# ----------------- Load or create master file -----------------
if os.path.exists(MASTER_FILE):
    master_df = pd.read_excel(MASTER_FILE, engine="openpyxl")
    logging.info(f"Loaded master file: {MASTER_FILE} ({len(master_df)} rows)")
else:
    master_df = pd.DataFrame()
    logging.info("No existing master file found. Starting new DataFrame.")

# ----------------- Find latest 18WHE file -----------------
all_files = [f for f in os.listdir(SOURCE_FOLDER) if f.startswith("18WHE") and f.endswith(".xlsx")]

if not all_files:
    logging.warning("No matching files found in source folder.")
else:
    # Sort by last modified time (newest first) and pick only the latest file
    all_files.sort(key=lambda x: os.path.getmtime(os.path.join(SOURCE_FOLDER, x)), reverse=True)
    latest_file = os.path.join(SOURCE_FOLDER, all_files[0])
    logging.info(f"Processing latest file: {all_files[0]}")

    try:
        df = pd.read_excel(latest_file, engine="openpyxl")
        master_df = pd.concat([master_df, df], ignore_index=True).drop_duplicates()
        logging.info(f"Current master row count: {len(master_df)}")
        logging.info(f"Preview:\n{df.head()}")
    except Exception as e:
        logging.error(f"Failed to process file '{latest_file}': {e}")

# ----------------- Save updated master -----------------
master_df.to_excel(MASTER_FILE, index=False, engine="openpyxl")
logging.info(f"Master rows AFTER append: {len(master_df)}")
logging.info("Master file updated successfully.")
