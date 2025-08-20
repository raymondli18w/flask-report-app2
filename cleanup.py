import os
import shutil
import tempfile
import logging

# ---------------- Logging ----------------
logging.basicConfig(
    filename="cleanup.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logging.info("Cleanup script started")

# ---------------- Functions ----------------
def remove_temp_files(path):
    """Remove files and folders in the given path"""
    if not os.path.exists(path):
        return
    for root, dirs, files in os.walk(path):
        for f in files:
            try:
                os.remove(os.path.join(root, f))
            except Exception as e:
                logging.warning(f"Failed to remove file {f}: {e}")
        for d in dirs:
            try:
                shutil.rmtree(os.path.join(root, d))
            except Exception as e:
                logging.warning(f"Failed to remove dir {d}: {e}")

def clean_windows_temp():
    """Clean Windows temp folders"""
    temp_paths = [
        tempfile.gettempdir(),                   # Default temp
        os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Temp"),  # Windows Temp
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp")            # LocalAppData Temp
    ]
    for path in temp_paths:
        logging.info(f"Cleaning: {path}")
        remove_temp_files(path)

# ---------------- Run cleanup ----------------
clean_windows_temp()
logging.info("Cleanup completed successfully.")
print("Temporary files cleaned. Check cleanup.log for details.")
