import dropbox
import io
import os
from dotenv import load_dotenv

load_dotenv()

token = os.getenv("DROPBOX_ACCESS_TOKEN")
file_path = os.getenv("DROPBOX_FILE_PATH")

dbx = dropbox.Dropbox(token)

try:
    metadata, res = dbx.files_download(file_path)
    print("File downloaded:", metadata.name)
    print("File size (bytes):", len(res.content))
except Exception as e:
    print("Error downloading file:", e)
