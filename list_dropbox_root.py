import os
import dropbox
from dotenv import load_dotenv

load_dotenv()

DROPBOX_ACCESS_TOKEN = os.getenv("DROPBOX_ACCESS_TOKEN")
dbx = dropbox.Dropbox(DROPBOX_ACCESS_TOKEN)

try:
    result = dbx.files_list_folder("")
    print("Root folder contents:")
    for entry in result.entries:
        print("-", entry.name)
except Exception as e:
    print("Error listing root folder:", e)
