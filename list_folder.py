import os
import requests

ACCESS_TOKEN = os.environ.get("DROPBOX_ACCESS_TOKEN")
FOLDER_PATH = "/latest.xlsx"  # empty string for root folder or "/render 1" for subfolder

def list_folder():
    url = "https://api.dropboxapi.com/2/files/list_folder"
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }
    data = {
        "path": FOLDER_PATH,
        "recursive": False
    }
    resp = requests.post(url, headers=headers, json=data)
    resp.raise_for_status()
    return resp.json()

if __name__ == "__main__":
    folder_contents = list_folder()
    print(folder_contents)
