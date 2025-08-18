import os
import requests

ACCESS_TOKEN = 'sl.u.AF5Cox4GqlXDe58elsDAVi-_nl_YUgMPSsRLSBJLIv8mr03q-6rYLG7ZM3VPbNMhOdf7Z40CjP0lqSIzvz2G4O_PU0qhaAi8-nPf98Y8Oy1k_ahRYu0mV8hvD0KYuPdGR2DwF9mwEcKcPVEiU3qDTDnXuQ9T59CJ5lD73i7OeuuzRv3kWLDNtjjI3Kc3rf4wDnaAY2y7kwKfhspQ7So0MDOkoeKKmVQuGKulWlRu104ZOnW9zMRd5mm7-WwyAMCw9d32ms1F0eJNDqGfDXplhIDDgpYM7uKtkVB-eh1erKwxELXrPo_57DoimJlaaKhlXyv7618UKZTlboGrx1JEcT29iJ7nMk0alN8nRmsGFv8TfffhATIRLMXuGbTnenfru4aTB6ZEJgU181a5AJ8Rlk7KtGFtRyoJl8P-a4G-O8uqxh0OPhorwxSx5GDXGqsGaswEEjV9rO5gV3hYh_DWXNUkRmH9PCtozf5A7a-cWvJIGWj-TwMMXCYaUohz8XiGBEqPX0cW1B_ECeP5mWZPmF6QUiepZohBgVV59K-bLLc8mAkLNrzW3RhUU9ymqm-N-VcAjIIiwhSzPKUqi6lypIwY8a4bu68mFTh0eok3HOvBQ5mTaIFJSWRCAGBMT6CMuJyh3e7N1tLqWQm_dOTQzs5VO4T8ew9_jxd6oFvSdnGJmMsPQQG39eK57kmzU0CoOSRsCKW_rWoISSVJaPbgr6xXvR-jwWr64zWTfhsYEe0Ws90vmKaz7VzIqZcyo7J1Rwt5tSecrL2eQY2DBCdSe2yLBjyoBthpWtJOLt8NiOIGfSUYRDCCbIdRgeDAZfqHI7ARKWFWmYOxkuCuAQZOLc7XIl_NfBgY9tUMGLyKjbMQwOcpPF1ksxaYiOp41xvC6i1oSBb2omXnzJke_VtrE0wxNEDAdF4a5tqoiFdNrxl8YvhhyU-GlFGWCw88eruu4fnn3lvQwSf5Uy5I77aEU8hJg_WYmo29Zzcfij-DEJYkhNjb48vziFGbLGclW0JAnWuMNndnebK5dHjCnAzdylla-do3RtHueEQvyM3N2L4O0MZWa7HMGy96vsuNty0s54zhUxExjA2zJu4G-MP_0cb5tTCBZbRkfsgdBPKLHWgiBoYVwldDjsUON9XEVf_csIDJyNsVS3mfn4LQ2g4SAp1uKEhRr2LarXuUwdXPbxkroJ4doA17zMpk7yTMZtFgRKqGGzK6HEpv5mmLfnrL7nKbmSufLIRSXGUuq6-6OX2p1BvxFHmauls2Yn70BPiyitx35FE9OCBMs02vbjrJnCXnANZSQPlCejNbzAZSQZrES9eXwsWH3XlQVgj6GgGCV_LmZzNAWFp4HlbbWH27l1V-52BktgNKjQsGfFDnbs_Z2gYI9Aq6WsX7lUs7TtTxO_zjjvdUW0IIAUOnfRtB6LZF'
FILE_PATH = '/latest.xlsx'  # Dropbox file path

def get_temp_link():
    url = "https://api.dropboxapi.com/2/files/get_temporary_link"
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json"
    }
    resp = requests.post(url, headers=headers, json={"path": FILE_PATH})
    resp.raise_for_status()
    return resp.json()["link"]

def download_file(download_url):
    r = requests.get(download_url)
    r.raise_for_status()
    with open("latest.xlsx", "wb") as f:
        f.write(r.content)
    print("Downloaded latest.xlsx successfully.")

if __name__ == "__main__":
    temp_link = get_temp_link()
    download_file(temp_link)
