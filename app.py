from flask import Flask, request, send_file
import pandas as pd
import dropbox
import io
from datetime import datetime

app = Flask(__name__)

DROPBOX_ACCESS_TOKEN = "sl.u.AF7Wym5jOpKbTgpwvOjU38M1mbYtyv0I7YB8CBY2J51U5k9HjW10OLJqaQTlI3QOD_lFhRsWI-3PMyGqaOsZsBEpl5_sh-G96cmjNaKaB7Gjy4My5wiVY0e7tUGu5vrx3LfRZbOq8RfrtdkK77fTx1e8mIJnQg9r4LHbaVcue5qJzfVCc8VXgq5r0_4rH_tIPNmbehKKmUKAWysMeghACZgAvmMlk5ZT5Q4DCzjzD6z4Sb8Nh3d64JfPBibGx029ujck7LcnvsbpVRys5C40DLkeyjNVQdtrj4HU6LGRlOfxy2TWh_M41ygdhZVN2t2pilGqGewdI1CgPuRrzNVV0z4l9QhuW3r7AuK943TgA-sB92WAO5_nD3sqYktjkbCcv9QoNGtxILgjAv0zOj2IeiPQ6UBn-uNDBkNrpfBxN41hZQXT4it3oj7RLQp0FWUU49Uuq7UXYYJJEcLNEdTN5fQE-ty4Opw5jPNQGhCFBU3mdjtIUPqzZ-7i0Jh3-GTHIak3gy1gedxWbACDlVuxGWTla9-i3DgZE4G6TJ4wV0zylFD8dDgUxmBxNnFEuKxSe4gcOjPhrft2FHkH1OhmQvgRYqJuRRWXALCekSg1UtGumg4aklXD3df80_4FXPwb5V-28WbLB8M3VpP1Phx0SfwUtbuzuJM0fSrRCUqMCfLLLPy68n6El4x_Y8eSTpe8KD9Rd6tKidDn2LSIWo5EA7DT6XvqQB2H4MguZLh-W-ckEDCEI7F7E7Mwad5LDlqdiI_-rZwMxbO2LGhjS9QUPtRNfQlITafLCRIqjDl4JT_jSpqIyxfc8WyNq43VqySfyB1e1QXqiaw5BpyKyNcGs4KL9O_m0nokvQxxQzWvCsadL71ep8ktUXyuuRCxCCJ9SPKHezBQB1KuIbhywUjzOnAruEV26027OzkCCBcZbfDnP_or88c_EdFRVR85JeQpCq_pqJf0Jd7_WzneznvGHsXpauqmq7Dy8oh9o82wwcdgOk0DZLDxavbflTVdUsl_2uiABL6EJ4W7e15jS-jDaB5Np2NmZ6SuLVc3XHKOv7wrncxxdpc2EOQ0finEaCq0ELDUCWwatpeUPCE58CFC2MX354TT1E4lHdAuDNQPOy8Vh9R60UwphWizK7rhrztJHGQ9lrf9N78rzYUEgZcDLd1du8PTScTAlXQuuRa65Cot-6UisszfJ8Pn2SZARy6fSfwlIXE-poTouLntRK7N1xmvnDIBh1ai1LKc-Nlv9n_HUwf5j5vEVUybSYZPrtYEDEOuW7Fewn2QSq0m6tYN6YSArCdSYT1na5IkeXWd4JvaJzSg_on5_CM3FM_nRgO3N8G_pGPm-P36aDy4-QWG8itoKhWJ5rJHNozojcxsw4Ppazc1tYx9041KmzsDGqxiuCUUDYQaAggfAto1tANNPp7d"
DROPBOX_FILE_PATH = '/latest.xlsx'  # Dropbox file path

def download_latest_excel():
    dbx = dropbox.Dropbox(DROPBOX_ACCESS_TOKEN)
    try:
        metadata, res = dbx.files_download(DROPBOX_FILE_PATH)
        data = res.content
        return io.BytesIO(data)
    except Exception as e:
        print(f"Error downloading from Dropbox: {e}")
        return None

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        start_date_str = request.form.get("start_date")
        end_date_str = request.form.get("end_date")

        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
            if start_date > end_date:
                return "Error: Start date cannot be after end date.", 400
        except Exception:
            return "Error: Invalid date format.", 400

        excel_file = download_latest_excel()
        if not excel_file:
            return "Error: Could not download Excel file.", 500

        df = pd.read_excel(excel_file)
        df["Shipped Date"] = pd.to_datetime(df["Shipped Date"], errors='coerce')

        filtered = df[(df["Shipped Date"] >= start_date) & (df["Shipped Date"] <= end_date)]

        output = io.BytesIO()
        filtered.to_excel(output, index=False)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="filtered_latest.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    return """
    <form method="POST">
      Start Date: <input type="date" name="start_date" required>
      End Date: <input type="date" name="end_date" required>
      <input type="submit" value="Filter and Download Excel">
    </form>
    """

if __name__ == "__main__":
    app.run(debug=True)
