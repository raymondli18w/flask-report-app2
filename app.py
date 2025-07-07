from flask import Flask, request, send_file, render_template_string
import pandas as pd
import io
import requests

app = Flask(__name__)

# Replace this with your Dropbox direct download link (must end with ?dl=1)
DROPBOX_LINK = "https://www.dropbox.com/scl/fi/cfssje129vu9pbb8p4cjk/latest.xlsx?rlkey=gruo7e9iteu23rz2f5hdl2nbw&st=o3au5c5u&dl=1"

HTML_PAGE = """
<!doctype html>
<html>
<head>
  <title>Filter and Export Report</title>
</head>
<body>
  <h2>Filter Report by Date Range</h2>
  <form action="/filter" method="get">
    Start Date: <input type="date" name="start_date" required>
    End Date: <input type="date" name="end_date" required>
    <button type="submit">Export CSV</button>
  </form>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_PAGE)

@app.route('/filter')
def filter_report():
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    if not start_date or not end_date:
        return "Please provide start_date and end_date in YYYY-MM-DD format.", 400

    try:
        # Download Excel file from Dropbox
        response = requests.get(DROPBOX_LINK)
        response.raise_for_status()  # Raise error if request failed

        # Load Excel file into pandas dataframe
        excel_data = io.BytesIO(response.content)
        df = pd.read_excel(excel_data)

    except Exception as e:
        return f"Failed to fetch or read Excel file: {e}", 500

    date_column = 'Shipped Date'

    if date_column not in df.columns:
        return f"The column '{date_column}' is missing in the Excel file.", 500

    df[date_column] = pd.to_datetime(df[date_column], errors='coerce')

    mask = (df[date_column] >= pd.to_datetime(start_date)) & (df[date_column] <= pd.to_datetime(end_date))
    filtered_df = df.loc[mask]

    if filtered_df.empty:
        return f"No data found
