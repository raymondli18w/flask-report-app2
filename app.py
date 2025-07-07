from flask import Flask, request, send_file, render_template_string
import pandas as pd
import requests
import io

app = Flask(__name__)

# ✅ Direct download link to your OneDrive Excel file
ONEDRIVE_FILE_URL = "https://api.onedrive.com/v1.0/shares/u!EYwU9rsnIGJBlRaM3v_iRNkBaMrUrEkRCWjMrUA8wS4cJw/root/content"

# ✅ HTML for UI
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

# ✅ Helper function: download latest file from OneDrive
def get_latest_excel_file():
    try:
        response = requests.get(ONEDRIVE_FILE_URL)
        response.raise_for_status()
        return io.BytesIO(response.content)  # Load as in-memory file
    except Exception as e:
        print(f"Error downloading file: {e}")
        return None

# ✅ Home page
@app.route('/')
def index():
    return render_template_string(HTML_PAGE)

# ✅ Download filtered CSV
@app.route('/filter')
def filter_report():
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    if not start_date or not end_date:
        return "Please provide start_date and end_date in YYYY-MM-DD format.", 400

    excel_file = get_latest_excel_file()
    if not excel_file:
        return "Failed to fetch Excel file from OneDrive.", 500

    df = pd.read_excel(excel_file)

    if 'Shipped Date' not in df.columns:
        return "The column 'Shipped Date' is missing in the Excel file.", 500

    df['Shipped Date'] = pd.to_datetime(df['Shipped Date'], errors='coerce')
    mask = (df['Shipped Date'] >= pd.to_datetime(start_date)) & (df['Shipped Date'] <= pd.to_datetime(end_date))
    filtered_df = df.loc[mask]

    if filtered_df.empty:
        return f"No data found between {start_date} and {end_date}", 404

    # ✅ Convert to CSV in memory
    csv_buffer = io.StringIO()
    filtered_df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)

    filename = f"report_{start_date}_to_{end_date}.csv"
    return send_file(
        io.BytesIO(csv_buffer.getvalue().encode()),
        mimetype='text/csv',
        as_attachment=True,
        download_name=filename
    )

# ✅ Health check route (optional)
@app.route("/ping")
def ping():
    return "pong", 200

# ✅ Local run (Render ignores this block)
if __name__ == '__main__':
    app.run(debug=True)
