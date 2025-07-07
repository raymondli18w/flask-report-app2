from flask import Flask, request, send_file, render_template_string
import pandas as pd
import io
import requests

app = Flask(__name__)

# Dropbox direct download link to your latest Excel file
DROPBOX_FILE_URL = "https://www.dropbox.com/scl/fi/cfssje129vu9pbb8p4cjk/latest.xlsx?rlkey=1xn1ona4h5yv653yak4hxieoz&dl=1"

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

    # Download the Excel file from Dropbox
    try:
        response = requests.get(DROPBOX_FILE_URL)
        response.raise_for_status()
    except Exception as e:
        return f"Failed to download Excel file: {str(e)}", 500

    try:
        # Load Excel file into DataFrame from bytes
        df = pd.read_excel(io.BytesIO(response.content))
    except Exception as e:
        return f"Failed to read Excel file: {str(e)}", 500

    # Check if 'shipped date' column exists (case insensitive)
    df_columns_lower = [col.lower() for col in df.columns]
    if 'shipped date' not in df_columns_lower:
        return "The column 'shipped date' is missing in the Excel file.", 500

    # Normalize column name to access
    shipped_date_col = [col for col in df.columns if col.lower() == 'shipped date'][0]

    # Convert shipped date to datetime
    df[shipped_date_col] = pd.to_datetime(df[shipped_date_col], errors='coerce')

    # Filter by date range
    mask = (df[shipped_date_col] >= pd.to_datetime(start_date)) & (df[shipped_date_col] <= pd.to_datetime(end_date))
    filtered_df = df.loc[mask]

    if filtered_df.empty:
        return f"No data found between {start_date} and {end_date}.", 404

    # Export filtered data to CSV in memory
    csv_buffer = io.StringIO()
    filtered_df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)

    filename = f"report_{start_date}_to_{end_date}.csv"

    return send_file(
        io.BytesIO(csv_buffer.getvalue().encode('utf-8')),
        mimetype='text/csv',
        as_attachment=True,
        download_name=filename
    )

@app.route('/ping')
def ping():
    return "pong", 200

if __name__ == '__main__':
    app.run(debug=True)
