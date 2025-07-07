from flask import Flask, request, send_file, render_template_string
import pandas as pd
import io
import requests

app = Flask(__name__)

# Your OneDrive API link to the public Excel file
ONEDRIVE_FILE_URL = "https://api.onedrive.com/v1.0/shares/u!EYwU9rsnIGJBlRaM3v_iRNkBWRkE2tulp-oBFpTt3LUoUw/root/content"

# HTML page for date input
HTML_PAGE = """
<!doctype html>
<html>
<head><title>Filter Report</title></head>
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
        response = requests.get(ONEDRIVE_FILE_URL)
        response.raise_for_status()
        excel_bytes = io.BytesIO(response.content)
    except Exception as e:
        return f"Failed to fetch Excel file from OneDrive: {e}", 500

    try:
        df = pd.read_excel(excel_bytes)

        if 'Shipped Date' not in df.columns:
            return "The column 'Shipped Date' is missing in the Excel file.", 500

        df['Shipped Date'] = pd.to_datetime(df['Shipped Date'], errors='coerce')
        mask = (df['Shipped Date'] >= pd.to_datetime(start_date)) & (df['Shipped Date'] <= pd.to_datetime(end_date))
        filtered_df = df.loc[mask]

        if filtered_df.empty:
            return f"No data found between {start_date} and {end_date}", 404

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

    except Exception as e:
        return f"Error processing Excel file: {e}", 500

@app.route('/ping')
def ping():
    return "pong", 200

if __name__ == '__main__':
    app.run(debug=True)
