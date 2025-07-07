from flask import Flask, request, send_file, render_template_string
import pandas as pd
import io
import glob
import os

# Create the Flask app
app = Flask(__name__)

# Folder where your Excel files are stored
EXCEL_FOLDER = r"C:\Users\RaymondLi\OneDrive - 18wheels.ca\auto 1\ProcessedReports"

# HTML for the user interface
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

# Helper function to get the newest matching Excel file
def get_latest_excel_file():
    files = glob.glob(os.path.join(EXCEL_FOLDER, "18whe piece ledger view3*.xlsx"))
    if not files:
        return None
    latest_file = max(files, key=os.path.getmtime)
    return latest_file

# Home page with the date picker form
@app.route('/')
def index():
    return render_template_string(HTML_PAGE)

# Filter and export route
@app.route('/filter')
def filter_report():
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    if not start_date or not end_date:
        return "Please provide start_date and end_date in YYYY-MM-DD format.", 400

    excel_file = get_latest_excel_file()
    if not excel_file:
        return "No matching Excel files found in ProcessedReports folder.", 404

    df = pd.read_excel(excel_file)

    if 'Activity Date' not in df.columns:
        return "The column 'Activity Date' is missing in the Excel file.", 500

    df['Activity Date'] = pd.to_datetime(df['Activity Date'], errors='coerce')

    mask = (df['Activity Date'] >= pd.to_datetime(start_date)) & (df['Activity Date'] <= pd.to_datetime(end_date))
    filtered_df = df.loc[mask]

    if filtered_df.empty:
        return f"No data found between {start_date} and {end_date}", 404

    # Create CSV in memory
    csv_buffer = io.StringIO()
    filtered_df.to_csv(csv_buffer, index=False)
    csv_buffer.seek(0)

    filename = f"ledger_{start_date}_to_{end_date}.csv"

    return send_file(
        io.BytesIO(csv_buffer.getvalue().encode()),
        mimetype='text/csv',
        as_attachment=True,
        download_name=filename  # For Flask 2.x+
    )

# Run the app
if __name__ == '__main__':
    app.run(debug=True)
