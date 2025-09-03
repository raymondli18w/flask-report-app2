from flask import Flask, render_template, send_file, request
import pandas as pd
import io
import os
from datetime import datetime  # ADD THIS IMPORT

app = Flask(__name__)
MASTER_FILE = "master.xlsx"

# ADD VERSION CHECK ROUTE
@app.route("/version")
def version_check():
    """Check when the app was last updated"""
    return f"Last updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

# ADD EMERGENCY DOWNLOAD ROUTE  
@app.route('/latest')
def emergency_download():
    """Direct download without filtering - for urgent customer requests"""
    return send_file(MASTER_FILE, 
                    as_attachment=True, 
                    download_name=f'LATEST_DATA_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx')

def load_filtered_df(start_date=None, end_date=None):
    """Load Excel, clean dates, and filter if start/end provided."""
    if not os.path.exists(MASTER_FILE):
        return pd.DataFrame()

    try:
        df = pd.read_excel(MASTER_FILE, engine='openpyxl')  # ADD ENGINE FOR SAFETY
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return pd.DataFrame()

    # Normalize column names
    df.columns = [col.strip() for col in df.columns]

    # Find shipped date column
    shipped_col = None
    for col in df.columns:
        if col.lower().replace(" ", "") == "shippeddate":
            shipped_col = col
            break

    if shipped_col:
        # Clean & normalize dates
        df[shipped_col] = pd.to_datetime(
            df[shipped_col].astype(str).str.strip().str.split().str[0],
            errors="coerce"
        )

        if start_date and end_date:
            try:
                start = pd.to_datetime(start_date)
                end = pd.to_datetime(end_date)
                df = df[(df[shipped_col] >= start) & (df[shipped_col] <= end)]
            except Exception:
                pass  # Skip filtering on error

    return df

@app.route("/", methods=["GET", "POST"])
def index():
    start_date = request.values.get("start_date")
    end_date = request.values.get("end_date")

    df = load_filtered_df(start_date, end_date)

    # ADD EMPTY DATA MESSAGE
    if df.empty:
        message = "No data available" + (" for selected date range" if start_date and end_date else "")
    else:
        message = f"Showing {len(df)} records" + (" for selected date range" if start_date and end_date else "")

    return render_template(
        "index.html",
        tables=[df.to_html(classes="data", index=False)] if not df.empty else [],
        titles=df.columns.values if not df.empty else [],
        start_date=start_date,
        end_date=end_date,
        message=message  # ADD MESSAGE TO TEMPLATE
    )

@app.route("/download", methods=["GET", "POST"])
def download():
    start_date = request.values.get("start_date")
    end_date = request.values.get("end_date")

    df = load_filtered_df(start_date, end_date)

    # Create in-memory Excel
    output = io.BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')  # ADD ENGINE
    output.seek(0)

    filename = "filtered_master.xlsx" if (start_date and end_date) else "master.xlsx"
    return send_file(
        output,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
    )

if __name__ == "__main__":
    app.run(debug=True)