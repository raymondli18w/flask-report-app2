from flask import Flask, render_template, send_file, request
import pandas as pd
import io
import os

app = Flask(__name__)
MASTER_FILE = "master.xlsx"   # make sure this file is in the same folder as app.py


def load_filtered_df(start_date=None, end_date=None):
    """Load Excel, clean dates, and filter if start/end provided."""
    if not os.path.exists(MASTER_FILE):
        return pd.DataFrame()  # return empty if file missing

    df = pd.read_excel(MASTER_FILE)

    # Normalize column names (strip spaces, lowercased for consistency)
    df.columns = [col.strip() for col in df.columns]

    # Find the shipped date column, whether it's "Shipped Date" or "ShippedDate"
    shipped_col = None
    for col in df.columns:
        if col.lower().replace(" ", "") == "shippeddate":
            shipped_col = col
            break

    if shipped_col:
        # Clean & normalize the dates
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
                pass  # if bad input, skip filtering

    return df


@app.route("/", methods=["GET", "POST"])
def index():
    start_date = request.values.get("start_date")
    end_date = request.values.get("end_date")

    df = load_filtered_df(start_date, end_date)

    return render_template(
        "index.html",
        tables=[df.to_html(classes="data", index=False)] if not df.empty else [],
        titles=df.columns.values if not df.empty else [],
        start_date=start_date,
        end_date=end_date,
    )


@app.route("/download", methods=["GET", "POST"])
def download():
    start_date = request.values.get("start_date")
    end_date = request.values.get("end_date")

    df = load_filtered_df(start_date, end_date)

    # Save Excel in memory
    output = io.BytesIO()
    df.to_excel(output, index=False)
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
