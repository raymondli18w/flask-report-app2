from flask import Flask, render_template, send_file, request
import pandas as pd
import io

app = Flask(__name__)
MASTER_FILE = "master.xlsx"

def load_filtered_df(start_date=None, end_date=None):
    try:
        df = pd.read_excel(MASTER_FILE)
    except Exception as e:
        print(f"Error loading {MASTER_FILE}: {e}")
        return pd.DataFrame()

    col_name = "Shipped Date"  # exact column name in Excel
    if col_name not in df.columns:
        print(f"Warning: '{col_name}' column missing in Excel")
        return df

    # Convert to datetime, coerce errors, and keep only the date
    df[col_name] = pd.to_datetime(df[col_name], errors='coerce').dt.date

    if start_date and end_date:
        start = pd.to_datetime(start_date).date()
        end = pd.to_datetime(end_date).date()
        df = df[(df[col_name] >= start) & (df[col_name] <= end)]

    return df

@app.route("/", methods=["GET"])
def index():
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")

    df = load_filtered_df(start_date, end_date)

    return render_template(
        "index.html",
        tables=[df.to_html(classes='data')],
        titles=df.columns.values,
        start_date=start_date,
        end_date=end_date
    )

@app.route("/download", methods=["GET"])
def download():
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")

    df = load_filtered_df(start_date, end_date)

    # Create Excel in memory
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    filename = "filtered_master.xlsx" if (start_date and end_date) else "master.xlsx"
    return send_file(
        output,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True
    )

if __name__ == "__main__":
    app.run(debug=True)
