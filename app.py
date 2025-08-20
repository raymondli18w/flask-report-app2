from flask import Flask, render_template, send_file, request
import pandas as pd
import io

app = Flask(__name__)
MASTER_FILE = "master.xlsx"

def load_filtered_df(start_date=None, end_date=None):
    # Load master Excel
    df = pd.read_excel(MASTER_FILE)

    # Parse 'ShippedDate' safely; convert errors to NaT
    df['ShippedDate'] = pd.to_datetime(df.get('ShippedDate', pd.Series()), errors='coerce').dt.date

    # Filter by date if start/end are provided
    if start_date and end_date:
        try:
            start = pd.to_datetime(start_date).date()
            end = pd.to_datetime(end_date).date()
            df = df[(df['ShippedDate'] >= start) & (df['ShippedDate'] <= end)]
        except Exception:
            # If parsing fails, ignore filter
            pass

    return df

@app.route("/", methods=["GET"])
def index():
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")

    df = load_filtered_df(start_date, end_date)

    return render_template(
        "index.html",
        tables=[df.to_html(classes='data', index=False)],
        titles=df.columns.values,
        start_date=start_date,
        end_date=end_date
    )

@app.route("/download", methods=["GET"])
def download():
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")

    df = load_filtered_df(start_date, end_date)

    # Save Excel to in-memory buffer
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
