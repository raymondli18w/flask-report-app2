from flask import Flask, render_template, send_file, request
import pandas as pd
import io

app = Flask(__name__)
MASTER_FILE = "master.xlsx"

def load_filtered_df(start_date=None, end_date=None):
    df = pd.read_excel(MASTER_FILE)
    df['ShippedDate'] = pd.to_datetime(df['ShippedDate'], errors='coerce')
    if start_date and end_date:
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        df = df[(df['ShippedDate'] >= start) & (df['ShippedDate'] <= end)]
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

    # Stream Excel directly from memory
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
