from flask import Flask, render_template, send_file, request
import pandas as pd

app = Flask(__name__)
MASTER_FILE = "master.xlsx"

@app.route("/")
def index():
    # Optional: get date range from query parameters
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")
    
    df = pd.read_excel(MASTER_FILE)

    if start_date and end_date:
        df['ShippedDate'] = pd.to_datetime(df['ShippedDate'], errors='coerce')
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        df = df[(df['ShippedDate'] >= start) & (df['ShippedDate'] <= end)]

    # Render as HTML table
    return render_template("index.html", tables=[df.to_html(classes='data')], titles=df.columns.values)

@app.route("/download")
def download():
    return send_file(MASTER_FILE, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
