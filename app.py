from flask import Flask, render_template, send_file, request
import pandas as pd

app = Flask(__name__)
MASTER_FILE = "master.xlsx"  # Make sure this path is correct

@app.route("/", methods=["GET", "POST"])
def index():
    df = pd.read_excel(MASTER_FILE)
    
    # Handle form submission for date filtering
    if request.method == "POST":
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")
        if start_date and end_date:
            df["Shipped Date"] = pd.to_datetime(df["Shipped Date"])
            df = df[(df["Shipped Date"] >= start_date) & (df["Shipped Date"] <= end_date)]
    
    table = df.to_html(classes="data", index=False)
    return render_template("index.html", tables=[table])

@app.route("/download")
def download():
    return send_file(MASTER_FILE, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
