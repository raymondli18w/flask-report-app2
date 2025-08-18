from flask import Flask, request, send_file, render_template
import pandas as pd
from io import BytesIO

app = Flask(__name__)

MASTER_FILE = "master.xlsx"

@app.route('/')
def index():
    # Show a simple HTML page with date inputs
    return render_template('index.html')

@app.route('/download')
def download():
    start_date = request.args.get('start')
    end_date = request.args.get('end')

    df = pd.read_excel(MASTER_FILE)

    # Filter by shipped date
    df['Shipped Date'] = pd.to_datetime(df['Shipped Date'])
    if start_date and end_date:
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
        df = df[(df['Shipped Date'] >= start) & (df['Shipped Date'] <= end)]

    # Export to Excel in memory
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(output, download_name='filtered_master.xlsx', as_attachment=True)
