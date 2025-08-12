from flask import Flask, request, send_file, render_template_string
import pandas as pd
from datetime import datetime

app = Flask(__name__)

HTML_FORM = """
<!doctype html>
<title>Filter Shipped Date</title>
<h2>Filter data by shipped date range</h2>
<form method="POST">
  Start Date: <input type="date" name="start_date" required>
  End Date: <input type="date" name="end_date" required>
  <input type="submit" value="Filter and Download Excel">
</form>
{% if error %}
  <p style="color:red;">{{ error }}</p>
{% endif %}
{% if preview %}
  <h3>Filtered Data Preview (first 5 rows):</h3>
  {{ preview|safe }}
{% endif %}
"""

@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    preview_html = None
    if request.method == "POST":
        start_date_str = request.form.get("start_date")
        end_date_str = request.form.get("end_date")

        # Convert strings to datetime objects
        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
        except ValueError:
            error = "Invalid date format. Please use YYYY-MM-DD."
            return render_template_string(HTML_FORM, error=error)

        if start_date > end_date:
            error = "Start date cannot be after end date."
            return render_template_string(HTML_FORM, error=error)

        # Load data and filter
        df = pd.read_excel("latest.xlsx")
        df["Shipped Date"] = pd.to_datetime(df["Shipped Date"], errors='coerce')

        filtered = df[(df["Shipped Date"] >= start_date) & (df["Shipped Date"] <= end_date)]

        if filtered.empty:
            error = f"No data found between {start_date_str} and {end_date_str}."
            return render_template_string(HTML_FORM, error=error)

        # Save filtered to a file for download
        filtered_file = "filtered_latest.xlsx"
        filtered.to_excel(filtered_file, index=False)

        # Preview first 5 rows as HTML table
        preview_html = filtered.head().to_html(classes="table table-striped")

        # Instead of forcing download immediately, show preview and a download link
        return render_template_string(
            HTML_FORM + """
            <p>Filtered data contains {{ rows }} rows.</p>
            <a href="/download?start_date={{ start_date }}&end_date={{ end_date }}">Download Excel</a>
            """,
            error=error,
            preview=preview_html,
            rows=len(filtered),
            start_date=start_date_str,
            end_date=end_date_str
        )

    return render_template_string(HTML_FORM)

@app.route("/download")
def download_filtered():
    start_date_str = request.args.get("start_date")
    end_date_str = request.args.get("end_date")

    # Re-run filtering here to get fresh file
    try:
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
        end_date = datetime.strptime(end_date_str, "%Y-%m-%d")
    except Exception:
        return "Invalid date parameters.", 400

    df = pd.read_excel("latest.xlsx")
    df["Shipped Date"] = pd.to_datetime(df["Shipped Date"], errors='coerce')
    filtered = df[(df["Shipped Date"] >= start_date) & (df["Shipped Date"] <= end_date)]

    if filtered.empty:
        return f"No data found between {start_date_str} and {end_date_str}.", 404

    filtered_file = "filtered_latest.xlsx"
    filtered.to_excel(filtered_file, index=False)

    return send_file(
        filtered_file,
        as_attachment=True,
        download_name="filtered_latest.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    app.run(debug=True)
