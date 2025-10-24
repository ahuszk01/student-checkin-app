from flask import Flask, render_template_string, request, redirect, url_for
from openpyxl import load_workbook
from datetime import datetime
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os

app = Flask(__name__)

# Google Drive file ID for your Excel sheet
DRIVE_FILE_ID = '1DoMicOLFTH1pabvxCZTK98F9BUz4mMCs'
LOCAL_EXCEL_PATH = 'student_roster.xlsx'

GROUPS = {
    "Csiga": "🐌",
    "Suni": "🦔",
    "Katica": "🐞",
    "Katica halado": "🐞",
    "Pillango": "🦋",
    "Nyuszi": "🐇",
    "Baglyok": "🦉",
    "Sas": "🦅"
}

def fetch_excel_from_drive():
    """Download the Excel file from Google Drive."""
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()
    drive = GoogleDrive(gauth)
    file = drive.CreateFile({'id': DRIVE_FILE_ID})
    file.GetContentFile(LOCAL_EXCEL_PATH)

@app.route("/")
def home():
    return render_template_string("""
    <html>
    <head>
      <title>Select Group</title>
      <style>
        body { font-family: sans-serif; background: #f0f4f8; padding: 20px; text-align: center; }
        .group-btn {
          display: inline-block;
          margin: 10px;
          padding: 20px;
          font-size: 24px;
          background: #0078D4;
          color: white;
          border-radius: 10px;
          width: 180px;
          height: 100px;
          line-height: 1.2;
          text-decoration: none;
        }
      </style>
    </head>
    <body>
      <h2>Select a Group</h2>
      {% for name, icon in groups.items() %}
        <a class="group-btn" href="/group/{{ name }}">{{ icon }}<br>{{ name }}</a>
      {% endfor %}
    </body>
    </html>
    """, groups=GROUPS)

@app.route("/group/<group>", methods=["GET", "POST"])
def group_view(group):
    fetch_excel_from_drive()
    wb = load_workbook(LOCAL_EXCEL_PATH)
    if group not in wb.sheetnames:
        return f"<h3>Group '{group}' not found.</h3>", 404

    ws = wb[group]
    today_str = datetime.now().strftime("%d/%m/%Y")

    date_col = None
    for col in range(2, ws.max_column + 1):
        cell = ws.cell(row=1, column=col).value
        if isinstance(cell, datetime):
            cell = cell.strftime("%d/%m/%Y")
        if str(cell) == today_str:
            date_col = col
            break

    students = []
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row=row, column=1).value
        if name not in [None, ""]:
            checked_in = False
            if date_col:
                cell = ws.cell(row=row, column=date_col)
                checked_in = (cell.value == "✅")
            students.append({"name": name, "checked_in": checked_in})

    if request.method == "POST":
        student = request.form["student"]
        if date_col:
            for row in range(2, ws.max_row + 1):
                if ws.cell(row=row, column=1).value == student:
                    cell = ws.cell(row=row, column=date_col)
                    if cell.value != "✅":
                        cell.value = "✅"
                        wb.save(LOCAL_EXCEL_PATH)
                    break
        return redirect(url_for("group_view", group=group, checked=student))

    checked = request.args.get("checked", "")
    return render_template_string("""
    <html>
    <head>
      <title>{{ group }} Students</title>
      <style>
        body { font-family: sans-serif; background: #f0f4f8; padding: 20px; text-align: center; }
        .student-grid { display: flex; flex-wrap: wrap; justify-content: center; gap: 15px; }
        .student-card {
          width: 120px;
          height: 140px;
          border-radius: 10px;
          box-shadow: 0 0 5px rgba(0,0,0,0.1);
          background: white;
          padding: 0;
          opacity: 1;
        }
        .student-card.checked {
          opacity: 0.4;
        }
        .card-button {
          width: 100%;
          height: 100%;
          border: none;
          background: none;
          font-size: 24px;
          padding: 10px;
          cursor: pointer;
          display: flex;
          flex-direction: column;
          justify-content: center;
          align-items: center;
          color: #333;
        }
        .card-button:hover {
          background: #e0f0ff;
          border-radius: 10px;
        }
        .card-button p {
          margin: 5px 0 0;
          font-size: 16px;
        }
      </style>
    </head>
    <body>
      <h2>{{ group }} {{ icon }}</h2>
      {% if checked %}
        <p style="color: green;">{{ checked }} checked in successfully!</p>
      {% endif %}
      <div class="student-grid">
        {% for student in students %}
          <form method="POST" class="student-card {% if student.checked_in %}checked{% endif %}">
            <input type="hidden" name="student" value="{{ student.name }}">
            <button type="submit" class="card-button" {% if student.checked_in %}disabled{% endif %}>
              <div style="font-size: 40px;">👤</div>
              <p>{{ student.name }}</p>
            </button>
          </form>
        {% endfor %}
      </div>
      <p><a href="/">← Back to Groups</a></p>
    </body>
    </html>
    """, group=group, icon=GROUPS.get(group, ""), students=students, checked=checked)

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000)
