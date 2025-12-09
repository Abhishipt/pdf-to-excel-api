import os
import tempfile
import shutil
import camelot
import pdfplumber
import pandas as pd
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from PyPDF2 import PdfReader
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
import time
import threading

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

CLEANUP_INTERVAL = 180  # seconds (3 minutes)

def periodic_cleanup():
    while True:
        now = time.time()
        for filename in os.listdir(UPLOAD_FOLDER):
            path = os.path.join(UPLOAD_FOLDER, filename)
            if os.path.isfile(path) and now - os.path.getmtime(path) > CLEANUP_INTERVAL:
                os.remove(path)
        time.sleep(60)

threading.Thread(target=periodic_cleanup, daemon=True).start()

def fallback_pdfplumber(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        rows = []
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for row in table:
                    if any(cell is not None for cell in row):
                        rows.append([cell if cell else "" for cell in row])
        return pd.DataFrame(rows)

def save_to_excel(df, output_path):
    wb = Workbook()
    ws = wb.active

    # Styling setup
    bold_font = Font(bold=True, name="Mangal")
    normal_font = Font(name="Mangal")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    for i, row in df.iterrows():
        for j, value in enumerate(row):
            cell = ws.cell(row=i+1, column=j+1, value=value)
            cell.font = bold_font if i == 0 else normal_font
            cell.alignment = center_align
            cell.border = border

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[column].width = max_length + 2

    wb.save(output_path)

@app.route("/convert", methods=["POST"])
def convert():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    filename = file.filename.rsplit(".", 1)[0]
    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(input_path)

    try:
        tables = camelot.read_pdf(input_path, pages="all", flavor="lattice")
        if tables and tables.n > 0:
            dfs = [table.df for table in tables if table.df.shape[1] >= 5]
            final_df = pd.concat(dfs, ignore_index=True)
        else:
            final_df = fallback_pdfplumber(input_path)

        output_path = os.path.join(
            UPLOAD_FOLDER,
            f"{filename}_Tools_Subidha.xlsx"
        )
        save_to_excel(final_df, output_path)

        return send_file(output_path, as_attachment=True)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        if os.path.exists(input_path):
            os.remove(input_path)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
