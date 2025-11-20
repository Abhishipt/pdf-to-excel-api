from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import uuid
import threading
import time
import camelot
import pdfplumber
import openpyxl
from openpyxl.styles import Border, Side, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Track active files to avoid cleaning them mid-process
ACTIVE_FILES = set()

# Auto-delete files after delay
def delete_file_later(path, delay=300):
    def remove():
        time.sleep(delay)
        if os.path.exists(path) and os.path.basename(path) not in ACTIVE_FILES:
            try:
                os.remove(path)
                print(f"[CLEANUP] Deleted {path}")
            except Exception as e:
                print(f"[ERROR] Cleanup failed: {e}")
    threading.Thread(target=remove, daemon=True).start()

# Periodic cleanup every 3 minutes
def periodic_cleanup(interval=180):
    def cleanup():
        while True:
            time.sleep(interval)
            try:
                for fname in os.listdir(UPLOAD_FOLDER):
                    if fname in ACTIVE_FILES:
                        continue
                    fpath = os.path.join(UPLOAD_FOLDER, fname)
                    if os.path.isfile(fpath):
                        try:
                            os.remove(fpath)
                            print(f"[PERIODIC] Removed {fpath}")
                        except Exception as e:
                            print(f"[ERROR] Periodic cleanup failed: {e}")
            except Exception as e:
                print(f"[ERROR] Cleanup loop failed: {e}")
    threading.Thread(target=cleanup, daemon=True).start()

@app.route('/')
def home():
    return jsonify({'status': 'PDF to Excel (improved) API is running ✅'}), 200

# Ping route to keep API alive
@app.route('/ping')
def ping():
    return jsonify({'ping': 'pong'}), 200

@app.route('/convert', methods=['POST'])
def convert_pdf_to_excel():
    file = request.files.get('file')
    if not file:
        return 'No file uploaded', 400

    filename = secure_filename(file.filename)
    base_name = os.path.splitext(filename)[0]
    file_id = str(uuid.uuid4())
    input_pdf = os.path.join(UPLOAD_FOLDER, f"{file_id}_{filename}")
    output_excel = os.path.join(UPLOAD_FOLDER, f"{file_id}_{base_name}_Tools_Subidha.xlsx")
    file.save(input_pdf)

    ACTIVE_FILES.add(os.path.basename(input_pdf))
    ACTIVE_FILES.add(os.path.basename(output_excel))

    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        row_index = 1
        content_found = False

        border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        bold_font = Font(bold=True)

        def style_row(cells, is_header=False):
            for cell in cells:
                cell.border = border
                if is_header:
                    cell.font = bold_font

        # Try Camelot first
        try:
            camelot_tables = camelot.read_pdf(input_pdf, pages='all', flavor='stream')
            if camelot_tables.n > 0:
                for table in camelot_tables:
                    df = pd.DataFrame(table.df.values.tolist())
                    for r in dataframe_to_rows(df, index=False, header=True):
                        ws.append(r)
                        style_row(ws[row_index], is_header=(row_index == 1))
                        row_index += 1
                content_found = True
        except Exception as e:
            print(f"[WARN] Camelot failed: {e}")

        # Fallback to pdfplumber
        if not content_found:
            with pdfplumber.open(input_pdf) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            df = pd.DataFrame(table)
                            for r in dataframe_to_rows(df, index=False, header=True):
                                clean_row = [str(cell).encode('utf-8').decode('utf-8') if cell else '' for cell in r]
                                ws.append(clean_row)
                                style_row(ws[row_index], is_header=(row_index == 1))
                                row_index += 1
                        content_found = True
                    else:
                        text = page.extract_text()
                        if text:
                            for line in text.splitlines():
                                ws.append([line])
                                style_row(ws[row_index])
                                row_index += 1
                            content_found = True

        if not content_found:
            ACTIVE_FILES.discard(os.path.basename(input_pdf))
            ACTIVE_FILES.discard(os.path.basename(output_excel))
            delete_file_later(input_pdf)
            return 'No extractable content found in PDF.', 400

        # Auto-adjust column widths
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 2

        wb.save(output_excel)

    except Exception as e:
        print("❌ Error:", e)
        ACTIVE_FILES.discard(os.path.basename(input_pdf))
        ACTIVE_FILES.discard(os.path.basename(output_excel))
        return 'Conversion failed.', 500

    ACTIVE_FILES.discard(os.path.basename(input_pdf))
    ACTIVE_FILES.discard(os.path.basename(output_excel))
    delete_file_later(input_pdf)
    delete_file_later(output_excel)

    return send_file(
        output_excel,
        as_attachment=True,
        download_name=os.path.basename(output_excel),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    periodic_cleanup(interval=180)  # start background cleanup
    app.run(debug=False)
