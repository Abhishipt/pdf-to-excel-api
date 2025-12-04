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
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Clean up old files > 3 minutes
def cleanup_old_files(folder, age_limit=180):
    now = time.time()
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        if os.path.isfile(file_path):
            if now - os.path.getmtime(file_path) > age_limit:
                try:
                    os.remove(file_path)
                except Exception as e:
                    print(f"Failed to delete {file_path}: {e}")

# Schedule cleanup every time a conversion happens
def schedule_cleanup():
    threading.Thread(target=cleanup_old_files, args=(UPLOAD_FOLDER,)).start()

# Per-file cleanup after 1 minute
def delete_file_later(path, delay=60):
    def remove():
        time.sleep(delay)
        if os.path.exists(path):
            os.remove(path)
    threading.Thread(target=remove).start()

@app.route('/')
def home():
    return jsonify({'status': 'PDF to Excel (Hindi OCR Enhanced) API is running ✅'}), 200

@app.route('/convert', methods=['POST'])
def convert_pdf_to_excel():
    file = request.files.get('file')
    if not file:
        return 'No file uploaded', 400

    original_name = secure_filename(file.filename.rsplit('.', 1)[0])
    extension = file.filename.rsplit('.', 1)[-1]
    file_id = str(uuid.uuid4())

    input_pdf = os.path.join(UPLOAD_FOLDER, f"{file_id}_{original_name}.{extension}")
    output_excel = os.path.join(UPLOAD_FOLDER, f"{original_name}_Tools_Subidha.xlsx")
    file.save(input_pdf)

    schedule_cleanup()

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

        # Primary: Camelot
        try:
            camelot_tables = camelot.read_pdf(input_pdf, pages='all', flavor='stream')
            if camelot_tables.n > 0:
                for table in camelot_tables:
                    for i, row in enumerate(table.df.values.tolist()):
                        ws.append(row)
                        style_row(ws[row_index], is_header=(i == 0))
                        row_index += 1
                content_found = True
        except Exception as e:
            print("Camelot error:", e)

        # Fallback: pdfplumber
        if not content_found:
            with pdfplumber.open(input_pdf) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            for i, row in enumerate(table):
                                clean_row = [cell if cell else '' for cell in row]
                                ws.append(clean_row)
                                style_row(ws[row_index], is_header=(i == 0))
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

        # Final fallback: OCR Hindi
        if not content_found:
            images = convert_from_path(input_pdf)
            for img in images:
                text = pytesseract.image_to_string(img, lang='hin')
                if text.strip():
                    for line in text.strip().split('\n'):
                        if line.strip():
                            ws.append([line])
                            style_row(ws[row_index])
                            row_index += 1
                    content_found = True

        if not content_found:
            delete_file_later(input_pdf)
            return 'No readable content found.', 400

        wb.save(output_excel)

    except Exception as e:
        print("❌ Conversion error:", e)
        return 'Conversion failed.', 500

    delete_file_later(input_pdf)
    delete_file_later(output_excel)

    return send_file(
        output_excel,
        as_attachment=True,
        download_name=os.path.basename(output_excel),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=False)
