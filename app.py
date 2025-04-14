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

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Auto-delete files after 5 minutes
def delete_file_later(path, delay=300):
    def remove():
        time.sleep(delay)
        if os.path.exists(path):
            os.remove(path)
    threading.Thread(target=remove).start()

@app.route('/')
def home():
    return jsonify({'status': 'PDF to Excel (Smart - Camelot + pdfplumber) API is running ✅'}), 200

@app.route('/convert', methods=['POST'])
def convert_pdf_to_excel():
    file = request.files.get('file')
    if not file:
        return 'No file uploaded', 400

    filename = secure_filename(file.filename)
    file_id = str(uuid.uuid4())
    input_pdf = os.path.join(UPLOAD_FOLDER, f"{file_id}_{filename}")
    output_excel = os.path.join(UPLOAD_FOLDER, f"{file_id}_converted.xlsx")
    file.save(input_pdf)

    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        row_index = 1
        content_found = False

        # First try extracting with Camelot (stream mode)
        camelot_tables = camelot.read_pdf(input_pdf, pages='all', flavor='stream')
        if camelot_tables.n > 0:
            for table in camelot_tables:
                for row in table.df.values.tolist():
                    ws.append(row)
                    row_index += 1
            content_found = True

        # If Camelot fails to extract anything, fallback to pdfplumber
        if not content_found:
            with pdfplumber.open(input_pdf) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            for row in table:
                                ws.append([cell if cell is not None else '' for cell in row])
                                row_index += 1
                        content_found = True
                    else:
                        text = page.extract_text()
                        if text:
                            for line in text.splitlines():
                                ws.append([line])
                                row_index += 1
                            content_found = True

        if not content_found:
            delete_file_later(input_pdf)
            return 'No extractable content found in PDF.', 400

        wb.save(output_excel)

    except Exception as e:
        print("❌ Error:", e)
        return 'Conversion failed.', 500

    delete_file_later(input_pdf)
    delete_file_later(output_excel)

    return send_file(output_excel, as_attachment=True, download_name='converted.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=False)
