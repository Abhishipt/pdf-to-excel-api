from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import uuid
import threading
import time
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
    return jsonify({'status': 'PDF to Excel API is running ✅'}), 200

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

        with pdfplumber.open(input_pdf) as pdf:
            text_found = False
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    for row in table:
                        ws.append(row)
                        row_index += 1
                    text_found = True
                elif page.extract_text():
                    lines = page.extract_text().splitlines()
                    for line in lines:
                        ws.append([line])
                        row_index += 1
                    text_found = True

        if not text_found:
            delete_file_later(input_pdf)
            return 'No extractable text or tables found. Cannot convert scanned/image-only PDFs.', 400

        wb.save(output_excel)

    except Exception as e:
        print("❌ Error:", e)
        return 'Conversion failed.', 500

    delete_file_later(input_pdf)
    delete_file_later(output_excel)

    return send_file(output_excel, as_attachment=True, download_name='converted.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=False)
    
