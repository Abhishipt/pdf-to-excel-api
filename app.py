from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import uuid
import threading
import time
import camelot

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
    return jsonify({'status': 'PDF to Excel (Camelot) API is running ✅'}), 200

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
        tables = camelot.read_pdf(input_pdf, pages='all', flavor='lattice')
        if tables.n == 0:
            delete_file_later(input_pdf)
            return 'No tables found in the PDF.', 400

        tables.export(output_excel, f='excel')

    except Exception as e:
        print("❌ Error:", e)
        return 'Conversion failed.', 500

    delete_file_later(input_pdf)
    delete_file_later(output_excel)

    return send_file(output_excel, as_attachment=True, download_name='converted.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=False)
