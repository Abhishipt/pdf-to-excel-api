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

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def delete_file_later(path, delay=300):
    def remove():
        time.sleep(delay)
        if os.path.exists(path):
            os.remove(path)
    threading.Thread(target=remove).start()

@app.route('/')
def home():
    return jsonify({'status': 'PDF to Excel (Preserve All + Header Restore) API ✅'}), 200

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

        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
        bold_font = Font(bold=True)

        def style_row(cells, is_header=False):
            for cell in cells:
                cell.border = border
                if is_header:
                    cell.font = bold_font

        def write_table_to_ws(table_data):
            nonlocal row_index
            for i, row in enumerate(table_data):
                clean_row = [cell if cell is not None else '' for cell in row]
                ws.append(clean_row)
                style_row(ws[row_index], is_header=(i == 0))
                row_index += 1

        # Try Camelot first
        camelot_tables = camelot.read_pdf(input_pdf, pages='all', flavor='stream')
        if camelot_tables.n > 0:
            for table in camelot_tables:
                write_table_to_ws(table.df.values.tolist())
            content_found = True

        # Fallback to pdfplumber if Camelot fails
        if not content_found:
            with pdfplumber.open(input_pdf) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            write_table_to_ws(table)
                        content_found = True
                    else:
                        # Try using the first text line as header
                        lines = page.extract_text().splitlines() if page.extract_text() else []
                        if lines:
                            ws.append([cell.strip() for cell in lines[0].split()])  # use as header
                            style_row(ws[row_index], is_header=True)
                            row_index += 1
                            for line in lines[1:]:
                                ws.append([line.strip()])
                                style_row(ws[row_index])
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
