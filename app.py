from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import uuid
import threading
import time
import fitz  # PyMuPDF for rendering Hindi text as image
import pdfplumber
import openpyxl
from openpyxl.styles import Border, Side, Font
from openpyxl.drawing.image import Image as ExcelImage

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Auto-delete after 1 minute
def delete_file_later(path, delay=60):
    def remove():
        time.sleep(delay)
        if os.path.exists(path):
            os.remove(path)
    threading.Thread(target=remove).start()

@app.route('/')
def home():
    return jsonify({'status': 'PDF to Excel (text + image fallback + header fix) API is running ✅'}), 200

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
        max_cols = 0

        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
        bold_font = Font(bold=True)

        def style_row(cells, is_header=False):
            for cell in cells:
                cell.border = border
                if is_header:
                    cell.font = bold_font

        def normalize_and_add_row(text_line, is_header=False):
            nonlocal row_index, max_cols
            cols = text_line.split()
            max_cols = max(max_cols, len(cols))
            ws.append(cols)
            style_row(ws[row_index], is_header=is_header)
            row_index += 1

        text_found = False
        with pdfplumber.open(input_pdf) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    lines = text.splitlines()
                    for i, line in enumerate(lines):
                        normalize_and_add_row(line, is_header=(i == 0 and row_index == 1))
                    text_found = True

        if not text_found:
            doc = fitz.open(input_pdf)
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                image_path = os.path.join(UPLOAD_FOLDER, f"page_{page_num}_{file_id}.png")
                pix = page.get_pixmap(dpi=150)
                pix.save(image_path)

                img = ExcelImage(image_path)
                ws.add_image(img, f"A{row_index}")
                row_index += 25

                delete_file_later(image_path)

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            row_values = [cell.value if cell.value is not None else '' for cell in row]
            if len(row_values) < max_cols:
                for _ in range(max_cols - len(row_values)):
                    ws.cell(row=row[0].row, column=len(row) + 1, value='')

        wb.save(output_excel)

    except Exception as e:
        print("❌ Error:", e)
        return 'Conversion failed.', 500

    delete_file_later(input_pdf)
    delete_file_later(output_excel)

    return send_file(output_excel, as_attachment=True, download_name='converted.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(debug=False)
