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
from openpyxl.styles import Border, Side, Font, Alignment
from datetime import datetime, timedelta

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Cleanup old files older than 3 minutes
def periodic_cleanup():
    now = time.time()
    cutoff = now - 180  # 3 minutes
    for fname in os.listdir(UPLOAD_FOLDER):
        path = os.path.join(UPLOAD_FOLDER, fname)
        if os.path.isfile(path) and os.path.getmtime(path) < cutoff:
            try:
                os.remove(path)
            except Exception:
                pass

# Start background cleanup every time a conversion is triggered
def schedule_cleanup():
    threading.Thread(target=periodic_cleanup).start()

@app.route('/')
def home():
    return jsonify({'status': '✅ PDF to Excel Converter is running'}), 200

@app.route('/convert', methods=['POST'])
def convert_pdf_to_excel():
    file = request.files.get('file')
    if not file:
        return 'No file uploaded', 400

    filename = secure_filename(file.filename)
    file_base = os.path.splitext(filename)[0]
    file_id = str(uuid.uuid4())
    input_pdf = os.path.join(UPLOAD_FOLDER, f"{file_id}_{filename}")
    output_excel = os.path.join(UPLOAD_FOLDER, f"{file_base}_Tools_Subidha.xlsx")

    file.save(input_pdf)

    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Converted"

        row_index = 1
        content_found = False

        border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        bold_font = Font(bold=True)
        alignment = Alignment(wrap_text=True, vertical="top")

        def style_cells(cells, is_header=False):
            for cell in cells:
                cell.border = border
                cell.alignment = alignment
                if is_header:
                    cell.font = bold_font

        # Try with Camelot
        try:
            tables = camelot.read_pdf(input_pdf, pages='all', flavor='stream')
            if tables.n > 0:
                for table in tables:
                    data = table.df.values.tolist()
                    for i, row in enumerate(data):
                        clean_row = [cell.strip() for cell in row]
                        ws.append(clean_row)
                        style_cells(ws[row_index], is_header=(i == 0))
                        row_index += 1
                content_found = True
        except Exception:
            pass  # Fallback to pdfplumber if Camelot fails

        # Fallback: Try with pdfplumber
        if not content_found:
            with pdfplumber.open(input_pdf) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    if tables:
                        for table in tables:
                            for i, row in enumerate(table):
                                clean_row = [cell.strip() if cell else '' for cell in row]
                                ws.append(clean_row)
                                style_cells(ws[row_index], is_header=(i == 0))
                                row_index += 1
                        content_found = True
                    else:
                        # Extract lines of text
                        text = page.extract_text()
                        if text:
                            for line in text.splitlines():
                                ws.append([line])
                                style_cells(ws[row_index])
                                row_index += 1
                            content_found = True

        if not content_found:
            os.remove(input_pdf)
            return 'No extractable content found in PDF.', 400

        # Auto-adjust column width
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            adjusted_width = min((max_length + 2), 60)
            ws.column_dimensions[col_letter].width = adjusted_width

        wb.save(output_excel)

    except Exception as e:
        print("❌ Conversion error:", e)
        return 'Conversion failed due to an internal error.', 500

    # Schedule deletion
    schedule_cleanup()

    return send_file(
        output_excel,
        as_attachment=True,
        download_name=os.path.basename(output_excel),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=False)
