from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import uuid
import threading
import time
import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Border, Side, Font, PatternFill

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ACTIVE_FILES = set()

# Auto-delete files after delay
def delete_file_later(path, delay=60):
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
    return jsonify({'status': 'PDF to Excel (PyMuPDF) API is running ✅'}), 200

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

        border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        bold_font = Font(bold=True)
        header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")

        def style_row(cells, is_header=False):
            for cell in cells:
                cell.border = border
                if is_header:
                    cell.font = bold_font
                    cell.fill = header_fill

        # Extract text with PyMuPDF
        doc = fitz.open(input_pdf)
        for page in doc:
            text = page.get_text("text")
            if not text:
                continue

            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue

                # Split on colon or large spacing
                if ":" in line:
                    parts = [p.strip() for p in line.split(":", 1)]
                    ws.append(parts)
                elif "   " in line:  # multiple spaces
                    parts = [p.strip() for p in line.split("   ") if p.strip()]
                    ws.append(parts)
                else:
                    ws.append([line])

                style_row(ws[row_index])
                row_index += 1

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
