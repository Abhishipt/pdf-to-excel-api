from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import os
import uuid
import threading
import time
import openpyxl
from openpyxl.styles import Border, Side, Font, PatternFill

# Unicode-friendly text extraction
# pdfminer.six
from pdfminer.high_level import extract_text
# PyMuPDF (fallback for positional words)
import fitz

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ACTIVE_FILES = set()

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
    return jsonify({'status': 'PDF to Excel (Unicode-safe) API is running ✅'}), 200

@app.route('/ping')
def ping():
    return jsonify({'ping': 'pong'}), 200

def unicode_text_pdfminer(pdf_path):
    try:
        # pdfminer.six maps CID fonts to Unicode better than many libraries
        text = extract_text(pdf_path)  # returns full doc text
        return text or ""
    except Exception as e:
        print(f"[WARN] pdfminer failed: {e}")
        return ""

def unicode_lines_pymupdf(pdf_path):
    # Fallback: reconstruct lines from word positions
    try:
        doc = fitz.open(pdf_path)
        lines = []
        for page in doc:
            # Extract words: (x0, y0, x1, y1, "text", block_no, line_no, word_no)
            words = page.get_text("words")
            # Group by line_no
            from collections import defaultdict
            line_map = defaultdict(list)
            for (x0, y0, x1, y1, text, block, lno, wno) in words:
                line_map[(block, lno)].append((x0, text))
            # Sort by x0 and join
            for _, items in sorted(line_map.items(), key=lambda kv: (kv[0][0], kv[0][1])):
                items.sort(key=lambda t: t[0])
                line = " ".join(t[1] for t in items).strip()
                if line:
                    lines.append(line)
        return "\n".join(lines)
    except Exception as e:
        print(f"[WARN] PyMuPDF fallback failed: {e}")
        return ""

def parse_line_to_cells(line):
    # Prefer colon split
    if ":" in line:
        left, right = line.split(":", 1)
        return [left.strip(), right.strip()]
    # If spaced columns (≥3 spaces), split
    if "   " in line:
        parts = [p.strip() for p in line.split("   ") if p.strip()]
        if len(parts) > 1:
            return parts
    # Otherwise single cell
    return [line.strip()]

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
        current_row = 1

        border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        bold_font = Font(bold=True)
        header_fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")

        def style_row(row_idx, is_header=False):
            for cell in ws[row_idx]:
                cell.border = border
                if is_header:
                    cell.font = bold_font
                    cell.fill = header_fill

        # 1) Try pdfminer for clean Unicode
        text = unicode_text_pdfminer(input_pdf)

        # 2) Fallback to PyMuPDF word-reconstructed lines if empty
        if not text.strip():
            text = unicode_lines_pymupdf(input_pdf)

        if not text.strip():
            ACTIVE_FILES.discard(os.path.basename(input_pdf))
            ACTIVE_FILES.discard(os.path.basename(output_excel))
            delete_file_later(input_pdf)
            return 'No extractable content found in PDF.', 400

        # Heuristic: tag lines that look like section headers for styling
        header_keywords = [
            "Mutation Details", "Correction slip Generation",
            "Applicant Details", "Vendee Details", "Vendor Details",
            "Plot Details", "Document uploaded", "View of Mutation"
        ]
        devanagari_header_signals = ["विवरण", "दस्तावेज", "विक्रेता", "क्रेता", "विवरण", "हalka", "मौजा", "खाता", "क्षेत्रफल"]

        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                continue

            cells = parse_line_to_cells(line)
            ws.append(cells)
            # Style header-like lines (English keywords or certain Devanagari signals)
            is_header = any(k.lower() in line.lower() for k in header_keywords) or \
                        any(k in line for k in devanagari_header_signals)
            style_row(current_row, is_header=is_header)
            current_row += 1

        # Auto-fit columns
        for col in ws.columns:
            max_len = 0
            letter = col[0].column_letter
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[letter].width = min(max_len + 2, 60)

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
    periodic_cleanup(interval=180)
    app.run(debug=False)
