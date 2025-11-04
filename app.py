import os
import io
import zipfile
import shutil
import tempfile
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import pandas as pd

# ReportLab imports
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Arabic shaping imports
import arabic_reshaper
from bidi.algorithm import get_display

# Configuration
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}
MAX_CONTENT_LENGTH = 200 * 1024 * 1024  # 200MB

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "change-this-secret")
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# Font setup: put a good Arabic TTF in ./fonts/ (Amiri recommended)
FONT_DIR = os.path.join(os.path.dirname(__file__), "fonts")
ARABIC_FONT = os.path.join(FONT_DIR, "Amiri-Regular.ttf")  # change if needed
ARABIC_FONT_NAME = "Amiri"

# Register font if exists
if os.path.exists(ARABIC_FONT):
    pdfmetrics.registerFont(TTFont(ARABIC_FONT_NAME, ARABIC_FONT))
else:
    # If font missing, use built-in DejaVuSans (may not shape Arabic correctly)
    # You should add a proper Arabic TTF to ./fonts/ for best results.
    ARABIC_FONT_NAME = "Helvetica"

# Helpers
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def sanitize_filename(name):
    # Keep the original name but remove filesystem-illegal characters
    # preserve Arabic and spaces
    invalid = r'\/:*?"<>|'
    return "".join(c for c in name if c not in invalid).strip() or "sheet"

def reshape_rtl(text):
    """
    For Arabic/RTL language: reshape and bidi-reorder text for ReportLab rendering.
    If text is not str, convert to str.
    """
    if text is None:
        return ""
    s = str(text)
    # quick check: if there are Arabic characters, apply reshaper + bidi
    # Arabic unicode block: \u0600-\u06FF, extended: \u0750-\u077F, \u08A0-\u08FF
    if any('\u0600' <= ch <= '\u06FF' or '\u0750' <= ch <= '\u077F' or '\u08A0' <= ch <= '\u08FF' for ch in s):
        reshaped = arabic_reshaper.reshape(s)
        bidi_text = get_display(reshaped)
        return bidi_text
    else:
        # For non-Arabic, still return string (left-to-right)
        return s

def df_to_table_data(df):
    """
    Convert pandas DataFrame to a 2D list suitable for ReportLab Table
    Apply reshape_rtl to each cell, and convert NaN to empty string
    """
    # header
    headers = [reshape_rtl(col) for col in df.columns.tolist()]
    rows = []
    for _, r in df.iterrows():
        row = []
        for cell in r.tolist():
            if pd.isna(cell):
                row.append("")
            else:
                row.append(reshape_rtl(cell))
        rows.append(row)
    data = [headers] + rows
    return data

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/convert", methods=["POST"])
def convert():
    if 'excel' not in request.files:
        flash("No file part", "danger")
        return redirect(url_for('index'))

    file = request.files['excel']
    if file.filename == '':
        flash("No selected file", "danger")
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash("Unsupported file type. Upload .xls or .xlsx", "danger")
        return redirect(url_for('index'))

    filename = secure_filename(file.filename)

    workdir = tempfile.mkdtemp(prefix="xls2pdf_")
    try:
        input_path = os.path.join(workdir, filename)
        file.save(input_path)

        # read sheet names
        try:
            xls = pd.ExcelFile(input_path)
            sheet_names = xls.sheet_names
        except Exception as e:
            flash(f"Failed to read Excel file: {str(e)}", "danger")
            return redirect(url_for('index'))

        pdf_paths = []

        for sheet_name in sheet_names:
            # read sheet into df
            try:
                # openpyxl for xlsx, xlrd fallback for xls if needed
                engine = 'openpyxl' if filename.lower().endswith('xlsx') else None
                df = pd.read_excel(input_path, sheet_name=sheet_name, engine=engine)
            except Exception:
                df = pd.read_excel(input_path, sheet_name=sheet_name)

            # skip completely empty sheets
            if df.dropna(how='all').empty:
                continue

            # convert df -> table data (with reshaped Arabic text)
            data = df_to_table_data(df)

            # create PDF file named exactly as sheet (Option A)
            safe_name = sanitize_filename(sheet_name)
            pdf_filename = f"{safe_name}.pdf"
            pdf_path = os.path.join(workdir, pdf_filename)

            # Create PDF document
            # Using A4 landscape if many columns, else portrait
            page_size = A4
            try:
                # choose landscape if >6 columns
                if len(data[0]) > 6:
                    page_size = landscape(A4)
            except Exception:
                page_size = A4

            doc = SimpleDocTemplate(pdf_path, pagesize=page_size, rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=20)

            # styles
            title_style = ParagraphStyle(
                name="Title",
                fontName=ARABIC_FONT_NAME,
                fontSize=14,
                alignment=TA_CENTER,
                spaceAfter=8
            )
            cell_style = ParagraphStyle(
                name="Cell",
                fontName=ARABIC_FONT_NAME,
                fontSize=10,
                alignment=TA_RIGHT  # align right for RTL
            )

            # create table with Paragraphs so fonts & alignment apply
            table_data = []
            for row_index, row in enumerate(data):
                row_cells = []
                for cell in row:
                    # use Paragraph so long text wraps
                    p = Paragraph(cell.replace("\n", "<br/>"), cell_style)
                    row_cells.append(p)
                table_data.append(row_cells)

            # table styling
            tbl = Table(table_data, repeatRows=1)
            tbl.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2c3e50")),  # header bg
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('INNERGRID', (0,0), (-1,-1), 0.25, colors.HexColor("#444444")),
                ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor("#444444")),
                ('FONTNAME', (0,0), (-1,-1), ARABIC_FONT_NAME),
                ('FONTSIZE', (0,0), (-1,-1), 9),
                ('LEFTPADDING', (0,0), (-1,-1), 6),
                ('RIGHTPADDING', (0,0), (-1,-1), 6),
                ('TOPPADDING', (0,0), (-1,-1), 4),
                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ]))

            elements = []
            # title
            display_title = reshape_rtl(sheet_name)
            elements.append(Paragraph(display_title, title_style))
            elements.append(Spacer(1, 6))
            elements.append(tbl)

            # build PDF
            doc.build(elements)

            pdf_paths.append(pdf_path)

        if not pdf_paths:
            flash("No non-empty sheets found in the uploaded Excel file.", "warning")
            return redirect(url_for('index'))

        # zip the PDFs into memory
        memory_file = io.BytesIO()
        with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
            for p in pdf_paths:
                zf.write(p, arcname=os.path.basename(p))
        memory_file.seek(0)

        zip_name = os.path.splitext(filename)[0] + "_pdfs.zip"
        return send_file(memory_file, mimetype='application/zip',
                         as_attachment=True, download_name=zip_name)
    finally:
        # cleanup
        try:
            shutil.rmtree(workdir)
        except Exception:
            pass

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
