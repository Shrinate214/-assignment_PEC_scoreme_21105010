import os
import pdfplumber
import pytesseract
from pytesseract import image_to_string
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
import pandas as pd
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from openpyxl import Workbook
import fitz 
import cv2
import numpy as np
import pdfplumber
import pytesseract
from PIL import Image
import pandas as pd
from openpyxl import Workbook
import re

import pandas as pd
from pdf2image import convert_from_path
poppler_path = r"C:\Users\akshat shrinate\Release-24.08.0-0\poppler-24.08.0\Library\bin"


app = Flask(__name__)
# Specify your Poppler path
app.secret_key = "Akshat123"
import os
import pdfplumber
import pandas as pd
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from werkzeug.utils import secure_filename
from pytesseract import image_to_string



# Configure upload & output folders
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER
app.config["ALLOWED_EXTENSIONS"] = {"pdf"}


# -----------------------------------------------
#  Function: Extract Text and Tables from PDF
# -----------------------------------------------
import fitz
import pytesseract
from PIL import Image

from PIL import Image

def extract_tables_from_pdf(pdf_path):
    pdf_document = fitz.open(pdf_path)  # Open PDF
    tables = []
    last_headers = None

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        words = page.get_text("words")  # Extract words (text)

        #If no text detected, OCR is used to extract from images
        if not words:
            pix = page.get_pixmap()  # Convert page to image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            page_text = pytesseract.image_to_string(img)  # Perform OCR
            words = [(0, 0, 0, 0, word, 0, 0, 0) for word in page_text.split("\n") if word.strip()]
        
        lines = {}
        for word in words:
            x0, y0, x1, y1, word_text, _, line_no, _ = word
            if y0 not in lines:
                lines[y0] = []
            lines[y0].append((x0, word_text))

        sorted_lines = sorted(lines.items(), key=lambda x: x[0])

        table = []
        non_table_text = []  # To store paragraphs

        for _, line in sorted_lines:
            sorted_line = sorted(line, key=lambda x: x[0])
            row = [word_text for _, word_text in sorted_line]

            # If first row looks like a header, store it such that tables can be mergef
            if last_headers is None:
                last_headers = row  # Treat first row as header
            elif row == last_headers:
                continue  # Ignore repeated headers

            # Ignore textual data by checking for spaces greater tha n" "
            if any(" " in cell for cell in row):  
                non_table_text.append(" ".join(row))
                continue

            # store table rows, skipping headers
            if last_headers and len(row) == len(last_headers):  
                table.append(row)
            elif not last_headers:  
                last_headers = row  
                # Assign header dynamically

        if table:
            tables.append(table)

    return tables



# --------------------------------------------------
#  Function: Save Tables to Excel
# --------------------------------------------------
def save_tables_to_excel(tables, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Table_1"

    for i, table in enumerate(tables):
        df = pd.DataFrame(table)
        if i == 0:
            for row in df.values:
                ws.append(row.tolist())  # Append row-wise
        else:
            sheet = wb.create_sheet(title=f"Table_{i+1}")
            for row in df.values:
                sheet.append(row.tolist())

    wb.save(output_path)

# ----------------------------------------------
#  Flask App: File Upload & Processing
# ----------------------------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["pdf"]
        if not file:
            return render_template("index.html", error="No file selected.")

        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(pdf_path)

        output_filename = filename.replace(".pdf", ".xlsx")
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        tables = extract_tables_from_pdf(pdf_path)  # ✅ Extract Tables
        if not tables:
            return render_template("index.html", error="No tables found.")

        save_tables_to_excel(tables, output_path)  # ✅ Save to Excel

        return render_template("index.html", filename=output_filename)

    return render_template("index.html", filename=None)

# ----------------------------------------------
#  Flask App: File Download Route
# ----------------------------------------------
@app.route("/download/<filename>")
def download(filename):
    file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)

