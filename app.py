from flask import Flask, render_template, request, send_file
import os
from pdf2docx import Converter
from PIL import Image
from zipfile import ZipFile
from docx import Document
from fpdf import FPDF
import pandas as pd

# ------------------------
# Initialize Flask App
# ------------------------
app = Flask(__name__)

# ------------------------
# Upload folder setup
# ------------------------
UPLOAD_FOLDER = "uploads"
# Automatically create uploads folder at runtime
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# ------------------------
# Routes
# ------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("files")
        conversion_type = request.form.get("conversion_type")
        converted_files = []

        for uploaded_file in files:
            file_path = os.path.join(app.config["UPLOAD_FOLDER"], uploaded_file.filename)
            uploaded_file.save(file_path)

            # --- TEXT → UPPERCASE ---
            if conversion_type == "text_uppercase":
                with open(file_path, "r", encoding="utf-8") as f:
                    content = f.read()
                output_file = file_path.replace(".txt", "_uppercase.txt")
                with open(output_file, "w", encoding="utf-8") as f:
                    f.write(content.upper())

            # --- PDF → WORD ---
            elif conversion_type == "pdf_to_word":
                output_file = file_path.replace(".pdf", ".docx")
                cv = Converter(file_path)
                cv.convert(output_file)
                cv.close()

            # --- IMAGE → PDF ---
            elif conversion_type == "image_to_pdf":
                image = Image.open(file_path)
                output_file = file_path.rsplit(".", 1)[0] + ".pdf"
                image.save(output_file, "PDF")

            # --- WORD → PDF ---
            elif conversion_type == "word_to_pdf":
                doc = Document(file_path)
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                for para in doc.paragraphs:
                    pdf.multi_cell(0, 10, para.text)
                output_file = file_path.replace(".docx", ".pdf")
                pdf.output(output_file)

            # --- EXCEL → CSV ---
            elif conversion_type == "excel_to_csv":
                df = pd.read_excel(file_path)
                output_file = file_path.replace(".xlsx", ".csv")
                df.to_csv(output_file, index=False)

            # --- AUDIO MP3 ↔ WAV ---
            else:
                return "Invalid conversion type"

            converted_files.append(output_file)

        # ------------------------
        # If multiple files → zip
        # ------------------------
        if len(converted_files) > 1:
            zip_path = os.path.join(app.config["UPLOAD_FOLDER"], "converted_files.zip")
            with ZipFile(zip_path, "w") as zipf:
                for f in converted_files:
                    zipf.write(f, os.path.basename(f))
            return send_file(zip_path, as_attachment=True)
        else:
            return send_file(converted_files[0], as_attachment=True)

    return render_template("index.html")


# ------------------------
# Run App
# ------------------------
if __name__ == "__main__":
    # Use 0.0.0.0 for deployment
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
