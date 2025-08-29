from werkzeug.utils import secure_filename
from flask import Flask, render_template, request, send_file
import os
from pdf2docx import Converter
from PIL import Image
from zipfile import ZipFile
from docx import Document
from docx2pdf import convert
from fpdf import FPDF
import pandas as pd
from pydub import AudioSegment
from moviepy.editor import VideoFileClip

# Initialize Flask App
app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
CONVERTED_FOLDER = "converted"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["CONVERTED_FOLDER"] = CONVERTED_FOLDER


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("files")
        conversion_type = request.form.get("conversion_type")

        if not files or not conversion_type:
            return "No file or conversion type selected", 400

        converted_files = []

        for f in files:
            filename = secure_filename(f.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            f.save(filepath)

            try:
                # PDF → Word
                if conversion_type == "pdf_to_word" and filename.lower().endswith(".pdf"):
                    output_file = os.path.join(CONVERTED_FOLDER, filename.replace(".pdf", ".docx"))
                    cv = Converter(filepath)
                    cv.convert(output_file)
                    cv.close()
                    converted_files.append(output_file)

                # Image → PDF
                elif conversion_type == "image_to_pdf" and filename.lower().endswith((".jpg", ".jpeg", ".png")):
                    output_file = os.path.join(CONVERTED_FOLDER, filename.rsplit(".", 1)[0] + ".pdf")
                    image = Image.open(filepath).convert("RGB")
                    image.save(output_file, "PDF")
                    converted_files.append(output_file)

                # Word → PDF (using docx2pdf)
                elif conversion_type == "word_to_pdf" and filename.lower().endswith(".docx"):
                    output_file = os.path.join(CONVERTED_FOLDER, filename.replace(".docx", ".pdf"))
                    convert(filepath, output_file)
                    converted_files.append(output_file)

                # Excel → CSV
                elif conversion_type == "excel_to_csv" and filename.lower().endswith((".xls", ".xlsx")):
                    output_file = os.path.join(CONVERTED_FOLDER, filename.rsplit(".", 1)[0] + ".csv")
                    df = pd.read_excel(filepath)
                    df.to_csv(output_file, index=False)
                    converted_files.append(output_file)

                # MP3 → WAV
                elif conversion_type == "mp3_to_wav" and filename.lower().endswith(".mp3"):
                    output_file = os.path.join(CONVERTED_FOLDER, filename.replace(".mp3", ".wav"))
                    audio = AudioSegment.from_mp3(filepath)
                    audio.export(output_file, format="wav")
                    converted_files.append(output_file)

                # WAV → MP3
                elif conversion_type == "wav_to_mp3" and filename.lower().endswith(".wav"):
                    output_file = os.path.join(CONVERTED_FOLDER, filename.replace(".wav", ".mp3"))
                    audio = AudioSegment.from_wav(filepath)
                    audio.export(output_file, format="mp3")
                    converted_files.append(output_file)

                # MP4 → AVI
                elif conversion_type == "mp4_to_avi" and filename.lower().endswith(".mp4"):
                    output_file = os.path.join(CONVERTED_FOLDER, filename.replace(".mp4", ".avi"))
                    clip = VideoFileClip(filepath)
                    clip.write_videofile(output_file, codec="libx264")
                    clip.close()
                    converted_files.append(output_file)

                # AVI → MP4
                elif conversion_type == "avi_to_mp4" and filename.lower().endswith(".avi"):
                    output_file = os.path.join(CONVERTED_FOLDER, filename.replace(".avi", ".mp4"))
                    clip = VideoFileClip(filepath)
                    clip.write_videofile(output_file, codec="libx264")
                    clip.close()
                    converted_files.append(output_file)
                
                # MP4 → MP3
                elif conversion_type == "mp4_to_mp3" and filename.lower().endswith(".mp4"):
                    output_file = os.path.join(CONVERTED_FOLDER, filename.replace(".mp4", ".mp3"))
                    clip = VideoFileClip(filepath)
                    clip.audio.write_audiofile(output_file)   # Extract audio only
                    clip.close()
                    converted_files.append(output_file)


                else:
                    return f"Unsupported conversion for {filename}", 400

            except Exception as e:
                return f"Error converting {filename}: {str(e)}", 500

        # If multiple files, zip them
        if len(converted_files) > 1:
            zip_path = os.path.join(CONVERTED_FOLDER, "converted_files.zip")
            with ZipFile(zip_path, "w") as zipf:
                for f in converted_files:
                    zipf.write(f, os.path.basename(f))
            return send_file(zip_path, as_attachment=True)

        # Single file
        return send_file(converted_files[0], as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
