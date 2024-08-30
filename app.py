from flask import Flask, render_template, request, send_file
from pdf2docx import Converter
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'

# Create upload and output directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'pdf_file' not in request.files:
        return 'No file part'
    
    pdf_file = request.files['pdf_file']
    if pdf_file.filename == '':
        return 'No selected file'
    
    pdf_path = os.path.join(UPLOAD_FOLDER, pdf_file.filename)
    pdf_file.save(pdf_path)
    
    output_filename = f"{os.path.splitext(pdf_file.filename)[0]}.docx"
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)
    
    # Convert PDF to DOCX
    cv = Converter(pdf_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()
    
    # Provide download link for the converted file
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
