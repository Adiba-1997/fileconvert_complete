from flask import Flask, request, send_file, send_from_directory
from flask_cors import CORS
import os
import subprocess
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

# ✅ Serve index.html
@app.route('/')
def serve_index():
    return send_from_directory('.', 'index.html')

# ✅ Serve static files like script.js, style.css, etc.
@app.route('/<path:path>')
def serve_static_files(path):
    return send_from_directory('.', path)

# ✅ Word to PDF route
@app.route('/convert/word-to-pdf', methods=['POST'])
def convert_word_to_pdf():
    if 'file' not in request.files:
        return {'error': 'No file uploaded'}, 400

    file = request.files['file']
    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(input_path)

    output_filename = filename.rsplit('.', 1)[0] + '.pdf'
    output_path = os.path.join(CONVERTED_FOLDER, output_filename)

    try:
        subprocess.run([
            'libreoffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', CONVERTED_FOLDER,
            input_path
        ], check=True)
    except subprocess.CalledProcessError:
        return {'error': 'Conversion failed'}, 500

    return send_file(output_path, as_attachment=True)

# ✅ PDF to Word route
@app.route('/convert/pdf-to-word', methods=['POST'])
def convert_pdf_to_word():
    if 'file' not in request.files:
        return {'error': 'No file uploaded'}, 400

    file = request.files['file']
    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(input_path)

    output_filename = filename.rsplit('.', 1)[0] + '.docx'
    output_path = os.path.join(CONVERTED_FOLDER, output_filename)

    try:
        subprocess.run([
            'libreoffice',
            '--headless',
            '--convert-to', 'docx',
            '--outdir', CONVERTED_FOLDER,
            input_path
        ], check=True)
    except subprocess.CalledProcessError:
        return {'error': 'Conversion failed'}, 500

    return send_file(output_path, as_attachment=True)




