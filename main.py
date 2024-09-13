from flask import Flask, render_template, request, send_file, url_for
from PIL import Image
from docx2pdf import convert
import io
import os
import uuid
import tempfile
import fitz
from pdf2docx import Converter
import zipfile

app = Flask(__name__)

file_storage = {}

def jpg_to_png(image_file, original_filename):
    img = Image.open(image_file)
    img_io = io.BytesIO()
    img.save(img_io, format='PNG')
    img_io.seek(0)
    base_name = os.path.splitext(original_filename)[0]
    return img_io, f'{base_name}.png'

def word_to_pdf(word_file, original_filename):
    temp_pdf_path = tempfile.mktemp(suffix='.pdf')
    temp_word_path = tempfile.mktemp(suffix='.docx')
    word_file.save(temp_word_path)
    convert(temp_word_path, temp_pdf_path)
    with open(temp_pdf_path, 'rb') as f:
        pdf_io = io.BytesIO(f.read())
    base_name = os.path.splitext(original_filename)[0]
    return pdf_io, f'{base_name}.pdf'

def png_to_jpg(image_file, original_filename):
    img = Image.open(image_file)
    img_io = io.BytesIO()
    img.convert('RGB').save(img_io, format='JPEG')
    img_io.seek(0)
    base_name = os.path.splitext(original_filename)[0]
    return img_io, f'{base_name}.jpg'

def pdf_to_word(pdf_file, original_filename):
    temp_pdf_path = tempfile.mktemp(suffix='.pdf')
    with open(temp_pdf_path, 'wb') as f:
        f.write(pdf_file.read())
    temp_word_path = tempfile.mktemp(suffix='.docx')
    cv = Converter(temp_pdf_path)
    try:
        cv.convert(temp_word_path, start=0, end=None)
        cv.close()
        with open(temp_word_path, 'rb') as f:
            docx_io = io.BytesIO(f.read())
        base_name = os.path.splitext(original_filename)[0]
        return docx_io, f'{base_name}.docx'
    finally:
        try:
            if os.path.exists(temp_pdf_path):
                os.remove(temp_pdf_path)
            if os.path.exists(temp_word_path):
                os.remove(temp_word_path)
        except PermissionError as e:
            print(f"Permission error: {e}")

def pdf_to_jpg(pdf_file, original_filename):
    pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
    zip_io = io.BytesIO()
    with zipfile.ZipFile(zip_io, mode='w') as zip_file:
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_io = io.BytesIO()
            img.save(img_io, format='JPEG')
            img_io.seek(0)
            zip_file.writestr(f'{os.path.splitext(original_filename)[0]}_page_{page_num + 1}.jpg', img_io.read())
    zip_io.seek(0)
    return zip_io, f'{os.path.splitext(original_filename)[0]}.zip'

def pdf_to_png(pdf_file, original_filename):
    pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
    zip_io = io.BytesIO()
    with zipfile.ZipFile(zip_io, mode='w') as zip_file:
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_io = io.BytesIO()
            img.save(img_io, format='PNG')
            img_io.seek(0)
            zip_file.writestr(f'{os.path.splitext(original_filename)[0]}_page_{page_num + 1}.png', img_io.read())
    zip_io.seek(0)
    return zip_io, f'{os.path.splitext(original_filename)[0]}.zip'

def jpg_to_pdf(image_file, original_filename):
    img = Image.open(image_file)
    pdf_io = io.BytesIO()
    img.convert('RGB').save(pdf_io, format='PDF')
    pdf_io.seek(0)
    base_name = os.path.splitext(original_filename)[0]
    return pdf_io, f'{base_name}.pdf'

def png_to_pdf(image_file, original_filename):
    img = Image.open(image_file)
    pdf_io = io.BytesIO()
    img.convert('RGB').save(pdf_io, format='PDF')
    pdf_io.seek(0)
    base_name = os.path.splitext(original_filename)[0]
    return pdf_io, f'{base_name}.pdf'

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        file_type = request.form['file_type']
        original_filename = file.filename
        file_id = str(uuid.uuid4())
        
        if file_type == 'jpg_to_png':
            img_io, filename = jpg_to_png(file, original_filename)
            file_storage[file_id] = (img_io, 'image/png', filename)

        elif file_type == 'word_to_pdf':
            pdf_io, filename = word_to_pdf(file, original_filename)
            file_storage[file_id] = (pdf_io, 'application/pdf', filename)

        elif file_type == 'jpg_to_pdf':
            pdf_io, filename = jpg_to_pdf(file, original_filename)
            file_storage[file_id] = (pdf_io, 'application/pdf', filename)

        elif file_type == 'png_to_jpg':
            img_io, filename = png_to_jpg(file, original_filename)
            file_storage[file_id] = (img_io, 'image/jpeg', filename)

        elif file_type == 'pdf_to_word':
            docx_io, filename = pdf_to_word(file, original_filename)
            file_storage[file_id] = (docx_io, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', filename)

        elif file_type == 'png_to_pdf':
            pdf_io, filename = png_to_pdf(file, original_filename)
            file_storage[file_id] = (pdf_io, 'application/pdf', filename)

        elif file_type == 'pdf_to_jpg':
            zip_io, filename = pdf_to_jpg(file, original_filename)
            file_storage[file_id] = (zip_io, 'application/zip', filename)

        elif file_type == 'pdf_to_png':
            zip_io, filename = pdf_to_png(file, original_filename)
            file_storage[file_id] = (zip_io, 'application/zip', filename)

        return render_template('index.html', download_url=url_for('download', file_id=file_id), show_button=True)
    
    return render_template('index.html', show_button=False)

@app.route('/download/<file_id>')
def download(file_id):
    file_data = file_storage.get(file_id)
    if file_data:
        file_io, mime_type, filename = file_data
        response = send_file(file_io, mimetype=mime_type, as_attachment=True, download_name=filename)
        response.headers["Refresh"] = "1; url='/'"
        return response
    return 'File not found', 404

if __name__ == '__main__':
    app.run(debug=True)
