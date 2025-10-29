from flask import Flask, request, render_template, send_from_directory, redirect, url_for, flash, send_file, jsonify
import os
import uuid
from werkzeug.utils import secure_filename
from docx import Document as DocxDocument
import mammoth
from flask import Response
import zipfile

from word import generate_from_files


BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_BASE = os.path.join(BASE_DIR, "output")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_BASE, exist_ok=True)

ALLOWED_EXTENSIONS = {'.xlsx', '.xls', '.docx'}

app = Flask(__name__)
# Use environment variable for secret key in production
app.secret_key = os.environ.get('SECRET_KEY', 'change-this-to-a-secret')  # override via ENV


def allowed_file(filename):
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    if 'template_file' not in request.files or 'data_file' not in request.files:
        flash('Cần upload cả file template (.docx) và file dữ liệu (.xlsx)')
        return redirect(url_for('index'))

    template = request.files['template_file']
    data = request.files['data_file']

    if template.filename == '' or data.filename == '':
        flash('Tên file không hợp lệ')
        return redirect(url_for('index'))

    if not (allowed_file(template.filename) and allowed_file(data.filename)):
        flash('Chỉ chấp nhận .docx và .xlsx/.xls')
        return redirect(url_for('index'))

    job_id = uuid.uuid4().hex
    job_output = os.path.join(OUTPUT_BASE, job_id)
    os.makedirs(job_output, exist_ok=True)

    # save uploads
    tpl_name = secure_filename(template.filename)
    data_name = secure_filename(data.filename)
    tpl_path = os.path.join(UPLOAD_FOLDER, f"{job_id}_tpl_{tpl_name}")
    data_path = os.path.join(UPLOAD_FOLDER, f"{job_id}_data_{data_name}")
    template.save(tpl_path)
    data.save(data_path)

    # call generator
    try:
        created = generate_from_files(data_path, tpl_path, output_folder=job_output)
    except Exception as e:
        flash(f'Lỗi khi xử lý: {e}')
        return redirect(url_for('index'))

    # prepare relative paths for links
    rel_files = [os.path.basename(p) for p in created]

    # create a zip of all generated files for easy download
    zip_name = f"{job_id}.zip"
    zip_path = os.path.join(job_output, zip_name)
    try:
        with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
            for p in created:
                # arcname should be just the filename
                zf.write(p, arcname=os.path.basename(p))
    except Exception as e:
        flash(f'Không thể tạo ZIP: {e}')

    return render_template('result.html', files=rel_files, job_id=job_id, zip_name=zip_name)


@app.route('/download/<job_id>/<filename>')
def download(job_id, filename):
    # Serve a single generated file. Resolve the real filename server-side to avoid path issues.
    directory = os.path.join(OUTPUT_BASE, job_id)
    if not os.path.isdir(directory):
        return "Not Found", 404

    # normalize the requested filename and find a matching file in the directory
    requested_safe = secure_filename(filename)
    match = None
    for f in os.listdir(directory):
        if secure_filename(f) == requested_safe:
            match = f
            break

    if match is None:
        return "Not Found", 404

    abs_path = os.path.join(directory, match)
    return send_file(abs_path, as_attachment=True)


@app.route('/preview/<job_id>/<filename>')
def preview_docx(job_id, filename):
    """Return a simple HTML preview (text) of a generated .docx file."""
    directory = os.path.join(OUTPUT_BASE, job_id)
    if not os.path.isdir(directory):
        return jsonify({'error': 'Job not found'}), 404

    requested_safe = secure_filename(filename)
    match = None
    for f in os.listdir(directory):
        if secure_filename(f) == requested_safe:
            match = f
            break

    if match is None:
        return jsonify({'error': 'File not found'}), 404

    abs_path = os.path.join(directory, match)
    try:
        # Use mammoth to convert DOCX to HTML which preserves much of formatting
        with open(abs_path, 'rb') as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value  # the generated HTML
            # mammoth.messages may contain warnings
        return Response(html, mimetype='text/html')
    except Exception as e:
        return jsonify({'error': f'Không thể đọc file: {e}'}), 500


@app.route('/download_all/<job_id>')
def download_all(job_id):
    """Serve the zip containing all generated files for the job."""
    directory = os.path.join(OUTPUT_BASE, job_id)
    zip_name = f"{job_id}.zip"
    return send_from_directory(directory, zip_name, as_attachment=True)


if __name__ == '__main__':
    # When running locally for testing, allow override of port via PORT env var
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_DEBUG', '1') == '1'
    app.run(host='0.0.0.0', port=port, debug=debug)
