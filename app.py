"""
Flask app - Word document generator từ Excel + template.
Flow: Upload → Configure → Generate
"""

from flask import Flask, request, render_template, jsonify, redirect, url_for, flash, send_file, send_from_directory
import os
import json
import zipfile
from werkzeug.utils import secure_filename

from job_manager import JobManager
from excel_utils import get_sheet_names, read_excel_sheet
from word import get_placeholders_from_template, generate_from_mapping

# --- Cấu hình ---
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_BASE = os.path.join(BASE_DIR, "output")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_BASE, exist_ok=True)

ALLOWED_EXTENSIONS = {'.xlsx', '.xls', '.docx'}

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'change-this-secret-key-in-production')

job_mgr = JobManager(OUTPUT_BASE)


def allowed_file(filename):
    """Kiểm tra file extension hợp lệ."""
    _, ext = os.path.splitext(filename.lower())
    return ext in ALLOWED_EXTENSIONS


def num_to_excel_col(num):
    """Convert số (0-based) thành tên cột Excel: 0→A, 1→B, ..., 26→AA, 27→AB, ..."""
    result = ""
    num = num + 1  # convert to 1-based
    while num > 0:
        num -= 1
        result = chr(65 + (num % 26)) + result
        num //= 26
    return result


def validate_job(job_id):
    """Validate job tồn tại."""
    if not job_mgr.job_exists(job_id):
        return False, "Job không tồn tại"
    return True, None


# ============ ROUTES ============

@app.route('/')
def index():
    """Trang chủ - upload file."""
    return render_template('index.html')


@app.route('/api/analyze', methods=['POST'])
def analyze():
    """
    Upload file Word + Excel → phân tích.
    API endpoint trả JSON.
    """
    if 'template_file' not in request.files or 'data_file' not in request.files:
        return jsonify({'error': 'Cần upload cả file template (.docx) và file dữ liệu (.xlsx)'}), 400
    
    template = request.files['template_file']
    data = request.files['data_file']
    
    if template.filename == '' or data.filename == '':
        return jsonify({'error': 'Tên file không hợp lệ'}), 400
    
    if not (allowed_file(template.filename) and allowed_file(data.filename)):
        return jsonify({'error': 'Chỉ chấp nhận .docx và .xlsx/.xls'}), 400
    
    try:
        # Tạo job mới
        job_id = job_mgr.create_job()
        job_dir = job_mgr.get_job_dir(job_id)
        config = job_mgr.load_job_config(job_id)
        
        # Lưu file upload
        tpl_name = secure_filename(template.filename)
        data_name = secure_filename(data.filename)
        tpl_path = os.path.join(UPLOAD_FOLDER, f"{job_id}_tpl_{tpl_name}")
        data_path = os.path.join(UPLOAD_FOLDER, f"{job_id}_data_{data_name}")
        template.save(tpl_path)
        data.save(data_path)
        
        config['template_path'] = tpl_path
        config['excel_path'] = data_path
        
        # Đọc sheet names
        try:
            sheet_names = get_sheet_names(data_path)
            config['excel_sheets'] = sheet_names
            config['sheet_name'] = sheet_names[0]  # mặc định sheet đầu
        except Exception as e:
            return jsonify({'error': f'Không thể đọc Excel: {e}'}), 400
        
        # Đọc placeholder từ template
        try:
            placeholders = sorted(list(get_placeholders_from_template(tpl_path)))
            config['placeholders'] = [{'name': p, 'column': None} for p in placeholders]
        except Exception as e:
            return jsonify({'error': f'Không thể đọc template: {e}'}), 400
        
        # Đọc headers từ Excel (mặc định sheet đầu)
        try:
            headers, _ = read_excel_sheet(data_path, config['sheet_name'], 
                                          config['header_row'], config['data_start_row'])
            config['excel_headers'] = headers
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            app.logger.error(f"Error reading headers: {tb}")
            return jsonify({'error': f'Không thể đọc headers: {str(e)}\n\n{tb}'}), 400
        
        config['status'] = 'uploaded'
        job_mgr.save_job_config(job_id, config)
        
        return jsonify({
            'job_id': job_id,
            'sheets': sheet_names,
            'placeholders': config['placeholders'],
            'headers': headers
        })
    
    except Exception as e:
        return jsonify({'error': f'Lỗi server: {str(e)}'}), 500


@app.route('/configure/<job_id>')
def configure(job_id):
    """Trang configure - mapping placeholder + sheet settings."""
    ok, err = validate_job(job_id)
    if not ok:
        flash(err)
        return redirect(url_for('index')), 404
    
    try:
        config = job_mgr.load_job_config(job_id)
        
        # Tính preview data
        headers, rows = read_excel_sheet(
            config['excel_path'], 
            config['sheet_name'],
            config['header_row'],
            config['data_start_row']
        )
        
        preview_rows = rows[1:5]  # 4 dòng mẫu
        
        # Tạo column labels (A, B, C, ..., Z, AA, AB, ...)
        column_labels = [num_to_excel_col(i) for i in range(len(config['excel_headers']))]
        
        return render_template('configure.html', 
                             job_id=job_id,
                             config=config,
                             preview_rows=preview_rows,
                             column_labels=column_labels)
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        app.logger.error(f"[CONFIGURE ERROR] {tb}")
        error_msg = f"<h2>Lỗi tải Configure:</h2><pre>{tb}</pre>"
        return error_msg, 500


@app.route('/api/config/<job_id>', methods=['GET', 'POST'])
def api_config(job_id):
    """API để load/save config."""
    ok, err = validate_job(job_id)
    if not ok:
        return jsonify({'error': err}), 404
    
    if request.method == 'GET':
        try:
            config = job_mgr.load_job_config(job_id)
            return jsonify(config)
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    
    elif request.method == 'POST':
        try:
            data = request.json
            config = job_mgr.load_job_config(job_id)
            
            # Update allowed fields
            if 'sheet_name' in data:
                config['sheet_name'] = data['sheet_name']
            if 'header_row' in data:
                config['header_row'] = int(data['header_row'])
            if 'data_start_row' in data:
                config['data_start_row'] = int(data['data_start_row'])
            if 'placeholders' in data:
                config['placeholders'] = data['placeholders']
            if 'filename_column' in data:
                config['filename_column'] = data['filename_column']
            if 'required_columns' in data:
                config['required_columns'] = data['required_columns']
            if 'skip_empty_rows' in data:
                config['skip_empty_rows'] = data['skip_empty_rows']
            
            job_mgr.save_job_config(job_id, config)
            return jsonify({'success': True})
        except Exception as e:
            return jsonify({'error': str(e)}), 400


@app.route('/api/sheets/<job_id>', methods=['POST'])
def api_get_sheet_headers(job_id):
    """API để lấy headers khi user đổi sheet."""
    ok, err = validate_job(job_id)
    if not ok:
        return jsonify({'error': err}), 404
    
    try:
        sheet_name = request.json.get('sheet_name')
        config = job_mgr.load_job_config(job_id)
        
        headers, rows = read_excel_sheet(
            config['excel_path'],
            sheet_name,
            config['header_row'],
            config['data_start_row']
        )
        
        preview_rows = rows[:3]
        
        return jsonify({
            'headers': headers,
            'preview_rows': preview_rows,
            'row_count': len(rows)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 400


@app.route('/api/generate/<job_id>', methods=['POST'])
def api_generate(job_id):
    """API để generate files."""
    ok, err = validate_job(job_id)
    if not ok:
        return jsonify({'error': err}), 404
    
    try:
        config = job_mgr.load_job_config(job_id)
        
        # Validate config
        if not config['placeholders']:
            return jsonify({'error': 'Template không có placeholder'}), 400
        
        unmapped = [p for p in config['placeholders'] if not p['column']]
        if unmapped:
            unmapped_names = ', '.join([p['name'] for p in unmapped])
            return jsonify({'error': f'Placeholder chưa map: {unmapped_names}'}), 400
        
        # Build mapping dict
        mapping = {p['name']: p['column'] for p in config['placeholders']}
        
        job_dir = job_mgr.get_job_dir(job_id)
        output_folder = os.path.join(job_dir, 'generated')
        os.makedirs(output_folder, exist_ok=True)
        
        # Generate
        results, created_files = generate_from_mapping(
            config['excel_path'],
            config['template_path'],
            output_folder,
            mapping,
            sheet_name=config['sheet_name'],
            header_row=config['header_row'],
            data_start_row=config['data_start_row'],
            filename_column=config['filename_column'],
            required_columns=config['required_columns'],
            skip_empty_rows=config['skip_empty_rows']
        )
        
        # Tạo ZIP
        zip_path = os.path.join(job_dir, 'files.zip')
        with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
            for fpath in created_files:
                zf.write(fpath, arcname=os.path.basename(fpath))
        
        # Lưu results
        results_path = os.path.join(job_dir, 'results.json')
        with open(results_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        
        config['status'] = 'generated'
        job_mgr.save_job_config(job_id, config)
        
        # Summary
        success_count = sum(1 for r in results if r['status'] == 'success')
        skipped_count = sum(1 for r in results if r['status'] == 'skipped')
        error_count = sum(1 for r in results if r['status'] == 'error')
        
        return jsonify({
            'success': True,
            'results': results,
            'summary': {
                'total': len(results),
                'success': success_count,
                'skipped': skipped_count,
                'error': error_count,
                'files_count': len(created_files)
            }
        })
    
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        app.logger.error(f"Generate error: {tb}")
        return jsonify({'error': f"Lỗi khi tạo file: {str(e)}\n\n{tb}"}), 500


@app.route('/result/<job_id>')
def result(job_id):
    """Trang kết quả."""
    ok, err = validate_job(job_id)
    if not ok:
        flash(err)
        return redirect(url_for('index')), 404
    
    try:
        config = job_mgr.load_job_config(job_id)
        job_dir = job_mgr.get_job_dir(job_id)
        results_path = os.path.join(job_dir, 'results.json')
        
        if not os.path.exists(results_path):
            flash('Chưa generate')
            return redirect(url_for('configure', job_id=job_id)), 404
        
        with open(results_path, 'r', encoding='utf-8') as f:
            results = json.load(f)
        
        output_folder = os.path.join(job_dir, 'generated')
        generated_files = []
        if os.path.exists(output_folder):
            generated_files = sorted(os.listdir(output_folder))
        
        success_count = sum(1 for r in results if r['status'] == 'success')
        skipped_count = sum(1 for r in results if r['status'] == 'skipped')
        error_count = sum(1 for r in results if r['status'] == 'error')
        
        return render_template('result.html',
                             job_id=job_id,
                             results=results,
                             files=generated_files,
                             summary={
                                 'total': len(results),
                                 'success': success_count,
                                 'skipped': skipped_count,
                                 'error': error_count
                             })
    except Exception as e:
        flash(f'Lỗi: {str(e)}')
        return redirect(url_for('index')), 500


@app.route('/download/<job_id>/<filename>')
def download(job_id, filename):
    """Tải từng file."""
    ok, err = validate_job(job_id)
    if not ok:
        return err, 404
    
    job_dir = job_mgr.get_job_dir(job_id)
    output_folder = os.path.join(job_dir, 'generated')
    
    # Decodse URL-encoded filename
    from urllib.parse import unquote
    filename = unquote(filename)
    
    # Find file in output folder (case-insensitive match)
    if not os.path.isdir(output_folder):
        return "Thư mục output không tồn tại", 404
    
    # List files and find match
    try:
        files = os.listdir(output_folder)
        file_match = None
        
        # Try exact match first
        if filename in files:
            file_match = filename
        else:
            # Try case-insensitive match
            for f in files:
                if f.lower() == filename.lower():
                    file_match = f
                    break
        
        if not file_match:
            return "File không tìm thấy", 404
        
        file_path = os.path.join(output_folder, file_match)
        
        # Double-check file is in the right folder (security)
        if not os.path.abspath(file_path).startswith(os.path.abspath(output_folder)):
            return "Không được phép truy cập", 403
        
        return send_file(file_path, as_attachment=True)
    
    except Exception as e:
        app.logger.error(f"Download error: {e}")
        return f"Lỗi: {str(e)}", 500


@app.route('/download_all/<job_id>')
def download_all(job_id):
    """Tải ZIP tất cả file."""
    ok, err = validate_job(job_id)
    if not ok:
        return err, 404
    
    job_dir = job_mgr.get_job_dir(job_id)
    zip_path = os.path.join(job_dir, 'files.zip')
    
    if not os.path.exists(zip_path):
        return "ZIP không tìm thấy", 404
    
    return send_file(zip_path, as_attachment=True, download_name='files.zip')


# ============ ERROR HANDLERS ============

@app.errorhandler(404)
def not_found(error):
    return "Not Found", 404


@app.errorhandler(500)
def server_error(error):
    return "Server Error", 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_DEBUG', '1') == '1'
    app.run(host='0.0.0.0', port=port, debug=debug)
