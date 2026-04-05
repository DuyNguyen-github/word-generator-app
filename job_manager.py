"""
Quản lý job state - lưu trữ thông tin phục vụ quá trình upload → configure → generate.
"""

import os
import json
import uuid


class JobManager:
    """Quản lý job state với lưu trữ file JSON."""
    
    def __init__(self, base_output_dir):
        self.base_output_dir = base_output_dir
        os.makedirs(base_output_dir, exist_ok=True)
    
    def create_job(self):
        """Tạo job mới. Trả về job_id."""
        job_id = uuid.uuid4().hex
        job_dir = os.path.join(self.base_output_dir, job_id)
        os.makedirs(job_dir, exist_ok=True)
        
        # Tạo file config trống
        config = {
            'job_id': job_id,
            'template_path': None,
            'excel_path': None,
            'sheet_name': None,
            'header_row': 1,
            'data_start_row': 2,
            'placeholders': [],  # [{"name": "A", "column": "Họ và tên"}, ...]
            'filename_column': None,
            'required_columns': [],  # ["Họ và tên", ...]
            'skip_empty_rows': True,
            'excel_headers': [],  # cached from analyze
            'excel_sheets': [],  # cached sheet names
            'status': 'created'  # created, uploaded, configured, generated
        }
        self.save_job_config(job_id, config)
        return job_id
    
    def save_job_config(self, job_id, config):
        """Lưu config job."""
        job_dir = os.path.join(self.base_output_dir, job_id)
        config_path = os.path.join(job_dir, 'config.json')
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    
    def load_job_config(self, job_id):
        """Tải config job."""
        job_dir = os.path.join(self.base_output_dir, job_id)
        config_path = os.path.join(job_dir, 'config.json')
        
        if not os.path.exists(config_path):
            raise ValueError(f"Job {job_id} không tồn tại")
        
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        return config
    
    def get_job_dir(self, job_id):
        """Lấy thư mục job."""
        return os.path.join(self.base_output_dir, job_id)
    
    def job_exists(self, job_id):
        """Kiểm tra job có tồn tại không."""
        return os.path.exists(os.path.join(self.base_output_dir, job_id, 'config.json'))
