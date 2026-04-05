# Word Document Generator

Một ứng dụng Flask đơn giản để tạo hàng loạt file Word từ template + dữ liệu Excel.

## ✨ Tính năng

- **3 bước rõ ràng**: Upload → Configure → Generate
- **Mapping linh hoạt**: Nối placeholder `{{A}}`, `{{B}}`, `{{C}}` với cột Excel bất kỳ
- **Chọn Sheet & Header**: Hỗ trợ nhiều sheet, chọn dòng header, dòng bắt đầu dữ liệu
- **Validation**: Cột bắt buộc, bỏ qua dòng trống
- **Log chi tiết**: Xem kết quả từng dòng (thành công/bỏ qua/lỗi)
- **Tải ZIP**: Tải tất cả file cùng lúc
- **Dễ debug**: Code rõ ràng, không over-engineering

## 🚀 Quick Start

### Prerequisites
- Python 3.7+
- Không cần database, không cần queue, chỉ cần Flask

### Installation & Run

**Windows (PowerShell):**
```powershell
# Tạo virtual environment (optional nhưng nên dùng)
python -m venv word_app
.\word_app\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt

# Run app
python app.py
```

**Linux/Mac:**
```bash
python -m venv word_app
source word_app/bin/activate
pip install -r requirements.txt
python app.py
```

App sẽ chạy tại: **http://localhost:5000**

## 📋 Cách dùng

### Bước 1: Upload
- Upload file Word template (`.docx`) có placeholder `{{A}}`, `{{B}}`, ...
- Upload file Excel (`.xlsx`, `.xls`) có dữ liệu

### Bước 2: Configure
- **Chọn Sheet**: Chọn sheet Excel cần dùng
- **Chọn Header Row & Data Start Row**: Chỉ định dòng header và dòng dữ liệu bắt đầu
- **Mapping Placeholder**: Nối mỗi placeholder với cột Excel tương ứng
  - Ví dụ: `{{A}}` → "Họ và tên", `{{B}}` → "Ngày sinh"
- **Tùy chọn**:
  - Chọn cột làm tên file output
  - Đánh dấu cột bắt buộc
  - Bỏ qua dòng trống

### Bước 3: Generate
- Click "Tạo file"
- Xem kết quả chi tiết (thành công/bỏ qua/lỗi)
- Tải file hoặc tải ZIP

## 📁 Cấu trúc thư mục

```
word-generator-app/
├── app.py              # Routes Flask chính
├── word.py             # Xử lý Word (placeholder, fill)
├── excel_utils.py      # Đọc Excel
├── job_manager.py      # Quản lý job state
├── requirements.txt
├── README.md
├── templates/
│   ├── index.html      # Bước 1: Upload
│   ├── configure.html  # Bước 2: Configure
│   └── result.html     # Bước 3: Result
├── static/
│   └── style.css       # Stylesheet
├── output/             # File kết quả (auto-created)
└── uploads/            # File upload tạm (auto-created)
```

## ⚙️ Cấu hình

Tùy chọn môi trường:
```bash
# Set port
$env:PORT = 8000
python app.py

# Disable debug mode
$env:FLASK_DEBUG = 0
python app.py

# Set secret key cho production
$env:SECRET_KEY = "your-secret-key"
python app.py
```

## 🔧 Cách hoạt động

### Flow chính:
1. **Upload**: Lưu file, đọc sheet names + placeholder từ template
2. **Configure**: Lưu config (mapping, settings) trong JSON
3. **Generate**: 
   - Đọc Excel theo config
   - Loop từng dòng, build data dict từ mapping
   - Fill template, save file
   - Ghi log (success/skipped/error), zip file

### Job State:
- Mỗi job có folder riêng: `output/<job_id>/`
- Config lưu trong `output/<job_id>/config.json`
- File output: `output/<job_id>/generated/`
- Results: `output/<job_id>/results.json`
- ZIP: `output/<job_id>/files.zip`

## 📝 Placeholder trong Word Template

Dùng format: `{{A}}`, `{{B}}`, `{{C}}` (placeholder đơn giản)

Template sẽ tìm `{{...}}` trong:
- Paragraphs
- Table cells

Xử lý an toàn cho placeholder bị tách giữa nhiều runs.

## 🛠️ Development

### Run tests:
```bash
python -m pytest tests/
```

### Debug mode:
```bash
$env:FLASK_DEBUG = 1
python app.py
```

### Cleaning:
```bash
# Xóa các job cũ
Remove-Item -Recurse output/*
```

## 🚢 Deploy trên Render

### 1. Prepare
```bash
# Create Procfile
echo "web: gunicorn app:app" > Procfile
```

### 2. Push to GitHub
```bash
git add .
git commit -m "Update app"
git push origin main
```

### 3. Deploy
- Vào Render.com → New → Web Service
- Connect GitHub repo
- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app`
- Set PORT env: 10000 (Render mặc định)

## 📦 Dependencies

- **Flask 2.0+** - Web framework
- **openpyxl 3.0+** - Đọc Excel
- **python-docx 0.8** - Xử lý Word
- **gunicorn** - WSGI server

## ⚠️ Limitations

- **Temp files**: File được lưu local, không dùng cloud storage
- **Memory**: Không thích hợp cho file Excel > 10K rows hoặc Word ngành lớn
- **No auth**: Không có authentication
- **Cleanup**: Cần xóa thủ công hoặc dùng cron job để cleanup
- **Single user**: Dùng được cho 1 user hoặc small team

## 🤔 FAQ

**Q: Placeholder không được replace?**  
A: Kiểm tra placeholder name (case-sensitive) có match với cột Excel không. Dùng trang Configure để verify.

**Q: Sao file output không được tạo?**  
A: Kiểm tra error message trên result page. Thường do Excel format lỗi hoặc template không hợp lệ.

**Q: Có thể customize file name?**  
A: Có, chọn cột làm tên file trên trang Configure. Hoặc để trống để dùng `row_N`.

**Q: Dữ liệu có được lưu?**  
A: Chỉ lưu tạm trong `output/<job_id>`. Cleanup tuỳ operator.

## 📞 Support

- Check error message trên UI
- Xem log trong terminal
- Check `output/<job_id>/results.json` cho chi tiết

## 📄 License

MIT


- There's also a ZIP download available on the result page to download all generated files at once ("Tải tất cả (ZIP)").

If you want, I can:
- Add a ZIP download of all generated files.
- Add progress feedback for large jobs.
- Validate columns and show a preview mapping UI.

Deploying to GitHub and Render.com
---------------------------------

1) Initialize Git and push to GitHub

```powershell
git init
git add .
git commit -m "Initial commit: Word generator web app"
# create a repo on GitHub and then add remote, example:
git remote add origin https://github.com/<your-username>/<repo-name>.git
git branch -M main
git push -u origin main
```

2) Configure Render.com

- Sign in to Render (https://render.com) and click "New" -> "Web Service".
- Connect your GitHub account and select the repository you pushed.
- For the build command, use (Render will auto-detect Python; you can leave build blank):

	Build command: pip install -r requirements.txt

- For the start command, use Gunicorn pointing to the Flask app object:

	Start command: gunicorn app:app --workers 4 --bind 0.0.0.0:$PORT

- Use the default branch (main), choose a plan, and create the service. Render will build and deploy.

Notes and tips
- Ensure `requirements.txt` includes all dependencies (we added `gunicorn` already).
- Keep `SECRET_KEY` and any sensitive config out of source control; set them as environment variables in Render's dashboard (SERVICE -> Environment -> Add SECRET_KEY).
- By default the app writes to `uploads/` and `output/` on the instance filesystem; Render ephemeral filesystem is fine for small runs, but for persistent storage or large jobs consider using S3 or attaching an external storage.

If you want, I can create a small `render.yaml` or add GitHub Actions to automatically push tags/releases. I can also add instructions to set secrets and an example `Procfile` if you prefer.
