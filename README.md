# Simple web UI to generate Word files from a .docx template and an Excel file

This small project provides a minimal Flask web app to upload a Word template (`.docx`) and an Excel file (`.xlsx`) and generate one Word document per row in the spreadsheet.

Quick start (Windows PowerShell):

1. Create and activate a virtual environment (optional but recommended):

```powershell
python -m venv .venv; .\.venv\Scripts\Activate.ps1
```

2. Install dependencies:

```powershell
pip install -r requirements.txt
```

3. Run the app:

```powershell
python app.py
```

4. Open http://127.0.0.1:5000 in your browser, upload `template.docx` and `data.xlsx`, then click "Tạo file".

Notes:
- The generator uses `{{ColumnName}}` placeholders in the Word template (must match exact Excel header names).
- Generated files are stored under `output/<job_id>/` and can be downloaded from the result page.

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
