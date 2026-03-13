# Comprehensive Metadata Extractor

A beginner-friendly Python project that extracts **filesystem metadata** and **document metadata** from common Office and PDF files, with both a **CLI tool** and a **Flask web interface**.

## Why this project is useful

When you're auditing documents, troubleshooting office files, or building digital forensics workflows, metadata is often more valuable than the visible content. This project provides a simple way to inspect that metadata locally or through a browser UI.

## Supported file types

- `.docx` (Word Open XML)
- `.xlsx` (Excel Open XML)
- `.pptx` (PowerPoint Open XML)
- `.pdf`
- Legacy OLE formats:
  - `.doc`
  - `.xls`
  - `.ppt`

## Features

- Extracts file system metadata (size, timestamps, permissions, path).
- Extracts format-specific document metadata (author, title, revision, etc.).
- CLI mode for scripts, automation, and terminal users.
- Web UI mode for non-technical users.
- Safe JSON serialization for datetime/bytes values.
- Graceful error handling for unsupported files and missing dependencies.
- Render-ready deployment configuration.

## Screenshots

> Add screenshots here after deployment/local run:
- Landing page screenshot
- Results page screenshot

## Architecture / How it works

1. `metadata_extractor.py` contains reusable extraction functions and CLI entry point.
2. `app.py` imports `extract_metadata()` and exposes Flask routes:
   - `GET /` → landing page + upload form
   - `POST /extract` → upload + metadata extraction
   - `GET /health` → health check
3. The UI uses Jinja templates and a small CSS file.
4. Uploaded files are saved in a temporary file, processed, and removed immediately.

## Local setup

### 1) Clone repository

```bash
git clone https://github.com/vrajpatell/Comprehensive-Metadata-Extractor.git
cd Comprehensive-Metadata-Extractor
```

### 2) Create and activate virtual environment

**macOS/Linux**
```bash
python3 -m venv .venv
source .venv/bin/activate
```

**Windows (PowerShell)**
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3) Install dependencies

```bash
pip install -r requirements.txt
```

## Run the CLI

```bash
python metadata_extractor.py path/to/document.pdf
```

### Example CLI commands

```bash
python metadata_extractor.py samples/report.docx
python metadata_extractor.py samples/slides.pptx
python metadata_extractor.py samples/financials.xlsx
python metadata_extractor.py samples/legacy.doc
```

### Example JSON output

```json
{
  "file_system": {
    "file_name": "report.docx",
    "file_path": "/absolute/path/report.docx",
    "size_bytes": 29314,
    "permissions": "-rw-r--r--",
    "last_modified": "2026-03-13T14:37:10.112233",
    "last_access": "2026-03-13T14:37:11.112233",
    "creation_time": "2026-03-13T14:37:10.112233"
  },
  "document_format": ".docx",
  "document_metadata": {
    "author": "Jane Doe",
    "title": "Quarterly Report",
    "revision": 4,
    "created": "2026-03-10T09:00:00",
    "modified": "2026-03-12T18:10:00"
  }
}
```

## Run the web UI

```bash
python app.py
```

Then open: `http://localhost:5000`

### Web workflow

1. Open landing page.
2. Upload one supported file.
3. View grouped metadata and formatted raw JSON.
4. Click “Upload another file” to repeat.

## Deploy on Render

This repository includes `render.yaml` for Render Blueprint deployment.

### Option A: Blueprint deploy (recommended)

1. Push your fork/branch to GitHub.
2. In Render, click **New +** → **Blueprint**.
3. Select this repository.
4. Render detects `render.yaml` and configures the web service.
5. Deploy.

### Option B: Manual web service

- **Environment:** Python
- **Build Command:** `pip install -r requirements.txt`
- **Start Command:** `gunicorn app:app`

## Project structure

```text
.
├── app.py
├── metadata_extractor.py
├── requirements.txt
├── render.yaml
├── static/
│   └── style.css
└── templates/
    ├── base.html
    ├── index.html
    └── result.html
```

## Future improvements

- Add unit tests for extractor functions and Flask routes.
- Add drag-and-drop upload zone enhancements.
- Add sample files for quick demo/testing.
- Improve metadata normalization for uncommon edge cases.
- Add Dockerfile for containerized deployment options.

## License

No license file is currently included. Add a `LICENSE` file (for example, MIT) if you want to permit open-source reuse explicitly.

## Author

Maintained by **Vraj Patel** and contributors.
