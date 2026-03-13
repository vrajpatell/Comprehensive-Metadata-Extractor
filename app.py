"""Flask web interface for the Comprehensive Metadata Extractor."""

import json
import os
import tempfile
from pathlib import Path

from flask import Flask, render_template, request
from werkzeug.utils import secure_filename

from metadata_extractor import SUPPORTED_EXTENSIONS, extract_metadata

MAX_FILE_SIZE_BYTES = 16 * 1024 * 1024  # 16MB

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_SIZE_BYTES


@app.get("/")
def index():
    """Render landing page with upload form."""
    return render_template("index.html", supported_extensions=sorted(SUPPORTED_EXTENSIONS))


@app.get("/health")
def health():
    """Simple health endpoint for uptime checks."""
    return {"status": "ok"}, 200


@app.post("/extract")
def extract():
    """Handle upload, run metadata extraction, and render results page."""
    uploaded_file = request.files.get("file")
    if not uploaded_file or not uploaded_file.filename:
        return render_template(
            "index.html",
            supported_extensions=sorted(SUPPORTED_EXTENSIONS),
            error="Please select a file before submitting.",
        ), 400

    safe_name = secure_filename(uploaded_file.filename)
    extension = Path(safe_name).suffix.lower()

    if extension not in SUPPORTED_EXTENSIONS:
        return render_template(
            "index.html",
            supported_extensions=sorted(SUPPORTED_EXTENSIONS),
            error=f"Unsupported file format: {extension or 'unknown'}.",
        ), 400

    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=extension) as temp_file:
            temp_path = temp_file.name
            uploaded_file.save(temp_file)

        metadata = extract_metadata(temp_path)
        metadata.setdefault("file_system", {})["original_upload_name"] = safe_name
        pretty_json = json.dumps(metadata, indent=2)

        return render_template(
            "result.html",
            metadata=metadata,
            pretty_json=pretty_json,
        )
    except Exception as exc:
        return render_template(
            "index.html",
            supported_extensions=sorted(SUPPORTED_EXTENSIONS),
            error=f"Unable to extract metadata: {exc}",
        ), 500
    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port)
