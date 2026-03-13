#!/usr/bin/env python3
"""Comprehensive metadata extraction for Office and PDF documents.

This module can be used both as:
- a command line script (`python metadata_extractor.py <file>`)
- an importable library (`from metadata_extractor import extract_metadata`)
"""

import argparse
import json
import os
import stat
import sys
from datetime import datetime
from typing import Dict

SUPPORTED_EXTENSIONS = {".docx", ".xlsx", ".pptx", ".pdf", ".doc", ".xls", ".ppt"}


def _dependency_import_error(package_name: str, install_name: str | None = None) -> None:
    """Raise a consistent error when an optional dependency is missing."""
    package_to_install = install_name or package_name
    raise RuntimeError(
        f"Missing optional dependency '{package_name}'. Install it with: pip install {package_to_install}"
    )


def _make_json_safe(value):
    """Convert nested metadata values into JSON-serializable primitives."""
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    if isinstance(value, dict):
        return {str(k): _make_json_safe(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_make_json_safe(item) for item in value]
    return str(value)


def get_filesystem_metadata(file_path: str) -> Dict[str, str]:
    """Retrieve filesystem metadata (size, timestamps, permissions)."""
    st = os.stat(file_path)
    return {
        "file_name": os.path.basename(file_path),
        "file_path": os.path.abspath(file_path),
        "size_bytes": st.st_size,
        "permissions": stat.filemode(st.st_mode),
        "last_modified": datetime.fromtimestamp(st.st_mtime).isoformat(),
        "last_access": datetime.fromtimestamp(st.st_atime).isoformat(),
        "creation_time": datetime.fromtimestamp(st.st_ctime).isoformat(),
    }


def extract_docx_metadata(file_path: str) -> Dict:
    """Extract metadata from DOCX files using python-docx."""
    try:
        from docx import Document
    except ImportError:
        _dependency_import_error("python-docx")

    doc = Document(file_path)
    core_props = doc.core_properties
    return {
        "author": core_props.author,
        "category": core_props.category,
        "comments": core_props.comments,
        "content_status": core_props.content_status,
        "created": core_props.created.isoformat() if core_props.created else None,
        "identifier": core_props.identifier,
        "keywords": core_props.keywords,
        "language": core_props.language,
        "last_modified_by": core_props.last_modified_by,
        "last_printed": core_props.last_printed,
        "modified": core_props.modified.isoformat() if core_props.modified else None,
        "revision": core_props.revision,
        "subject": core_props.subject,
        "title": core_props.title,
        "detailed": {
            "property_names": [prop for prop in dir(core_props) if not prop.startswith("_")],
        },
    }


def extract_xlsx_metadata(file_path: str) -> Dict:
    """Extract metadata from XLSX files using openpyxl."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        _dependency_import_error("openpyxl")

    wb = load_workbook(file_path, read_only=True)
    props = wb.properties
    return {
        "title": props.title,
        "subject": props.subject,
        "creator": props.creator,
        "last_modified_by": props.lastModifiedBy,
        "created": props.created.isoformat() if props.created else None,
        "modified": props.modified.isoformat() if props.modified else None,
        "keywords": props.keywords,
        "category": props.category,
        "description": props.description,
        "detailed": {
            "application": props.appVersion if hasattr(props, "appVersion") else None,
        },
    }


def extract_pptx_metadata(file_path: str) -> Dict:
    """Extract metadata from PPTX files using python-pptx."""
    try:
        from pptx import Presentation
    except ImportError:
        _dependency_import_error("python-pptx")

    prs = Presentation(file_path)
    core_props = prs.core_properties
    return {
        "author": core_props.author,
        "category": core_props.category,
        "comments": core_props.comments,
        "content_status": core_props.content_status,
        "created": core_props.created.isoformat() if core_props.created else None,
        "identifier": core_props.identifier,
        "keywords": core_props.keywords,
        "language": core_props.language,
        "last_modified_by": core_props.last_modified_by,
        "last_printed": core_props.last_printed,
        "modified": core_props.modified.isoformat() if core_props.modified else None,
        "revision": core_props.revision,
        "subject": core_props.subject,
        "title": core_props.title,
        "detailed": {
            "property_names": [prop for prop in dir(core_props) if not prop.startswith("_")],
        },
    }


def extract_pdf_metadata(file_path: str) -> Dict:
    """Extract metadata from PDF files using PyPDF2."""
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        _dependency_import_error("PyPDF2")

    reader = PdfReader(file_path)
    raw_metadata = reader.metadata
    metadata = {}
    if raw_metadata:
        for key, value in raw_metadata.items():
            clean_key = key.lstrip("/")
            metadata[clean_key] = value.isoformat() if isinstance(value, datetime) else value

    metadata["number_of_pages"] = len(reader.pages)
    metadata["pdf_version"] = reader.pdf_header if hasattr(reader, "pdf_header") else None
    return metadata


def extract_ole_metadata(file_path: str) -> Dict:
    """Extract metadata from legacy OLE Office files (DOC, XLS, PPT)."""
    try:
        import olefile
    except ImportError:
        _dependency_import_error("olefile")

    if not olefile.isOleFile(file_path):
        raise ValueError("Not a valid OLE file.")

    ole = olefile.OleFileIO(file_path)
    meta = ole.get_metadata()
    metadata = {
        "author": meta.author,
        "title": meta.title,
        "subject": meta.subject,
        "keywords": meta.keywords,
        "comments": meta.comments,
        "last_saved_by": meta.last_saved_by,
        "revision_number": meta.revision,
        "total_edit_time": meta.total_edit_time,
        "create_time": meta.create_time.isoformat() if meta.create_time else None,
        "last_printed": meta.last_printed.isoformat() if meta.last_printed else None,
        "last_saved_time": meta.last_save_time.isoformat() if meta.last_save_time else None,
        "detailed": {
            "properties": [prop for prop in dir(meta) if not prop.startswith("_")],
        },
    }
    ole.close()
    return metadata


def extract_metadata(file_path: str) -> Dict:
    """Extract filesystem + document metadata for a supported file path."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File does not exist: {file_path}")

    _, ext = os.path.splitext(file_path)
    ext = ext.lower()

    if ext == ".docx":
        doc_metadata = extract_docx_metadata(file_path)
    elif ext == ".xlsx":
        doc_metadata = extract_xlsx_metadata(file_path)
    elif ext == ".pptx":
        doc_metadata = extract_pptx_metadata(file_path)
    elif ext == ".pdf":
        doc_metadata = extract_pdf_metadata(file_path)
    elif ext in {".doc", ".xls", ".ppt"}:
        doc_metadata = extract_ole_metadata(file_path)
    else:
        raise ValueError(f"Unsupported file format: {ext}")

    combined_metadata = {
        "file_system": get_filesystem_metadata(file_path),
        "document_format": ext,
        "document_metadata": doc_metadata,
    }
    return _make_json_safe(combined_metadata)


def main() -> None:
    """CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Extract comprehensive metadata from Office and PDF documents."
    )
    parser.add_argument("file", type=str, help="Path to the document")
    args = parser.parse_args()

    try:
        metadata = extract_metadata(args.file)
        print(json.dumps(metadata, indent=4))
    except Exception as exc:
        print("Error extracting metadata:", str(exc))
        sys.exit(1)


if __name__ == "__main__":
    main()
