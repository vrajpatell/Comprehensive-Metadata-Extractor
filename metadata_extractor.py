#!/usr/bin/env python3
"""
metadata_extractor.py

A comprehensive metadata extraction tool for various document types.
This tool extracts detailed metadata including both file system attributes and document-specific properties.
Supported file types include:
  - DOCX files using python-docx
  - XLSX files using openpyxl
  - PPTX files using python-pptx
  - PDF files using PyPDF2
  - Legacy OLE files (DOC, XLS, PPT) using olefile

Usage:
    python metadata_extractor.py path/to/document
"""

import sys
import os
import argparse
import json
import stat
from datetime import datetime

def get_filesystem_metadata(file_path):
    """
    Retrieves file system metadata such as file size, permissions, creation and modification times.
    """
    try:
        st = os.stat(file_path)
        metadata = {
            "file_name": os.path.basename(file_path),
            "file_path": os.path.abspath(file_path),
            "size_bytes": st.st_size,
            "permissions": stat.filemode(st.st_mode),
            "last_modified": datetime.fromtimestamp(st.st_mtime).isoformat(),
            "last_access": datetime.fromtimestamp(st.st_atime).isoformat(),
            "creation_time": datetime.fromtimestamp(st.st_ctime).isoformat(),
        }
        return metadata
    except Exception as e:
        return {"error": f"Error retrieving filesystem metadata: {str(e)}"}

def extract_docx_metadata(file_path):
    """
    Extracts detailed metadata from a DOCX file using python-docx.
    """
    try:
        from docx import Document
    except ImportError:
        sys.exit("Error: python-docx is not installed. Install it using: pip install python-docx")
    
    doc = Document(file_path)
    core_props = doc.core_properties
    metadata = {
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
        # Additional detailed info can include custom properties if needed.
        "detailed": {
            "property_names": [prop for prop in dir(core_props) if not prop.startswith("_")],
        }
    }
    return metadata

def extract_xlsx_metadata(file_path):
    """
    Extracts detailed metadata from an XLSX file using openpyxl.
    """
    try:
        from openpyxl import load_workbook
    except ImportError:
        sys.exit("Error: openpyxl is not installed. Install it using: pip install openpyxl")
    
    wb = load_workbook(file_path, read_only=True)
    props = wb.properties
    metadata = {
        "title": props.title,
        "subject": props.subject,
        "creator": props.creator,
        "last_modified_by": props.lastModifiedBy,
        "created": props.created.isoformat() if props.created else None,
        "modified": props.modified.isoformat() if props.modified else None,
        "keywords": props.keywords,
        "category": props.category,
        "description": props.description,
        # Additional properties can be extended here
        "detailed": {
            "application": props.appVersion if hasattr(props, 'appVersion') else None,
        }
    }
    return metadata

def extract_pptx_metadata(file_path):
    """
    Extracts detailed metadata from a PPTX file using python-pptx.
    """
    try:
        from pptx import Presentation
    except ImportError:
        sys.exit("Error: python-pptx is not installed. Install it using: pip install python-pptx")
    
    prs = Presentation(file_path)
    core_props = prs.core_properties
    metadata = {
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
        }
    }
    return metadata

def extract_pdf_metadata(file_path):
    """
    Extracts detailed metadata from a PDF file using PyPDF2.
    """
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        sys.exit("Error: PyPDF2 is not installed. Install it using: pip install PyPDF2")
    
    try:
        reader = PdfReader(file_path)
        raw_metadata = reader.metadata
        metadata = {}
        if raw_metadata:
            for key, value in raw_metadata.items():
                # Removing leading '/' from keys and converting datetime if needed
                clean_key = key.lstrip("/")
                if isinstance(value, datetime):
                    metadata[clean_key] = value.isoformat()
                else:
                    metadata[clean_key] = value
        # Additional details like number of pages and PDF version
        metadata["number_of_pages"] = len(reader.pages)
        metadata["pdf_version"] = reader.pdf_header if hasattr(reader, 'pdf_header') else None
        return metadata
    except Exception as e:
        raise ValueError("Error reading PDF metadata: " + str(e))

def extract_ole_metadata(file_path):
    """
    Extracts detailed metadata from legacy OLE Office files (DOC, XLS, PPT) using olefile.
    """
    try:
        import olefile
    except ImportError:
        sys.exit("Error: olefile is not installed. Install it using: pip install olefile")
    
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
            "properties": [prop for prop in dir(meta) if not prop.startswith("_")]
        }
    }
    ole.close()
    return metadata

def extract_metadata(file_path):
    """
    Determines file type based on extension and extracts metadata accordingly.
    Combines filesystem metadata and document-specific metadata into a detailed structure.
    """
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()

    doc_metadata = {}
    if ext == ".docx":
        doc_metadata = extract_docx_metadata(file_path)
    elif ext == ".xlsx":
        doc_metadata = extract_xlsx_metadata(file_path)
    elif ext == ".pptx":
        doc_metadata = extract_pptx_metadata(file_path)
    elif ext == ".pdf":
        doc_metadata = extract_pdf_metadata(file_path)
    elif ext in [".doc", ".xls", ".ppt"]:
        doc_metadata = extract_ole_metadata(file_path)
    else:
        raise ValueError("Unsupported file format: " + ext)

    filesystem_metadata = get_filesystem_metadata(file_path)
    combined_metadata = {
        "file_system": filesystem_metadata,
        "document_format": ext,
        "document_metadata": doc_metadata,
    }
    return combined_metadata

def main():
    parser = argparse.ArgumentParser(description="Extract comprehensive metadata from Office and PDF documents.")
    parser.add_argument("file", type=str, help="Path to the document")
    args = parser.parse_args()

    file_path = args.file

    if not os.path.exists(file_path):
        print("File does not exist:", file_path)
        sys.exit(1)

    try:
        metadata = extract_metadata(file_path)
        # Print the structured metadata as a JSON string
        print(json.dumps(metadata, indent=4))
    except Exception as e:
        print("Error extracting metadata:", str(e))
        sys.exit(1)

if __name__ == "__main__":
    main()
