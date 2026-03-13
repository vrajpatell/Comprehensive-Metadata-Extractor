# Comprehensive Metadata Extractor

This project is a Python CLI tool that extracts detailed metadata from common Office and PDF document formats. It combines:

- **Document metadata** (author, title, revision, etc.)
- **Filesystem metadata** (size, timestamps, permissions)

## Supported Document Types

- **Word Documents**
  - DOCX (using [python-docx](https://pypi.org/project/python-docx/))
  - Legacy DOC (OLE-based using [olefile](https://pypi.org/project/olefile/))
- **Excel Spreadsheets**
  - XLSX (using [openpyxl](https://pypi.org/project/openpyxl/))
  - Legacy XLS (OLE-based using [olefile](https://pypi.org/project/olefile/))
- **PowerPoint Presentations**
  - PPTX (using [python-pptx](https://pypi.org/project/python-pptx/))
  - Legacy PPT (OLE-based using [olefile](https://pypi.org/project/olefile/))
- **PDF Documents**
  - PDF (using [PyPDF2](https://pypi.org/project/PyPDF2/))

## Installation

Install dependencies:

```bash
pip install python-docx python-pptx openpyxl olefile PyPDF2
```

## Usage

Run the extractor against a supported file:

```bash
python metadata_extractor.py path/to/document.pdf
```

Successful output is JSON:

```json
{
  "file_system": {
    "file_name": "document.pdf",
    "file_path": "/abs/path/document.pdf",
    "size_bytes": 12345
  },
  "document_format": ".pdf",
  "document_metadata": {
    "Author": "Example",
    "number_of_pages": 10
  }
}
```

## Error Handling Behavior

- Missing files raise a clear file-not-found error.
- Unsupported extensions raise a clear unsupported-format error.
- Missing optional parsers raise actionable install guidance.
- Non-JSON-native metadata values are normalized into JSON-safe output.
