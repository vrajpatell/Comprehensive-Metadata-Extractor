# Comprehensive Metadata Extractor

This project is a Python tool designed to extract detailed metadata from a variety of document types. In addition to retrieving document-specific properties, the tool also gathers file system metadata such as file size, permissions, and timestamps.

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

Install the required dependencies using pip:

```bash
pip install python-docx python-pptx openpyxl olefile PyPDF2
