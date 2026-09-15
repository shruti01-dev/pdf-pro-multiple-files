# SSA PDF Studio

Offline desktop PDF toolkit for **Windows and Mac**. Merge, split, convert, compress, watermark, and protect PDFs without uploading files to the internet.

## About

SSA PDF Studio is a local Python app for accounting and office work. It keeps source PDFs on your machine, lets you set merge order yourself, and converts bank statements and payslips to Excel / Word as close to the original layout as possible.

## Features

**Organize**
- Merge PDFs in the order you choose (Move up / Move down, then Next)
- Split, extract, or remove pages

**Optimize & edit**
- Compress PDF
- Rotate pages
- Text watermark
- Page numbers

**Convert**
- PDF to images
- Images to PDF
- PDF to Word (layout copy)
- PDF to Excel (bank statements, payslips, tables; OCR for scans)

**Security**
- Password-protect a PDF
- Unlock a PDF if you know the password

## Run

```bash
python -m pip install pypdf pillow reportlab python-docx pandas openpyxl pymupdf pdfplumber
python Pdf.py
```

Optional (scanned PDFs): install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki). On Windows you can run `install_tesseract.bat`.

## Build a desktop app

```bash
python -m pip install pyinstaller
pyinstaller pdf.spec
```

## License

Private / internal use unless you add a license file.
