@echo off
echo Installing Tesseract OCR (required for scanned PDFs)...
echo.
winget install --id UB-Mannheim.TesseractOCR -e --accept-package-agreements --accept-source-agreements
echo.
echo Done. Close SSA PDF Studio completely, then run: python Pdf.py
echo Then use PDF to Excel again.
pause
