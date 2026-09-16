#!/usr/bin/env bash
# macOS / Linux equivalent of install_tesseract.bat.
# Installs Tesseract OCR, which SSA PDF Studio needs only for scanned PDFs.
set -euo pipefail

echo "Installing Tesseract OCR (required for scanned PDFs)..."
echo

if [[ "$(uname -s)" == "Darwin" ]]; then
  if ! command -v brew >/dev/null 2>&1; then
    echo "Homebrew is not installed."
    echo "Install it from https://brew.sh and then run this script again."
    exit 1
  fi
  brew install tesseract
elif command -v apt >/dev/null 2>&1; then
  sudo apt update && sudo apt install -y tesseract-ocr
elif command -v dnf >/dev/null 2>&1; then
  sudo dnf install -y tesseract
else
  echo "No supported package manager found. Install tesseract manually."
  exit 1
fi

echo
echo "Done. Verifying..."
tesseract --version | head -n 1
echo
echo "Quit SSA PDF Studio completely, reopen it, then use PDF to Excel again."
