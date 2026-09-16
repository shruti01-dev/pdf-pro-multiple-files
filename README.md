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

## Install

Download the file for your computer from the [latest release](https://github.com/shruti01-dev/pdf-pro-multiple-files/releases/latest).

| Your computer | Download | Then |
| --- | --- | --- |
| Windows | `SSA PDF Studio.exe` | Double-click to run. No install needed. |
| Mac (Apple Silicon: M1–M4) | `SSA-PDF-Studio-macos-arm64.dmg` | Open the `.dmg`, drag the app to Applications. |
| Mac (older Intel) | `SSA-PDF-Studio-macos-x86_64.dmg` | Open the `.dmg`, drag the app to Applications. |

Not sure which Mac you have? Click the Apple menu > About This Mac. "Apple M1/M2/M3/M4" means
Apple Silicon; "Intel" means the Intel download.

### First launch on a Mac

The app is not signed with a paid Apple Developer certificate, so macOS blocks it the first time
and may say the app "cannot be opened" or is "damaged". This is expected. To open it:

1. In Applications, **right-click** (or Control-click) the app and choose **Open**.
2. Click **Open** again in the dialog.

macOS remembers the choice, so normal double-clicking works from then on. If it still refuses,
run this once in Terminal:

```bash
xattr -dr com.apple.quarantine "/Applications/SSA PDF Studio.app"
```

## Run from source

```bash
python -m pip install -r requirements.txt
python Pdf.py
```

Optional (scanned PDFs only): install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki).
On Windows run `install_tesseract.bat`; on Mac run `bash install_tesseract.sh` (or `brew install tesseract`).

## Build a desktop app

PyInstaller [is not a cross-compiler](https://pyinstaller.org/en/latest/), so a Windows machine
cannot produce a Mac app and vice versa. Each build must run on its own operating system:

```bash
python -m pip install -r requirements.txt pyinstaller
pyinstaller pdf.spec --clean --noconfirm
```

This produces `dist/SSA PDF Studio.exe` on Windows and `dist/SSA PDF Studio.app` on macOS.

### Releasing both platforms without owning a Mac

`.github/workflows/build.yml` builds both on GitHub's runners. To publish a release:

```bash
git tag v1.1
git push origin v1.1
```

The workflow builds the Windows `.exe` and both Mac `.dmg` files and attaches them to the release
for that tag. To test a build without releasing, open the **Actions** tab, pick
**Build desktop apps**, and choose **Run workflow** — the results appear as downloadable artifacts.

## License

Private / internal use unless you add a license file.
