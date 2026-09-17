"""
SSA PDF Studio - Desktop PDF Toolkit
------------------------------------
A Python desktop application inspired by common online PDF toolboxes.
It does NOT use iLovePDF branding, code, images, or proprietary UI assets.

Copyright (C) 2026 SSA

This program is free software: you can redistribute it and/or modify it under
the terms of the GNU Affero General Public License as published by the Free
Software Foundation, either version 3 of the License, or (at your option) any
later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY
WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A
PARTICULAR PURPOSE. See the GNU Affero General Public License for more details.
You should have received a copy of the license along with this program; if not,
see <https://www.gnu.org/licenses/>.

AGPL applies because this application bundles PyMuPDF, which is AGPL-licensed.
See THIRD-PARTY-NOTICES.txt for the full dependency list.

Core features included:
1. Merge PDF files
2. Split PDF into separate range-wise files (example: 1,3-5,8-10)
3. Extract selected pages into one PDF
4. Remove selected pages
5. Rotate selected/all pages
6. Add text watermark
7. Add page numbers
8. Protect/encrypt PDF
9. Unlock/decrypt PDF when you know the password
10. Compress/optimize PDF using PyMuPDF when available
11. Convert PDF pages to images
12. Convert images/JPG/PNG to PDF
13. Extract PDF text to Word
14. Extract PDF tables/text to Excel

Install dependencies:
    python3 -m pip install -r requirements.txt
    # Optional (scanned/image PDFs): install Tesseract OCR
    #   Windows: winget install --id UB-Mannheim.TesseractOCR
    #   macOS:   brew install tesseract

Run:
    python3 Pdf.py

Packaging (PyInstaller is not a cross-compiler, so run this on the target OS):
    python3 -m pip install pyinstaller
    pyinstaller pdf.spec
"""

import os
import re
import sys
import traceback
from collections import Counter
from pathlib import Path
from typing import List, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    from pypdf import PdfReader, PdfWriter
except Exception:  # pragma: no cover
    PdfReader = None
    PdfWriter = None

try:
    from PIL import Image
except Exception:  # pragma: no cover
    Image = None

try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
except Exception:  # pragma: no cover
    canvas = None
    A4 = None
    mm = None

APP_NAME = "SSA PDF Studio"
APP_VERSION = "1.0"


# --------------------------- Utility helpers ---------------------------

def require_package(condition, install_hint: str):
    if not condition:
        raise RuntimeError(f"Required package missing. Please install: {install_hint}")


def safe_filename(name: str) -> str:
    name = re.sub(r"[^a-zA-Z0-9._ -]", "_", name).strip()
    return name or "output"


def _is_file_lock_error(exc: BaseException) -> bool:
    if isinstance(exc, PermissionError):
        return True
    errno = getattr(exc, "errno", None)
    return errno in (13, 32) or "denied" in str(exc).lower() or "being used" in str(exc).lower()


def _unlocked_save_path(path: str) -> str:
    """If path is locked/open, return path_1, path_2, ... in the same folder."""
    path = os.path.normpath(path)
    root, ext = os.path.splitext(path)
    candidates = [path] + [f"{root}_{n}{ext}" for n in range(1, 40)]
    return candidates


def _save_docx_unlocked(docx, dest: str) -> str:
    last_err: BaseException | None = None
    for path in _unlocked_save_path(dest):
        try:
            docx.save(path)
            return path
        except OSError as exc:
            if _is_file_lock_error(exc):
                last_err = exc
                continue
            raise
    raise RuntimeError(
        "Cannot save the Word file because it is open in another program:\n"
        f"{os.path.normpath(dest)}\n\n"
        "Close converted.docx in Microsoft Word, then run PDF to Word again."
    ) from last_err


def _copy_file_unlocked(src: str, dest: str) -> str:
    import shutil

    last_err: BaseException | None = None
    for path in _unlocked_save_path(dest):
        try:
            shutil.copy2(src, path)
            return path
        except OSError as exc:
            if _is_file_lock_error(exc):
                last_err = exc
                continue
            raise
    raise RuntimeError(
        "Cannot save the file because it is open in another program:\n"
        f"{os.path.normpath(dest)}\n\n"
        "Close that file, then try again."
    ) from last_err


def _save_workbook_unlocked(wb, dest: str) -> str:
    last_err: BaseException | None = None
    for path in _unlocked_save_path(dest):
        try:
            wb.save(path)
            return path
        except OSError as exc:
            if _is_file_lock_error(exc):
                last_err = exc
                continue
            raise
    raise RuntimeError(
        "Cannot save the Excel file because it is open in another program:\n"
        f"{os.path.normpath(dest)}\n\n"
        "Close that workbook, then run PDF to Excel again."
    ) from last_err


def ensure_pdf_tools():
    require_package(PdfReader is not None and PdfWriter is not None, "python3 -m pip install pypdf")


def _dialog_parent():
    root = getattr(tk, "_default_root", None)
    if root is None:
        return None
    try:
        root.update_idletasks()
        root.lift()
        root.attributes("-topmost", True)
        root.after(80, lambda: root.attributes("-topmost", False))
        root.focus_force()
    except Exception:
        pass
    return root


def ask_output_folder() -> str:
    folder = filedialog.askdirectory(title="Select folder to save output", parent=_dialog_parent())
    if not folder:
        raise RuntimeError("No output folder selected.")
    return folder


def ask_pdf_file(title="Select PDF file") -> str:
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        parent=_dialog_parent(),
    )
    if not file_path:
        raise RuntimeError("No PDF file selected.")
    return file_path


def ask_pdf_files(title="Select PDF files") -> List[str]:
    files = filedialog.askopenfilenames(
        title=title,
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        parent=_dialog_parent(),
    )
    if not files:
        raise RuntimeError("No PDF files selected.")
    return list(files)


def ask_ordered_pdf_files(title: str = "Select PDFs to merge") -> List[str]:
    """
    Pick PDFs, then let the user set merge order.
    Windows file dialogs ignore click order and return names sorted by the folder.
    """
    files = ask_pdf_files(title)
    ordered = _reorder_files_dialog(
        files,
        window_title="Set merge order",
        hint="Page 1 of the merged PDF will be the first file in this list.",
    )
    if not ordered:
        raise RuntimeError("Merge cancelled.")
    return ordered


def _reorder_files_dialog(paths: List[str], window_title: str, hint: str) -> List[str]:
    parent = _dialog_parent()
    win = tk.Toplevel(parent) if parent is not None else tk.Toplevel()
    win.title(window_title)
    win.transient(parent)
    win.grab_set()
    win.resizable(True, True)
    win.geometry("560x420")
    result: List[str] = []
    items = list(paths)

    ttk.Label(
        win,
        text=hint + "\nSelect a file, then Move up or Move down. Number 1 is first in the merged PDF.",
        wraplength=520,
        justify="left",
    ).pack(anchor="w", padx=14, pady=(12, 6))

    list_frame = ttk.Frame(win)
    list_frame.pack(fill="both", expand=True, padx=14, pady=4)
    scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
    box = tk.Listbox(list_frame, selectmode="browse", height=12, yscrollcommand=scrollbar.set)
    scrollbar.config(command=box.yview)
    box.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def refresh(select_index: int | None = None):
        box.delete(0, tk.END)
        for i, p in enumerate(items, start=1):
            box.insert(tk.END, f"{i}.  {os.path.basename(p)}")
        if items:
            idx = 0 if select_index is None else max(0, min(select_index, len(items) - 1))
            box.selection_set(idx)
            box.see(idx)

    def selected_index() -> int:
        sel = box.curselection()
        return int(sel[0]) if sel else 0

    def move(delta: int):
        if len(items) < 2:
            return
        i = selected_index()
        j = i + delta
        if j < 0 or j >= len(items):
            return
        items[i], items[j] = items[j], items[i]
        refresh(j)

    def add_more():
        extra = filedialog.askopenfilenames(
            title="Add more PDFs",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            parent=win,
        )
        if not extra:
            return
        for p in extra:
            if p not in items:
                items.append(p)
        refresh(len(items) - 1)

    def remove_sel():
        if not items:
            return
        i = selected_index()
        items.pop(i)
        refresh(min(i, len(items) - 1) if items else None)

    def confirm():
        if len(items) < 2:
            messagebox.showerror("Merge PDF", "Keep at least 2 PDF files.", parent=win)
            return
        result.extend(items)
        win.destroy()

    def cancel():
        win.destroy()

    edit_row = ttk.Frame(win)
    edit_row.pack(fill="x", padx=14, pady=(8, 4))
    ttk.Button(edit_row, text="Move up", command=lambda: move(-1)).pack(side="left", padx=(0, 6))
    ttk.Button(edit_row, text="Move down", command=lambda: move(1)).pack(side="left", padx=(0, 6))
    ttk.Button(edit_row, text="Add PDFs", command=add_more).pack(side="left", padx=(0, 6))
    ttk.Button(edit_row, text="Remove", command=remove_sel).pack(side="left")

    next_row = ttk.Frame(win)
    next_row.pack(fill="x", padx=14, pady=(8, 14))
    ttk.Button(next_row, text="Cancel", command=cancel).pack(side="right")
    ttk.Button(next_row, text="Next", command=confirm).pack(side="right", padx=(0, 8))
    ttk.Label(next_row, text="When the list order is correct, click Next.").pack(side="left")

    refresh(0)
    win.protocol("WM_DELETE_WINDOW", cancel)
    win.wait_window()
    return result


def ask_text(title: str, prompt: str, initialvalue: str = "") -> str:
    value = simpledialog.askstring(
        title, prompt, initialvalue=initialvalue, parent=_dialog_parent()
    )
    if value is None or value.strip() == "":
        raise RuntimeError("No input provided.")
    return value.strip()


def ask_int(title: str, prompt: str, minvalue=None, maxvalue=None) -> int:
    value = simpledialog.askinteger(title, prompt, minvalue=minvalue, maxvalue=maxvalue)
    if value is None:
        raise RuntimeError("No input provided.")
    return int(value)


def parse_ranges(range_text: str, total_pages: int) -> List[Tuple[int, int]]:
    """
    Parses user page range input.
    Example: "1, 3-5, 8-10" returns [(1,1), (3,5), (8,10)]
    Page numbers are 1-based for user convenience.
    """
    if not range_text or not range_text.strip():
        raise ValueError("Page range cannot be blank.")

    ranges = []
    parts = [p.strip() for p in range_text.split(",") if p.strip()]
    for part in parts:
        if "-" in part:
            start_s, end_s = [x.strip() for x in part.split("-", 1)]
            if not start_s.isdigit() or not end_s.isdigit():
                raise ValueError(f"Invalid page range: {part}")
            start, end = int(start_s), int(end_s)
        else:
            if not part.isdigit():
                raise ValueError(f"Invalid page number: {part}")
            start = end = int(part)

        if start < 1 or end < 1 or start > total_pages or end > total_pages:
            raise ValueError(f"Page range {part} is outside 1-{total_pages}.")
        if start > end:
            raise ValueError(f"Page range start cannot be greater than end: {part}")
        ranges.append((start, end))

    if not ranges:
        raise ValueError("No valid page ranges found.")
    return ranges


def add_pages_by_ranges(reader, writer, ranges: List[Tuple[int, int]]):
    for start, end in ranges:
        for page_no in range(start - 1, end):
            writer.add_page(reader.pages[page_no])


def write_pdf(writer, output_path: str):
    with open(output_path, "wb") as f:
        writer.write(f)


def open_folder(path: str):
    try:
        if sys.platform == "darwin":
            os.system(f'open "{path}"')
        elif os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            os.system(f'xdg-open "{path}"')
    except Exception:
        pass


# --------------------------- PDF Operations ---------------------------

def _temp_pdf_path(prefix: str = "ssa_pdf_") -> str:
    import tempfile

    tmp = tempfile.NamedTemporaryFile(prefix=prefix, suffix=".pdf", delete=False)
    tmp.close()
    return tmp.name


def _unlock_pymupdf(doc, src_path: str) -> None:
    if not (getattr(doc, "needs_pass", False) or doc.is_encrypted):
        return
    if doc.authenticate(""):
        return
    pwd = ask_text(
        "PDF password",
        f"This PDF is password-protected.\nEnter password for:\n{os.path.basename(src_path)}",
    )
    if not doc.authenticate(pwd):
        raise RuntimeError(f"Could not unlock PDF (wrong password):\n{src_path}")


def _plain_copy_for_merge(src_path: str) -> str:
    """
    Make a local unencrypted copy. insert_pdf() fails on PDFs that stay
    marked encrypted even after authenticate('') — common for bank statements.
    """
    use_path, is_net_tmp = _local_pdf_copy(src_path)
    dest = _temp_pdf_path("ssa_plain_")
    try:
        try:
            import pymupdf

            doc = pymupdf.open(use_path)
            try:
                _unlock_pymupdf(doc, src_path)
                enc_none = getattr(pymupdf, "PDF_ENCRYPT_NONE", 1)
                saved = False
                for kwargs in (
                    {"garbage": 4, "deflate": True, "encryption": enc_none},
                    {"encryption": enc_none},
                    {"garbage": 4, "deflate": True},
                    {},
                ):
                    try:
                        doc.save(dest, **kwargs)
                        saved = True
                        break
                    except Exception:
                        continue
                if not saved:
                    raise RuntimeError("pymupdf could not write a decrypted copy")
            finally:
                doc.close()
            return dest
        except RuntimeError as exc:
            msg = str(exc).lower()
            if "password" in msg or "unlock" in msg:
                raise
        except Exception:
            pass

        reader = PdfReader(use_path)
        if reader.is_encrypted:
            try:
                ok = reader.decrypt("")
            except Exception:
                ok = False
            if not ok:
                pwd = ask_text(
                    "PDF password",
                    f"This PDF is password-protected.\nEnter password for:\n{os.path.basename(src_path)}",
                )
                try:
                    ok = reader.decrypt(pwd)
                except Exception:
                    ok = False
            if not ok:
                raise RuntimeError(f"Could not unlock PDF (wrong password):\n{src_path}")
        writer = PdfWriter()
        if hasattr(writer, "append"):
            writer.append(reader)
        else:
            for page in reader.pages:
                writer.add_page(page)
        write_pdf(writer, dest)
        return dest
    except Exception:
        try:
            os.remove(dest)
        except OSError:
            pass
        raise
    finally:
        if is_net_tmp:
            try:
                os.remove(use_path)
            except OSError:
                pass


def _save_merged_to(path: str, saver) -> None:
    import shutil

    target = path
    tmp = None
    if _is_network_path(path):
        tmp = _temp_pdf_path("ssa_merged_")
        target = tmp
    try:
        saver(target)
        if tmp:
            shutil.copy2(tmp, path)
    finally:
        if tmp:
            try:
                os.remove(tmp)
            except OSError:
                pass


def _write_merged_pdfs(plain_paths: List[str], out: str) -> None:
    last_err = None
    try:
        import pymupdf

        merged = pymupdf.open()
        try:
            for p in plain_paths:
                src = pymupdf.open(p)
                try:
                    merged.insert_pdf(src)
                finally:
                    src.close()

            def _save(target: str):
                try:
                    merged.save(target, garbage=4, deflate=True)
                except Exception:
                    merged.save(target)

            _save_merged_to(out, _save)
            return
        finally:
            merged.close()
    except Exception as exc:
        last_err = exc

    try:
        writer = PdfWriter()
        for p in plain_paths:
            if hasattr(writer, "append"):
                writer.append(p)
            else:
                reader = PdfReader(p)
                for page in reader.pages:
                    writer.add_page(page)

        def _save_writer(target: str):
            write_pdf(writer, target)

        _save_merged_to(out, _save_writer)
    except Exception as exc:
        detail = f"{last_err}\n{exc}" if last_err else str(exc)
        raise RuntimeError(f"Could not merge PDFs.\n{detail}") from exc


def merge_pdfs():
    ensure_pdf_tools()
    pdf_files = ask_ordered_pdf_files("Select PDFs to merge")
    if len(pdf_files) < 2:
        raise RuntimeError("Please select at least 2 PDF files to merge.")
    output_folder = ask_output_folder()
    output_name = ask_text("Output file name", "Enter output PDF file name:", "merged.pdf")
    if not output_name.lower().endswith(".pdf"):
        output_name += ".pdf"
    out = os.path.join(output_folder, safe_filename(output_name))

    plains: List[str] = []
    try:
        for pdf in pdf_files:
            plains.append(_plain_copy_for_merge(pdf))
        _write_merged_pdfs(plains, out)
        return f"Merged PDF created:\n{out}", output_folder
    finally:
        for temp_path in plains:
            try:
                os.remove(temp_path)
            except OSError:
                pass



def split_pdf():
    ensure_pdf_tools()
    pdf_path = ask_pdf_file("Select PDF to split")
    reader = PdfReader(pdf_path)
    if reader.is_encrypted:
        raise RuntimeError("Please unlock/decrypt this PDF first.")
    total = len(reader.pages)
    ranges = parse_ranges(
        ask_text("Split page ranges", f"PDF has {total} pages. Enter ranges, e.g. 1,3-5,8-10:"),
        total,
    )
    output_folder = ask_output_folder()
    base = Path(pdf_path).stem

    created = []
    for index, page_range in enumerate(ranges, start=1):
        writer = PdfWriter()
        add_pages_by_ranges(reader, writer, [page_range])
        start, end = page_range
        out = os.path.join(output_folder, f"{safe_filename(base)}_split_{index}_{start}-{end}.pdf")
        write_pdf(writer, out)
        created.append(out)

    return "Split PDF files created:\n" + "\n".join(created), output_folder


def extract_pages():
    ensure_pdf_tools()
    pdf_path = ask_pdf_file("Select PDF to extract pages from")
    reader = PdfReader(pdf_path)
    if reader.is_encrypted:
        raise RuntimeError("Please unlock/decrypt this PDF first.")
    total = len(reader.pages)
    ranges = parse_ranges(
        ask_text("Extract pages", f"PDF has {total} pages. Enter pages/ranges, e.g. 1,3-5,8-10:"),
        total,
    )
    output_folder = ask_output_folder()
    output_name = ask_text("Output file name", "Enter output PDF file name:", "extracted_pages.pdf")
    if not output_name.lower().endswith(".pdf"):
        output_name += ".pdf"

    writer = PdfWriter()
    add_pages_by_ranges(reader, writer, ranges)
    out = os.path.join(output_folder, safe_filename(output_name))
    write_pdf(writer, out)
    return f"Extracted pages PDF created:\n{out}", output_folder


def remove_pages():
    ensure_pdf_tools()
    pdf_path = ask_pdf_file("Select PDF to remove pages from")
    reader = PdfReader(pdf_path)
    if reader.is_encrypted:
        raise RuntimeError("Please unlock/decrypt this PDF first.")
    total = len(reader.pages)
    ranges = parse_ranges(
        ask_text("Remove pages", f"PDF has {total} pages. Enter pages/ranges to remove, e.g. 2,4-6:"),
        total,
    )
    remove_set = set()
    for start, end in ranges:
        remove_set.update(range(start - 1, end))

    output_folder = ask_output_folder()
    output_name = ask_text("Output file name", "Enter output PDF file name:", "pages_removed.pdf")
    if not output_name.lower().endswith(".pdf"):
        output_name += ".pdf"

    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if i not in remove_set:
            writer.add_page(page)
    out = os.path.join(output_folder, safe_filename(output_name))
    write_pdf(writer, out)
    return f"PDF with selected pages removed created:\n{out}", output_folder


def rotate_pdf():
    ensure_pdf_tools()
    pdf_path = ask_pdf_file("Select PDF to rotate")
    reader = PdfReader(pdf_path)
    if reader.is_encrypted:
        raise RuntimeError("Please unlock/decrypt this PDF first.")
    total = len(reader.pages)
    angle = ask_int("Rotation angle", "Enter rotation angle: 90, 180 or 270", minvalue=0, maxvalue=360)
    if angle not in (90, 180, 270):
        raise RuntimeError("Rotation angle must be 90, 180 or 270.")
    range_text = ask_text("Pages to rotate", f"PDF has {total} pages. Enter ranges or 'all':", "all")
    if range_text.lower() == "all":
        rotate_set = set(range(total))
    else:
        ranges = parse_ranges(range_text, total)
        rotate_set = set()
        for start, end in ranges:
            rotate_set.update(range(start - 1, end))

    output_folder = ask_output_folder()
    output_name = ask_text("Output file name", "Enter output PDF file name:", "rotated.pdf")
    if not output_name.lower().endswith(".pdf"):
        output_name += ".pdf"

    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if i in rotate_set:
            page.rotate(angle)
        writer.add_page(page)
    out = os.path.join(output_folder, safe_filename(output_name))
    write_pdf(writer, out)
    return f"Rotated PDF created:\n{out}", output_folder


def make_overlay_pdf(text: str, output_overlay: str, page_width: float, page_height: float, mode: str):
    require_package(canvas is not None, "python3 -m pip install reportlab")
    c = canvas.Canvas(output_overlay, pagesize=(page_width, page_height))
    c.setFont("Helvetica", 12 if mode == "page_number" else 48)
    c.setFillGray(0.45, alpha=0.25 if mode == "watermark" else 0.85)

    if mode == "watermark":
        c.saveState()
        c.translate(page_width / 2, page_height / 2)
        c.rotate(35)
        c.drawCentredString(0, 0, text)
        c.restoreState()
    else:
        c.drawRightString(page_width - 18 * mm, 12 * mm, text)

    c.save()


def add_text_watermark():
    ensure_pdf_tools()
    require_package(canvas is not None, "python3 -m pip install reportlab")
    pdf_path = ask_pdf_file("Select PDF for watermark")
    text = ask_text("Watermark", "Enter watermark text:", "CONFIDENTIAL")
    output_folder = ask_output_folder()
    output_name = ask_text("Output file name", "Enter output PDF file name:", "watermarked.pdf")
    if not output_name.lower().endswith(".pdf"):
        output_name += ".pdf"

    reader = PdfReader(pdf_path)
    if reader.is_encrypted:
        raise RuntimeError("Please unlock/decrypt this PDF first.")
    writer = PdfWriter()

    for idx, page in enumerate(reader.pages):
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        overlay_path = os.path.join(output_folder, f"__overlay_watermark_{idx}.pdf")
        make_overlay_pdf(text, overlay_path, width, height, "watermark")
        overlay = PdfReader(overlay_path).pages[0]
        page.merge_page(overlay)
        writer.add_page(page)
        try:
            os.remove(overlay_path)
        except OSError:
            pass

    out = os.path.join(output_folder, safe_filename(output_name))
    write_pdf(writer, out)
    return f"Watermarked PDF created:\n{out}", output_folder


def add_page_numbers():
    ensure_pdf_tools()
    require_package(canvas is not None, "python3 -m pip install reportlab")
    pdf_path = ask_pdf_file("Select PDF for page numbering")
    output_folder = ask_output_folder()
    output_name = ask_text("Output file name", "Enter output PDF file name:", "page_numbered.pdf")
    if not output_name.lower().endswith(".pdf"):
        output_name += ".pdf"

    reader = PdfReader(pdf_path)
    if reader.is_encrypted:
        raise RuntimeError("Please unlock/decrypt this PDF first.")
    total = len(reader.pages)
    writer = PdfWriter()

    for idx, page in enumerate(reader.pages, start=1):
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        overlay_path = os.path.join(output_folder, f"__overlay_page_no_{idx}.pdf")
        make_overlay_pdf(f"Page {idx} of {total}", overlay_path, width, height, "page_number")
        overlay = PdfReader(overlay_path).pages[0]
        page.merge_page(overlay)
        writer.add_page(page)
        try:
            os.remove(overlay_path)
        except OSError:
            pass

    out = os.path.join(output_folder, safe_filename(output_name))
    write_pdf(writer, out)
    return f"Page-numbered PDF created:\n{out}", output_folder


def protect_pdf():
    ensure_pdf_tools()
    pdf_path = ask_pdf_file("Select PDF to protect")
    password = ask_text("Password", "Enter password for PDF:")
    output_folder = ask_output_folder()
    output_name = ask_text("Output file name", "Enter output PDF file name:", "protected.pdf")
    if not output_name.lower().endswith(".pdf"):
        output_name += ".pdf"

    reader = PdfReader(pdf_path)
    if reader.is_encrypted:
        raise RuntimeError("PDF is already encrypted. Unlock first if needed.")
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(password)
    out = os.path.join(output_folder, safe_filename(output_name))
    write_pdf(writer, out)
    return f"Protected PDF created:\n{out}", output_folder


def unlock_pdf():
    ensure_pdf_tools()
    pdf_path = ask_pdf_file("Select password-protected PDF")
    password = ask_text("Password", "Enter existing PDF password:")
    output_folder = ask_output_folder()
    output_name = ask_text("Output file name", "Enter output PDF file name:", "unlocked.pdf")
    if not output_name.lower().endswith(".pdf"):
        output_name += ".pdf"

    reader = PdfReader(pdf_path)
    if reader.is_encrypted:
        ok = reader.decrypt(password)
        if not ok:
            raise RuntimeError("Password is incorrect or PDF cannot be decrypted.")
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    out = os.path.join(output_folder, safe_filename(output_name))
    write_pdf(writer, out)
    return f"Unlocked PDF created:\n{out}", output_folder


def _fmt_kb(n: int) -> str:
    return f"{(n / 1024.0):.1f} KB"


def _strip_pdf_metadata(doc) -> None:
    try:
        doc.set_metadata({})
    except Exception:
        pass
    try:
        doc.del_xml_metadata()
    except Exception:
        pass


def _save_optimized_pdf(doc, path: str) -> None:
    attempts = (
        dict(
            garbage=4,
            deflate=True,
            deflate_images=True,
            deflate_fonts=True,
            clean=True,
            pretty=False,
            use_objstms=1,
        ),
        dict(garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True),
        dict(garbage=4, deflate=True, deflate_images=True, deflate_fonts=True),
        dict(garbage=4, deflate=True),
    )
    last_err = None
    for kwargs in attempts:
        try:
            doc.save(path, **kwargs)
            return
        except Exception as exc:
            last_err = exc
    if last_err:
        raise last_err
    raise RuntimeError("Could not save compressed PDF.")


def _pixmap_for_jpeg(pix, pymupdf):
    """RGB or gray pixmap without alpha, suitable for JPEG."""
    if pix.alpha:
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    elif pix.n >= 4:
        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
    return pix


def _recompress_page_images(doc, pymupdf, max_dpi: int = 150, jpeg_quality: int = 50) -> int:
    """
    Downsample images sharper than max_dpi and recompress as JPEG.
    Returns how many images were replaced.
    """
    import io

    replaced = 0
    seen = set()
    for page in doc:
        for img in page.get_images(full=True):
            xref = img[0]
            if xref in seen or xref <= 0:
                continue
            seen.add(xref)
            try:
                pix = pymupdf.Pixmap(doc, xref)
            except Exception:
                continue
            if pix.width < 32 or pix.height < 32:
                continue
            # 1-bit / tiny fax streams often get bigger as JPEG
            try:
                info = doc.extract_image(xref) or {}
                raw = info.get("image") or b""
            except Exception:
                try:
                    raw = doc.xref_stream(xref) or b""
                except Exception:
                    raw = b""
            if pix.n == 1 and len(raw) < 8000:
                continue

            display_w = display_h = 0.0
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            if rects:
                display_w = float(rects[0].width)
                display_h = float(rects[0].height)
            dpi_x = (pix.width * 72.0 / display_w) if display_w > 1 else float(max_dpi)
            dpi_y = (pix.height * 72.0 / display_h) if display_h > 1 else float(max_dpi)
            dpi = max(dpi_x, dpi_y)
            scale = 1.0
            if dpi > max_dpi + 5:
                scale = max_dpi / dpi

            try:
                pix = _pixmap_for_jpeg(pix, pymupdf)
                new_w = max(1, int(round(pix.width * scale)))
                new_h = max(1, int(round(pix.height * scale)))
                jpeg = b""
                if Image is not None:
                    mode = "L" if pix.n == 1 else "RGB"
                    if pix.n not in (1, 3):
                        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                        mode = "RGB"
                    img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
                    if (new_w, new_h) != (pix.width, pix.height):
                        img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                    if img.mode not in ("RGB", "L"):
                        img = img.convert("RGB")
                    buf = io.BytesIO()
                    img.save(buf, format="JPEG", quality=jpeg_quality, optimize=True)
                    jpeg = buf.getvalue()
                else:
                    if scale < 0.98:
                        n = 0
                        w, h = pix.width, pix.height
                        while n < 4 and (w // 2) >= new_w and (h // 2) >= new_h:
                            n += 1
                            w //= 2
                            h //= 2
                        if n:
                            try:
                                pix.shrink(n)
                            except Exception:
                                pass
                    jpeg = pix.tobytes("jpeg", jpg_quality=jpeg_quality)
                if not jpeg:
                    continue
                if raw and len(jpeg) >= len(raw) and scale >= 0.98:
                    continue
                page.replace_image(xref, stream=jpeg)
                replaced += 1
            except Exception:
                continue
    return replaced


def _rebuild_pdf_as_jpeg_pages(src_path: str, pymupdf, dpi: int = 150, jpeg_quality: int = 45) -> str:
    """Last-resort shrink for scanned/image-only PDFs. Drops selectable text."""
    src = pymupdf.open(src_path)
    out_path = _temp_pdf_path("ssa_jpgpdf_")
    dst = pymupdf.open()
    try:
        zoom = dpi / 72.0
        matrix = pymupdf.Matrix(zoom, zoom)
        for page in src:
            pix = page.get_pixmap(matrix=matrix, alpha=False, colorspace=pymupdf.csRGB)
            jpeg = pix.tobytes("jpeg", jpg_quality=jpeg_quality)
            new_page = dst.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(new_page.rect, stream=jpeg)
        _save_optimized_pdf(dst, out_path)
    finally:
        dst.close()
        src.close()
    return out_path


def compress_pdf():
    """Compress images, drop unused objects, subset fonts, strip metadata."""
    try:
        import pymupdf
    except Exception:
        raise RuntimeError("Please install PyMuPDF: python3 -m pip install pymupdf")

    pdf_path = ask_pdf_file("Select PDF to compress")
    output_folder = ask_output_folder()
    output_name = ask_text("Output file name", "Enter output PDF file name:", "compressed.pdf")
    if not output_name.lower().endswith(".pdf"):
        output_name += ".pdf"
    out = os.path.join(output_folder, safe_filename(output_name))

    use_path, is_temp = _local_pdf_copy(pdf_path)
    temps = [use_path] if is_temp else []
    original_size = os.path.getsize(pdf_path)
    notes = []
    try:
        work = _temp_pdf_path("ssa_cmpwork_")
        temps.append(work)
        doc = pymupdf.open(use_path)
        try:
            _unlock_pymupdf(doc, pdf_path)
            enc_none = getattr(pymupdf, "PDF_ENCRYPT_NONE", 1)
            try:
                doc.save(work, garbage=4, deflate=True, encryption=enc_none)
            except Exception:
                doc.save(work)
        finally:
            doc.close()

        doc = pymupdf.open(work)
        try:
            rewritten = 0
            try:
                rewritten = doc.rewrite_images(
                    dpi_threshold=200,
                    dpi_target=150,
                    quality=50,
                    lossy=True,
                    lossless=True,
                    bitonal=True,
                    color=True,
                    gray=True,
                )
                if rewritten:
                    notes.append("images recompressed to 150 DPI / JPEG 50")
            except Exception:
                rewritten = 0
            manual = _recompress_page_images(doc, pymupdf, max_dpi=150, jpeg_quality=50)
            if manual and not rewritten:
                notes.append("images downsampled to 150 DPI and JPEG-compressed")
            elif manual:
                notes.append(f"{manual} extra image(s) recompressed")

            try:
                doc.subset_fonts()
                notes.append("fonts subset")
            except Exception:
                try:
                    doc.subset_fonts(fallback=True)
                    notes.append("fonts subset")
                except Exception:
                    pass

            _strip_pdf_metadata(doc)
            notes.append("metadata stripped")
            notes.append("duplicate/unused objects removed")

            compact = _temp_pdf_path("ssa_cmpout_")
            temps.append(compact)
            _save_optimized_pdf(doc, compact)
        finally:
            doc.close()

        best = compact
        best_size = os.path.getsize(compact)

        # Scanned PDFs: rebuild pages as 150 DPI JPEGs if that is smaller
        try:
            chars = _pdf_extractable_char_count(work)
        except Exception:
            chars = 50
        if chars < 80 and best_size >= original_size * 0.98:
            try:
                raster = _rebuild_pdf_as_jpeg_pages(work, pymupdf, dpi=150, jpeg_quality=45)
                temps.append(raster)
                raster_size = os.path.getsize(raster)
                if raster_size < best_size:
                    best = raster
                    best_size = raster_size
                    notes.append("pages rebuilt as 150 DPI JPEG")
            except Exception:
                pass

        import shutil

        if best_size >= original_size:
            shutil.copy2(pdf_path, out)
            msg = (
                f"Could not shrink this PDF (already compact).\n"
                f"{out}\n\n"
                f"Original:    {_fmt_kb(original_size)}\n"
                f"Compressed:  {_fmt_kb(best_size)} (not used)\n"
                f"Kept original size: {_fmt_kb(original_size)}"
            )
            return msg, output_folder

        shutil.copy2(best, out)
        saved = original_size - best_size
        pct = (saved * 100.0 / original_size) if original_size else 0
        applied = ", ".join(notes) if notes else "standard optimization"
        msg = (
            f"Compressed PDF created:\n{out}\n\n"
            f"Original:    {_fmt_kb(original_size)}\n"
            f"Compressed:  {_fmt_kb(best_size)}\n"
            f"Saved:       {_fmt_kb(saved)} ({pct:.0f}%)\n\n"
            f"Applied: {applied}."
        )
        return msg, output_folder
    finally:
        for temp_path in temps:
            try:
                if os.path.abspath(temp_path) != os.path.abspath(out):
                    os.remove(temp_path)
            except OSError:
                pass


def pdf_to_images():
    try:
        import pymupdf  # PyMuPDF
    except Exception:
        raise RuntimeError("Please install PyMuPDF: python3 -m pip install pymupdf")

    pdf_path = ask_pdf_file("Select PDF to convert to images")
    output_folder = ask_output_folder()
    dpi = ask_int("Image quality", "Enter DPI, e.g. 150 or 200:", minvalue=72, maxvalue=600)
    doc = pymupdf.open(pdf_path)
    zoom = dpi / 72
    matrix = pymupdf.Matrix(zoom, zoom)
    base = safe_filename(Path(pdf_path).stem)
    created = []
    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        out = os.path.join(output_folder, f"{base}_page_{i}.png")
        pix.save(out)
        created.append(out)
    doc.close()
    return "Images created:\n" + "\n".join(created), output_folder


def images_to_pdf():
    require_package(Image is not None, "python3 -m pip install pillow")
    image_files = filedialog.askopenfilenames(
        title="Select image files",
        filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff")],
    )
    if not image_files:
        raise RuntimeError("No image files selected.")
    output_folder = ask_output_folder()
    output_name = ask_text("Output file name", "Enter output PDF file name:", "images_to_pdf.pdf")
    if not output_name.lower().endswith(".pdf"):
        output_name += ".pdf"
    out = os.path.join(output_folder, safe_filename(output_name))

    image_list = []
    for f in image_files:
        img = Image.open(f).convert("RGB")
        image_list.append(img)
    first, rest = image_list[0], image_list[1:]
    first.save(out, save_all=True, append_images=rest)
    return f"Image PDF created:\n{out}", output_folder


def _map_pdf_font(font_name: str) -> str:
    raw = (font_name or "").split("+")[-1]
    low = raw.lower()
    if "times" in low:
        return "Times New Roman"
    if "helv" in low or "arial" in low:
        return "Arial"
    if "courier" in low:
        return "Courier New"
    if "calibri" in low:
        return "Calibri"
    if "georgia" in low:
        return "Georgia"
    if "garamond" in low:
        return "Garamond"
    if "verdana" in low:
        return "Verdana"
    if "cambria" in low:
        return "Cambria"
    if "palatino" in low:
        return "Palatino Linotype"
    cleaned = re.sub(
        r"[-,]?(Bold|Italic|Oblique|Regular|Medium|Light|Black|SemiBold|MT)$",
        "",
        raw,
        flags=re.I,
    ).strip()
    return cleaned or "Calibri"


def _pdf_span_style(span: dict) -> Tuple[str, float, bool, bool, int]:
    font = _map_pdf_font(span.get("font") or "")
    size = float(span.get("size") or 11)
    flags = int(span.get("flags") or 0)
    fname = (span.get("font") or "").lower()
    bold = bool(flags & 16) or ("bold" in fname)
    italic = bool(flags & 2) or ("italic" in fname) or ("oblique" in fname)
    color = int(span.get("color") or 0)
    return font, size, bold, italic, color


def _apply_run_style(run, font: str, size: float, bold: bool, italic: bool, color: int) -> None:
    from docx.shared import Pt, RGBColor

    run.bold = bold
    run.italic = italic
    try:
        run.font.name = font
        run.font.size = Pt(max(8, min(size, 48)))
        r, g, b = (color >> 16) & 255, (color >> 8) & 255, color & 255
        if (r, g, b) != (0, 0, 0):
            run.font.color.rgb = RGBColor(r, g, b)
    except Exception:
        pass


def _word_pt(size: float, body: bool = True) -> float:
    """PDF point sizes look larger in Word; keep body text statement-sized."""
    if body:
        return max(7.0, min(float(size) * 0.88, 11.0))
    return max(9.0, min(float(size) * 0.9, 14.0))


def _compact_paragraph(para) -> None:
    from docx.shared import Pt

    fmt = para.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(1)
    fmt.line_spacing = 1.0


def _word_token_is_noise(text: str) -> bool:
    t = (text or "").strip()
    if not t:
        return True
    low = t.lower()
    if "uncerstand" in low:
        return True
    if low in {"h+", "h +", "'yy,", "yy,", "yy", "'yy", "di", "un", "syno\\", "syno"}:
        return True
    if re.fullmatch(r"[^\w]+", t) and len(t) <= 3:
        return True
    # OCR / broken-font junk (logo fragments, ffff runs, brackets)
    if "\\" in t and not re.search(r"\d", t):
        return True
    if re.search(r"(.)\1{3,}", t):
        return True
    if re.search(r"f{3,}", t, re.I):
        return True
    if re.search(r"[\[\]]", t) and not re.search(r"\d", t):
        return True
    return False


def _word_is_footer_line(text: str) -> bool:
    low = (text or "").lower()
    hints = (
        "computer generated",
        "registered office",
        "corporate office",
        "unless the constituent",
        "authorised signatory",
        "authorized signatory",
        "this is a system generated",
        "for hdfc bank",
        "hdfc bank ltd",
        "cin:",
        "cin :",
        "gstin",
        "website:",
        "www.hdfcbank",
    )
    return any(h in low for h in hints)


def _set_word_table_borders(table, visible: bool, horizontal_only: bool = False) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    tbl = table._tbl
    tbl_pr = tbl.tblPr
    if tbl_pr is None:
        return
    for child in list(tbl_pr):
        if child.tag == qn("w:tblBorders"):
            tbl_pr.remove(child)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        if not visible:
            el.set(qn("w:val"), "nil")
            el.set(qn("w:sz"), "0")
        elif horizontal_only and edge in ("left", "right", "insideV"):
            el.set(qn("w:val"), "nil")
            el.set(qn("w:sz"), "0")
        else:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tbl_pr.append(borders)


def _style_word_cell(cell, font_pt: float = 8.0, bold: bool = False, align=None) -> None:
    from docx.shared import Pt

    for para in cell.paragraphs:
        _compact_paragraph(para)
        if align is not None:
            try:
                para.alignment = align
            except Exception:
                pass
        if not para.runs and (cell.text or "").strip():
            para.add_run(cell.text)
        for run in para.runs:
            _apply_word_run_font(run, run.text or cell.text, font_pt, bold=bold)


def _split_visual_line_two_cols(line, page_width: float):
    if not line:
        return None
    if len(line) == 1:
        y, x0, x1, text = line[0]
        if x0 > page_width * 0.48:
            return "", text
        return text, ""
    best_i, best_gap = 0, -1.0
    for i in range(len(line) - 1):
        gap = line[i + 1][1] - line[i][2]
        if gap > best_gap:
            best_gap, best_i = gap, i
    if best_gap >= 28 and line[best_i][2] < page_width * 0.70:
        left = " ".join(t[3] for t in line[: best_i + 1]).strip()
        right = " ".join(t[3] for t in line[best_i + 1 :]).strip()
        return left, right
    text = " ".join(t[3] for t in line).strip()
    x0, x1 = line[0][1], line[-1][2]
    if x0 > page_width * 0.48:
        return "", text
    return text, ""


def _add_flow_lines(docx, lines, font_pt: float = 9.0) -> None:
    """One visual PDF line → one Word paragraph. Do not split into left/right columns."""
    for line in lines or []:
        text = " ".join(t[3] for t in line).strip()
        if text:
            _add_plain_line(docx, text, font_pt=font_pt)


def _find_simple_txn_header(lines):
    """Looser header detect for HDFC-style Date / Narration / Amount / Balance rows."""
    for i, line in enumerate(lines):
        joined = " ".join(t[3] for t in line)
        low = joined.lower()
        if not any(k in low for k in ("date", "value dt", "txn dt", "tran dt")):
            continue
        if not any(
            k in low
            for k in (
                "narration",
                "particular",
                "description",
                "remark",
                "detail",
                "chq",
                "ref",
            )
        ):
            if not any(k in low for k in ("debit", "credit", "withdraw", "deposit")):
                continue
        if not any(k in low for k in ("withdrawal", "deposit", "balance", "debit", "credit", "amount", "amt")):
            continue
        titles, bounds = _learn_columns_from_header_lines([line])
        if len(titles) >= 3 and len(bounds) >= 4:
            return i, i + 1, titles, bounds
    return -1, -1, [], []


def _row_starts_new_txn(cells: List[str], titles: List[str]) -> bool:
    for i, h in enumerate(titles or []):
        hl = (h or "").lower()
        if "date" in hl and "value" not in hl and i < len(cells):
            if _is_date_text(cells[i]):
                return True
    if cells and _is_date_text(cells[0]):
        return True
    return False


def _add_two_col_layout(docx, lines, page_width: float, font_pt: float = 8.5) -> None:
    if not lines:
        return
    rows = [_split_visual_line_two_cols(line, page_width) for line in lines]
    rows = [r for r in rows if r and ((r[0] or "").strip() or (r[1] or "").strip())]
    if not rows:
        return
    table = docx.add_table(rows=len(rows), cols=2)
    try:
        table.autofit = True
    except Exception:
        pass
    _set_word_table_borders(table, visible=False)
    for i, (left, right) in enumerate(rows):
        table.rows[i].cells[0].text = left or ""
        table.rows[i].cells[1].text = right or ""
        _style_word_cell(table.rows[i].cells[0], font_pt=font_pt, bold=False)
        _style_word_cell(table.rows[i].cells[1], font_pt=font_pt, bold=False)
        try:
            table.rows[i].cells[1].paragraphs[0].alignment = 2  # right
        except Exception:
            pass


def _filter_cluster_lines(lines, skip_bboxes=None):
    skip_bboxes = skip_bboxes or []
    out = []
    for line in lines:
        kept = []
        for y, x0, x1, text in line:
            if _word_token_is_noise(text):
                continue
            box = (x0, y, x1, y + 8)
            if any(_bbox_overlap_frac(box, sb) > 0.45 for sb in skip_bboxes):
                continue
            kept.append((y, x0, x1, text))
        if kept:
            out.append(kept)
    return out


def _statement_table_from_words(page, words=None, layout=None):
    """Build a Word table from a bank-statement header row when find_tables fails."""
    if words is None:
        words = page.get_text("words") or []
    lines = _cluster_page_word_lines(words)
    if not lines:
        return None
    start, end, titles, bounds = _find_header_window(lines)
    if not titles or len(titles) < 4 or len(bounds) < 5:
        start, end, titles, bounds = _find_simple_txn_header(lines)
    carried = layout or {}
    if (not titles or len(bounds) < 4) and carried.get("titles") and carried.get("bounds"):
        titles = list(carried["titles"])
        bounds = list(carried["bounds"])
        start, end = 0, 0
        for li, line in enumerate(lines):
            if any(_is_date_text(t[3]) for t in line):
                start = li
                end = li
                break
        header_lines = lines[:start]
    elif titles and bounds:
        header_lines = lines[:start]
        if layout is not None:
            layout["titles"] = list(titles)
            layout["bounds"] = list(bounds)
    else:
        return None

    rows = [list(titles)]
    footer_idx = len(lines)
    body_start = end if end > start else start
    for li, line in enumerate(lines[body_start:], start=body_start):
        joined = " ".join(t[3] for t in line)
        if _word_is_footer_line(joined) and li > body_start + 1:
            footer_idx = li
            break
        cells = [""] * len(titles)
        for _y, x0, x1, text in line:
            if _word_token_is_noise(text):
                continue
            ci = _col_index_for_x((x0 + x1) / 2.0, bounds)
            if 0 <= ci < len(cells):
                cells[ci] = (cells[ci] + " " + text).strip()
        if not any(cells):
            continue
        if len(rows) > 1 and not _row_starts_new_txn(cells, titles):
            prev = rows[-1]
            for j, c in enumerate(cells):
                if c:
                    prev[j] = (prev[j] + " " + c).strip()
        else:
            rows.append(cells)
    if len(rows) < 2:
        return None
    header_y = min((w[0] for w in (lines[start] if start < len(lines) else lines[0])), default=0.0)
    body_lines = lines[body_start:footer_idx]
    last_y = max((w[0] for line in body_lines for w in line), default=header_y) + 10
    bbox = (0.0, header_y - 6.0, float(page.rect.width), last_y)
    return {
        "bbox": bbox,
        "rows": rows,
        "header_lines": header_lines,
        "footer_lines": lines[footer_idx:],
    }


def _heading_level_for_size(size: float, bold: bool, median_size: float) -> int:
    # Kept for compatibility; Word Heading styles inflate fonts, so conversion
    # no longer uses them for body pages.
    return 0


def _bbox_overlap_frac(a, b) -> float:
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    ix0, iy0 = max(ax0, bx0), max(ay0, by0)
    ix1, iy1 = min(ax1, bx1), min(ay1, by1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    inter = (ix1 - ix0) * (iy1 - iy0)
    area = max(1.0, (ax1 - ax0) * (ay1 - ay0))
    return inter / area


def _word_table_fill_ratio(rows: List[List[str]]) -> float:
    cells = [c for r in (rows or []) for c in r]
    if not cells:
        return 0.0
    filled = sum(1 for c in cells if (c or "").strip())
    return filled / float(len(cells))


def _word_table_is_junk(rows: List[List[str]], bbox, page_area: float) -> bool:
    """Reject PyMuPDF grids that smashed the page into empty / merged cells."""
    if not rows or len(rows) < 2:
        return True
    cols = max((len(r) for r in rows), default=0)
    if cols < 2:
        return True
    fill = _word_table_fill_ratio(rows)
    area = max(1.0, (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]))
    area_frac = area / max(1.0, page_area)
    if fill < 0.20:
        return True
    if area_frac > 0.50 and fill < 0.38:
        return True
    if cols >= 10 and fill < 0.32:
        return True
    if _word_table_header_is_fragmented(rows):
        return True
    if not _table_extract_is_faithful(rows):
        return True
    return False


def _word_table_header_is_fragmented(rows: List[List[str]]) -> bool:
    """True when find_tables split Particulars → P | articular | s."""
    if not rows:
        return True
    header = [(c or "").strip() for c in rows[0]]
    if any(len(c) == 1 and c.isalpha() for c in header):
        return True
    joined = " ".join(header).lower()
    compact = re.sub(r"[^a-z]", "", joined)
    if "articular" in compact and "particulars" not in joined.replace(" ", ""):
        return True
    if any(c.lower() in {"cred", "credi", "debi", "o.", "s"} for c in header) and len(header) >= 6:
        return True
    if len(rows) > 1:
        row = [(c or "").strip() for c in rows[1]]
        for a, b in zip(row, row[1:]):
            if not a or not b:
                continue
            if a[-1].isalpha() and b[:1].islower():
                return True
            if a.endswith(" B") and b.upper().startswith("ANK"):
                return True
            if a.endswith(" P") and b.upper().startswith("URCH"):
                return True
    return False


def _drop_overlapping_word_tables(items: list, page_area: float) -> list:
    scored = []
    for item in items:
        bbox = item[3]["bbox"]
        rows = item[3]["rows"]
        fill = _word_table_fill_ratio(rows)
        area = max(1.0, (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]))
        score = fill * 12.0 - (area / max(1.0, page_area)) * 3.0 + min(len(rows), 25) * 0.05
        scored.append((score, item))
    scored.sort(key=lambda t: t[0], reverse=True)
    kept = []
    for _score, item in scored:
        bbox = item[3]["bbox"]
        if any(_bbox_overlap_frac(bbox, k[3]["bbox"]) > 0.55 for k in kept):
            continue
        kept.append(item)
    kept.sort(key=lambda t: (t[0], t[1]))
    return kept


def _table_looks_like_line_items(rows: List[List[str]]) -> bool:
    if not rows or len(rows) < 3:
        return False
    header = " ".join(c or "" for c in rows[0]).lower()
    if "date" not in header:
        return False
    if not any(k in header for k in ("description", "particular", "narration", "details")):
        return False
    return any(k in header for k in ("debit", "credit", "amount", "qty", "quantity"))


def _line_item_stop_text(text: str) -> bool:
    low = (text or "").lower()
    return any(
        k in low
        for k in (
            "subtotal",
            "total amount",
            "grand total",
            "amount in words",
            "thank you",
            "earn - redeem",
            "place of supply",
            "guest email",
            "cashier",
        )
    )


def _invoice_header_columns(line, page_width: float):
    """Date | Description | Qty | Debit | Credit — SAC/HSN stays inside Description."""
    words = sorted((t for t in (line or []) if (t[3] or "").strip()), key=lambda t: t[1])
    if len(words) < 4:
        return None
    joined = " ".join(t[3] for t in words).lower()
    if "date" not in joined or "description" not in joined:
        return None
    if not any(k in joined for k in ("debit", "credit", "amount")):
        return None
    glue = 12.0
    groups = [[words[0]]]
    for w in words[1:]:
        g_right = max(t[2] for t in groups[-1])
        if w[1] - g_right < glue:
            groups[-1].append(w)
        else:
            groups.append([w])
    cols = []
    for g in groups:
        title = re.sub(r"\s+", " ", " ".join(t[3] for t in g)).strip()
        low = title.lower()
        if re.search(r"\b(sac|hsn)\b", low) or low in {"code", "sac code", "hsn/sac"}:
            continue
        cols.append((title, min(t[1] for t in g), max(t[2] for t in g)))
    if len(cols) < 4:
        return None
    titles = [c[0] for c in cols]
    lefts = [c[1] for c in cols]
    rights = [c[2] for c in cols]
    bounds = [0.0]
    for i in range(len(cols) - 1):
        low = titles[i].lower()
        if any(k in low for k in ("desc", "particular", "narrat")):
            bounds.append(max(rights[i] + 4.0, lefts[i + 1] - 8.0))
        else:
            bounds.append((rights[i] + lefts[i + 1]) / 2.0)
    bounds.append(max(float(page_width), rights[-1] + 40.0))
    return titles, bounds


def _tally_ledger_table_from_lines(lines, page_width: float):
    """Tally ledger: Date | Dr/Cr | Particulars | Vch Type | Vch No. | Debit | Credit."""
    header_i = -1
    header_line = []
    for i, line in enumerate(lines):
        joined = " ".join(t[3] for t in line).lower()
        if (
            "date" in joined
            and "particular" in joined
            and "debit" in joined
            and "credit" in joined
            and ("vch" in joined or "voucher" in joined)
        ):
            header_i = i
            header_line = line
            break
    if header_i < 0 or len(header_line) < 5:
        return None
    words = sorted(header_line, key=lambda t: t[1])
    glue = 10.0
    groups = [[words[0]]]
    for w in words[1:]:
        if w[1] - max(t[2] for t in groups[-1]) < glue:
            groups[-1].append(w)
        else:
            groups.append([w])
    named = []
    for g in groups:
        title = re.sub(r"\s+", " ", " ".join(t[3] for t in g)).strip()
        named.append((title, min(t[1] for t in g), max(t[2] for t in g)))
    by = {re.sub(r"[^a-z]", "", t.lower()): (t, l, r) for t, l, r in named}

    def _col(*keys):
        for k in keys:
            if k in by:
                return by[k]
        for key, val in by.items():
            if any(k in key for k in keys):
                return val
        return None

    date_c = _col("date")
    part_c = _col("particulars", "particular")
    vtype_c = _col("vchtype", "vouchertype")
    vno_c = _col("vchno", "voucherno")
    debit_c = _col("debit")
    credit_c = _col("credit")
    if not all((date_c, part_c, vtype_c, vno_c, debit_c, credit_c)):
        return None
    titles = ["Date", "", "Particulars", "Vch Type", "Vch No.", "Debit", "Credit"]
    bounds = [
        0.0,
        float(date_c[2]) + 4.0,
        float(part_c[1]) + 4.0,
        float(vtype_c[1]) - 6.0,
        (float(vtype_c[2]) + float(vno_c[1])) / 2.0,
        min(float(debit_c[1]) - 18.0, float(vno_c[2]) + 22.0),
        (float(debit_c[2]) + float(credit_c[1])) / 2.0,
        max(float(page_width), float(credit_c[2]) + 40.0),
    ]
    for i in range(1, len(bounds)):
        if bounds[i] <= bounds[i - 1] + 8:
            bounds[i] = bounds[i - 1] + 12.0
    rows = []
    skip_idx = set()
    page_label = ""
    page_y = header_y = min(t[0] for t in header_line)
    for i, line in enumerate(lines[:header_i]):
        joined_pre = " ".join(t[3] for t in line).strip()
        if re.fullmatch(r"Page\s+\d+", joined_pre, re.I):
            page_label = joined_pre
            skip_idx.add(i)
            page_y = min(t[0] for t in line)
    if page_label:
        rows.append([""] * 6 + [page_label])
    rows.append(list(titles))
    last_y = min(t[0] for t in header_line)
    drcr = {"dr", "cr", "DR", "CR", "Dr", "Cr"}
    for line in lines[header_i + 1 :]:
        joined = " ".join(t[3] for t in line).strip()
        if not joined:
            continue
        cells = [""] * len(titles)
        for _y, x0, x1, text in line:
            tok = (text or "").strip()
            if not tok:
                continue
            ci = _col_index_for_x((x0 + x1) / 2.0, bounds)
            if 0 <= ci < len(cells):
                cells[ci] = (cells[ci] + " " + tok).strip()
        if not any(cells):
            continue
        is_new = _is_date_text(cells[0]) or ((cells[1] or "").strip() in drcr and any(cells[2:]))
        if len(rows) > 1 and not is_new and not any(_is_amount_text(c) for c in cells):
            prev = rows[-1]
            for j, c in enumerate(cells):
                if c:
                    prev[j] = (prev[j] + " " + c).strip()
            continue
        rows.append(cells)
        last_y = max(t[0] for t in line)
    if len(rows) < 4:
        return None
    header_y = min(t[0] for t in header_line)
    return {
        "bbox": (0.0, min(page_y, header_y) - 4.0, float(page_width), last_y + 12.0),
        "rows": rows,
        "header_index": header_i,
        "line_indexes": skip_idx,
        "kind": "tally",
        "col_fracs": [0.12, 0.05, 0.28, 0.14, 0.13, 0.14, 0.14],
    }


def _line_item_table_from_lines(lines, page_width: float):
    """Rebuild Date/Description/Debit invoice grids when find_tables invents junk."""
    tally = _tally_ledger_table_from_lines(lines, page_width)
    if tally:
        return tally
    start, end, titles, bounds = -1, -1, [], []
    for i, line in enumerate(lines):
        parsed = _invoice_header_columns(line, page_width)
        if parsed:
            titles, bounds = parsed
            start, end = i, i + 1
            break
    if not titles or len(titles) < 4 or len(bounds) < 5:
        start, end, titles, bounds = _find_simple_txn_header(lines)
    if not titles or len(titles) < 4 or len(bounds) < 5:
        return None
    debit_i = next(
        (i for i, t in enumerate(titles) if any(k in (t or "").lower() for k in ("debit", "amount", "credit"))),
        -1,
    )
    looks_invoice = any("desc" in (t or "").lower() for t in titles) and not any(
        "particular" in (t or "").lower() for t in titles
    )
    if looks_invoice and debit_i >= 2 and not any("qty" in (t or "").lower() for t in titles):
        qty_xs = []
        for line in lines[end if end > start else start + 1 : start + 24]:
            for _y, x0, x1, text in line:
                mid = (x0 + x1) / 2.0
                if re.fullmatch(r"\d{1,4}", (text or "").strip()) and 270 <= mid < bounds[debit_i] - 12:
                    qty_xs.append(mid)
        if len(qty_xs) >= 3:
            qx = sorted(qty_xs)[len(qty_xs) // 2]
            titles = list(titles)
            bounds = list(bounds)
            titles.insert(debit_i, "Qty.")
            bounds.insert(debit_i, qx - 16.0)
    header_line = lines[start] if 0 <= start < len(lines) else []
    body_start = end if end > start else start + 1
    rows = [list(titles)]
    last_y = min((t[0] for t in header_line), default=0.0)
    desc_i = 1
    for i, title in enumerate(titles):
        low = (title or "").lower()
        if any(k in low for k in ("desc", "particular", "narrat")):
            desc_i = i
            break
    footer_from = len(lines)
    for li, line in enumerate(lines[body_start:], start=body_start):
        joined = " ".join(t[3] for t in line).strip()
        if not joined:
            continue
        if re.fullmatch(r"[-_=. ]{6,}", joined):
            continue
        if _line_item_stop_text(joined):
            footer_from = li
            break
        cells = [""] * len(titles)
        for _y, x0, x1, text in line:
            tok = (text or "").strip()
            if _word_token_is_noise(text) and not re.fullmatch(r"[-/*@*&]+", tok):
                continue
            ci = _col_index_for_x((x0 + x1) / 2.0, bounds)
            if 0 <= ci < len(cells):
                cells[ci] = (cells[ci] + " " + tok).strip()
        if not any(cells):
            continue
        first = cells[0]
        if len(rows) > 1 and not _is_date_text(first):
            prev = rows[-1]
            chunk = " ".join(c for c in cells if c).strip()
            if chunk and 0 <= desc_i < len(prev):
                prev[desc_i] = (prev[desc_i] + " " + chunk).strip()
            continue
        rows.append(cells)
        last_y = max(t[0] for t in line)
    if len(rows) < 3:
        return None
    header_y = min((t[0] for t in header_line), default=0.0)
    bbox = (0.0, header_y - 4.0, float(page_width), last_y + 10.0)
    return {"bbox": bbox, "rows": rows, "header_index": start, "footer_from": footer_from}


def _two_col_header_block(lines, page_width: float, stop_y: float):
    """Guest | invoice-field block used by hotel/tax invoices (iLovePDF-style)."""
    idxs: List[int] = []
    rows: List[List[str]] = []
    for i, line in enumerate(lines):
        y = min(t[0] for t in line)
        if y >= stop_y - 2:
            break
        joined = " ".join(t[3] for t in line).strip()
        low = joined.lower()
        if "date" in low and "description" in low:
            break
        if re.fullmatch(r"[-_=. ]{6,}", joined):
            break
        pair = _split_visual_line_two_cols(line, page_width)
        if not pair:
            continue
        left, right = (pair[0] or "").strip(), (pair[1] or "").strip()
        if left and right:
            idxs.append(i)
            rows.append([left, right])
            continue
        if idxs and rows and (left or right):
            idxs.append(i)
            rows.append([left, right])
    if len(rows) < 4:
        return None
    y0 = min(min(t[0] for t in lines[i]) for i in idxs)
    y1 = max(min(t[0] for t in lines[i]) for i in idxs) + 12.0
    return {
        "bbox": (0.0, y0 - 2.0, float(page_width), y1),
        "rows": rows,
        "line_indexes": set(idxs),
        "skip_bbox": False,
    }


def _page_table_items(page) -> list:
    items = []
    try:
        finder = page.find_tables()
        tables = list(finder.tables) if finder else []
    except Exception:
        tables = []
    if not tables:
        try:
            finder = page.find_tables(strategy="text")
            tables = list(finder.tables) if finder else []
        except Exception:
            tables = []
    page_area = max(1.0, page.rect.width * page.rect.height)
    for table in tables:
        try:
            bbox = tuple(table.bbox)
            data = table.extract() or []
        except Exception:
            continue
        rows = [[_cell_as_exact_text(c) for c in (row or [])] for row in data]
        rows = [r for r in rows if any(c for c in r)]
        if _word_table_is_junk(rows, bbox, page_area):
            continue
        items.append((bbox[1], bbox[0], "table", {"bbox": bbox, "rows": rows}))
    return _drop_overlapping_word_tables(items, page_area)


def _word_table_header_index(rows: List[List[str]]) -> int:
    for i, row in enumerate(rows[:3]):
        joined = " ".join(c or "" for c in row).lower()
        if "date" in joined and any(k in joined for k in ("debit", "credit", "particular", "description")):
            return i
    return 0


def _word_col_alignment(title: str, col_index: int, n_cols: int):
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    hl = (title or "").lower().strip()
    if hl in {"", "dr", "cr", "dr/cr"}:
        return WD_ALIGN_PARAGRAPH.CENTER
    if any(k in hl for k in ("debit", "credit", "amount", "qty", "vch no", "voucher no")):
        return WD_ALIGN_PARAGRAPH.RIGHT
    if "date" in hl:
        return WD_ALIGN_PARAGRAPH.RIGHT
    if n_cols >= 5 and col_index >= n_cols - 2 and not hl:
        return WD_ALIGN_PARAGRAPH.RIGHT
    return WD_ALIGN_PARAGRAPH.LEFT


def _set_word_table_col_fracs(table, fracs: List[float]) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Twips

    if not fracs:
        return
    total = sum(fracs) or 1.0
    # ~6.9" usable on A4 with ~0.4" margins
    usable_twips = 9936
    widths = [max(200, int(usable_twips * (f / total))) for f in fracs]
    try:
        table.autofit = False
    except Exception:
        pass
    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        for child in list(grid):
            grid.remove(child)
        for w in widths:
            gc = OxmlElement("w:gridCol")
            gc.set(qn("w:w"), str(w))
            grid.append(gc)
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths):
                try:
                    cell.width = Twips(widths[i])
                except Exception:
                    pass


def _add_word_table(docx, rows: List[List[str]], font_pt: float = 8.0, layout=None) -> None:
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    cols = max(len(r) for r in rows)
    if cols < 1:
        return
    table = docx.add_table(rows=len(rows), cols=cols)
    layout = layout or {}
    tally = layout.get("kind") == "tally" or (
        "particular" in " ".join(rows[_word_table_header_index(rows)]).lower()
        and "vch" in " ".join(rows[_word_table_header_index(rows)]).lower()
    )
    try:
        table.style = "Table Grid"
        table.autofit = not tally
    except Exception:
        pass
    _set_word_table_borders(table, visible=True, horizontal_only=bool(tally))
    try:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else tbl.get_or_add_tblPr()
        tblW = tblPr.find(qn("w:tblW"))
        if tblW is None:
            tblW = OxmlElement("w:tblW")
            tblPr.append(tblW)
        tblW.set(qn("w:w"), "5000")
        tblW.set(qn("w:type"), "pct")
    except Exception:
        pass
    hi = _word_table_header_index(rows)
    header = [(c or "") for c in rows[hi]] + [""] * cols
    fracs = layout.get("col_fracs")
    if tally and not fracs:
        fracs = [0.12, 0.05, 0.28, 0.14, 0.13, 0.14, 0.14][:cols]
    if fracs:
        _set_word_table_col_fracs(table, list(fracs)[:cols])
    for i, row in enumerate(rows):
        is_header = i == hi
        only_money = (not is_header) and not any(
            (row[j] if j < len(row) else "").strip()
            for j in range(min(4, cols))
        ) and any(_is_amount_text(row[j] if j < len(row) else "") for j in range(cols))
        for j in range(cols):
            text = row[j] if j < len(row) else ""
            cell = table.rows[i].cells[j]
            cell.text = text
            title = header[j] if j < len(header) else ""
            align = _word_col_alignment(title, j, cols)
            if "page" in (text or "").lower() and re.search(r"page\s+\d+", text or "", re.I):
                align = WD_ALIGN_PARAGRAPH.RIGHT
            bold = is_header or only_money or (tally and "balance" in text.lower())
            _style_word_cell(cell, font_pt=font_pt, bold=bold, align=align)


def _image_bytes_from_block(doc, block: dict):
    data = block.get("image")
    ext = (block.get("ext") or "png").lower()
    if data:
        return data, ext
    xref = int(block.get("xref") or 0)
    if xref <= 0:
        return None, ext
    try:
        info = doc.extract_image(xref)
        return info.get("image"), (info.get("ext") or ext).lower()
    except Exception:
        return None, ext


def _word_image_stream(data: bytes, ext: str):
    import io

    if not data:
        return None, ext
    if ext in {"jb2", "jbig2", "jpx", "jp2", "tif", "tiff", "bmp"}:
        if Image is None:
            return None, ext
        img = Image.open(io.BytesIO(data))
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format="PNG")
        data, ext = buf.getvalue(), "png"
    stream = io.BytesIO(data)
    stream.name = f"image.{ext if ext in {'png', 'jpg', 'jpeg', 'gif', 'wmf'} else 'png'}"
    return stream, ext


def _add_word_image(docx, data: bytes, ext: str, width_pt: float) -> None:
    from docx.shared import Inches

    stream, ext = _word_image_stream(data, ext)
    if stream is None:
        return
    width_in = max(0.8, min(6.3, (width_pt or 200) / 72.0))
    para = docx.add_paragraph()
    _compact_paragraph(para)
    run = para.add_run()
    run.add_picture(stream, width=Inches(width_in))


def _add_word_images_row(docx, images: list) -> None:
    """Place logos/banners that share a Y band on one Word row, like iLovePDF."""
    from docx.shared import Inches

    if not images:
        return
    para = docx.add_paragraph()
    _compact_paragraph(para)
    for i, img in enumerate(images):
        stream, _ext = _word_image_stream(img["data"], img["ext"])
        if stream is None:
            continue
        if i:
            para.add_run("  ")
        run = para.add_run()
        width_in = max(0.7, min(4.8, (img["width_pt"] or 160) / 72.0))
        run.add_picture(stream, width=Inches(width_in))


def _text_has_devanagari(text: str) -> bool:
    return any("\u0900" <= ch <= "\u097f" for ch in (text or ""))


def _text_is_garbled(text: str) -> bool:
    """True when PDF fonts did not map to real Unicode (Word shows empty boxes)."""
    s = text or ""
    if len(s.strip()) < 8:
        return False
    bad = 0
    good = 0
    for ch in s:
        if ch.isspace():
            continue
        o = ord(ch)
        if ch == "\ufffd" or 0xE000 <= o <= 0xF8FF or 0xF0000 <= o <= 0xFFFFD:
            bad += 1
            continue
        if ch.isalnum() or ("\u0900" <= ch <= "\u097f") or ch in ".,;:()/\\-%&'\"₹":
            good += 1
            continue
        if o < 32 or o > 255:
            bad += 1
    if bad + good < 10:
        return False
    return bad >= 8 and bad >= good * 0.35


def _word_font_for_text(text: str) -> str:
    return "Nirmala UI" if _text_has_devanagari(text) else "Calibri"


def _apply_word_run_font(run, text: str, font_pt: float, bold: bool = False) -> None:
    from docx.oxml.ns import qn
    from docx.shared import Pt

    font = _word_font_for_text(text)
    run.font.name = font
    run.font.size = Pt(font_pt)
    run.bold = bold
    try:
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            from docx.oxml import OxmlElement

            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        rFonts.set(qn("w:ascii"), font)
        rFonts.set(qn("w:hAnsi"), font)
        rFonts.set(qn("w:eastAsia"), font)
        rFonts.set(qn("w:cs"), font)
    except Exception:
        pass


def _add_plain_line(docx, text: str, font_pt: float = 8.5, bold: bool = False, align=None) -> None:
    text = (text or "").strip()
    if not text:
        return
    para = docx.add_paragraph()
    _compact_paragraph(para)
    if align is not None:
        try:
            para.alignment = align
        except Exception:
            pass
    run = para.add_run(text)
    _apply_word_run_font(run, text, font_pt, bold=bold)


def _image_covers_most_of_page(bbox, page) -> bool:
    pw = max(1.0, float(page.rect.width))
    ph = max(1.0, float(page.rect.height))
    w = max(0.0, float(bbox[2]) - float(bbox[0]))
    h = max(0.0, float(bbox[3]) - float(bbox[1]))
    return (w * h) / (pw * ph) >= 0.22 or (w >= pw * 0.82 and h >= ph * 0.35)


def _is_small_logo(bbox, page) -> bool:
    pw = max(1.0, float(page.rect.width))
    ph = max(1.0, float(page.rect.height))
    w = max(0.0, float(bbox[2]) - float(bbox[0]))
    h = max(0.0, float(bbox[3]) - float(bbox[1]))
    if bbox[1] > ph * 0.22:
        return False
    if _image_covers_most_of_page(bbox, page):
        return False
    return w < pw * 0.42 and h < ph * 0.16 and w >= 36 and h >= 14


def _page_words_and_text(page, textpage=None):
    kwargs = {}
    if textpage is not None:
        kwargs["textpage"] = textpage
    try:
        text = page.get_text("text", **kwargs) or ""
    except TypeError:
        text = page.get_text("text") or ""
    try:
        words = page.get_text("words", **kwargs) or []
    except TypeError:
        words = page.get_text("words") or []
    return text, words


def _try_ocr_textpage(page):
    tessdata = _prepare_tesseract_env()
    try:
        return page.get_textpage_ocr(
            flags=3, language="eng", dpi=200, full=True, tessdata=tessdata
        )
    except TypeError:
        return page.get_textpage_ocr(flags=3, language="eng", dpi=200, full=True)


def _page_top_images(doc, page, pymupdf=None):
    """Logos, banners, QR photos — skip full-page scans and broken drawing bboxes."""
    items = []
    page_area = max(1.0, float(page.rect.width) * float(page.rect.height))
    try:
        data = page.get_text("dict") or {}
    except Exception:
        return items
    seen = set()
    for block in data.get("blocks") or []:
        if block.get("type") != 1:
            continue
        bbox = tuple(block.get("bbox") or (0, 0, 0, 0))
        bw = float(bbox[2] - bbox[0])
        bh = float(bbox[3] - bbox[1])
        if bw < 24 or bh < 24:
            continue
        if (bw * bh) / page_area > 0.38:
            continue
        key = (round(bbox[0], 1), round(bbox[1], 1), round(bw, 1), round(bh, 1))
        if key in seen:
            continue
        seen.add(key)
        img_bytes, ext = _image_bytes_from_block(doc, block)
        if not img_bytes:
            continue
        items.append(
            {
                "bbox": bbox,
                "y": float(bbox[1]),
                "x": float(bbox[0]),
                "data": img_bytes,
                "ext": (ext or "png").lower(),
                "width_pt": bw,
            }
        )
    items.sort(key=lambda t: (t["y"], t["x"]))
    return items


def _group_images_by_row(images, y_tol: float = 28.0) -> list:
    rows = []
    for img in images or []:
        if not rows or abs(img["y"] - rows[-1][0]["y"]) > y_tol:
            rows.append([img])
        else:
            rows[-1].append(img)
            rows[-1].sort(key=lambda t: t["x"])
    return rows


def _ui_pulse(msg: str) -> None:
    """Keep the window painting so Windows does not show Not Responding."""
    root = getattr(tk, "_default_root", None)
    if root is None:
        return
    try:
        root.title(f"{APP_NAME} - {msg}")
        root.update_idletasks()
        root.update()
    except Exception:
        pass


def _ocr_page_words(page, tessdata: str = "", language: str = "eng"):
    """OCR one scanned/garbled page. dpi 150 is much faster than 200 on long PDFs."""
    lang = language or "eng"

    def _run(use_lang: str):
        if tessdata:
            return page.get_textpage_ocr(
                flags=3, language=use_lang, dpi=150, full=True, tessdata=tessdata
            )
        return page.get_textpage_ocr(flags=3, language=use_lang, dpi=150, full=True)

    try:
        tp = _run(lang)
    except TypeError:
        tp = page.get_textpage_ocr(flags=3, language=lang, dpi=150, full=True)
    except Exception:
        if lang != "eng":
            try:
                tp = _run("eng")
            except Exception:
                raise
        else:
            raise
    return _page_words_and_text(page, tp)


def _point_in_bbox(x: float, y: float, bbox, pad: float = 3.0) -> bool:
    return bbox[0] - pad <= x <= bbox[2] + pad and bbox[1] - pad <= y <= bbox[3] + pad


def _table_extract_is_faithful(rows: List[List[str]]) -> bool:
    """Reject tables that smashed two amounts into one cell (invented layout)."""
    if not rows or len(rows) < 2:
        return False
    cols = max(len(r) for r in rows)
    if cols < 2:
        return False
    two_amts = re.compile(r"[\d,]+\.\d{2}.+[\d,]+\.\d{2}")
    for row in rows:
        for cell in row:
            if two_amts.search(cell or ""):
                return False
    return True


def _add_positioned_line(docx, line, page_width: float, font_pt: float = 9.0) -> None:
    """Keep PDF words in X order with tab stops — do not invent columns or values."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt

    tokens = [
        t
        for t in sorted(line or [], key=lambda x: x[1])
        if not _word_token_is_noise(t[3])
    ]
    if not tokens:
        return
    glue = 6.0
    groups: List[List[tuple]] = [[tokens[0]]]
    for w in tokens[1:]:
        if w[1] - groups[-1][-1][2] >= glue:
            groups.append([w])
        else:
            groups[-1].append(w)

    usable = 7.3
    pw = max(1.0, float(page_width))
    para = docx.add_paragraph()
    _compact_paragraph(para)
    first_in = max(0.0, (groups[0][0][1] / pw) * usable)
    if first_in >= 0.35:
        try:
            para.paragraph_format.left_indent = Inches(min(first_in, 3.8))
        except Exception:
            pass

    texts: List[str] = []
    tab_pos: List[float] = []
    for i, g in enumerate(groups):
        txt = " ".join(t[3] for t in g).strip()
        if not txt:
            continue
        if not texts:
            texts.append(txt)
            continue
        x_in = (g[0][1] / pw) * usable
        x_in = max(first_in + 0.45, min(usable - 0.15, x_in))
        tab_pos.append(x_in)
        texts.append(txt)

    if tab_pos:
        try:
            pPr = para._p.get_or_add_pPr()
            tabs_el = OxmlElement("w:tabs")
            for pos in tab_pos:
                tab = OxmlElement("w:tab")
                tab.set(qn("w:val"), "left")
                tab.set(qn("w:pos"), str(int(pos * 1440)))
                tabs_el.append(tab)
            pPr.append(tabs_el)
        except Exception:
            pass

    body = texts[0]
    if len(texts) > 1:
        body = texts[0] + "".join("\t" + t for t in texts[1:])
    run = para.add_run(body)
    _apply_word_run_font(run, body, font_pt)


def _horizontal_word_items(page) -> List[Tuple[float, float, float, str]]:
    """Selectable horizontal words only — skip rotated stamp watermarks."""
    items: List[Tuple[float, float, float, str]] = []
    try:
        data = page.get_text("dict") or {}
    except Exception:
        data = {}
    for block in data.get("blocks") or []:
        if block.get("type") != 0:
            continue
        for line in block.get("lines") or []:
            dx, dy = line.get("dir") or (1, 0)
            if abs(float(dy)) > abs(float(dx)) * 0.4:
                continue
            for span in line.get("spans") or []:
                text = (span.get("text") or "").strip()
                if not text:
                    continue
                bbox = span.get("bbox") or (0, 0, 0, 0)
                y0, x0, x1 = float(bbox[1]), float(bbox[0]), float(bbox[2])
                parts = text.split()
                if len(parts) <= 1:
                    if not _word_token_is_noise(text):
                        items.append((y0, x0, x1, text))
                    continue
                width = max(1.0, x1 - x0)
                letters = sum(max(1, len(p)) for p in parts) + max(0, len(parts) - 1)
                cx = x0
                for p in parts:
                    w = width * (max(1, len(p)) / letters)
                    if not _word_token_is_noise(p):
                        items.append((y0, cx, cx + w, p))
                    cx += w
    if items:
        items.sort(key=lambda t: (round(t[0] / 3.5), t[1]))
        return items
    _, words = _page_words_and_text(page)
    for w in words or []:
        if len(w) < 5:
            continue
        tok = (w[4] or "").strip()
        if not tok or _word_token_is_noise(tok):
            continue
        x0, y0, x1 = float(w[0]), float(w[1]), float(w[2])
        # Tall thin boxes are usually vertical watermark glyphs
        if (float(w[3]) - y0) > (x1 - x0) * 1.8 and (x1 - x0) < 14:
            continue
        items.append((y0, x0, x1, tok))
    items.sort(key=lambda t: (round(t[0] / 3.5), t[1]))
    return items


def _cluster_xy_tokens(items, y_tol: float = 3.8) -> List[List[Tuple[float, float, float, str]]]:
    if not items:
        return []
    items = sorted(items, key=lambda t: (round(t[0] / y_tol), t[1]))
    lines: List[List[Tuple[float, float, float, str]]] = []
    for it in items:
        if not lines or abs(it[0] - lines[-1][0][0]) > y_tol:
            lines.append([it])
        else:
            lines[-1].append(it)
    for line in lines:
        line.sort(key=lambda t: t[1])
    return lines


def _line_to_form_kv(line, page_width: float):
    """Split a visual line into (label, value) only on an explicit colon."""
    tokens = [t for t in sorted(line, key=lambda x: x[1]) if (t[3] or "").strip()]
    if not tokens:
        return None
    joined = " ".join(t[3] for t in tokens).strip()
    if ":" not in joined:
        return None
    left, right = joined.split(":", 1)
    left, right = left.strip(" ."), right.strip()
    if left and right and 2 <= len(left) <= 80:
        return left, right
    return None


def _form_table_from_page(page, items=None):
    """Build Field/Value rows from a certificate or form page."""
    page_width = float(page.rect.width)
    items = items if items is not None else _horizontal_word_items(page)
    lines = _cluster_xy_tokens(items)
    titles: List[str] = []
    rows: List[List[str]] = []
    for line in lines:
        kv = _line_to_form_kv(line, page_width)
        if kv:
            rows.append([kv[0], kv[1]])
            continue
        text = " ".join(t[3] for t in line).strip()
        if not text or _word_token_is_noise(text):
            continue
        if not rows and len(text) <= 90:
            titles.append(text)
        elif not rows:
            titles.append(text[:120])
    if len(rows) < 4:
        return None
    return {"titles": titles[:8], "rows": [["Field", "Value"]] + rows}


def _page_needs_visual_copy(page, words=None) -> bool:
    """
    True when Word text-flow cannot keep the PDF layout:
    stamp papers, rotated watermarks, CID-font junk, form certificates.
    Those pages must be inserted as a picture of the PDF page.
    """
    blob = (page.get_text("text") or "").lower()
    if any(
        k in blob
        for k in (
            "non judicial",
            "stamp duty",
            "certificate of stamp",
            "e-stamp",
            "estamp",
            "notary govt",
            "india non judicial",
        )
    ):
        return True
    vert_lines = 0
    try:
        data = page.get_text("dict") or {}
        for block in data.get("blocks") or []:
            if block.get("type") != 0:
                continue
            for line in block.get("lines") or []:
                dx, dy = line.get("dir") or (1, 0)
                if abs(float(dy)) > abs(float(dx)) * 0.55:
                    vert_lines += 1
    except Exception:
        vert_lines = 0
    if vert_lines >= 6:
        return True
    words = words or []
    if len(words) >= 40:
        short = 0
        for w in words:
            tok = (w[4] if len(w) > 4 else "") or ""
            tok = tok.strip()
            if tok and len(tok) <= 2:
                short += 1
        if short / len(words) >= 0.40:
            return True
    return False


def _insert_pdf_page_image(docx, page) -> None:
    """One PDF page → one Word page that looks like the original."""
    import io
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt

    try:
        pix = page.get_pixmap(dpi=140, alpha=False)
    except TypeError:
        import pymupdf

        pix = page.get_pixmap(matrix=pymupdf.Matrix(1.85, 1.85), alpha=False)
    stream = io.BytesIO(pix.tobytes("png"))
    stream.name = "page.png"
    pw_in = max(4.0, float(page.rect.width) / 72.0)
    ph_in = max(4.0, float(page.rect.height) / 72.0)
    try:
        section = docx.sections[-1]
        section.page_width = Inches(pw_in)
        section.page_height = Inches(ph_in)
        section.left_margin = Inches(0.2)
        section.right_margin = Inches(0.2)
        section.top_margin = Inches(0.2)
        section.bottom_margin = Inches(0.2)
    except Exception:
        pass
    para = docx.add_paragraph()
    try:
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    except Exception:
        pass
    _compact_paragraph(para)
    run = para.add_run()
    run.add_picture(stream, width=Inches(max(3.5, pw_in - 0.45)))


def _fit_word_section_to_page(docx, page) -> None:
    from docx.shared import Inches

    try:
        section = docx.sections[-1]
        section.page_width = Inches(max(5.5, float(page.rect.width) / 72.0))
        section.page_height = Inches(max(7.0, float(page.rect.height) / 72.0))
        section.left_margin = Inches(0.45)
        section.right_margin = Inches(0.32)
        section.top_margin = Inches(0.40)
        section.bottom_margin = Inches(0.40)
    except Exception:
        pass


def _line_is_centered_on_page(line, page_width: float) -> bool:
    if not line:
        return False
    x0 = min(t[1] for t in line)
    x1 = max(t[2] for t in line)
    return abs(((x0 + x1) / 2.0) - (page_width / 2.0)) <= page_width * 0.14


def _add_word_page_line(docx, line, page_width: float) -> None:
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    text = " ".join(t[3] for t in (line or [])).strip()
    if not text:
        return
    y = min(t[0] for t in line)
    if re.fullmatch(r"Page\s+\d+", text, re.I):
        _add_plain_line(docx, text, font_pt=9.0, bold=False, align=WD_ALIGN_PARAGRAPH.RIGHT)
        return
    low = text.lower()
    if "date" in low and any(k in low for k in ("debit", "credit", "particular", "description")):
        _add_positioned_line(docx, line, page_width, font_pt=10.0)
        return
    if y < 170 and _line_is_centered_on_page(line, page_width) and ":" not in text:
        if y < 68:
            _add_plain_line(docx, text, font_pt=12.0, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        elif any(k in text.lower() for k in ("ledger", "account", "invoice", "statement")):
            _add_plain_line(docx, text, font_pt=10.0, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        else:
            _add_plain_line(docx, text, font_pt=9.0, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER)
        return
    _add_positioned_line(docx, line, page_width, font_pt=10.0)


def _line_is_page_heading(line, page_width: float) -> bool:
    if not line:
        return False
    text = " ".join(t[3] for t in line).strip()
    if not text or len(text) > 70 or ":" in text:
        return False
    if _is_date_text(text):
        return False
    y = min(t[0] for t in line)
    if y > 190:
        return False
    return len(text.split()) <= 12


def _convert_pdf_page_to_word(
    docx, doc, page, pymupdf, text: str = "", words=None, layout=None, allow_tables: bool = True
) -> None:
    """Copy this page: real tables stay tables, logos stay images, leftover text stays text."""
    if words is None:
        text, words = _page_words_and_text(page)
    _fit_word_section_to_page(docx, page)
    page_width = float(page.rect.width)
    word_lines = _cluster_page_word_lines(words or [])
    # Word boxes keep Date | Description | Debit columns. Span-split text
    # glues those headers into one cell, so iLovePDF-style tables disappear.
    if len(word_lines) >= 6:
        lines = word_lines
    else:
        horiz = _horizontal_word_items(page)
        lines = _cluster_xy_tokens(horiz) if horiz else word_lines

    table_payloads = []
    if allow_tables:
        try:
            raw_tables = _page_table_items(page)
        except Exception:
            raw_tables = []
        for y, _x, _kind, payload in raw_tables:
            rows = payload.get("rows") or []
            if rows and any(any((c or "").strip() for c in row) for row in rows):
                table_payloads.append((y, payload))
        rebuilt = _line_item_table_from_lines(lines, page_width)
        if rebuilt:
            rbbox = rebuilt.get("bbox")
            if rbbox:
                table_payloads = [
                    (y, p)
                    for y, p in table_payloads
                    if _bbox_overlap_frac(p.get("bbox") or (0, 0, 0, 0), rbbox) < 0.40
                ]
            table_payloads.append((rebuilt["bbox"][1], rebuilt))
        stop_y = min((p["bbox"][1] for _y, p in table_payloads if p.get("bbox")), default=page.rect.height)
        header_block = _two_col_header_block(lines, page_width, stop_y)
        if header_block:
            table_payloads.append((header_block["bbox"][1], header_block))

    skip_bboxes = [
        p["bbox"]
        for _y, p in table_payloads
        if p.get("bbox") and p.get("skip_bbox", True)
    ]
    skip_line_idx = set()
    for _y, payload in table_payloads:
        skip_line_idx |= set(payload.get("line_indexes") or [])
    kept_lines = []
    for i, line in enumerate(lines):
        if i in skip_line_idx:
            continue
        if skip_bboxes:
            inside = 0
            for t in line:
                cx = (t[1] + t[2]) / 2.0
                if any(_point_in_bbox(cx, t[0], b) for b in skip_bboxes):
                    inside += 1
            if inside >= max(1, int(len(line) * 0.6)):
                continue
        kept_lines.append(line)

    events = []
    for line in kept_lines:
        events.append((min(t[0] for t in line), 1, ("line", line)))
    for y, payload in table_payloads:
        events.append((float(y), 2, ("table", payload)))
    try:
        image_rows = _group_images_by_row(_page_top_images(doc, page, pymupdf))
    except Exception:
        image_rows = []
    for row in image_rows:
        events.append((float(row[0]["y"]), 0, ("images", row)))
    events.sort(key=lambda e: (e[0], e[1]))

    wrote = False
    for _y, _order, payload in events:
        kind, data = payload
        if kind == "images":
            _add_word_images_row(docx, data)
            wrote = True
        elif kind == "table":
            _add_word_table(docx, data["rows"], font_pt=9.0, layout=data)
            wrote = True
        else:
            _add_word_page_line(docx, data, page_width)
            wrote = True
    if wrote:
        return

    leftover = (text or "").strip()
    if leftover:
        for para in leftover.splitlines():
            _add_plain_line(docx, para, font_pt=10.0)


def _account_header_lines_from_page(page, max_lines: int = 12) -> List[str]:
    """Bank/branch/account lines above the transaction header — for Word intro."""
    lines = _cluster_page_word_lines(page.get_text("words") or [])
    out: List[str] = []
    for line in lines[:28]:
        joined = " ".join(t[3] for t in line).strip()
        if not joined:
            continue
        low = joined.lower()
        if _header_line_score(joined) >= 4 and any(
            k in low for k in ("narration", "withdrawal", "deposit", "balance", "particular")
        ):
            break
        out.append(joined)
        if len(out) >= max_lines:
            break
    return out


def _pdf_looks_like_bank_statement(doc) -> bool:
    blob = ""
    try:
        n = min(3, doc.page_count)
        for i in range(n):
            blob += (doc[i].get_text("text") or "")[:5000]
    except Exception:
        return False
    low = blob.lower()
    hits = sum(
        1
        for k in (
            "narration",
            "withdrawal",
            "deposit",
            "closing balance",
            "value dt",
            "chq./ref",
            "ifsc",
        )
        if k in low
    )
    return hits >= 3


def _xml_escape_text(text: str) -> str:
    return (
        (text or "")
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _word_font_from_pdf(font: str) -> str:
    f = (font or "").lower()
    if "nirmala" in f or "mangal" in f or "devanag" in f:
        return "Nirmala UI"
    if "times" in f or "nimbusrom" in f:
        return "Times New Roman"
    if "courier" in f or "mono" in f:
        return "Courier New"
    return "Arial"


def _pdf_color_hex(color) -> str:
    try:
        if color is None:
            return "000000"
        if isinstance(color, int):
            return f"{color & 0xFFFFFF:06X}"
        if isinstance(color, (tuple, list)) and len(color) >= 3:
            r, g, b = color[:3]
            if max(r, g, b) <= 1.01:
                r, g, b = int(r * 255), int(g * 255), int(b * 255)
            return f"{int(r):02X}{int(g):02X}{int(b):02X}"
    except Exception:
        pass
    return "000000"


def _fit_word_section_exact(section, page) -> None:
    from docx.shared import Inches, Twips

    section.page_width = Inches(float(page.rect.width) / 72.0)
    section.page_height = Inches(float(page.rect.height) / 72.0)
    section.left_margin = Twips(0)
    section.right_margin = Twips(0)
    section.top_margin = Twips(0)
    section.bottom_margin = Twips(0)
    try:
        section.header_distance = Twips(0)
        section.footer_distance = Twips(0)
    except Exception:
        pass


def _word_page_anchor_paragraph(docx, page, first_page: bool):
    if first_page:
        section = docx.sections[0]
        para = docx.paragraphs[0] if docx.paragraphs else docx.add_paragraph()
    else:
        section = docx.add_section()
        para = docx.add_paragraph()
    _fit_word_section_exact(section, page)
    _compact_paragraph(para)
    try:
        para.paragraph_format.space_before = 0
        para.paragraph_format.space_after = 0
    except Exception:
        pass
    return para


def _emu(pt: float) -> int:
    return max(0, int(round(float(pt) * 12700)))


def _next_word_shape_id(docx) -> int:
    n = int(getattr(docx, "_ssa_shape_id", 1) or 1)
    docx._ssa_shape_id = n + 1
    return n


def _append_drawing_xml(para, xml: str) -> None:
    from docx.oxml import parse_xml

    run = para.add_run()
    run._r.append(parse_xml(xml))


def _wp_anchor_open(docx, left_pt, top_pt, width_pt, height_pt, behind: bool, z: int, name: str) -> str:
    i = _next_word_shape_id(docx)
    behind_v = "1" if behind else "0"
    return (
        '<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{max(1, z)}" '
        f'behindDoc="{behind_v}" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{_emu(left_pt)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{_emu(top_pt)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{_emu(width_pt)}" cy="{_emu(height_pt)}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        "<wp:wrapNone/>"
        f'<wp:docPr id="{i}" name="{_xml_escape_text(name)}"/>'
        "<wp:cNvGraphicFramePr>"
        '<a:graphicFrameLocks noChangeAspect="1"/>'
        "</wp:cNvGraphicFramePr>"
    )


def _append_word_textbox(docx, para, text, left, top, width, height, font, size_pt, bold, italic, color, z) -> None:
    sz = max(8, int(round(size_pt * 2)))
    b = "<w:b/>" if bold else ""
    it = "<w:i/>" if italic else ""
    w, h = max(width, 3.0), max(height, size_pt * 1.15)
    xml = (
        _wp_anchor_open(docx, left, top, w, h, False, z, f"t{_next_word_shape_id(docx)}")
        + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        + '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
        + f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{_emu(w)}" cy="{_emu(h)}"/></a:xfrm>'
        + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>'
        + "</wps:spPr><wps:txbx><w:txbxContent><w:p>"
        + '<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        + '<w:ind w:left="0" w:right="0"/></w:pPr><w:r><w:rPr>'
        + f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
        + f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>{b}{it}'
        + f'<w:color w:val="{color}"/></w:rPr>'
        + f'<w:t xml:space="preserve">{_xml_escape_text(text)}</w:t>'
        + "</w:r></w:p></w:txbxContent></wps:txbx>"
        + '<wps:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/>'
        + "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>"
    )
    _append_drawing_xml(para, xml)


def _append_word_line_shape(docx, para, x1, y1, x2, y2, color, weight, z) -> None:
    left, top = min(x1, x2), min(y1, y2)
    width, height = abs(x2 - x1), abs(y2 - y1)
    if width < 0.6:
        width = max(weight, 0.6)
        left -= width / 2.0
    if height < 0.6:
        height = max(weight, 0.6)
        top -= height / 2.0
    xml = (
        _wp_anchor_open(docx, left, top, max(width, 0.6), max(height, 0.6), True, z, f"l{_next_word_shape_id(docx)}")
        + '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        + '<wps:wsp><wps:cNvSpPr/><wps:spPr>'
        + f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{_emu(max(width, 0.6))}" cy="{_emu(max(height, 0.6))}"/></a:xfrm>'
        + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        + f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        + "</wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>"
    )
    _append_drawing_xml(para, xml)


def _append_word_image_anchor(docx, para, r_id, left, top, width, height, z) -> None:
    w, h = max(width, 4.0), max(height, 4.0)
    xml = (
        _wp_anchor_open(docx, left, top, w, h, True, z, f"i{_next_word_shape_id(docx)}")
        + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        + "<pic:pic><pic:nvPicPr>"
        + f'<pic:cNvPr id="0" name="Image"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/>'
        + "</pic:cNvPicPr></pic:nvPicPr><pic:blipFill>"
        + f'<a:blip r:embed="{r_id}"/><a:stretch><a:fillRect/></a:stretch>'
        + "</pic:blipFill><pic:spPr>"
        + f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{_emu(w)}" cy="{_emu(h)}"/></a:xfrm>'
        + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        + "</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>"
    )
    _append_drawing_xml(para, xml)


def _pdf_spans_for_word(page) -> list:
    """Every visible PDF span, exact glyphs — no merging, no rewrite."""
    out = []
    try:
        data = page.get_text("dict") or {}
    except Exception:
        return out
    for block in data.get("blocks") or []:
        if block.get("type") != 0:
            continue
        for line in block.get("lines") or []:
            dx, dy = line.get("dir") or (1, 0)
            rotate = abs(float(dy)) > abs(float(dx)) * 0.55
            for span in line.get("spans") or []:
                text = span.get("text") or ""
                if not (text or "").strip():
                    continue
                bbox = span.get("bbox") or (0, 0, 0, 0)
                flags = int(span.get("flags") or 0)
                out.append(
                    {
                        "text": text,
                        "bbox": tuple(float(x) for x in bbox),
                        "size": float(span.get("size") or 9.0),
                        "font": span.get("font") or "",
                        "bold": bool(flags & 16),
                        "italic": bool(flags & 2),
                        "color": span.get("color"),
                        "rotate": rotate,
                    }
                )
    return out


def _pdf_images_for_word(doc, page) -> list:
    items = []
    try:
        data = page.get_text("dict") or {}
    except Exception:
        return items
    seen = set()
    page_area = max(1.0, float(page.rect.width) * float(page.rect.height))
    for block in data.get("blocks") or []:
        if block.get("type") != 1:
            continue
        bbox = tuple(block.get("bbox") or (0, 0, 0, 0))
        bw = float(bbox[2] - bbox[0])
        bh = float(bbox[3] - bbox[1])
        if bw < 8 or bh < 8:
            continue
        if (bw * bh) / page_area > 0.92:
            continue
        key = (round(bbox[0], 1), round(bbox[1], 1), round(bw, 1), round(bh, 1))
        if key in seen:
            continue
        seen.add(key)
        img_bytes, ext = _image_bytes_from_block(doc, block)
        if not img_bytes:
            continue
        items.append({"bbox": bbox, "data": img_bytes, "ext": (ext or "png").lower()})
    return items


def _pt_twips(pt: float) -> str:
    return str(max(0, int(round(float(pt) * 20.0))))


def _span_text_width_pt(text: str, size_pt: float, bold: bool = False) -> float:
    """Conservative width so Word does not wrap inside the frame and cover the next line."""
    t = text or ""
    if not t:
        return 4.0
    factor = 0.62 if bold else 0.56
    return max(4.0, len(t) * float(size_pt) * factor + 3.0)


def _same_word_line(a_bbox, b_bbox) -> bool:
    ay0, ay1 = float(a_bbox[1]), float(a_bbox[3])
    by0, by1 = float(b_bbox[1]), float(b_bbox[3])
    mid = (ay0 + ay1) / 2.0
    return by0 - 1.0 <= mid <= by1 + 1.0


def _begin_exact_word_page(docx, page, first_page: bool) -> None:
    if first_page:
        section = docx.sections[0]
        if docx.paragraphs:
            p0 = docx.paragraphs[0]
            if not (p0.text or "").strip() and not any(p0.runs):
                parent = p0._element.getparent()
                if parent is not None:
                    parent.remove(p0._element)
    else:
        section = docx.add_section()
    _fit_word_section_exact(section, page)


def _add_exact_placed_text(
    docx, text, x, y, width, height, font, size_pt, bold=False, italic=False, color="000000"
) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt

    para = docx.add_paragraph()
    fmt = para.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    try:
        fmt.line_spacing = Pt(max(6.0, float(size_pt)))
    except Exception:
        pass
    pPr = para._p.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    line_twips = max(120, int(round(float(size_pt) * 20.0)))
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), str(line_twips))
    spacing.set(qn("w:lineRule"), "exact")
    word_wrap = OxmlElement("w:wordWrap")
    word_wrap.set(qn("w:val"), "0")
    pPr.append(word_wrap)

    frame_h = min(max(float(size_pt) * 1.25, 8.0), max(8.0, float(height)))
    frame = OxmlElement("w:framePr")
    frame.set(qn("w:w"), _pt_twips(max(8.0, width)))
    frame.set(qn("w:h"), _pt_twips(frame_h))
    frame.set(qn("w:hRule"), "exact")
    frame.set(qn("w:x"), _pt_twips(x))
    frame.set(qn("w:y"), _pt_twips(y))
    frame.set(qn("w:hAnchor"), "page")
    frame.set(qn("w:vAnchor"), "page")
    frame.set(qn("w:wrap"), "none")
    frame.set(qn("w:anchorLock"), "1")
    pPr.append(frame)

    run = para.add_run(text)
    face = font or "Arial"
    run.font.size = Pt(max(6.0, float(size_pt)))
    run.bold = bool(bold)
    if italic:
        run.italic = True
    try:
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        rFonts.set(qn("w:ascii"), face)
        rFonts.set(qn("w:hAnsi"), face)
        rFonts.set(qn("w:eastAsia"), face)
        rFonts.set(qn("w:cs"), face)
        run.font.name = face
    except Exception:
        pass
    try:
        rPr = run._element.get_or_add_rPr()
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color or "000000")
        rPr.append(c)
    except Exception:
        pass


def _add_exact_placed_image(docx, data, ext, x, y, width, height) -> None:
    import io
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Emu

    if not data:
        return
    para = docx.add_paragraph()
    _compact_paragraph(para)
    pPr = para._p.get_or_add_pPr()
    frame = OxmlElement("w:framePr")
    frame.set(qn("w:w"), _pt_twips(max(8.0, width)))
    frame.set(qn("w:h"), _pt_twips(max(8.0, height)))
    frame.set(qn("w:x"), _pt_twips(x))
    frame.set(qn("w:y"), _pt_twips(y))
    frame.set(qn("w:hAnchor"), "page")
    frame.set(qn("w:vAnchor"), "page")
    frame.set(qn("w:wrap"), "none")
    frame.set(qn("w:anchorLock"), "1")
    pPr.append(frame)
    stream = io.BytesIO(data)
    stream.name = f"img.{ext if ext in {'png', 'jpg', 'jpeg', 'gif'} else 'png'}"
    run = para.add_run()
    run.add_picture(stream, width=Inches(max(0.15, width / 72.0)), height=Inches(max(0.15, height / 72.0)))


def _convert_pdf_page_exact_word(docx, doc, page, first_page: bool = True, words=None) -> None:
    """Place each PDF value at its original X/Y. Never build tables."""
    _begin_exact_word_page(docx, page, first_page)

    for img in _pdf_images_for_word(doc, page):
        bbox = img["bbox"]
        x0, y0, x1, y1 = (float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3]))
        _add_exact_placed_image(
            docx,
            img["data"],
            img.get("ext") or "png",
            x0,
            y0,
            max(8.0, x1 - x0),
            max(8.0, y1 - y0),
        )

    spans = _pdf_spans_for_word(page)
    if not spans and words:
        for w in words:
            if len(w) < 5:
                continue
            tok = (w[4] or "").strip()
            if not tok:
                continue
            spans.append(
                {
                    "text": tok,
                    "bbox": (float(w[0]), float(w[1]), float(w[2]), float(w[3])),
                    "size": 9.0,
                    "font": "Arial",
                    "bold": False,
                    "italic": False,
                    "color": 0,
                }
            )

    page_w = float(page.rect.width)
    ordered = sorted(
        [sp for sp in spans if (sp.get("text") or "").strip()],
        key=lambda s: (round(float(s["bbox"][1]), 1), float(s["bbox"][0])),
    )
    for i, sp in enumerate(ordered):
        text = sp["text"]
        x0, y0, x1, y1 = sp["bbox"]
        if x1 <= x0:
            x1 = x0 + 4.0
        if y1 <= y0:
            y1 = y0 + max(8.0, sp["size"])
        next_x = page_w - 6.0
        for nxt in ordered[i + 1 :]:
            if not _same_word_line(sp["bbox"], nxt["bbox"]):
                continue
            nx0 = float(nxt["bbox"][0])
            if nx0 > x0 + 1.0:
                next_x = nx0
                break
        size_pt = float(sp["size"])
        bbox_w = max(4.0, x1 - x0)
        need_w = _span_text_width_pt(text, size_pt, bool(sp.get("bold")))
        room = max(bbox_w, next_x - x0 - 2.5)
        width = min(max(bbox_w, need_w, 10.0), room)
        if need_w > room and len(text.strip()) > 1:
            size_pt = max(6.0, size_pt * (room / need_w) * 0.98)
        _add_exact_placed_text(
            docx,
            text,
            x0,
            y0,
            width,
            max(size_pt * 1.2, y1 - y0),
            _word_font_from_pdf(sp["font"]),
            size_pt,
            bool(sp.get("bold")),
            bool(sp.get("italic")),
            _pdf_color_hex(sp.get("color")),
        )


def pdf_to_word():
    try:
        import pymupdf
        from docx import Document
        from docx.shared import Pt, Inches
    except Exception:
        raise RuntimeError("Please install: python3 -m pip install pymupdf python-docx")

    pdf_path = ask_pdf_file("Select PDF to convert into Word")
    output_folder = ask_output_folder()
    output_name = ask_text("Output file name", "Enter output Word file name:", "converted.docx")
    if not output_name.lower().endswith(".docx"):
        output_name += ".docx"
    out = os.path.join(output_folder, safe_filename(output_name))

    use_path, is_temp = _local_pdf_copy(pdf_path)
    temps = [use_path] if is_temp else []
    try:
        doc = pymupdf.open(use_path)
        try:
            _unlock_pymupdf(doc, pdf_path)
            total = doc.page_count
            probe_empty = 0
            probe_n = min(3, total)
            for si in range(probe_n):
                _t, _w = _page_words_and_text(doc[si])
                if len(_w) + len((_t or "").strip()) < 40:
                    probe_empty += 1
            scanned = probe_n > 0 and probe_empty == probe_n
            garbled = False
            for si in range(probe_n):
                _t, _w = _page_words_and_text(doc[si])
                if _text_is_garbled(_t or ""):
                    garbled = True
                    break

            tessdata = ""
            if scanned:
                try:
                    tessdata = _prepare_tesseract_env()
                except Exception as exc:
                    if scanned:
                        raise RuntimeError(
                            "This PDF is a SCAN (each page is a photo), so Word has no text to copy.\n\n"
                            "Install OCR once, restart SSA PDF Studio, then run PDF to Word again:\n\n"
                            f"  {_tesseract_install_command()}\n\n"
                            f"({exc})"
                        ) from exc
                    tessdata = ""
                if tessdata and scanned:
                    minutes = max(2, int(round(total * 0.08)))
                    ok = messagebox.askyesno(
                        APP_NAME,
                        f"This PDF looks scanned ({total} pages).\n\n"
                        "OCR will read photos into Word text. Pages that already "
                        "have selectable text are copied as-is.\n"
                        f"It may take about {minutes} minutes.\n\n"
                        "Continue?",
                        parent=_dialog_parent(),
                    )
                    if not ok:
                        raise RuntimeError("No input provided.")
            else:
                tessdata = ""

            docx = Document()
            try:
                docx.styles["Normal"].font.name = "Calibri"
                docx.styles["Normal"].font.size = Pt(9)
                docx.styles["Normal"].paragraph_format.space_before = Pt(0)
                docx.styles["Normal"].paragraph_format.space_after = Pt(1)
                for section in docx.sections:
                    section.left_margin = Inches(0.5)
                    section.right_margin = Inches(0.5)
                    section.top_margin = Inches(0.5)
                    section.bottom_margin = Inches(0.4)
            except Exception:
                pass

            ocr_used = 0
            for i, page in enumerate(doc):
                _ui_pulse(f"PDF to Word: page {i + 1}/{total}")
                text, words = _page_words_and_text(page)
                native_ok = len(words) >= 12
                if tessdata and not native_ok:
                    try:
                        text, words = _ocr_page_words(page, tessdata, "hin+eng")
                        ocr_used += 1
                    except Exception:
                        pass
                _convert_pdf_page_exact_word(
                    docx, doc, page, first_page=(i == 0), words=words
                )
            used = f"PDF layout copy, {total} pages"
            if ocr_used:
                used += f" (OCR on {ocr_used} empty pages)"
            _ui_pulse("PDF to Word: saving...")
            save_path = os.path.normpath(out)
            if _is_network_path(out):
                tmp_out = os.path.splitext(_temp_pdf_path("ssa_docx_"))[0] + ".docx"
                temps.append(tmp_out)
                _save_docx_unlocked(docx, tmp_out)
                out = _copy_file_unlocked(tmp_out, save_path)
            else:
                out = _save_docx_unlocked(docx, save_path)
        finally:
            doc.close()
        _ui_pulse("Ready. Select any tool below.")
        return (
            f"Word file created:\n{out}\n\n"
            f"Pages: {total}\n"
            f"Converted as: {used}.\n\n"
            "Each PDF value is placed at the same spot on the page.\n"
            "No Word tables are built, and nothing is rewritten."
        ), output_folder
    finally:
        for temp_path in temps:
            try:
                os.remove(temp_path)
            except OSError:
                pass


def _cell_as_exact_text(cell) -> str:
    """Keep PDF cell content exactly as text. Never convert/reformat amounts."""
    if cell is None:
        return ""
    if isinstance(cell, str):
        return re.sub(r"[\r\n]+", " ", cell).strip()
    if isinstance(cell, bool):
        return "TRUE" if cell else "FALSE"
    if isinstance(cell, int):
        return str(cell)
    if isinstance(cell, float):
        # Preserve 2-decimal money style when value is a float from the parser
        if abs(cell - round(cell, 2)) < 1e-9:
            return f"{cell:.2f}"
        text = format(cell, "f").rstrip("0").rstrip(".")
        return text if text else "0"
    return str(cell)


def _normalize_table_exact(table) -> List[List[str]]:
    """
    Convert cells to exact text and pad short rows with empty cells at the END only.
    Never reorder columns, never swap/replace existing cell values.
    Always pad to at least 2 columns so narration merge never IndexErrors.
    """
    rows = [row for row in (table or []) if row is not None]
    if not rows:
        return []

    exact_rows = [[_cell_as_exact_text(cell) for cell in row] for row in rows]
    max_cols = max(max((len(row) for row in exact_rows), default=0), 2)
    normalized: List[List[str]] = []
    for row in exact_rows:
        if len(row) < max_cols:
            row = row + [""] * (max_cols - len(row))
        normalized.append(row)
    return normalized


def _cell_at(row: List[str], index: int) -> str:
    """Safe cell read — returns '' when column is missing."""
    if index < 0 or index >= len(row):
        return ""
    return row[index] or ""


_DATE_TOKEN = (
    r"\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}"  # 30.04.2025 / 01/11/24 / 03/10/25
    r"|"
    r"\d{4}-\d{2}-\d{2}"  # 2026-08-10 (YES Bank etc.)
    r"|"
    r"\d{1,2}[-/ ](?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-/ ]\d{2,4}"  # 19-Mar-24 (Tally)
    r"|"
    r"\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?,?\s+\d{4}"  # 05 Aug 2026
    r"|"
    r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4}"  # Jan 31, 2025
)
_DATE_RE = re.compile(rf"^(?:{_DATE_TOKEN})$", re.I)
_DATE_LEAD_RE = re.compile(rf"^({_DATE_TOKEN})", re.I)
_DATE_SEARCH_RE = re.compile(_DATE_TOKEN, re.I)
_PAGE_META_WORDS = {
    "account",
    "status",
    "regular",
    "operative",
    "active",
    "inactive",
    "dormant",
    "page",
    "no",
    "of",
    "statement",
    "city",
    "branch",
}
_AMOUNT_RE = re.compile(r"^[+-]?[\d,]+\.\d{2}$")
_REF_RE = re.compile(r"^[A-Za-z]*\d{8,}$")
_SNO_RE = re.compile(r"^\d{1,5}$")
_SMASHED_DATE_AMT_RE = re.compile(
    r"^("
    r"\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}"
    r"|"
    r"\d{4}-\d{2}-\d{2}"
    r")\s+([+-]?[\d,]+\.\d{2})$"
)
_COLUMN_HEADER_RE = re.compile(r"^Column_?\d+$", re.I)

# Bank statement footers / disclaimers (never treat as txn text)
_FOOTER_LINE_MARKERS = (
    "computer generated",
    "require a signature",
    "does not require any signature",
    "does not require a signature",
    "this is a computer",
    "this is a system",
    "system generated statement",
    "**this",
    "statement generated",
    "never share your otp",
    "dial your bank",
    "your base branch",
    "please call from your registered",
    "legends for transactions",
    "team icici",
    "sincerly",
    "sincerely,",
    "statement of transactions in",
)


def _is_statement_footer_line(text: str) -> bool:
    """True for disclaimer / 'computer generated' / '**This is a...' footer lines."""
    raw = (text or "").strip()
    low = re.sub(r"\s+", " ", raw.lower())
    if not low:
        return False
    # HDFC POS REF narration contains WWW.MYNTRA — that is a txn, not a footer.
    first = raw.split()[0] if raw.split() else ""
    if _is_date_text(first) or _DATE_LEAD_RE.match(raw):
        return False
    if low.startswith("**this") or low.startswith("this is a"):
        return True
    # Page-footer websites ("www.icici.bank.in Dial your Bank…") — not txn wrap
    if len(low) < 120 and (low.startswith("www.") or " www." in f" {low}"):
        return True
    return any(m in low for m in _FOOTER_LINE_MARKERS)


_FOOTER_PHRASE_RE = re.compile(
    r"(?:\*{1,2}\s*)?(?:this\s+is\s+a\s+)?"
    r"(?:computer|system)\s+generated\b"
    r"[^.]*?(?:signature\.?|statement\.?)?",
    re.I,
)
_FOOTER_FRAGMENT_RE = re.compile(
    r"(?:\*{1,2}\s*This\b)|(?:\bis a\b)|(?:\brequire a signature\.?)|"
    r"(?:\band does not\b)|(?:\bstatement and does not\b)",
    re.I,
)


def _strip_footer_phrases(text: str) -> str:
    """Remove leaked footer / disclaimer fragments from a cell."""
    t = (text or "").strip()
    if not t:
        return ""
    t = _FOOTER_PHRASE_RE.sub(" ", t)
    t = _FOOTER_FRAGMENT_RE.sub(" ", t)
    t = re.sub(
        r"(?:never share your otp|dial your bank|please call from your registered"
        r"|your base branch|www\.\S+)[^.]*",
        " ",
        t,
        flags=re.I,
    )
    t = re.sub(r"\s+", " ", t).strip(" \t-|,;")
    return t


def _is_date_text(value: str) -> bool:
    text = (value or "").strip()
    if _DATE_RE.match(text):
        return True
    # Polluted date cell ("03/10/25 INW...") — still a date for txn detection
    return bool(_DATE_LEAD_RE.match(text) and len(text) <= 40)


def _split_leading_date(value: str) -> Tuple[str, str]:
    """Split '03/10/25 INW ...' → ('03/10/25', 'INW ...')."""
    text = (value or "").strip()
    m = _DATE_LEAD_RE.match(text)
    if not m:
        return "", text
    return m.group(1), text[m.end() :].strip()


def _header_is_date_col(name: str) -> bool:
    """True for Date / Value Date / Value Dt (HDFC) columns."""
    low = (name or "").lower()
    if "date" in low:
        return True
    return bool(re.search(r"\bdt\b", low))


def _is_statement_page_meta_text(text: str) -> bool:
    """Letterhead crumbs reprinted on every page ('Account Status : Regular')."""
    raw = re.sub(r"\s+", " ", (text or "").strip())
    if not raw:
        return False
    low = raw.lower()
    if "account status" in low:
        return True
    parts = re.findall(r"[a-z0-9]+", low)
    if not parts:
        return False
    if not any(p in ("status", "regular") for p in parts):
        return False
    return all(p in _PAGE_META_WORDS or p.isdigit() for p in parts)


def _peel_date_cell(value: str) -> Tuple[str, str]:
    """
    Keep the real date in a date column.
    'Status 17/04/21' (HDFC page-2 letterhead + value date) → ('17/04/21', 'Status').
    """
    text = (value or "").strip()
    if not text:
        return "", ""
    if _DATE_RE.match(text):
        return text, ""
    d, rest = _split_leading_date(text)
    if d:
        return d, rest
    m = _DATE_SEARCH_RE.search(text)
    if not m:
        return "", text
    date = m.group(0)
    rest = f"{text[: m.start()]} {text[m.end() :]}".strip()
    rest = re.sub(r"\s+", " ", rest).strip(" :|-")
    return date, rest


def _drop_status_glued_to_date(text: str) -> str:
    """
    HDFC continuation pages glue letterhead 'Status' onto Value Dt.
    'Status 17/04/21' / '17/04/21 Status' → '17/04/21'. Other text is unchanged.
    """
    t = (text or "").strip()
    if not t:
        return ""
    if not re.search(r"(?i)\bstatus\b", t):
        return t
    stripped = re.sub(r"(?i)\bstatus\b", " ", t)
    stripped = re.sub(r"\s+", " ", stripped).strip(" :|-")
    if not stripped:
        return ""
    if _DATE_RE.match(stripped):
        return stripped
    d, rest = _peel_date_cell(t)
    if not d:
        return t
    rest = re.sub(r"(?i)\bstatus\b", " ", rest)
    rest = re.sub(r"\s+", " ", rest).strip(" :|-")
    if not rest or _is_statement_page_meta_text(rest):
        return d
    return t


def _is_amount_text(value: str) -> bool:
    text = (value or "").strip().replace(" ", "")
    text = text.replace("$", "").replace("₹", "")
    text = re.sub(r"^(?:Rs\.?|INR|USD)", "", text, flags=re.I)
    text = re.sub(r"(?:Dr|Cr|DR|CR)$", "", text)
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    return bool(_AMOUNT_RE.match(text))


def _is_sno_text(value: str) -> bool:
    return bool(_SNO_RE.match((value or "").strip()))


def _looks_like_party_label(text: str) -> bool:
    """Short counterparty / VPA printed above an ICICI sno+date line — not narration wrap."""
    t = re.sub(r"\s+", " ", (text or "").strip())
    if not t or len(t) > 32 or len(t.split()) > 4:
        return False
    if "/" in t or t.count("-") >= 2:
        return False
    if re.search(r"\d{10,}", t):
        return False
    if re.match(r"^(?:UPI-|IMPS-|NEFT-|RTGS-|PCD:|INB/|ATM-|POS\b)", t, re.I):
        return False
    if "XXXX" in t.upper():
        return False
    if "@" in t:
        return True
    return bool(re.search(r"[A-Za-z]{3,}", t))


def _looks_like_bank_txn_code_prefix(text: str) -> bool:
    """Axis/ICICI particulars printed above the next date (ATM-/UPI-), not HDFC wrap."""
    t = re.sub(r"\s+", " ", (text or "").strip())
    if not t:
        return False
    return bool(
        re.match(
            r"^(?:ATM-|INB/|IMPS/|UPI(?:-|/)|NEFT\s|RTGS\s|ECS/|BY\s+CASH|POS\s|"
            r"NACH|TRF|FT/|MOB/|PUR/)",
            t,
            re.I,
        )
    )


def _looks_like_new_particulars_line(text: str) -> bool:
    """Description printed on the line above date (Axis ATM-/INB/) or ICICI party label."""
    t = re.sub(r"\s+", " ", (text or "").strip())
    if not t:
        return False
    if _looks_like_party_label(t):
        return True
    return _looks_like_bank_txn_code_prefix(t)


def _looks_like_leading_overlay(text: str) -> bool:
    """
    ICICI / similar: counterparty or cheque label is printed on the line ABOVE
    S.No + date + amounts (e.g. 'blinkit.payu@hd', 'SEEMA TRIPATHI', 'Credit trxn').
    Not a UPI wrap fragment, time, date, or amount.
    """
    t = re.sub(r"\s+", " ", (text or "").strip())
    if not t or len(t) > 32 or len(t.split()) > 4:
        return False
    if "/" in t:
        return False
    if _is_amount_text(t) or _is_date_text(t):
        return False
    if re.match(r"^\d{1,2}:\d{2}\b", t):
        return False
    if t.upper() in ("AM", "PM") or re.match(r"^\d{1,2}:\d{2}\s*[AP]M$", t, re.I):
        return False
    if not _looks_like_party_label(t):
        return False
    return True


def _looks_like_cheque_number(value: str) -> bool:
    """True only for an actual cheque/DD number, not UPI ids or party names."""
    t = re.sub(r"\s+", " ", (value or "").strip())
    if not t or "@" in t or "/" in t:
        return False
    if _is_amount_text(t) or _is_date_text(t):
        return False
    if re.search(
        r"\b(UPI|IMPS|NEFT|RTGS|PAYTM|GPAY|PHONEPE|BLINKIT|TRXN|CREDIT|DEBIT)\b",
        t,
        re.I,
    ):
        return False
    m = re.match(r"^(?:CHQ|CHEQUE|CHEQUE\s*NO\.?|DD)[:\s#\-]*(\d{6,9})$", t, re.I)
    if m:
        return True
    compact = re.sub(r"[\s,]", "", t)
    return bool(re.fullmatch(r"\d{6,9}", compact))


def _looks_like_chq_column_token(text: str) -> bool:
    """True only for a real cheque/DD number. ATM locators and INB wraps are Particulars."""
    return _looks_like_cheque_number(text)


def _chq_column_accepts(header: str, text: str) -> bool:
    """
    What may sit in Cheque/Chq/Ref. Empty PDF cheque cells stay empty.
    UPI ids, party names, ATM locators, and wrap fragments never go here.
    """
    h = (header or "").lower()
    t = (text or "").strip()
    if not t:
        return False
    compact = re.sub(r"[\s,]", "", t)
    if "ref" in h:
        if re.fullmatch(r"\d{6,}", compact):
            return True
        if _looks_like_ref_token(t):
            return True
        return _looks_like_cheque_number(t)
    if _looks_like_cheque_number(t):
        return True
    # Axis "Chq No" printed one 5-digit cheque (92450). Not ICICI "Cheque Number".
    if "chq" in h and "cheque" not in h and re.fullmatch(r"\d{5}", compact):
        return True
    return False


def _text_belongs_in_narration(text: str) -> bool:
    """Particulars/remarks that must not be invented into an empty Cheque column."""
    t = (text or "").strip()
    if not t or _is_amount_text(t) or _is_date_text(t):
        return False
    if _looks_like_chq_column_token(t):
        return False
    if re.match(
        r"^(?:INB/|IMPS/|UPI|NEFT|RTGS|ECS/|ATM-|POS\b|NACH|BY\s|FT/|MOB/|PUR/|"
        r"TRF|CHQ|PCD:|LTD\.?/?$|DEPOSIT|SHAREKHAN|SERVICE|CONSOLIDATED)",
        t,
        re.I,
    ):
        return True
    if t.count("/") >= 2:
        return True
    if re.search(
        r"SHAREKHAN|DEPOSIT|TRANSFER|CASH|UPI-|IMPS|NEFT|LTD",
        t,
        re.I,
    ):
        return True
    return False


def _looks_like_ref_token(value: str) -> bool:
    """Cheque / UTR / RRN sitting alone — not narration, date, or amount."""
    t = re.sub(r"\s+", " ", (value or "").strip())
    if not t or _is_amount_text(t) or _is_date_text(t):
        return False
    if re.search(r"\b(UPI|IMPS|NEFT|RTGS|PHONEPE|GPAY|PAYTM|TRANSFER|FROM|TO|BANK)\b", t, re.I):
        return False
    compact = re.sub(r"[\s\-/]", "", t)
    if re.fullmatch(r"\d{6,9}", compact):
        return True
    if re.fullmatch(r"[A-Za-z]{1,8}\d{6,}", compact):
        return True
    if _REF_RE.match(compact):
        return True
    return False


def _split_smashed_date_amount(value: str) -> Tuple[str, str]:
    """If one cell has '01/11/24    1,352.00', split without changing either value."""
    text = (value or "").strip()
    match = _SMASHED_DATE_AMT_RE.match(text)
    if match:
        return match.group(1), match.group(2)
    return text, ""


def _looks_like_header_row(row: List[str]) -> bool:
    joined = " ".join((c or "").lower() for c in row)
    hints = (
        "date",
        "narration",
        "remark",
        "withdrawal",
        "deposit",
        "balance",
        "value",
        "chq",
        "cheque",
        "ref",
        "description",
        "running",
        "s no",
        "sno",
        "credits",
        "debits",
        "beginning",
    )
    return sum(1 for h in hints if h in joined) >= 2


def _is_junk_or_legend_row(row: List[str]) -> bool:
    """Drop legends / footer / disclaimer rows that are not transactions."""
    joined = " ".join((c or "").strip() for c in row)
    if not joined.strip():
        return True
    has_date = any(_is_date_text(c or "") for c in row)
    has_amt = any(_is_amount_text(c or "") for c in row)
    # Last txn on a page often picks up footer chrome in remarks. Keep the txn.
    if has_date and has_amt:
        return False
    if _is_statement_footer_line(joined):
        return True
    low = joined.lower()
    markers = (
        "legends for transactions",
        "pay any visa",
        "system generated statement",
        "does not require any signature",
        "computer generated",
        "require a signature",
        "never share your otp",
        "sincerly",
        "sincerely",
        "team icici",
        "your base branch",
        "dial your bank",
        "www.",
        "statement generated",
        "account statement",
    )
    return any(m in low for m in markers)


def _count_date_cells(rows: List[List[str]]) -> int:
    return sum(1 for row in rows for cell in row if _is_date_text(cell or ""))


def _header_roles(header: List[str]) -> dict:
    """Map logical roles to column indexes from a header row."""
    roles = {"date": 0, "narration": 1, "sno": None}
    for i, cell in enumerate(header):
        low = (cell or "").lower().replace("\n", " ")
        if "s no" in low or low.replace(".", "") in ("sno", "s no"):
            roles["sno"] = i
        if "date" in low and "value" not in low:
            roles["date"] = i
        if any(k in low for k in ("narration", "remark", "particular", "description")):
            roles["narration"] = i
    return roles


def _fix_smashed_cells_in_row(row: List[str]) -> List[str]:
    """
    Split date+amount jammed in one cell into separate empty slots.
    Never changes the amount/date text themselves — only moves them into empty columns.
    """
    fixed = list(row)
    # Grow by one column if needed when splitting into an amount-only trailing cell
    for i, cell in enumerate(list(fixed)):
        date_part, amt_part = _split_smashed_date_amount(cell)
        if not amt_part:
            continue
        fixed[i] = date_part
        # Prefer putting amount into next empty cell; else append new column
        placed = False
        for j in range(i + 1, len(fixed)):
            if not (fixed[j] or "").strip():
                fixed[j] = amt_part
                placed = True
                break
        if not placed:
            fixed.append(amt_part)
    return fixed


def _merge_multiline_transaction_rows(rows: List[List[str]]) -> List[List[str]]:
    """
    Bank-statement fix:
    PDF often splits one transaction across many rows (multi-line narration).
    Merge continuation rows into ONE row per transaction.
    Uses header to find Date / Narration columns (not hard-coded col 0/1).
    """
    if not rows:
        return rows

    rows = [_fix_smashed_cells_in_row(r) for r in rows]
    max_cols = max(max(len(r) for r in rows), 2)
    padded = [r + [""] * (max_cols - len(r)) for r in rows]

    header = None
    body = padded
    if padded and _looks_like_header_row(padded[0]):
        header = padded[0]
        body = padded[1:]

    roles = _header_roles(header) if header else {"date": 0, "narration": 1, "sno": None}
    date_col = roles["date"]
    narr_col = roles["narration"]
    sno_col = roles["sno"]
    ref_cols: List[int] = []
    money_cols: List[int] = []
    if header:
        for i, h in enumerate(header):
            hl = (h or "").lower()
            hl_ns = hl.replace(" ", "")
            if any(k in hl for k in ("ref", "chq", "cheque", "check no")):
                ref_cols.append(i)
            is_drcr_flag = any(
                k in hl_ns for k in ("debit/credit", "credit/debit", "dr/cr", "cr/dr")
            )
            if (not is_drcr_flag) and any(
                k in hl for k in ("withdraw", "deposit", "debit", "credit", "balance", "amount")
            ):
                money_cols.append(i)

    # Drop legend / footer noise before merge
    body = [r for r in body if not _is_junk_or_legend_row(r)]

    merged: List[List[str]] = []
    current: List[str] | None = None

    def start_new(row: List[str]):
        nonlocal current
        if current is not None:
            merged.append(current)
        current = list(row)

    def ensure_width(target: List[str], width: int):
        if len(target) < width:
            target.extend([""] * (width - len(target)))

    def fill_without_overwrite(target: List[str], source: List[str]):
        ensure_width(target, max(len(source), narr_col + 1, date_col + 1))
        ensure_width(source, max(len(source), narr_col + 1))
        if len(source) > len(target):
            target.extend([""] * (len(source) - len(target)))
        for i, val in enumerate(source):
            text = (val or "").strip()
            if not text:
                continue
            existing = _cell_at(target, i).strip()
            if i == date_col:
                if not existing and _is_date_text(text):
                    target[i] = text
                continue
            if i == narr_col:
                if existing:
                    if text not in existing:
                        target[i] = existing + " " + text
                else:
                    target[i] = text
                continue
            if sno_col is not None and i == sno_col:
                if not existing:
                    target[i] = text
                continue
            if i in money_cols:
                # Wrap lines must not park a number in an empty withdrawal/deposit
                if not existing and _is_amount_text(text):
                    target[i] = text
                elif not _is_amount_text(text):
                    ensure_width(target, narr_col + 1)
                    narr = _cell_at(target, narr_col).strip()
                    if text not in narr:
                        target[narr_col] = (narr + " " + text).strip()
                continue
            if i in ref_cols:
                # Never fill an empty cheque/ref from a wrap line (that's another column).
                if not existing:
                    ensure_width(target, narr_col + 1)
                    narr = _cell_at(target, narr_col).strip()
                    if text not in narr:
                        target[narr_col] = (narr + " " + text).strip()
                elif text.upper() not in existing.upper():
                    ensure_width(target, narr_col + 1)
                    narr = _cell_at(target, narr_col).strip()
                    if text not in narr:
                        target[narr_col] = (narr + " " + text).strip()
                continue
            # Other columns: fill ONLY if empty — never replace
            if not existing:
                target[i] = text
            else:
                if (
                    not _is_date_text(text)
                    and not _is_amount_text(text)
                    and not _REF_RE.match(text.replace(" ", ""))
                ):
                    ensure_width(target, narr_col + 1)
                    narr = _cell_at(target, narr_col).strip()
                    if text not in narr:
                        target[narr_col] = (narr + " " + text).strip()

    for row in body:
        ensure_width(row, max(date_col + 1, narr_col + 1, 2))
        date_cell = _cell_at(row, date_col).strip()
        has_date = _is_date_text(date_cell)
        # Some extracts put date in another cell — scan row
        if not has_date:
            has_date = any(_is_date_text((c or "").strip()) for c in row)
        has_sno = False
        if sno_col is not None:
            has_sno = _is_sno_text(_cell_at(row, sno_col))
        else:
            # First cell looks like serial and another cell is a date
            has_sno = _is_sno_text(_cell_at(row, 0)) and has_date
        has_any_amount = any(_is_amount_text((c or "").strip()) for c in row)
        narration_empty = not _cell_at(row, narr_col).strip()

        if current is None:
            start_new(row)
            continue

        current_has_date = _is_date_text(_cell_at(current, date_col).strip()) or any(
            _is_date_text((c or "").strip()) for c in current
        )
        current_has_amount = any(_is_amount_text((c or "").strip()) for c in current)

        # New txn: serial number row, or a new date while current already has a date
        if has_sno and has_date:
            start_new(row)
        elif has_date and current_has_date and has_any_amount:
            start_new(row)
        elif has_date and not current_has_date:
            fill_without_overwrite(current, row)
        elif (
            current_has_date
            and current_has_amount
            and has_any_amount
            and narration_empty
            and has_sno
        ):
            start_new(row)
        else:
            fill_without_overwrite(current, row)

    if current is not None:
        merged.append(current)

    # Keep only rows that look like real transactions (have a date or amount)
    def _is_txn_row(r: List[str]) -> bool:
        if _is_junk_or_legend_row(r):
            return False
        return any(_is_date_text(c or "") for c in r) or any(
            _is_amount_text(c or "") for c in r
        )

    merged = [r for r in merged if _is_txn_row(r)]

    if not merged:
        # Fall back to original padded body (still without legends) so we don't invent empty
        kept = [r for r in body if _is_txn_row(r)]
        if not kept:
            return [header] + body if header else body
        width = max(max(len(r) for r in kept), 2)
        if header:
            if len(header) < width:
                header = header + [""] * (width - len(header))
            return [header] + [r + [""] * (width - len(r)) for r in kept]
        return [r + [""] * (width - len(r)) for r in kept]

    width = max(max(len(r) for r in merged), 2)
    if header:
        if len(header) < width:
            header = header + [""] * (width - len(header))
        elif len(header) > width:
            width = len(header)
            merged = [r + [""] * (width - len(r)) for r in merged]
        return [header] + [r + [""] * (width - len(r)) for r in merged]
    return [r + [""] * (width - len(r)) for r in merged]


def _parse_bank_text_transactions(text: str) -> List[List[str]]:
    """
    Parse ICICI-style statement text:
      <SNo>
      <DD.MM.YYYY>
      <remark lines...>
      <amount(s)>
      <balance>
    Returns header + transaction rows with exact cell text.
    """
    lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    if not lines:
        return []

    # Start at header if present
    start = 0
    for i, ln in enumerate(lines):
        low = ln.lower()
        if "s no" in low or low in ("s no.", "sno", "s.no."):
            start = i
            break
        if _is_date_text(ln) and i > 0 and _is_sno_text(lines[i - 1]):
            start = i - 1
            break

    header = [
        "S No.",
        "Transaction Date",
        "Cheque Number",
        "Transaction Remarks",
        "Withdrawal Amount (INR)",
        "Deposit Amount (INR)",
        "Balance (INR)",
    ]

    body_lines = lines[start:]
    # Skip header label lines until first serial+date
    i = 0
    while i < len(body_lines) - 1:
        if _is_sno_text(body_lines[i]) and _is_date_text(body_lines[i + 1]):
            break
        i += 1
    body_lines = body_lines[i:]

    txns: List[List[str]] = []
    i = 0
    prev_balance: float | None = None
    time_re = re.compile(r"^\d{1,2}:\d{2}(?:\s*[AaPp][Mm])?$")
    ampm_re = re.compile(r"^[AaPp][Mm]$")

    def _to_float(a: str) -> float | None:
        try:
            return float(a.replace(",", "").replace("+", ""))
        except Exception:
            return None

    while i < len(body_lines):
        sno = body_lines[i]
        if not (_is_sno_text(sno) and i + 1 < len(body_lines) and _is_date_text(body_lines[i + 1])):
            # Stop at legends / footer
            if _is_junk_or_legend_row([body_lines[i]]) or body_lines[i].lower().startswith(
                "legends"
            ):
                break
            i += 1
            continue

        date = body_lines[i + 1]
        i += 2
        # Kotak/ICICI: time then value-date sit on their own lines — not remarks
        if i < len(body_lines) and time_re.match(body_lines[i]):
            date = f"{date} {body_lines[i]}".strip()
            i += 1
            if i < len(body_lines) and ampm_re.match(body_lines[i]):
                date = f"{date} {body_lines[i]}".strip()
                i += 1
        elif i < len(body_lines) and ampm_re.match(body_lines[i]):
            date = f"{date} {body_lines[i]}".strip()
            i += 1
        if i < len(body_lines) and _is_date_text(body_lines[i]):
            i += 1  # value date; same as txn date on Kotak — don't dump into remarks
        remark_parts: List[str] = []
        amounts: List[str] = []
        cheque = ""
        while i < len(body_lines):
            ln = body_lines[i]
            if _is_sno_text(ln) and i + 1 < len(body_lines) and _is_date_text(body_lines[i + 1]):
                break
            if ln.lower().startswith("legends") or _is_junk_or_legend_row([ln]):
                break
            if _is_amount_text(ln):
                amounts.append(ln)
            elif time_re.match(ln) or ampm_re.match(ln) or _is_date_text(ln):
                i += 1
                continue
            elif (
                not cheque
                and re.fullmatch(r"\d{6,9}", ln)
                and not _is_sno_text(ln)
            ):
                cheque = ln
            else:
                # Amounts already started — ignore trailing non-amount junk after balance
                if amounts and len(amounts) >= 2:
                    break
                remark_parts.append(ln)
            i += 1

        if not amounts:
            continue

        remark = " ".join(remark_parts).strip()
        remark = re.sub(r"^(\d{4})([A-Za-z])", r"\1 \2", remark)
        balance = amounts[-1]
        mids = amounts[:-1]
        withdrawal = ""
        deposit = ""
        bal_f = _to_float(balance)
        if len(mids) == 1:
            mid = mids[0].strip()
            # Signed Kotak-style amount: exact text, correct column, no balance guessing
            if mid.startswith("-") or (mid.startswith("(") and mid.endswith(")")):
                withdrawal = mid
            elif mid.startswith("+"):
                deposit = mid
            else:
                mid_f = _to_float(mid)
                if prev_balance is not None and bal_f is not None and mid_f is not None:
                    if bal_f < prev_balance - 0.001:
                        withdrawal = mid
                    else:
                        deposit = mid
                else:
                    peek = remark.lower()
                    if any(k in peek for k in ("credit trxn", "int.pd", "interest", "salary")):
                        deposit = mid
                    else:
                        withdrawal = mid
        elif len(mids) >= 2:
            withdrawal = mids[0]
            deposit = mids[1]
        if bal_f is not None:
            prev_balance = bal_f

        txns.append([sno, date, cheque, remark, withdrawal, deposit, balance])

    if not txns:
        return []
    return [header] + txns


_PASSBOOK_HEADER = [
    "Date",
    "Cheque / Ref",
    "Description",
    "Withdrawal",
    "Deposit",
    "Balance",
]


def _cluster_amount_x_bands(xs: List[float]) -> List[Tuple[float, float]]:
    """Group right-aligned amount X positions into 1–3 money columns."""
    if not xs:
        return []
    xs = sorted(xs)
    groups: List[List[float]] = [[xs[0]]]
    for x in xs[1:]:
        if x - groups[-1][-1] < 24:
            groups[-1].append(x)
        else:
            groups.append([x])
    bands = [(min(g) - 10.0, max(g) + 10.0) for g in groups if len(g) >= 2 or len(groups) <= 3]
    return bands[-3:]


def _amount_band_index(x: float, bands: List[Tuple[float, float]]) -> int:
    if not bands:
        return -1
    for i, (a, b) in enumerate(bands):
        if a <= x <= b:
            return i
    best_i, best_d = 0, abs(x - (bands[0][0] + bands[0][1]) / 2.0)
    for i, (a, b) in enumerate(bands):
        d = abs(x - (a + b) / 2.0)
        if d < best_d:
            best_i, best_d = i, d
    return best_i


def _parse_passbook_joined(line: str, prev_balance: float | None) -> List[str] | None:
    """
    Karnataka Bank / passbook line:
      02-04-2025  986715  TRANSFER-...  20,870.00           30,102.08
      01-09-2025  UPI:...  13,440.00                        3,44,749.47
    """
    raw = re.sub(r"\s+", " ", (line or "").strip())
    if not raw:
        return None
    m = _DATE_LEAD_RE.match(raw)
    if not m:
        return None
    date = m.group(1)
    rest = raw[m.end() :].strip()
    if not rest:
        return None
    amts: List[str] = []
    while True:
        mm = re.search(r"([+-]?[\d,]+\.\d{2})$", rest)
        if not mm or not _is_amount_text(mm.group(1)):
            break
        amts.insert(0, mm.group(1))
        rest = rest[: mm.start()].rstrip()
        if len(amts) >= 3:
            break
    if not amts:
        return None
    cheque = ""
    cm = re.match(r"^(\d{6,7})\b\s*(.*)$", rest)
    if cm:
        cheque = cm.group(1)
        rest = cm.group(2).strip()
    desc = rest
    balance = amts[-1]
    mids = amts[:-1]
    withdrawal = ""
    deposit = ""
    bal_f = _parse_excel_number(balance)
    if len(mids) >= 2:
        withdrawal, deposit = mids[0], mids[1]
    elif len(mids) == 1:
        mid = mids[0]
        mid_f = _parse_excel_number(mid)
        if prev_balance is not None and bal_f is not None and mid_f is not None:
            if bal_f < prev_balance - 0.001:
                withdrawal = mid
            else:
                deposit = mid
        else:
            withdrawal = mid
    return [date, cheque, desc, withdrawal, deposit, balance]


def _extract_passbook_statement(pdf_path: str, status_cb=None) -> List[List[List[str]]]:
    """
    Fixed-width passbook (Karnataka Bank etc.): Date | Chq | Particulars | Dr | Cr | Bal.
    Does not need a header row. Uses word X for amount columns when possible.
    """
    try:
        import pymupdf
    except Exception:
        return []

    header = list(_PASSBOOK_HEADER)
    rows: List[List[str]] = []
    prev_balance: float | None = None
    doc = pymupdf.open(pdf_path)
    try:
        sample = ""
        for page in doc[: min(2, doc.page_count)]:
            sample += (page.get_text("text") or "")[:2500]
        low = sample.lower()
        looks_kbl = "karnataka bank" in low
        looks_stmt_grid = (
            "balance" in low
            and any(
                k in low
                for k in (
                    "particulars",
                    "narration",
                    "transaction remarks",
                    "tran date",
                    "transaction date",
                    "withdrawal amt",
                    "value date",
                    "chq no",
                )
            )
        )
        if looks_stmt_grid and not looks_kbl:
            return []
        total = doc.page_count
        dated_hits = 0
        for pi, page in enumerate(doc, start=1):
            if status_cb:
                status_cb(f"Passbook extract: page {pi}/{total}...")
            words = page.get_text("words") or []
            lines = _cluster_page_word_lines(words)
            amt_xs = [
                x1
                for line in lines
                for _y, _x0, x1, t in line
                if _is_amount_text(t)
            ]
            bands = _cluster_amount_x_bands(amt_xs)
            for line in lines:
                joined = " ".join(t[3] for t in line).strip()
                low = joined.lower()
                if not joined:
                    continue
                if _is_statement_footer_line(joined) or _is_junk_or_legend_row([joined]):
                    continue
                if any(
                    k in low
                    for k in (
                        "the karnataka bank",
                        "transaction details",
                        "page ",
                        "account no",
                        "account name",
                        "branch :",
                        "unless constituent",
                        "computer generated",
                    )
                ) and not _DATE_LEAD_RE.match(joined):
                    if "opening balance" in low:
                        amts = [t[3] for _y, _x0, _x1, t in line if _is_amount_text(t)]
                        if amts:
                            prev_balance = _parse_excel_number(amts[-1])
                    continue
                if "opening balance" in low:
                    amts = [t[3] for _y, _x0, _x1, t in line if _is_amount_text(t)]
                    if amts:
                        prev_balance = _parse_excel_number(amts[-1])
                    continue
                left = line[0][3] if line else ""
                if not (_is_date_text(left) or _DATE_LEAD_RE.match(left)):
                    if rows and not any(_is_amount_text(t[3]) for t in line):
                        extra = " ".join(t[3] for t in line if not _is_date_text(t[3]))
                        extra = extra.strip()
                        if extra:
                            rows[-1][2] = (rows[-1][2] + " " + extra).strip()
                    continue
                dated_hits += 1
                parsed = None
                if len(bands) >= 2:
                    date = left
                    cheque = ""
                    desc_parts: List[str] = []
                    withdrawal = ""
                    deposit = ""
                    balance = ""
                    nband = len(bands)
                    for _y, x0, x1, t in line[1:]:
                        mid = (x0 + x1) / 2.0
                        if _is_amount_text(t):
                            bi = _amount_band_index(x1, bands)
                            if nband >= 3:
                                if bi <= 0:
                                    withdrawal = t
                                elif bi == 1:
                                    deposit = t
                                else:
                                    balance = t
                            else:
                                # 2 bands: left = txn amount, right = balance
                                if bi >= nband - 1:
                                    balance = t
                                elif not withdrawal and not deposit:
                                    withdrawal = t
                                else:
                                    deposit = t
                        elif _looks_like_ref_token(t) and not cheque and not desc_parts:
                            cheque = t
                        else:
                            desc_parts.append(t)
                    if not balance:
                        parsed = _parse_passbook_joined(joined, prev_balance)
                    else:
                        if withdrawal and not deposit and prev_balance is not None:
                            bal_f = _parse_excel_number(balance)
                            mid_f = _parse_excel_number(withdrawal)
                            if (
                                bal_f is not None
                                and mid_f is not None
                                and bal_f > prev_balance + 0.001
                            ):
                                deposit = withdrawal
                                withdrawal = ""
                        parsed = [
                            date,
                            cheque,
                            " ".join(desc_parts).strip(),
                            withdrawal,
                            deposit,
                            balance,
                        ]
                else:
                    parsed = _parse_passbook_joined(joined, prev_balance)
                if not parsed:
                    continue
                bal_f = _parse_excel_number(parsed[5])
                if bal_f is not None:
                    prev_balance = bal_f
                rows.append(parsed)
        if dated_hits < 3 or len(rows) < 3:
            return []
    finally:
        doc.close()
    return [_normalize_table_exact([header] + rows)]


def _extract_bank_statement_pymupdf(pdf_path: str, status_cb=None) -> List[List[List[str]]]:
    """
    ICICI / similar statements: geometry tables often only find the header.
    Parse page text into exact transaction rows; also try find_tables(strategy='text').
    """
    try:
        import pymupdf
    except Exception:
        return []

    all_tables: List[List[List[str]]] = []
    doc = pymupdf.open(pdf_path)
    try:
        total = doc.page_count
        for i, page in enumerate(doc, start=1):
            if status_cb:
                status_cb(f"Bank-statement parse: page {i}/{total}...")

            # 1) Text parser (best for this ICICI layout)
            text = page.get_text("text") or ""
            parsed = _parse_bank_text_transactions(text)
            if parsed and _count_date_cells(parsed) >= 1:
                all_tables.append(_normalize_table_exact(parsed))
                continue

            # 2) PyMuPDF text-strategy tables
            try:
                finder = page.find_tables(strategy="text")
            except TypeError:
                try:
                    finder = page.find_tables()
                except Exception:
                    finder = None
            except Exception:
                finder = None
            tables = getattr(finder, "tables", None) or []
            best = None
            best_score = 0
            for tab in tables:
                try:
                    raw = tab.extract()
                except Exception:
                    continue
                exact = _normalize_table_exact(raw)
                score = _count_date_cells(exact)
                if score > best_score:
                    best_score = score
                    best = exact
            if best and best_score >= 1:
                all_tables.append(best)
    finally:
        doc.close()
    return all_tables


def _fragment_cell_ratio(rows: List[List[str]]) -> float:
    """Share of non-empty cells that are tiny 1–2 char fragments (over-split words)."""
    nonempty = [(c or "").strip() for r in rows for c in r if (c or "").strip()]
    if not nonempty:
        return 1.0
    frags = sum(1 for c in nonempty if len(c) <= 2)
    return frags / len(nonempty)


def _header_is_column_numbered(header: List[str]) -> bool:
    cells = [(c or "").strip() for c in (header or []) if (c or "").strip()]
    if len(cells) < 3:
        return False
    return sum(1 for c in cells if _COLUMN_HEADER_RE.match(c)) >= max(3, int(0.5 * len(cells)))


def _header_looks_like_account_meta(header: List[str]) -> bool:
    """Address / account-info blocks mistaken for a statement table header."""
    joined = " ".join((c or "") for c in (header or [])).lower()
    joined = re.sub(r"\s+", " ", joined)
    if not joined.strip():
        return True
    if _looks_like_header_row(header or []):
        return False
    # Over-split "A ccount Na me" style headers
    compact = re.sub(r"\s+", "", joined)
    if (
        "accountname" in compact
        or "accountno" in compact
        or "accountbranch" in compact
        or "address" in compact[:40]
    ):
        return True
    markers = (
        "account branch",
        "account no",
        "account name",
        "page no",
        "statement of account",
        "joint holders",
        "cust id",
        "od limit",
        "bestech",
        "registered office",
        "preferred customer",
        "branch details",
        "period :",
        "your branch",
        "scheme :",
        "customer no",
    )
    return any(m in joined for m in markers)


def _extract_is_garbage(tables: List[List[List[str]]]) -> bool:
    """
    Structurally unusable extract: Column_* headers, over-split fragments,
    or account-address blocks posing as tables. Never prefer these over adaptive.
    """
    if not tables or not tables[0]:
        return True
    block = tables[0]
    header = block[0] if block else []
    if _header_is_column_numbered(header):
        return True
    if _fragment_cell_ratio(block) > 0.30:
        return True
    if _header_looks_like_account_meta(header):
        return True
    if not _looks_like_header_row(header):
        nonempty = [c for c in header if (c or "").strip()]
        if len(header) >= 6 and len(nonempty) <= 2:
            return True
        joined_h = " ".join((c or "") for c in header).lower()
        if any(
            m in joined_h
            for m in (
                "branch details",
                "period :",
                "mrs.",
                "joint holder",
                "account branch",
            )
        ):
            return True
    # Single-field vertical dump (every row has ≤1 non-empty cell) — Column_* precursor
    body = block[1:] if _looks_like_header_row(header) else block
    if len(body) >= 8:
        skinny = sum(
            1 for r in body if sum(1 for c in r if (c or "").strip()) <= 1
        )
        if skinny / len(body) >= 0.85:
            return True
    return False


def _tables_are_weak(page_tables: List[List[List[str]]]) -> bool:
    """True when geometry extract only found headers/legends, not real txns."""
    if not page_tables:
        return True
    flat = [row for t in page_tables for row in t]
    if not flat:
        return True

    # Garbage from over-split word/pdfplumber extracts (Column_1..N, single letters)
    first = flat[0]
    if _header_is_column_numbered(first):
        return True
    if _extract_is_garbage(page_tables):
        return True
    avg_cols = sum(len(r) for r in flat) / max(len(flat), 1)
    if avg_cols >= 10:
        single_char = sum(
            1
            for r in flat
            for c in r
            if (c or "").strip() and len((c or "").strip()) == 1
        )
        if single_char >= max(8, len(flat)):
            return True
    if _fragment_cell_ratio(flat) > 0.30:
        return True

    date_count = sum(_count_date_cells(t) for t in page_tables)
    # Prefer real ISO/full dates; ignore tiny fragmented year leftovers
    full_dates = sum(
        1
        for r in flat
        for c in r
        if re.match(r"^\d{4}-\d{2}-\d{2}$", (c or "").strip())
        or re.match(r"^\d{1,2}[/\-.]\d{1,2}[/\-.]\d{4}$", (c or "").strip())
        or re.match(
            r"^\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4}",
            (c or "").strip(),
            re.I,
        )
    )
    if full_dates >= 2:
        # Still weak if almost no amounts (header-only / broken tables)
        amt = sum(1 for r in flat for c in r if _is_amount_text(c or ""))
        return amt < 2

    useful_with_date = [
        r
        for r in flat
        if not _looks_like_header_row(r)
        and not _is_junk_or_legend_row(r)
        and any(_is_date_text(c or "") for c in r)
    ]
    if date_count >= 2 and len(useful_with_date) >= 2:
        amt = sum(1 for r in useful_with_date for c in r if _is_amount_text(c or ""))
        return amt < 2
    return True


def _txn_score(tables: List[List[List[str]]]) -> int:
    """How many real transaction rows (date + amount) are in the tables."""
    n = 0
    for block in tables or []:
        for row in block:
            if _looks_like_header_row(row) or _is_junk_or_legend_row(row):
                continue
            has_date = any(_is_date_text(c or "") for c in row)
            has_amt = any(_is_amount_text(c or "") for c in row)
            if has_date and has_amt:
                n += 1
    return n


def _table_quality(tables: List[List[List[str]]]) -> int:
    """
    Prefer tables with fewer empty columns and more filled cells.
    Used as tie-break when txn counts are equal (adaptive vs kotak etc.).
    """
    if not tables:
        return -10_000
    block = tables[0]
    if len(block) < 2:
        return -5_000
    header, body = block[0], block[1:]
    # Hard reject Column_* / fragment / account-meta garbage
    if _header_is_column_numbered(header):
        return -25_000
    if _fragment_cell_ratio(block) > 0.30:
        return -20_000
    if _header_looks_like_account_meta(header):
        return -18_000
    width = len(header)
    empty_cols = 0
    for c in range(width):
        if not any((r[c] if c < len(r) else "").strip() for r in body):
            empty_cols += 1
    filled = sum(1 for r in body for cell in r if (cell or "").strip())
    # Penalize year-spill into details: "2026 SentIMPS..."
    spill = 0
    for r in body:
        for cell in r:
            if re.match(r"^\d{4}\s+\S", (cell or "").strip()):
                spill += 1
                break
    # Description wrap text landed in Reference (e.g. "FIN 10", "C BANK")
    ref_spill = 0
    money_mess = 0
    has_wd = has_dep = False
    for i, h in enumerate(header):
        hl = (h or "").lower()
        hl_ns = hl.replace(" ", "")
        # Axis "Debit/Credit" holds CR/DR flags — not an amount column
        is_drcr_flag = any(
            k in hl_ns for k in ("debit/credit", "credit/debit", "dr/cr", "cr/dr")
        ) or hl_ns in ("drcr", "crdr")
        is_ref = any(k in hl for k in ("ref", "chq", "cheque", "check no"))
        is_money = (not is_drcr_flag) and any(
            k in hl for k in ("withdraw", "deposit", "debit", "credit", "balance", "amount")
        )
        for r in body:
            cell = (r[i] if i < len(r) else "") or ""
            if is_ref and re.search(
                r"\b(FIN|BANK|LLP|VISION|LOAN|CHAWLA|PRIVIHKA|IWORLD)\b", cell, re.I
            ):
                ref_spill += 1
            if is_money and cell.strip() and not _is_amount_text(cell):
                money_mess += 1
            if "withdraw" in hl and _is_amount_text(cell):
                has_wd = True
            if "deposit" in hl and _is_amount_text(cell):
                has_dep = True
    # YES-style statements use both Withdrawals and Deposits — reward separation
    both_money = 40 if (has_wd and has_dep) else 0
    # Prefer canonical "Transaction Date" over truncated "Transaction"
    header_bonus = 0
    joined_h = " ".join(header).lower()
    joined_ns = joined_h.replace(" ", "")
    if "transaction date" in joined_h and "value date" in joined_h:
        header_bonus = 25
    # Real statement headers (Date + Narration/Withdrawal/Deposit/Balance)
    if _looks_like_header_row(header):
        header_bonus += 80
    if any(k in joined_h for k in ("narration", "withdrawal amt", "closing balance", "chq")):
        header_bonus += 40
    # Native PDF headers (Axis Particulars/Amount/Debit-Credit) beat ICICI remaps
    native_hits = sum(
        1
        for k in (
            "particular",
            "amount(inr)",
            "debit/credit",
            "balance(inr)",
            "cheque",
            "branch",
            "s.no",
        )
        if k in joined_ns or k in joined_h
    )
    if native_hits >= 3:
        header_bonus += 120
    # Penalize remapped Withdrawal/Deposit when Particulars/Value Date were collapsed
    if "withdrawal amount" in joined_h and "deposit amount" in joined_h:
        if "particular" not in joined_h and "value date" not in joined_h:
            header_bonus -= 80
    # Prefer denser full grids over sparse blocks with the same txn score
    body_bonus = min(len(body), 200) * 2
    return (
        filled
        - empty_cols * 200
        - spill * 50
        - ref_spill * 80
        - money_mess * 60
        + both_money
        + header_bonus
        + body_bonus
    )


_ENGINE_PREF = {
    "adaptive": 6,
    "kotak": 5,
    "column-anchored": 4,
    "passbook": 3,
    "bank-text": 2,
    "stacked-text": 2,
    "pymupdf": 1,
}


def _cleanup_statement_table(header: List[str], rows: List[List[str]]) -> Tuple[List[str], List[List[str]]]:
    """
    Fix minor layout mistakes common across banks:
    - drop fully empty columns
    - move orphan year (2026) from details back onto Value/Txn date ("15 Aug" → "15 Aug 2026")
    - strip footer/disclaimer leaks; peel non-date junk from date columns
    - keep exact text otherwise
    """
    if not header or not rows:
        return header, rows

    width = max(len(header), max((len(r) for r in rows), default=0))
    header = header + [""] * (width - len(header))
    rows = [r + [""] * (width - len(r)) for r in rows]

    def _find_col(*needles: str) -> int:
        for i, h in enumerate(header):
            low = (h or "").lower()
            if all(n in low for n in needles):
                return i
        for i, h in enumerate(header):
            low = (h or "").lower()
            if any(n in low for n in needles):
                return i
        return -1

    txn_date_i = _find_col("transaction", "date")
    if txn_date_i < 0:
        txn_date_i = _find_col("txn", "date")
    if txn_date_i < 0:
        txn_date_i = _find_col("date")
    value_date_i = _find_col("value", "date")
    if value_date_i < 0:
        value_date_i = _find_col("value", "dt")
    details_i = _find_col("detail")
    if details_i < 0:
        details_i = _find_col("narration")
    if details_i < 0:
        details_i = _find_col("description")
    if details_i < 0:
        details_i = _find_col("remark")
    ref_i = _find_col("reference")
    if ref_i < 0:
        ref_i = _find_col("ref")
    if ref_i < 0:
        ref_i = _find_col("cheque")
    if ref_i < 0:
        ref_i = _find_col("chq")

    partial_date = re.compile(
        r"^\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*$",
        re.I,
    )
    year_prefix = re.compile(r"^(\d{4})\s+(.*)$")

    for r in rows:
        # Fix "15 Aug" + details starting with "2026 ..."
        for date_i in (value_date_i, txn_date_i):
            if date_i < 0 or details_i < 0:
                continue
            dcell = (r[date_i] or "").strip()
            det = (r[details_i] or "").strip()
            glued = re.match(r"^(\d{4})([A-Za-z].+)$", det)
            if glued:
                det = f"{glued.group(1)} {glued.group(2)}"
                r[details_i] = det
            if partial_date.match(dcell) and det:
                m = year_prefix.match(det)
                if m:
                    r[date_i] = f"{dcell} {m.group(1)}"
                    r[details_i] = m.group(2).strip()
            # Also: full txn date ok but value date missing year
            elif date_i == value_date_i and partial_date.match(dcell):
                # try take year from txn date
                if txn_date_i >= 0:
                    tm = re.search(r"(\d{4})$", (r[txn_date_i] or "").strip())
                    if tm:
                        r[date_i] = f"{dcell} {tm.group(1)}"

        # Before dropping empty cols: peel leading date out of description into Value Date
        if details_i >= 0:
            det = (r[details_i] or "").strip()
            d, rest = _split_leading_date(det)
            if d and rest:
                if value_date_i >= 0 and not (r[value_date_i] or "").strip():
                    r[value_date_i] = d
                    r[details_i] = rest
                elif txn_date_i >= 0 and (r[txn_date_i] or "").strip() == d:
                    r[details_i] = rest
                elif txn_date_i >= 0 and not (r[txn_date_i] or "").strip():
                    r[txn_date_i] = d
                    r[details_i] = rest

        # Strip footer phrases early so empty-col detection is accurate
        for i in range(len(r)):
            r[i] = _strip_footer_phrases(r[i] or "")
        for date_i in (txn_date_i, value_date_i):
            if date_i < 0:
                continue
            cell = (r[date_i] or "").strip()
            if not cell:
                continue
            cell = _drop_status_glued_to_date(cell)
            d, rest = _peel_date_cell(cell)
            if d:
                r[date_i] = d
                if rest and _is_statement_page_meta_text(rest):
                    rest = ""
                if rest and details_i >= 0:
                    r[details_i] = (rest + " " + (r[details_i] or "")).strip()
            elif not _is_date_text(cell):
                r[date_i] = ""
            else:
                r[date_i] = cell

    # Drop columns that are empty in every data row — but NEVER drop real
    # statement columns (Cheque/Debit/Credit/Balance/…) just because sparse.
    _keep_empty_if = (
        "date",
        "particular",
        "narrat",
        "remark",
        "desc",
        "detail",
        "cheque",
        "chq",
        "check",
        "ref",
        "debit",
        "credit",
        "withdraw",
        "deposit",
        "balance",
        "amount",
        "branch",
        "type",
        "s.no",
        "s no",
        "sno",
        "transaction",
        "value",
    )
    keep = []
    for c in range(width):
        has_data = any((r[c] or "").strip() for r in rows)
        has_header = bool((header[c] or "").strip())
        hlow = (header[c] or "").lower()
        important = has_header and any(k in hlow for k in _keep_empty_if)
        if has_data or important:
            keep.append(c)
        elif has_header and not has_data:
            # empty phantom gap column (no meaningful header keyword) — drop
            continue
        else:
            continue
    if not keep:
        keep = list(range(width))

    new_header = [header[c] for c in keep]
    new_rows = [[r[c] for c in keep] for r in rows]

    # Remap indexes after column drop
    def _remap(old_i: int) -> int:
        if old_i < 0 or old_i not in keep:
            return -1
        return keep.index(old_i)

    txn_date_i = _remap(txn_date_i)
    value_date_i = _remap(value_date_i)
    details_i = _remap(details_i)
    ref_i = _remap(ref_i)
    date_cols = [i for i in (txn_date_i, value_date_i) if i >= 0]
    for i, h in enumerate(new_header):
        if i not in date_cols and _header_is_date_col(h):
            date_cols.append(i)

    # Drop junk preamble / footer rows
    first_txn = 0
    for i, r in enumerate(new_rows):
        has_date = False
        for c in r:
            d, _rest = _split_leading_date(c or "")
            if d or _is_date_text(c or ""):
                has_date = True
                break
        has_amt = any(_is_amount_text(c or "") for c in r)
        if has_date and has_amt and not _is_junk_or_legend_row(r):
            first_txn = i
            break
    new_rows = new_rows[first_txn:]
    new_rows = [r for r in new_rows if not _is_junk_or_legend_row(r)]

    # Strip footer leaks + peel non-date junk from date columns
    narr_i = details_i
    for r in new_rows:
        for i in range(len(r)):
            r[i] = _strip_footer_phrases(r[i] if i < len(r) else "")
            r[i] = _drop_status_glued_to_date(r[i])

        for di in date_cols:
            if di >= len(r):
                continue
            cell = (r[di] or "").strip()
            if not cell:
                continue
            cell = _drop_status_glued_to_date(cell)
            d, rest = _peel_date_cell(cell)
            if d:
                r[di] = d
                rest = _strip_footer_phrases(rest)
                if rest and _is_statement_page_meta_text(rest):
                    rest = ""
                if rest and narr_i >= 0 and narr_i < len(r):
                    r[narr_i] = (rest + " " + (r[narr_i] or "")).strip()
                elif rest and not _is_statement_footer_line(rest):
                    # spill to first non-date col
                    for j in range(len(r)):
                        if j in date_cols:
                            continue
                        r[j] = ((rest + " " + (r[j] or "")).strip())
                        break
            else:
                # Pure footer / page-header fragment in a date column ("Status", "**This")
                cleaned = _strip_footer_phrases(cell)
                if (
                    not cleaned
                    or _is_statement_footer_line(cleaned)
                    or _is_statement_page_meta_text(cleaned)
                    or not _is_date_text(cleaned)
                ):
                    r[di] = ""
                else:
                    r[di] = cleaned

        # Rebalance narration that sat under Ref ("TRANSFER TO/FROM ...")
        # Keep cheque/ref tokens in Ref — do NOT move "TRANSFER 48977..." wholesale.
        if narr_i >= 0 and ref_i >= 0 and narr_i < len(r) and ref_i < len(r):
            ref = (r[ref_i] or "").strip()
            # Lone "/" or "-" is a column delimiter placeholder, not a ref
            if ref in {"/", "-", "/-", "-/", ".", "—", "–"}:
                r[ref_i] = ""
                ref = ""
            else:
                # Strip delimiter-only tokens and trailing " /"
                ref = re.sub(r"(?:^|\s)/+(?=\s|$)", " ", ref)
                ref = re.sub(r"\s+", " ", ref).strip(" /")
                r[ref_i] = ref
            if ref:
                # Move only TRANSFER / TO / FROM words; leave ref-like tokens
                parts = ref.split()
                moved_parts: List[str] = []
                keep_parts: List[str] = []
                for p in parts:
                    pl = p.upper()
                    if pl in {"TRANSFER", "TO", "FROM", "TRANSFER-", "TRANSFER,"}:
                        moved_parts.append(p.rstrip("-,"))
                    elif pl in {"/", "-"}:
                        continue
                    else:
                        keep_parts.append(p)
                if moved_parts:
                    r[narr_i] = ((r[narr_i] or "") + " " + " ".join(moved_parts)).strip()
                    r[ref_i] = " ".join(keep_parts).strip()
                # Deduplicate identical ref tokens (SBI reprints cheque on wrap line)
                ref2 = (r[ref_i] or "").strip()
                if ref2:
                    toks = ref2.split()
                    dedup: List[str] = []
                    seen = set()
                    for t in toks:
                        key = t.upper()
                        if key in seen:
                            continue
                        seen.add(key)
                        dedup.append(t)
                    r[ref_i] = " ".join(dedup)

        # Branch crumbs like "FROM 4430 /" / "TO /" belong with narration
        branch_i = -1
        for i, h in enumerate(new_header):
            if "branch" in (h or "").lower():
                branch_i = i
                break
        if narr_i >= 0 and branch_i >= 0 and branch_i < len(r):
            br = (r[branch_i] or "").strip()
            if re.match(r"^(FROM|TO)\b", br, re.I):
                # Keep pure branch codes if present after FROM/TO
                bm = re.match(r"^((?:FROM|TO)\b\s*)(.*)$", br, re.I)
                if bm:
                    lead, rest = bm.group(1).strip(), bm.group(2).strip()
                    # If rest is only "/" or empty, move all; else move FROM/TO word only
                    if not rest or rest in {"/", "-"}:
                        r[narr_i] = ((r[narr_i] or "") + " " + br).strip()
                        r[branch_i] = ""
                    elif re.match(r"^[\d/.\-]+$", rest.replace(" ", "")):
                        # "FROM 4430 /" — keep code in branch, move FROM to narr
                        r[narr_i] = ((r[narr_i] or "") + " " + lead.rstrip()).strip()
                        r[branch_i] = rest.strip(" /")
                    else:
                        r[narr_i] = ((r[narr_i] or "") + " " + br).strip()
                        r[branch_i] = ""
            else:
                # "99922 TO /" — keep numeric branch, move trailing TO/FROM into narr
                bm = re.match(r"^([\dA-Za-z]+)\s+((?:TO|FROM)\b.*)$", br, re.I)
                if bm:
                    r[branch_i] = bm.group(1)
                    r[narr_i] = ((r[narr_i] or "") + " " + bm.group(2).strip()).strip()
                else:
                    r[branch_i] = re.sub(r"\s+/+\s*$", "", br).strip()

        # Description leading date → Value Date when empty
        if narr_i >= 0 and narr_i < len(r):
            d, rest = _split_leading_date(r[narr_i] or "")
            if d and rest:
                if value_date_i >= 0 and value_date_i < len(r) and not (r[value_date_i] or "").strip():
                    r[value_date_i] = d
                    r[narr_i] = rest
                elif txn_date_i >= 0 and value_date_i < 0:
                    # no value col — drop duplicate date if txn already has it
                    if txn_date_i < len(r) and (r[txn_date_i] or "").strip() == d:
                        r[narr_i] = rest
                    else:
                        r[narr_i] = rest

        # Strip trailing " /" crumbs. Only peel a cheque when narration itself
        # says CHQ/CHEQUE — never promote UPI/IMPS ids into the cheque column.
        if narr_i >= 0 and narr_i < len(r):
            det = (r[narr_i] or "").strip()
            det = re.sub(r"\s*/+\s*$", "", det).strip()
            det = re.sub(r"\s+/\s+", " ", det).strip()
            if ref_i >= 0 and ref_i < len(r):
                ref_now = (r[ref_i] or "").strip()
                if ref_now in {"/", "-", ""}:
                    mchq = re.search(
                        r"(?:\b(?:CHQ|CHEQUE|CHEQUE\s*NO\.?|DD)\b[:\s#\-]*)(\d{6,9})\s*$",
                        det,
                        re.I,
                    )
                    if mchq:
                        r[ref_i] = mchq.group(1)
                        det = (det[: mchq.start()] + det[mchq.end() :]).strip()
            r[narr_i] = det

        # Cheque Number (not Chq/Ref): keep only real cheque digits. UPI/party
        # labels that X-mapped into this blank column belong in remarks.
        if narr_i >= 0 and ref_i >= 0 and narr_i < len(r) and ref_i < len(r):
            hlow = (new_header[ref_i] or "").lower() if ref_i < len(new_header) else ""
            strict_chq = (
                any(k in hlow for k in ("cheque", "chq", "check no", "check number"))
                and "ref" not in hlow
            )
            chq_now = (r[ref_i] or "").strip()
            if strict_chq and chq_now:
                det = (r[narr_i] or "").strip()
                real_chq_narr = bool(re.search(r"\b(?:CHQ|CHEQUE)\b", det, re.I))
                upi_like = bool(re.search(r"(?:UPI/|NEFT-|IMPS|RTGS)", det, re.I))
                if (not _looks_like_cheque_number(chq_now)) or (upi_like and not real_chq_narr):
                    if chq_now.upper() not in det.upper():
                        r[narr_i] = (chq_now + " " + det).strip()
                    r[ref_i] = ""

        # Final pass strip after moves
        for i in range(len(r)):
            r[i] = _strip_footer_phrases(r[i] or "")
            if (r[i] or "").strip() in {"/", "-", "/-", "-/"}:
                r[i] = ""

        # Cheque printed once in the PDF must not repeat on the immediately
        # following wrap row. Do NOT blank repeating UPI/party labels (same
        # merchant on many real transactions).
        if ref_i >= 0:
            prev = ""
            for r in new_rows:
                if ref_i >= len(r):
                    prev = ""
                    continue
                v = re.sub(r"\s+", " ", (r[ref_i] or "")).strip()
                if not v or v in {"/", "-"}:
                    prev = ""
                    continue
                compact = re.sub(r"[\s\-/]", "", v)
                consecutive_cheque = (
                    v == prev
                    and _looks_like_ref_token(v)
                    and re.fullmatch(r"\d{6,9}", compact)
                )
                if consecutive_cheque:
                    r[ref_i] = ""
                prev = v

    # Drop rows that lost all substance after footer strip
    kept = []
    for r in new_rows:
        if _is_junk_or_legend_row(r):
            continue
        has_date = any(_is_date_text(c or "") or _split_leading_date(c or "")[0] for c in r)
        has_amt = any(_is_amount_text(c or "") for c in r)
        if has_date and has_amt:
            kept.append(r)
        elif has_date or has_amt:
            kept.append(r)
    new_rows = kept

    return new_header, new_rows


def _cluster_page_word_lines(words, y_tol: float = 3.5) -> List[List[Tuple[float, float, float, str]]]:
    """Cluster PyMuPDF words into visual lines: each item is (y0, x0, x1, text)."""
    items = []
    for w in words:
        if len(w) < 5:
            continue
        text = (w[4] or "").strip()
        if not text:
            continue
        items.append((float(w[1]), float(w[0]), float(w[2]), text))
    if not items:
        return []
    items.sort(key=lambda t: (round(t[0] / y_tol), t[1]))
    lines: List[List[Tuple[float, float, float, str]]] = []
    for it in items:
        if not lines:
            lines.append([it])
            continue
        if abs(it[0] - lines[-1][0][0]) <= y_tol:
            lines[-1].append(it)
        else:
            lines.append([it])
    for line in lines:
        line.sort(key=lambda t: t[1])
    return lines


_HEADER_KEYWORDS = (
    "date",
    "value",
    "narration",
    "remark",
    "particular",
    "description",
    "detail",
    "transaction",
    "withdrawal",
    "deposit",
    "debit",
    "credit",
    "balance",
    "amount",
    "cheque",
    "chq",
    "ref",
    "reference",
    "running",
    "s no",
    "sno",
    "#",
)


def _header_line_score(text: str) -> int:
    low = (text or "").lower()
    return sum(1 for k in _HEADER_KEYWORDS if k in low)


def _is_statement_period_line(text: str) -> bool:
    """True for 'From : 01 Apr 2025 To : 02 Dec 2025' (not a column header)."""
    joined = (text or "").strip()
    if not joined:
        return False
    low = joined.lower()
    # Must look like a period range, not "Date ... Debit ... Balance"
    if _header_line_score(joined) >= 3:
        return False
    has_from = bool(re.search(r"\bfrom\b", low))
    has_to = bool(re.search(r"\bto\b", low))
    date_hits = len(_DATE_RE.findall(joined)) + len(
        re.findall(
            r"\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4}",
            joined,
            re.I,
        )
    )
    if has_from and has_to and date_hits >= 1:
        return True
    if has_from and date_hits >= 1 and ":" in joined and _header_line_score(joined) < 2:
        return True
    return False


def _is_non_header_noise_line(text: str) -> bool:
    """Address / pincode / account-meta lines that sit near the real table header."""
    joined = (text or "").strip()
    if not joined:
        return True
    low = joined.lower()
    score = _header_line_score(joined)
    # Never treat real txn lines as header-noise (6-digit cheque/refs look like PINs)
    looks_like_txn = bool(
        _DATE_RE.search(joined)
        or _DATE_LEAD_RE.match(joined)
        or _AMOUNT_RE.search(joined.replace(" ", ""))
        or _is_sno_text(joined.split()[0] if joined.split() else "")
    )
    # Statement period line glued beside IndusInd/IDFC-style headers
    if _is_statement_period_line(joined):
        return True
    # Indian PIN / city-only lines above Kotak-style headers
    if re.search(r"\b\d{6}\b", joined) and score < 2 and not looks_like_txn:
        return True
    if (
        score == 0
        and not looks_like_txn
        and re.search(
            r"\b(gurugram|gurgaon|mumbai|delhi|bengaluru|bangalore|chennai|hyderabad)\b",
            low,
        )
    ):
        return True
    # Account / address blocks above or below the txn table (HDFC etc.)
    if score < 3 and any(
        k in low
        for k in (
            "account branch",
            "account no",
            "account status",
            "cust id",
            "joint holders",
            "od limit",
            "statement from",
            "nomination",
            "preferred customer",
            "a/c open date",
            "rtgs/neft",
            "branch code",
            "registered office",
            "page no",
            "statement of account",
            "email",
            "phone no",
            "closing balance includes",
            "address :",
            "state :",
            "city :",
            "account type",
        )
    ):
        return True
    return False


def _is_statement_letterhead_line(text: str) -> bool:
    """Top-of-page address / company / account-meta, not a txn wrap fragment."""
    t = (text or "").strip()
    if not t or t in {".", ":", "-", "—"}:
        return True
    if _is_non_header_noise_line(t) or _is_statement_page_meta_text(t):
        return True
    first = t.split()[0]
    if _is_date_text(first) or _DATE_LEAD_RE.match(t):
        return False
    if re.search(
        r"\b(NEFT|IMPS|UPI-|RTGS|POS\s|INW\s|ATM-|INB/|HDFCN|SALARY|NETBANK|MUM-)\b",
        t,
        re.I,
    ):
        return False
    low = t.lower()
    return bool(
        re.search(
            r"\b(account type|m/s\.|shop no|private limited|chambers|plaza|"
            r"golf course|sector\s+\d+|preferred customer)\b",
            low,
        )
    )


def _learn_columns_from_header_lines(
    header_lines: List[List[Tuple[float, float, float, str]]],
) -> Tuple[List[str], List[float]]:
    """
    From 1–3 header visual lines, build column titles and x-bounds.
    Vertically stacked header words (overlapping X) merge into one column;
    only a clear horizontal gap starts a new column.
    """
    # (y, x0, x1, text)
    words: List[Tuple[float, float, float, str]] = []
    for line in header_lines:
        for y, x0, x1, text in line:
            words.append((y, x0, x1, text))
    if not words:
        return [], []
    words.sort(key=lambda t: t[1])

    # Glue threshold must stay small: multi-word titles ("Cheque Number") have
    # ~2pt gaps, while distinct columns (Txn Date → Value Date) are often ~8–15pt.
    # Using median word-width * 1.35 previously merged Withdrawal/Deposit/Balance.
    glue_thresh = 6.5

    groups: List[List[Tuple[float, float, float, str]]] = [[words[0]]]
    for w in words[1:]:
        g = groups[-1]
        g_right = max(x1 for _y, _x0, x1, _t in g)
        gap = w[1] - g_right
        if gap < glue_thresh:
            g.append(w)
        else:
            groups.append([w])

    titles: List[str] = []
    lefts: List[float] = []
    rights: List[float] = []
    for g in groups:
        # Reading order: top-to-bottom then left-to-right within the stack
        ordered = sorted(g, key=lambda t: (t[0], t[1]))
        title = " ".join(t[3] for t in ordered)
        title = re.sub(r"\s+", " ", title).strip()
        title = title.replace(" (₹)", "").replace("(₹)", "").strip()
        # Drop stacked format crumbs glued under Axis/ICICI headers
        title = re.sub(r"\s*\(?\s*dd\s*/\s*mm\s*/\s*yyyy\s*\)?", "", title, flags=re.I)
        title = re.sub(r"\s+", " ", title).strip(" -")
        if not title:
            continue
        titles.append(title)
        lefts.append(min(t[1] for t in g))
        rights.append(max(t[2] for t in g))

    if len(titles) < 2:
        return titles, []

    _MONEY = ("withdraw", "deposit", "debit", "credit", "balance", "amount", "amt")

    def _is_money_title(t: str) -> bool:
        low = t.lower()
        return any(k in low for k in _MONEY)

    def _is_txn_date_title(t: str) -> bool:
        low = t.lower().strip()
        if "value" in low:
            return False
        if _is_money_title(low):
            return False
        return low in ("date", "txn date", "tran date", "transaction date") or (
            "date" in low and "value" not in low
        )

    def _is_narr_title(t: str) -> bool:
        low = t.lower()
        return any(k in low for k in ("desc", "narrat", "particular", "remark", "detail"))

    def _is_type_title(t: str) -> bool:
        low = t.lower().strip()
        return low in ("type", "txn type", "tran type", "transaction type", "mode", "dr/cr")

    def _is_ref_title(t: str) -> bool:
        low = t.lower()
        return any(k in low for k in ("ref", "chq", "cheque", "check"))

    def _cut(i: int) -> float:
        """X boundary between column i and i+1."""
        a, b = lefts[i], lefts[i + 1]
        ra = rights[i]
        la, lb = titles[i].lower(), titles[i + 1].lower()
        # Narrow Date → wide Narration: narration text starts left of its header word.
        # Cut just after the Date header (not midpoint of lefts).
        if _is_txn_date_title(titles[i]) and _is_narr_title(titles[i + 1]):
            return ra + 6.0
        # Value Date → Description/Narration (same idea — keep the date out of Desc)
        if "date" in la and _is_narr_title(titles[i + 1]):
            return ra + 6.0
        # Type → Description: values like "Transfer Debit" are wider than "Type"
        # and sit left of Description — midpoint steals "Debit" into Description.
        if _is_type_title(titles[i]) and _is_narr_title(titles[i + 1]):
            return max(ra + 8.0, b - 8.0)
        # Chq/Ref | Particulars/Narration/Remarks (Axis, ICICI):
        # Cheque is a narrow column. Particulars text is left-aligned and starts
        # well left of the Particulars header word. Cutting near Particulars
        # (b-8) dumps INB/ATM/M/THANE wrap into Chq.
        if _is_ref_title(titles[i]) and _is_narr_title(titles[i + 1]):
            return ra + 14.0
        # Two adjacent date cols (Txn Date | Value Date): wide DD/MM/YYYY midpoints
        # spill into Value — cut near Value header left so txn date stays left.
        if _is_txn_date_title(titles[i]) and "date" in lb:
            return max(ra + 4.0, b - 10.0)
        # Chq/Ref values are wider than the header; cut near Value Dt left.
        if _is_ref_title(titles[i]) and (
            "value" in lb or (_is_txn_date_title(titles[i + 1]) is False and "date" in lb)
        ):
            return max(ra + 2.0, b - 2.0)
        # Description | Ref: tune by actual header gap.
        # Wide gap (SBI): Ref values sit under "Ref" — cut in the gap so cheque/ref
        # tokens are not swallowed into Description.
        # Tight/overlap (YES): narration wraps under Ref header — bias cut toward Ref.
        desc_l = _is_narr_title(titles[i])
        ref_r = _is_ref_title(titles[i + 1])
        if desc_l and ref_r:
            gap = b - ra
            if gap >= 12.0:
                return ra + gap * 0.55
            return a * 0.20 + b * 0.80
        # Ref | Branch (or next non-money): Ref values are wider than the "Ref" word;
        # cut near next header left so long cheque/ref numbers stay in Ref.
        if _is_ref_title(titles[i]) and not _is_money_title(titles[i + 1]):
            return max(ra + 2.0, b - 8.0)
        # Non-money → money (Branch|Debit etc.): codes are right-aligned; midpoint
        # steals short branch codes (e.g. "1") into Debit.
        if not _is_money_title(titles[i]) and _is_money_title(titles[i + 1]):
            return max(ra + 4.0, b - 14.0)
        # Money columns are right-aligned under headers — midpoint steals withdrawals
        # into Deposit (HDFC 538.40 sits at the Withdrawal/Deposit midpoint).
        if _is_money_title(titles[i]) and _is_money_title(titles[i + 1]):
            return b - 4.0
        return (a + b) / 2.0

    bounds = [0.0]
    for i in range(len(lefts) - 1):
        bounds.append(_cut(i))
    page_right = max(w[2] for w in words) + 80
    bounds.append(max(page_right, lefts[-1] + 80))
    return titles, bounds


def _header_titles_look_polluted(titles: List[str]) -> int:
    """
    Count titles that look like data (dates/amounts/period crumbs) rather than
    column names. Used to reject 'From : 01 Apr ... Date Type Debit' merges.
    """
    bad = 0
    for t in titles:
        text = (t or "").strip()
        if not text:
            continue
        if _is_amount_text(text) or _is_date_text(text):
            bad += 1
            continue
        if _DATE_LEAD_RE.match(text):
            bad += 1
            continue
        if _is_statement_period_line(text):
            bad += 1
            continue
        # "01 Apr 2025 Type" / ": 02 Dec 2025 Description"
        if re.search(
            r"\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{4}",
            text,
            re.I,
        ):
            bad += 1
            continue
        if re.fullmatch(r"[:.\-]+", text):
            bad += 1
    return bad


def _find_header_window(
    lines: List[List[Tuple[float, float, float, str]]],
) -> Tuple[int, int, List[str], List[float]]:
    """
    Find best 1–3 consecutive lines that look like a table header.
    Returns (start_idx, end_idx_exclusive, titles, bounds).
    """
    best = (-1, -1, [], [])
    best_score = 0
    n = len(lines)
    for i in range(n):
        for width in (1, 2, 3):
            if i + width > n:
                break
            raw_window = lines[i : i + width]
            # Drop address / pincode / period lines that sit beside real header rows
            window = [
                line
                for line in raw_window
                if not _is_non_header_noise_line(" ".join(t[3] for t in line))
            ]
            if not window:
                continue
            joined = " ".join(t[3] for line in window for t in line)
            score = _header_line_score(joined)
            # Prefer lines that mention both a date-ish and money-ish concept
            low = joined.lower()
            if "date" in low:
                score += 2
            if any(k in low for k in ("balance", "withdrawal", "deposit", "debit", "credit")):
                score += 2
            if score < 3:
                continue
            titles, bounds = _learn_columns_from_header_lines(window)
            if len(titles) < 4 or len(bounds) < 5:
                continue
            polluted = _header_titles_look_polluted(titles)
            if polluted:
                # Period+header merges invent extra cols and poison X bounds
                continue
            # Clean keyword titles (Date/Type/Debit…) beat raw column count
            clean_titles = sum(
                1
                for t in titles
                if _header_line_score(t) >= 1
                or (t or "").strip().lower() in ("type", "particulars", "remarks", "mode")
            )
            # Prefer more columns + higher keyword score; slight bias to fewer noise lines
            rank = score * 10 + clean_titles * 5 + len(titles) * 2 - len(window)
            # Prefer a single pure header line over multi-line merges of equal quality
            if len(window) == 1 and clean_titles >= 4:
                rank += 8
            # Axis-style: primary header + stacked "Date" / "Number" continuation
            # under Transaction / Cheque — reward keeping that continuation line.
            if len(window) >= 2:
                cont = " ".join(t[3] for t in window[-1]).lower()
                if re.search(r"\b(date|number|dd/mm|yyyy)\b", cont) and _header_line_score(
                    cont
                ) <= 2:
                    rank += 15
                    # Prefer completed titles like "Transaction Date" / "Cheque Number"
                    if any("date" in (t or "").lower() and "transaction" in (t or "").lower() for t in titles):
                        rank += 10
                    if any("cheque" in (t or "").lower() and "number" in (t or "").lower() for t in titles):
                        rank += 6
            if rank > best_score:
                best_score = rank
                best = (i, i + width, titles, bounds)
    return best


def _col_index_for_x(x: float, bounds: List[float]) -> int:
    for i in range(len(bounds) - 1):
        if bounds[i] <= x < bounds[i + 1]:
            return i
    return max(0, len(bounds) - 2)


def _normalize_date_cell(text: str) -> str:
    """Join split date tokens already in one cell; leave exact characters."""
    return re.sub(r"\s+", " ", (text or "").strip())


def _row_has_money(row: List[str], money_cols: List[int]) -> bool:
    for i in money_cols:
        if i < len(row) and _is_amount_text(row[i]):
            return True
    # fallback: any amount in row
    return any(_is_amount_text(c or "") for c in row)


def _row_has_date(row: List[str], date_cols: List[int]) -> bool:
    for i in date_cols:
        if i < len(row) and _is_date_text(_normalize_date_cell(row[i])):
            return True
    return any(_is_date_text(_normalize_date_cell(c or "")) for c in row)


def _reseat_row_amounts(titles: List[str], row: List[str], money_cols: List[int]) -> None:
    """Move amounts that landed in Init/Br back into Debit/Credit/Balance."""
    bal_i = -1
    dr_cr: List[int] = []
    for i, h in enumerate(titles):
        hl = (h or "").lower()
        if "balance" in hl:
            bal_i = i
        elif i in money_cols and "balance" not in hl:
            dr_cr.append(i)
    for i, h in enumerate(titles):
        if i >= len(row):
            continue
        hl = (h or "").lower().strip()
        if not any(k in hl for k in ("init", "branch")) and hl not in ("br", "br."):
            continue
        raw = (row[i] or "").strip()
        if not raw:
            continue
        keep: List[str] = []
        amts: List[str] = []
        for p in raw.split():
            if _is_amount_text(p):
                amts.append(p)
            else:
                keep.append(p)
        if not amts:
            continue
        if keep and keep[0].lower().rstrip(".") in ("br", "branch"):
            keep = keep[1:]
        if bal_i >= 0 and not (row[bal_i] or "").strip() and amts:
            row[bal_i] = amts[-1]
            amts = amts[:-1]
        for amt in amts:
            placed = False
            for mi in dr_cr:
                if mi < len(row) and not (row[mi] or "").strip():
                    row[mi] = amt
                    placed = True
                    break
            if not placed and bal_i >= 0:
                row[bal_i] = amt
        row[i] = " ".join(keep).strip()


def _adaptive_line_starts_txn(
    cells: List[str],
    date_cols: List[int],
    sno_cols: List[int],
    money_cols: List[int],
) -> bool:
    """True when a visual line is the start of a transaction (sno/date, often with amounts)."""
    del money_cols
    has_date_col = False
    for di in date_cols:
        if di < len(cells) and _is_date_text(_normalize_date_cell(cells[di])):
            has_date_col = True
            break
    has_sno = False
    if sno_cols:
        has_sno = any(_is_sno_text(cells[i]) for i in sno_cols if i < len(cells))
    else:
        has_sno = _is_sno_text(cells[0]) if cells else False
    return has_date_col or has_sno


def _extract_adaptive_statement(pdf_path: str, status_cb=None) -> List[List[List[str]]]:
    """
    Generic bank/statement extractor — NOT bank-specific.
    1) Detect header row(s) from keywords
    2) Learn column names + X bounds from that header
    3) Map every word on every page into those columns
    4) Start a new Excel row when a line looks like a new txn (date/sno + amount)
    5) Append continuation lines into the current txn
    Result: one table, all pages stacked, real header names.
    """
    try:
        import pymupdf
    except Exception:
        return []

    doc = pymupdf.open(pdf_path)
    try:
        total = doc.page_count
        titles: List[str] = []
        bounds: List[float] = []
        header_end_y_by_page: dict = {}

        # Learn columns from the first page that has a clear header
        for pi, page in enumerate(doc):
            if status_cb:
                status_cb(f"Adaptive: learning header on page {pi + 1}/{total}...")
            lines = _cluster_page_word_lines(page.get_text("words") or [])
            start, end, t, b = _find_header_window(lines)
            if not t or len(t) < 4 or len(b) < 5:
                s2, e2, t2, b2 = _find_simple_txn_header(lines)
                if t2 and b2:
                    start, end, t, b = s2, e2, t2, b2
            if t and b:
                titles, bounds = t, b
                # y below header window
                header_end_y_by_page[pi] = max(w[0] for line in lines[start:end] for w in line) + 2
                break
        if not titles or len(titles) < 3 or len(bounds) < 4:
            return []

        ncols = len(titles)
        money_cols = [
            i
            for i, name in enumerate(titles)
            if any(
                k in name.lower()
                for k in (
                    "withdrawal",
                    "deposit",
                    "debit",
                    "credit",
                    "balance",
                    "amount",
                    "running",
                )
            )
            # Axis "Debit/Credit" is CR/DR flag, not an amount column
            and "debit/credit" not in name.lower().replace(" ", "")
            and name.lower().strip() not in ("debit/credit", "dr/cr", "cr/dr")
        ]
        date_cols = [i for i, name in enumerate(titles) if _header_is_date_col(name)]
        # Axis multi-line header often labels txn-date col as bare "Transaction"
        if not date_cols:
            for i, name in enumerate(titles):
                if name.strip().lower() in ("transaction", "txn", "tran"):
                    date_cols.append(i)
        if not date_cols:
            date_cols = [0, 1][:ncols]
        sno_cols = [
            i
            for i, name in enumerate(titles)
            if name.strip() in ("#", "S No.", "S No", "SNo", "Sl", "Sl.", "S.NO", "S.No")
            or "s no" in name.lower()
            or "s.no" in name.lower()
            or re.sub(r"[\s.]+", "", name.lower()) in ("sno", "srno", "slno", "sr", "sl")
        ]
        chq_idxs = [
            i
            for i, name in enumerate(titles)
            if any(k in name.lower() for k in ("cheque", "chq", "check"))
        ]

        all_rows: List[List[str]] = []
        current: List[str] | None = None
        pending_money: List[str] | None = None
        pending_prefix: List[str] | None = None

        for pi, page in enumerate(doc):
            if status_cb:
                status_cb(f"Adaptive extract: page {pi + 1}/{total}...")
            lines = _cluster_page_word_lines(page.get_text("words") or [])
            # Skip repeated header on later pages
            start_line = 0
            s, e, t2, b2 = _find_header_window(lines)
            if not t2:
                s, e, t2, b2 = _find_simple_txn_header(lines)
            if t2:
                txn_before_header = False
                for line in lines[:s]:
                    joined = " ".join(w[3] for w in line).strip()
                    first = joined.split()[0] if joined.split() else ""
                    if _is_date_text(first) or _DATE_LEAD_RE.match(joined):
                        txn_before_header = True
                        break
                if not txn_before_header:
                    start_line = e
                # Prefer first-page bounds; optionally refresh if column count matches
                if len(t2) == ncols and pi > 0:
                    pass
            if pi > 0 and start_line == 0:
                # Skip reprinted letterhead, but keep wrap from the previous page
                # (HDFC "BANK, MUM-..." / Axis "INB/..." at the top of the next page).
                while start_line < len(lines) and _is_statement_letterhead_line(
                    " ".join(w[3] for w in lines[start_line])
                ):
                    start_line += 1

            # Do not reset `current` per page — HDFC wrap continues onto the next page.
            past_txn_block = False

            def flush():
                nonlocal current
                if current is None:
                    return
                # Peel narration that leaked into date cells before validating
                for di in date_cols:
                    if di >= len(current):
                        continue
                    d, rest = _peel_date_cell(current[di])
                    if d:
                        current[di] = d
                        if rest and _is_statement_page_meta_text(rest):
                            rest = ""
                        if rest:
                            spilled = False
                            for j, name in enumerate(titles):
                                nl = name.lower()
                                if any(
                                    k in nl
                                    for k in (
                                        "detail",
                                        "narration",
                                        "remark",
                                        "description",
                                        "particular",
                                    )
                                ):
                                    current[j] = (rest + " " + (current[j] or "")).strip()
                                    spilled = True
                                    break
                            if not spilled:
                                for j in range(len(current)):
                                    if j in date_cols or j in money_cols:
                                        continue
                                    current[j] = (rest + " " + (current[j] or "")).strip()
                                    break
                    elif _is_statement_page_meta_text(current[di] or ""):
                        current[di] = ""
                # Amounts that landed in Cheque (wrong X cut) → real money column,
                # never leave rupees sitting in Chq/Ref.
                for ci in chq_idxs:
                    if ci >= len(current):
                        continue
                    raw = (current[ci] or "").strip()
                    if not raw:
                        continue
                    amt_parts: List[str] = []
                    keep_parts: List[str] = []
                    for p in raw.split():
                        if _is_amount_text(p):
                            amt_parts.append(p)
                        else:
                            keep_parts.append(p)
                    if not amt_parts:
                        continue
                    current[ci] = " ".join(keep_parts).strip()
                    prefer: List[int] = []
                    bal_only: List[int] = []
                    for mi in money_cols:
                        if mi >= len(current):
                            continue
                        name = titles[mi].lower() if mi < len(titles) else ""
                        if "balance" in name:
                            bal_only.append(mi)
                        else:
                            prefer.append(mi)
                    for amt in amt_parts:
                        placed = False
                        for mi in prefer + bal_only:
                            if not (current[mi] or "").strip():
                                current[mi] = amt
                                placed = True
                                break
                        if not placed and bal_only:
                            current[bal_only[0]] = amt
                # Require some substance
                if _row_has_date(current, date_cols) and _row_has_money(current, money_cols):
                    # Normalize date cells
                    for di in date_cols:
                        if di < len(current):
                            current[di] = _normalize_date_cell(current[di])
                    _reseat_row_amounts(titles, current, money_cols)
                    # Empty Chq in the PDF must stay empty — never keep another column's text
                    for ci in chq_idxs:
                        if ci >= len(current):
                            continue
                        raw = (current[ci] or "").strip()
                        if not raw:
                            continue
                        hdr = titles[ci] if ci < len(titles) else ""
                        if _chq_column_accepts(hdr, raw):
                            continue
                        for j, name in enumerate(titles):
                            nl = name.lower()
                            if any(
                                k in nl
                                for k in (
                                    "particular",
                                    "narrat",
                                    "remark",
                                    "desc",
                                    "detail",
                                )
                            ):
                                current[j] = (raw + " " + (current[j] or "")).strip()
                                current[ci] = ""
                                break
                    all_rows.append(list(current))
                current = None

            def put_word(row: List[str], x0: float, x1: float, text: str):
                ci = min(_col_index_for_x((x0 + x1) / 2.0, bounds), ncols - 1)
                # Never drop Particulars/UPI into an empty Cheque column.
                if ci in chq_idxs:
                    hdr = titles[ci] if ci < len(titles) else ""
                    if not _chq_column_accepts(hdr, text):
                        for j, name in enumerate(titles):
                            nl = name.lower()
                            if any(
                                k in nl
                                for k in (
                                    "particular",
                                    "narrat",
                                    "remark",
                                    "desc",
                                    "detail",
                                )
                            ):
                                ci = j
                                break
                if row[ci]:
                    row[ci] = row[ci] + " " + text
                else:
                    row[ci] = text

            def _merge_orphan_into(target: List[str], source: List[str], money_first: bool):
                """Merge Axis orphan amount-line into a date/sno line (or vice versa)."""
                for i, val in enumerate(source):
                    text = (val or "").strip()
                    if not text:
                        continue
                    if i in chq_idxs:
                        hdr = titles[i] if i < len(titles) else ""
                        if target[i] or not _chq_column_accepts(hdr, text):
                            continue
                        target[i] = text
                        continue
                    if i in money_cols and _is_amount_text(text):
                        if not target[i] or not _is_amount_text(target[i]):
                            target[i] = text
                        continue
                    if i in date_cols or i in sno_cols:
                        if not target[i]:
                            target[i] = text
                        continue
                    if target[i]:
                        if money_first:
                            target[i] = (text + " " + target[i]).strip()
                        else:
                            target[i] = (target[i] + " " + text).strip()
                    else:
                        target[i] = text

            def _spill_prefix_to_detail(target: List[str], text: str):
                for j, name in enumerate(titles):
                    nl = name.lower()
                    if any(
                        k in nl
                        for k in (
                            "detail",
                            "narration",
                            "remark",
                            "description",
                            "particular",
                        )
                    ):
                        if text.upper() not in (target[j] or "").upper():
                            target[j] = (text + " " + (target[j] or "")).strip()
                        return
                for j, name in enumerate(titles):
                    if j in date_cols or j in sno_cols or j in money_cols or j in chq_idxs:
                        continue
                    if text.upper() not in (target[j] or "").upper():
                        target[j] = (text + " " + (target[j] or "")).strip()
                    return

            def _merge_pending_prefix(target: List[str], source: List[str]):
                """Attach the label printed above sno+date (usually the first remarks line)."""
                for i, val in enumerate(source):
                    text = (val or "").strip()
                    if not text:
                        continue
                    if _is_statement_page_meta_text(text):
                        continue
                    if i in date_cols or i in sno_cols or i in money_cols:
                        continue
                    if i in chq_idxs:
                        # Prefix line is remarks/party sitting above the date — not cheque fill-down.
                        _spill_prefix_to_detail(target, text)
                        continue
                    if target[i]:
                        if text.upper() not in target[i].upper():
                            target[i] = (text + " " + target[i]).strip()
                    else:
                        target[i] = text

            body_lines = lines[start_line:]
            for li, line in enumerate(body_lines):
                if past_txn_block:
                    break
                joined = " ".join(t[3] for t in line)
                low = joined.lower()
                if _header_line_score(joined) >= 3 and "date" in low:
                    continue
                if _is_non_header_noise_line(joined):
                    continue
                if _is_statement_letterhead_line(joined):
                    continue
                if _is_statement_page_meta_text(joined):
                    continue
                if _is_statement_footer_line(joined) or joined.strip().upper() in {
                    "HDFC BANK LIMITED",
                    "HDFC BANK LTD",
                }:
                    pending_money = None
                    pending_prefix = None
                    past_txn_block = True
                    break
                end_markers = (
                    "statement generated",
                    "legends for",
                    "system generated",
                    "unless constituent",
                    "registered office",
                    "this is a system",
                    "sincerly",
                    "sincerely",
                    "closing balance includes",
                    "contents of this statement",
                    "hdfc bank gstin",
                    "computer generated",
                    "require a signature",
                    "statement summary",
                    "transaction total",
                    "++++ end of report",
                    "end of report",
                    "opening balance:",
                    "closing balance:",
                    "closing balance ",
                    "legend :",
                    "iconn -",
                )
                if any(k in low for k in end_markers) or (
                    "page " in low and "of " in low and len(joined) < 40
                ):
                    pending_money = None
                    pending_prefix = None
                    # Keep current txn open so page-break wrap can attach.
                    # Only end the whole statement on a true totals/end marker.
                    if any(
                        k in low
                        for k in (
                            "transaction total",
                            "end of report",
                            "unless constituent",
                            "legend :",
                        )
                    ):
                        flush()
                        current = None
                        past_txn_block = True
                        break
                    past_txn_block = True
                    break

                cells = [""] * ncols
                for _y, x0, x1, text in line:
                    put_word(cells, x0, x1, text)

                has_money = _row_has_money(cells, money_cols)
                has_date_col = False
                for di in date_cols:
                    if di < len(cells) and _is_date_text(_normalize_date_cell(cells[di])):
                        has_date_col = True
                        break
                has_sno = False
                if sno_cols:
                    has_sno = any(_is_sno_text(cells[i]) for i in sno_cols if i < len(cells))
                else:
                    has_sno = _is_sno_text(cells[0]) if cells else False

                # Date inside narration is wrap text, not a new transaction.
                has_identity = has_date_col or has_sno

                # Complete starter: date/sno + amount on same visual line
                starts_new = has_money and has_identity
                # Axis staggered: sno+date line without amounts (amount was on prior line)
                if not starts_new and has_identity and not has_money:
                    starts_new = True
                    if current is not None and has_date_col and not has_sno:
                        cur_dates = {
                            _normalize_date_cell(current[di])
                            for di in date_cols
                            if di < len(current)
                            and _is_date_text(_normalize_date_cell(current[di]))
                        }
                        new_dates = {
                            _normalize_date_cell(cells[di])
                            for di in date_cols
                            if di < len(cells)
                            and _is_date_text(_normalize_date_cell(cells[di]))
                        }
                        if cur_dates and new_dates and (cur_dates & new_dates):
                            # Same date reprinted on a wrap line — keep one row
                            starts_new = False

                if starts_new:
                    flush()
                    current = cells
                    if pending_prefix is not None:
                        _merge_pending_prefix(current, pending_prefix)
                        pending_prefix = None
                    if pending_money is not None:
                        _merge_orphan_into(current, pending_money, money_first=True)
                        pending_money = None
                    continue

                # Orphan amount line (Axis: amounts print above sno/date)
                if has_money and not has_identity:
                    if current is not None and not _row_has_money(current, money_cols):
                        _merge_orphan_into(current, cells, money_first=False)
                    else:
                        pending_money = cells
                    continue

                # Party / cheque / Axis particulars printed on the line above the next txn
                if (
                    not has_identity
                    and not has_money
                    and (
                        _looks_like_leading_overlay(joined)
                        or _looks_like_new_particulars_line(joined)
                    )
                ):
                    nxt_is_txn = False
                    if li + 1 < len(body_lines):
                        nxt = body_lines[li + 1]
                        nxt_join = " ".join(t[3] for t in nxt)
                        if (
                            not _is_statement_footer_line(nxt_join)
                            and not _is_non_header_noise_line(nxt_join)
                        ):
                            nxt_cells = [""] * ncols
                            for _y, x0, x1, text in nxt:
                                put_word(nxt_cells, x0, x1, text)
                            nxt_is_txn = _adaptive_line_starts_txn(
                                nxt_cells, date_cols, sno_cols, money_cols
                            )
                    if nxt_is_txn or current is None:
                        steal_for_next = _looks_like_bank_txn_code_prefix(joined) or (
                            current is None and _looks_like_leading_overlay(joined)
                        )
                        if (
                            current is not None
                            and _row_has_money(current, money_cols)
                            and not steal_for_next
                        ):
                            # HDFC wrap sitting above the next date — keep on current.
                            pass
                        else:
                            pending_prefix = cells
                            continue

                if current is None:
                    # Narration fragment before any txn — ignore
                    continue

                # Continuation line — merge into current without corrupting
                # already-valid date / sno / money cells (e.g. Kotak "02:33 PM"
                # landing in the date column would break flush validation).
                def _spill_to_detail(text: str):
                    for j, name in enumerate(titles):
                        nl = name.lower()
                        if any(
                            k in nl
                            for k in (
                                "detail",
                                "narration",
                                "remark",
                                "description",
                                "particular",
                            )
                        ):
                            current[j] = (current[j] + " " + text).strip()
                            return True
                    return False

                for i, val in enumerate(cells):
                    text = (val or "").strip()
                    if not text:
                        continue
                    if _is_statement_page_meta_text(text):
                        continue
                    if i in money_cols:
                        if current[i] and _is_amount_text(current[i]):
                            if not _is_amount_text(text):
                                _spill_to_detail(text)
                            continue
                        # Wrap lines must not fill an empty Debit/Credit/Withdrawal.
                        # Amounts for this txn already sit on the date line (or orphan merge).
                        _spill_to_detail(text)
                        continue
                    if i in chq_idxs:
                        hdr = titles[i] if i < len(titles) else ""
                        if not (current[i] or "").strip() and _chq_column_accepts(hdr, text):
                            current[i] = text
                            continue
                        # Wrap/UPI must not fill an empty Cheque column.
                        _spill_to_detail(text)
                        continue
                    if i in date_cols and current[i] and _is_date_text(
                        _normalize_date_cell(current[i])
                    ):
                        if not _is_date_text(_normalize_date_cell(text)):
                            _spill_to_detail(text)
                        continue
                    if i in sno_cols and current[i] and _is_sno_text(current[i]):
                        if not _is_sno_text(text):
                            _spill_to_detail(text)
                        continue
                    if current[i]:
                        current[i] = (current[i] + " " + text).strip()
                    else:
                        current[i] = text
            # keep `current` open across pages
        flush()
    finally:
        doc.close()

    if not all_rows:
        return []
    # Clean header titles (remove newlines artifacts)
    clean_header = [re.sub(r"\s+", " ", h).strip() for h in titles]
    return [_normalize_table_exact([clean_header] + all_rows)]


_STACKED_TYPE_RE = re.compile(
    r"^(?:transfer\s+)?(?:debit|credit|withdrawal|deposit|neft|rtgs|imps|upi|ach|charge|interest)\b"
    r"|^(?:dr|cr)\b",
    re.I,
)
_STACKED_DASH_RE = re.compile(r"^[-–—]$")


def _lines_look_like_vertical_field_stack(lines: List[str]) -> bool:
    """
    True when many lines are a lone exact date or lone amount — typical of
    IndusInd/Axis text exports. Rejects SBI/YES multi-token lines.
    """
    exact_dates = 0
    exact_amts = 0
    solo_header_hits = 0
    for raw in lines:
        t = (raw or "").strip()
        if not t:
            continue
        if _DATE_RE.match(t):
            exact_dates += 1
        elif _is_amount_text(t) or _STACKED_DASH_RE.match(t):
            exact_amts += 1
        elif len(t) <= 16 and _header_line_score(t) >= 1 and " " not in t.strip():
            solo_header_hits += 1
    if exact_dates < 3 or exact_amts < max(6, exact_dates):
        return False
    # Prefer layouts that also expose header words one-per-line
    return solo_header_hits >= 3 or exact_dates >= 5


def _parse_stacked_statement_text(lines: List[str]) -> List[List[str]]:
    """
    Rebuild rows when PDF text stores each field on its own line (vertical stack).

    Common IndusInd / Axis-export pattern after get_text():
      Date
      Type
      Description
      Debit
      Credit
      Balance
      08 May 2025
      Transfer Credit
      N/SBIN.../M
      rs KAMLESH KHARB
      -
      60,000.00
      67,684.97

    Or without Type: Date → narration lines → amount(s) → balance.
    """
    if not _lines_look_like_vertical_field_stack(lines):
        return []

    cleaned: List[str] = []
    for raw in lines:
        t = (raw or "").strip()
        if not t:
            continue
        if _is_statement_footer_line(t):
            continue
        low = t.lower()
        if low.startswith("page ") and " of " in low:
            continue
        if _is_statement_period_line(t):
            continue
        cleaned.append(t)

    if not cleaned:
        return []

    # Locate header keywords appearing as consecutive single-token/short lines
    header_idx = -1
    header_cols: List[str] = []
    best_hdr_score = 0
    for i in range(len(cleaned)):
        window: List[str] = []
        for j in range(i, min(i + 10, len(cleaned))):
            piece = cleaned[j]
            if _is_date_text(piece) or _is_amount_text(piece):
                break
            if len(piece) > 40:
                break
            window.append(piece)
            joined = " ".join(window)
            score = _header_line_score(joined)
            low = joined.lower()
            if "date" in low:
                score += 2
            if any(k in low for k in ("balance", "debit", "credit", "withdrawal", "deposit")):
                score += 2
            if (
                score >= 5
                and len(window) >= 3
                and score >= best_hdr_score
                and all(not _is_amount_text(w) for w in window)
            ):
                best_hdr_score = score
                header_idx = i
                header_cols = list(window)
        # Headers appear near the top of the statement body
        if header_idx >= 0 and i > header_idx + 30:
            break

    if header_idx < 0 or len(header_cols) < 3:
        # Fallback header guess for IndusInd-style stacks
        header_cols = ["Date", "Type", "Description", "Debit", "Credit", "Balance"]
        body_start = 0
        for i, t in enumerate(cleaned):
            if _is_date_text(t):
                body_start = i
                break
        else:
            return []
    else:
        body_start = header_idx + len(header_cols)

    money_names = ("debit", "credit", "withdrawal", "deposit", "balance", "amount")
    n_money = sum(1 for h in header_cols if any(k in h.lower() for k in money_names))
    if n_money < 1:
        n_money = 2
    has_type_col = any(h.lower().strip() == "type" for h in header_cols)
    narr_idx = next(
        (
            i
            for i, h in enumerate(header_cols)
            if any(k in h.lower() for k in ("desc", "narrat", "particular", "remark", "detail"))
        ),
        1 if has_type_col else 1,
    )
    type_idx = next((i for i, h in enumerate(header_cols) if h.lower().strip() == "type"), None)
    date_idx = next((i for i, h in enumerate(header_cols) if "date" in h.lower()), 0)
    money_idxs = [
        i for i, h in enumerate(header_cols) if any(k in h.lower() for k in money_names)
    ]
    if not money_idxs:
        money_idxs = list(range(len(header_cols) - n_money, len(header_cols)))

    ncols = len(header_cols)
    rows: List[List[str]] = []
    current: List[str] | None = None
    money_buf: List[str] = []
    narr_parts: List[str] = []
    type_val = ""

    def flush():
        nonlocal current, money_buf, narr_parts, type_val
        if current is None:
            return
        row = list(current)
        if type_idx is not None and type_val:
            row[type_idx] = type_val
        if narr_parts and narr_idx < ncols:
            row[narr_idx] = " ".join(narr_parts).strip()
        # Map buffered money tokens into money columns (left→right)
        vals = list(money_buf)
        # Prefer aligning from the right when we have balance + optional debit/credit
        if len(vals) > len(money_idxs):
            vals = vals[-len(money_idxs) :]
        # Pad left with blanks if short
        while len(vals) < len(money_idxs):
            vals.insert(0, "")
        for mi, mv in zip(money_idxs, vals):
            if mi < ncols:
                # Lone dash means empty amount cell
                row[mi] = "" if _STACKED_DASH_RE.match(mv or "") else (mv or "")
        if _is_date_text(row[date_idx] if date_idx < len(row) else "") and any(
            _is_amount_text(row[i]) for i in money_idxs if i < len(row)
        ):
            rows.append(row)
        current = None
        money_buf = []
        narr_parts = []
        type_val = ""

    for t in cleaned[body_start:]:
        if _looks_like_header_row([t]) and _header_line_score(t) >= 2:
            # Repeated page header word — skip single keyword lines
            if t.lower() in {h.lower() for h in header_cols}:
                continue
        if _is_date_text(t):
            flush()
            current = [""] * ncols
            current[date_idx] = t
            continue
        if current is None:
            continue
        if _is_amount_text(t) or _STACKED_DASH_RE.match(t):
            money_buf.append(t)
            # Enough money tokens for this layout → ready for next date
            if len(money_buf) >= len(money_idxs):
                flush()
            continue
        # Type line (only once, before amounts)
        if (
            has_type_col
            and not type_val
            and not money_buf
            and _STACKED_TYPE_RE.match(t)
            and len(t) <= 40
        ):
            type_val = t
            continue
        # Narration / continuation
        if not money_buf:
            narr_parts.append(t)
        else:
            # Text after partial amounts is unusual — treat as narration spill
            narr_parts.append(t)

    flush()
    if not rows:
        return []
    return [header_cols] + rows


def _extract_stacked_statement(pdf_path: str, status_cb=None) -> List[List[List[str]]]:
    """
    Generic fallback when words are stacked one field per line (no useful X layout
    from text-mode extracts). Rebuilds Date / Type / Description / Debit / Credit / Balance.
    """
    try:
        import pymupdf
    except Exception:
        return []

    all_lines: List[str] = []
    doc = pymupdf.open(pdf_path)
    try:
        total = doc.page_count
        for pi, page in enumerate(doc):
            if status_cb:
                status_cb(f"Stacked-text rebuild: page {pi + 1}/{total}...")
            text = page.get_text("text") or ""
            all_lines.extend(text.splitlines())
    finally:
        doc.close()

    table = _parse_stacked_statement_text(all_lines)
    if not table or len(table) < 2:
        return []
    return [_normalize_table_exact(table)]


def _extract_column_anchored_bank(pdf_path: str, status_cb=None) -> List[List[List[str]]]:
    """
    YES Bank / similar: map words into fixed columns by X position.
    All pages → one logical table (caller merges). Exact text kept as-is.
    Columns:
      Transaction Date | Value Date | Description | Reference Number |
      Withdrawals | Deposits | Running Balance
    """
    try:
        import pymupdf
    except Exception:
        return []

    HEADER = [
        "Transaction Date",
        "Value Date",
        "Description",
        "Reference Number",
        "Withdrawals",
        "Deposits",
        "Running Balance",
    ]
    # Default YES Bank-ish anchors (midpoints / left edges from real statement)
    # Desc|Ref cut ~240 so wrapped description ("FIN 10", "BANK") stays out of Ref.
    default_bounds = [0, 70, 120, 240, 355, 430, 495, 900]

    def col_index(x: float, bounds: List[float]) -> int:
        for i in range(len(bounds) - 1):
            if bounds[i] <= x < bounds[i + 1]:
                return i
        return len(bounds) - 2

    def detect_bounds(words) -> List[float]:
        """Learn column bounds from a header line with Date / Withdrawals / Balance."""
        lines: dict = {}
        for w in words:
            if len(w) < 5:
                continue
            y = round(float(w[1]) / 3.0) * 3.0
            lines.setdefault(y, []).append(w)
        for _y, lw in sorted(lines.items()):
            texts = " ".join((w[4] or "") for w in lw).lower()
            if "withdrawal" not in texts or "deposit" not in texts:
                continue
            if "balance" not in texts and "running" not in texts:
                continue
            xs = {}
            for w in sorted(lw, key=lambda z: float(z[0])):
                t = (w[4] or "").lower()
                x0 = float(w[0])
                if "value" in t:
                    xs["vd"] = x0
                elif t in ("transaction", "date") and "vd" not in xs and "td" not in xs:
                    xs["td"] = x0
                elif "description" in t:
                    xs["desc"] = x0
                elif "reference" in t:
                    xs["ref"] = x0
                elif "withdrawal" in t:
                    xs["wd"] = x0
                elif "deposit" in t:
                    xs["dep"] = x0
                # Prefer leftmost of Running/Balance (do not overwrite Running with Balance)
                elif ("running" in t or "balance" in t) and "bal" not in xs:
                    xs["bal"] = x0
            lefts = [
                xs.get("td", 30.0),
                xs.get("vd", 78.0),
                xs.get("desc", 132.0),
                xs.get("ref", 255.0),
                xs.get("wd", 364.0),
                xs.get("dep", 441.0),
                xs.get("bal", 501.0),
            ]
            # Bias Description→Reference cut toward Ref (desc text wraps right).
            bounds = [0.0]
            for i in range(len(lefts) - 1):
                a, b = lefts[i], lefts[i + 1]
                if i == 2:  # desc → ref
                    bounds.append(a * 0.05 + b * 0.95)
                else:
                    bounds.append((a + b) / 2.0)
            bounds.append(900.0)
            return bounds
        # No YES-style Withdrawals/Deposits header — do not force YES defaults
        # onto other banks (SBI Debit/Credit etc.).
        return None

    all_rows: List[List[str]] = []
    doc = pymupdf.open(pdf_path)
    try:
        total = doc.page_count
        bounds = None
        for pi, page in enumerate(doc, start=1):
            if status_cb:
                status_cb(f"Column-anchored extract: page {pi}/{total}...")
            words = page.get_text("words") or []
            if pi == 1:
                bounds = detect_bounds(words)
                if not bounds:
                    return []
            if not bounds:
                return []

            # Cluster into visual lines
            items = []
            for w in words:
                if len(w) < 5:
                    continue
                text = (w[4] or "").strip()
                if not text:
                    continue
                items.append((float(w[1]), float(w[0]), float(w[2]), text))
            if not items:
                continue
            items.sort(key=lambda t: (round(t[0] / 3.0), t[1]))

            lines: List[List[Tuple[float, float, float, str]]] = []
            for it in items:
                if not lines:
                    lines.append([it])
                    continue
                if abs(it[0] - lines[-1][0][0]) <= 3.5:
                    lines[-1].append(it)
                else:
                    lines.append([it])

            current: List[str] | None = None

            def flush():
                nonlocal current
                if current is None:
                    return
                # Keep rows that have a transaction date and at least one amount/balance
                if _is_date_text(current[0]) and (
                    _is_amount_text(current[4])
                    or _is_amount_text(current[5])
                    or _is_amount_text(current[6])
                ):
                    all_rows.append(list(current))
                current = None

            for line in lines:
                line.sort(key=lambda t: t[1])
                # leftmost word
                left_text = line[0][3]
                left_x = line[0][1]
                starts_txn = _is_date_text(left_text) and left_x < bounds[2]

                if starts_txn:
                    flush()
                    current = [""] * 7
                    for _y0, x0, x1, text in line:
                        mid = (x0 + x1) / 2.0
                        ci = col_index(mid, bounds)
                        ci = min(max(ci, 0), 6)
                        if current[ci]:
                            current[ci] = current[ci] + " " + text
                        else:
                            current[ci] = text
                    continue

                if current is None:
                    continue

                # Continuation / multi-line description & reference
                joined = " ".join(t[3] for t in line)
                if _is_statement_footer_line(joined):
                    continue
                low = joined.lower()
                if any(
                    k in low
                    for k in (
                        "page ",
                        "statement of account",
                        "customer name",
                        "unless constituent",
                        "registered office",
                        "yes bank",
                        "transaction date",
                        "value date",
                        "running balance",
                        "computer generated",
                        "require a signature",
                    )
                ):
                    continue

                for _y0, x0, x1, text in line:
                    mid = (x0 + x1) / 2.0
                    ci = col_index(mid, bounds)
                    ci = min(max(ci, 0), 6)
                    # Don't overwrite filled amount columns with non-amounts
                    if ci in (4, 5, 6):
                        if current[ci]:
                            continue
                        if not _is_amount_text(text):
                            # stray text near amount cols → description
                            ci = 2
                    if current[ci]:
                        current[ci] = current[ci] + " " + text
                    else:
                        current[ci] = text
            flush()
    finally:
        doc.close()

    if not all_rows:
        return []
    return [_normalize_table_exact([HEADER] + all_rows)]


def _extract_kotak_bank(pdf_path: str, status_cb=None) -> List[List[List[str]]]:
    """
    Kotak statements: dates like '15 Aug 2026', signed amounts -10,500.00 / +24,65,000.00,
    multi-line details. Geometry tables only find the header — parse words by X bands.
    All pages stacked into one table.
    """
    try:
        import pymupdf
    except Exception:
        return []

    HEADER = [
        "S No.",
        "Transaction Date",
        "Value Date",
        "Transaction Details",
        "Chq / Ref No.",
        "Debit/Credit",
        "Balance",
    ]
    # x cutoffs from Kotak layout
    # # <70 | date 70-145 | value 145-195 | details 195-320 | ref 320-400 | amt 400-470 | bal >=470
    bounds = [0.0, 70.0, 145.0, 195.0, 320.0, 400.0, 470.0, 900.0]

    def col_index(x: float) -> int:
        for i in range(len(bounds) - 1):
            if bounds[i] <= x < bounds[i + 1]:
                return i
        return 6

    def looks_kotak(doc) -> bool:
        sample = ""
        for page in doc[: min(2, doc.page_count)]:
            sample += (page.get_text("text") or "")[:2500]
        low = sample.lower()
        return (
            ("debit/credit" in low and "transaction date" in low)
            or "kkbk" in low
            or "kotak" in low
        )

    all_rows: List[List[str]] = []
    doc = pymupdf.open(pdf_path)
    try:
        if not looks_kotak(doc):
            return []
        total = doc.page_count
        stacked_rows: List[List[str]] = []
        stacked_header: List[str] = []
        for pi, page in enumerate(doc, start=1):
            if status_cb:
                status_cb(f"Kotak statement extract: page {pi}/{total}...")
            parsed = _parse_bank_text_transactions(page.get_text("text") or "")
            if parsed and len(parsed) > 1:
                if not stacked_header:
                    stacked_header = list(parsed[0])
                stacked_rows.extend(parsed[1:])
        if len(stacked_rows) >= 3:
            return [_normalize_table_exact([stacked_header or HEADER] + stacked_rows)]

        for pi, page in enumerate(doc, start=1):
            if status_cb:
                status_cb(f"Kotak statement extract: page {pi}/{total}...")
            words = page.get_text("words") or []
            items = []
            for w in words:
                if len(w) < 5:
                    continue
                text = (w[4] or "").strip()
                if not text:
                    continue
                items.append((float(w[1]), float(w[0]), float(w[2]), text))
            if not items:
                continue
            items.sort(key=lambda t: (round(t[0] / 3.0), t[1]))

            lines: List[List[Tuple[float, float, float, str]]] = []
            for it in items:
                if not lines:
                    lines.append([it])
                    continue
                if abs(it[0] - lines[-1][0][0]) <= 3.5:
                    lines[-1].append(it)
                else:
                    lines.append([it])

            current: List[str] | None = None

            def flush():
                nonlocal current
                if current is None:
                    return
                if (
                    _is_sno_text(current[0])
                    and _is_date_text(current[1])
                    and (_is_amount_text(current[5]) or _is_amount_text(current[6]))
                ):
                    all_rows.append(list(current))
                current = None

            for line in lines:
                line.sort(key=lambda t: t[1])
                joined = " ".join(t[3] for t in line)
                low = joined.lower()
                if any(
                    k in low
                    for k in (
                        "transaction date",
                        "value date",
                        "transaction details",
                        "debit/credit",
                        "statement generated",
                        "page ",
                        "account statement",
                        "account #",
                        "privihka finiworld",
                    )
                ) and not any(_is_amount_text(t[3]) for t in line):
                    # header / footer / title lines
                    if not (_is_sno_text(line[0][3]) and any(_is_amount_text(t[3]) for t in line)):
                        continue

                # New txn: serial at left + amount on same visual band
                left = line[0]
                has_sno = left[1] < max(90.0, float(page.rect.width) * 0.12) and _is_sno_text(
                    left[3]
                )
                has_amt = any(_is_amount_text(t[3]) for t in line)
                if has_sno and has_amt:
                    flush()
                    current = [""] * 7
                    for _y, x0, x1, text in line:
                        ci = col_index((x0 + x1) / 2.0)
                        if current[ci]:
                            current[ci] = current[ci] + " " + text
                        else:
                            current[ci] = text
                    continue

                if current is None:
                    continue

                # Continuation: time + extra details/ref (no new serial+amount)
                for _y, x0, x1, text in line:
                    ci = col_index((x0 + x1) / 2.0)
                    # time sits under date column in PDF — keep date clean, put time in details
                    if ci == 1 and (
                        re.match(r"^\d{1,2}:\d{2}$", text)
                        or text.upper() in ("AM", "PM")
                    ):
                        current[3] = (current[3] + " " + text).strip()
                        continue
                    if ci in (5, 6):
                        if current[ci] or not _is_amount_text(text):
                            ci = 3
                    if current[ci]:
                        current[ci] = current[ci] + " " + text
                    else:
                        current[ci] = text
            flush()
    finally:
        doc.close()

    if not all_rows:
        return []
    return [_normalize_table_exact([HEADER] + all_rows)]


def _same_header(a: List[str], b: List[str]) -> bool:
    if not a or not b:
        return False
    left = [x.strip().lower() for x in a if x.strip()]
    right = [x.strip().lower() for x in b if x.strip()]
    if not left or not right:
        return False
    return left == right or (
        len(set(left) & set(right)) >= min(3, min(len(left), len(right)))
    )


def _unique_sheet_name(used: set, base: str) -> str:
    name = re.sub(r"[:\\/?*\[\]]", "_", base)[:31] or "Sheet"
    if name not in used:
        used.add(name)
        return name
    i = 2
    while True:
        suffix = f"_{i}"
        candidate = name[: 31 - len(suffix)] + suffix
        if candidate not in used:
            used.add(candidate)
            return candidate
        i += 1


def _is_network_path(path: str) -> bool:
    """True for UNC paths like \\\\server\\share\\... which are slow to process."""
    p = path.replace("/", "\\")
    return p.startswith("\\\\")


def _local_pdf_copy(pdf_path: str) -> Tuple[str, bool]:
    """
    Copy PDF to local temp when on a network drive so extraction is faster.
    Returns (path_to_use, delete_when_done).
    """
    import shutil
    import tempfile

    if not _is_network_path(pdf_path):
        return pdf_path, False
    tmp = tempfile.NamedTemporaryFile(prefix="ssa_pdf_", suffix=".pdf", delete=False)
    tmp.close()
    shutil.copy2(pdf_path, tmp.name)
    return tmp.name, True


def _extract_page_tables(page) -> list:
    """
    Fast text-based table detection only.
    Never use pdfplumber 'lines' strategy — it can hang for hours on bank PDFs.
    """
    text_settings = {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "min_words_vertical": 1,
        "min_words_horizontal": 1,
        "intersection_tolerance": 5,
        "snap_tolerance": 3,
        "join_tolerance": 3,
    }
    tables = page.extract_tables(table_settings=text_settings) or []
    return [
        t
        for t in tables
        if t and any(any(c not in (None, "") for c in (row or [])) for row in t)
    ]


def _raw_table_has_stacked_cells(raw) -> bool:
    """True when find_tables mashed several PDF rows into one cell with newlines."""
    for row in (raw or [])[1:]:
        for cell in row or []:
            if cell is None:
                continue
            text = cell if isinstance(cell, str) else str(cell)
            if "\n" in text and sum(1 for p in text.splitlines() if p.strip()) >= 2:
                return True
    return False


def _find_table_header_x_bounds(tab) -> List[float]:
    """X edges of the short header cells — not the tall stacked body cells."""
    bbox = getattr(tab, "bbox", None)
    cells = getattr(tab, "cells", None) or []
    if not bbox or not cells:
        return []
    y_top = float(bbox[1])
    header = []
    for c in cells:
        if not c:
            continue
        x0, y0, x1, y1 = (float(c[0]), float(c[1]), float(c[2]), float(c[3]))
        if y0 <= y_top + 3.0 and (y1 - y0) < 22.0:
            header.append((x0, x1))
    header.sort()
    if len(header) < 2:
        return []
    bounds = [header[0][0]]
    for i in range(len(header) - 1):
        bounds.append(header[i][1])
    bounds.append(max(header[-1][1], float(bbox[2])))
    return bounds


def _rebuild_stacked_table_from_words(page, tab, raw) -> List[List[str]]:
    """
    Payslip / invoice grids: find_tables returns 1 body row with newlines.
    Rebuild one Excel row per visual PDF line using column X from the header.
    """
    if not raw or not raw[0]:
        return []
    bounds = _find_table_header_x_bounds(tab)
    if len(bounds) < 3:
        return []
    header = [_cell_as_exact_text(c) for c in raw[0]]
    header = [h for h in header if (h or "").strip()]
    ncols = len(bounds) - 1
    if len(header) > ncols:
        header = header[:ncols]
    while len(header) < ncols:
        header.append("")
    bbox = tab.bbox
    x0, y0, x1, y1 = float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3])
    header_bottom = y0 + 12.0
    hcells = [
        c
        for c in (tab.cells or [])
        if c and float(c[1]) <= y0 + 3.0 and (float(c[3]) - float(c[1])) < 22.0
    ]
    if hcells:
        header_bottom = max(float(c[3]) for c in hcells) + 1.0
    words = []
    for w in page.get_text("words") or []:
        if len(w) < 5:
            continue
        tok = (w[4] or "").strip()
        if not tok:
            continue
        wx0, wy0, wx1, wy1 = float(w[0]), float(w[1]), float(w[2]), float(w[3])
        midy = (wy0 + wy1) / 2.0
        midx = (wx0 + wx1) / 2.0
        if midy < header_bottom or midy > y1 + 2.0:
            continue
        if midx < x0 - 4.0 or midx > x1 + 8.0:
            continue
        words.append(w)
    if not words:
        return []
    lines = _cluster_page_word_lines(words, y_tol=3.5)
    body: List[List[str]] = []
    for line in lines:
        cells = [""] * ncols
        for _y, wx0, wx1, tok in line:
            midx = (wx0 + wx1) / 2.0
            ci = _col_index_for_x(midx, bounds)
            ci = min(max(ci, 0), ncols - 1)
            cells[ci] = (cells[ci] + " " + tok).strip()
        joined = " ".join(cells).strip().lower()
        if joined.startswith("rupees:") or joined.startswith("this is a"):
            continue
        if any(c.strip() for c in cells):
            body.append(cells)
    if len(body) < 2:
        return []
    return [header] + body


def _extract_tables_pymupdf(pdf_path: str, status_cb=None) -> List[List[List[str]]]:
    """Fast table extraction with PyMuPDF (usually much faster than pdfplumber)."""
    try:
        import pymupdf  # PyMuPDF
    except Exception:
        return []

    all_tables: List[List[List[str]]] = []
    doc = pymupdf.open(pdf_path)
    try:
        total = doc.page_count
        for i, page in enumerate(doc, start=1):
            if status_cb:
                status_cb(f"PyMuPDF tables: page {i}/{total}...")
            try:
                finder = page.find_tables()
            except Exception:
                continue
            tables = getattr(finder, "tables", None) or []
            for tab in tables:
                try:
                    raw = tab.extract()
                except Exception:
                    continue
                if not raw:
                    continue
                if _raw_table_has_stacked_cells(raw):
                    rebuilt = _rebuild_stacked_table_from_words(page, tab, raw)
                    exact = _normalize_table_exact(rebuilt or raw)
                else:
                    exact = _normalize_table_exact(raw)
                if exact and any(any(c != "" for c in row) for row in exact):
                    all_tables.append(exact)
    finally:
        doc.close()
    return all_tables


def _peek_ruled_statement_ncols(pdf_path: str) -> int:
    """
    If page 1 has a drawn table grid (ICICI Detailed Statement etc.), return
    its column count. Adaptive guessing must not replace these cells.
    """
    try:
        import pymupdf
    except Exception:
        return 0
    doc = None
    try:
        doc = pymupdf.open(pdf_path)
        finder = doc[0].find_tables()
        tabs = getattr(finder, "tables", None) or []
        best = 0
        for tab in tabs:
            try:
                raw = tab.extract()
            except Exception:
                continue
            if not raw or not raw[0]:
                continue
            joined = re.sub(
                r"[\r\n]+", " ", " ".join((c or "") for c in raw[0])
            ).lower()
            if "balance" not in joined:
                continue
            if not any(
                k in joined
                for k in ("remark", "particular", "narrat", "description")
            ):
                continue
            best = max(best, len(raw[0]))
        return best
    except Exception:
        return 0
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass


def _usable_generic_table(tables: List[List[List[str]]]) -> bool:
    """True when the PDF already gave a real table — don't let guess parsers smash it."""
    if not tables or _extract_is_garbage(tables):
        return False
    rows = [r for b in tables for r in b]
    if len(rows) < 2:
        return False
    cols = max(len(r) for r in rows)
    filled = sum(1 for r in rows for c in r if (c or "").strip())
    return cols >= 3 and filled >= 8


def _tables_share_statement_header(page_tables: List[List[List[str]]]) -> bool:
    """True when every block is the same bank-statement grid (merge into one sheet)."""
    headers: List[List[str]] = []
    for block in page_tables or []:
        if block and _looks_like_header_row(block[0]):
            headers.append(
                [re.sub(r"\s+", " ", (c or "")).strip().lower() for c in block[0]]
            )
    if len(headers) < 1 or len(headers[0]) < 4:
        return False
    first = headers[0]
    same = 0
    for h in headers:
        if h == first:
            same += 1
            continue
        if len(h) != len(first):
            continue
        hit = sum(1 for a, b in zip(h, first) if a == b)
        if hit >= max(3, int(len(first) * 0.7)):
            same += 1
    return same >= max(1, int(len(headers) * 0.6))


def _cluster_words_into_rows(words, y_tol: float = 3.0, gap_factor: float = 1.6) -> List[List[str]]:
    """
    Build table-like rows from PyMuPDF words (x0,y0,x1,y1,word,...).
    Groups by vertical position, then splits columns on large horizontal gaps.
    Preserves exact word text — no reformatting.
    """
    if not words:
        return []

    items = []
    for w in words:
        if len(w) < 5:
            continue
        text = (w[4] or "").strip()
        if not text:
            continue
        x0, y0, x1, y1 = float(w[0]), float(w[1]), float(w[2]), float(w[3])
        items.append((y0, x0, x1, text))
    if not items:
        return []

    items.sort(key=lambda t: (round(t[0] / y_tol), t[1]))

    # Group into visual lines by y
    lines: List[List[Tuple[float, float, float, str]]] = []
    for item in items:
        if not lines:
            lines.append([item])
            continue
        prev_y = lines[-1][0][0]
        if abs(item[0] - prev_y) <= y_tol:
            lines[-1].append(item)
        else:
            lines.append([item])

    rows: List[List[str]] = []
    for line in lines:
        line.sort(key=lambda t: t[1])
        widths = [max(1.0, t[2] - t[1]) for t in line]
        avg_w = sum(widths) / len(widths)
        gap_thresh = max(8.0, avg_w * gap_factor)

        cells: List[str] = []
        buf = line[0][3]
        prev_x1 = line[0][2]
        for x0, _y0, x1, text in line[1:]:
            gap = x0 - prev_x1
            if gap >= gap_thresh:
                cells.append(buf.strip())
                buf = text
            else:
                # Same column — join with space (exact tokens)
                if buf and not buf.endswith(" ") and not text.startswith(" "):
                    buf = buf + " " + text
                else:
                    buf = buf + text
            prev_x1 = x1
        cells.append(buf.strip())
        if any(c for c in cells):
            rows.append(cells)
    return rows


def _extract_words_rows_pymupdf(pdf_path: str, status_cb=None) -> List[List[List[str]]]:
    """Position-based extract: uses word coordinates for accurate columns."""
    try:
        import pymupdf  # PyMuPDF
    except Exception:
        return []

    all_tables: List[List[List[str]]] = []
    doc = pymupdf.open(pdf_path)
    try:
        total = doc.page_count
        for i, page in enumerate(doc, start=1):
            if status_cb:
                status_cb(f"Word-position extract: page {i}/{total}...")
            try:
                words = page.get_text("words") or []
            except Exception:
                words = []
            rows = _cluster_words_into_rows(words)
            if rows:
                exact = _normalize_table_exact(rows)
                if exact and any(any(c != "" for c in row) for row in exact):
                    all_tables.append(exact)
    finally:
        doc.close()
    return all_tables


def _extract_text_rows_pymupdf(pdf_path: str, status_cb=None) -> List[List[List[str]]]:
    """
    Ultra-fast fallback: split each page's text into rows/columns.
    Finishes in seconds even on large statements (no table geometry scan).
    """
    try:
        import pymupdf  # PyMuPDF
    except Exception:
        return []

    all_tables: List[List[List[str]]] = []
    doc = pymupdf.open(pdf_path)
    try:
        total = doc.page_count
        for i, page in enumerate(doc, start=1):
            if status_cb:
                status_cb(f"Fast text extract: page {i}/{total}...")
            text = page.get_text("text") or ""
            rows: List[List[str]] = []
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue
                cells = [c.strip() for c in re.split(r"\s{2,}|\t", line) if c.strip()]
                if not cells:
                    cells = [line]
                rows.append(cells)
            if rows:
                exact = _normalize_table_exact(rows)
                if exact and any(any(c != "" for c in row) for row in exact):
                    all_tables.append(exact)
    finally:
        doc.close()
    return all_tables


def _pdf_extractable_char_count(pdf_path: str) -> int:
    """How many text characters are selectable in the PDF (0 ≈ scanned/image-only)."""
    try:
        import pymupdf
    except Exception:
        return 0
    total = 0
    doc = pymupdf.open(pdf_path)
    try:
        for page in doc:
            total += len((page.get_text("text") or "").strip())
            if total > 50:
                break
    finally:
        doc.close()
    return total


def _tesseract_candidates() -> List[str]:
    """Known install locations for the tesseract binary, per platform."""
    if os.name == "nt":
        return [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe"),
            os.path.expandvars(r"%USERPROFILE%\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
        ]
    if sys.platform == "darwin":
        # Launched from Finder a .app inherits a minimal PATH that excludes both
        # Homebrew prefixes, so these have to be probed explicitly.
        return [
            "/opt/homebrew/bin/tesseract",  # Apple Silicon Homebrew
            "/usr/local/bin/tesseract",  # Intel Homebrew
            "/opt/local/bin/tesseract",  # MacPorts
        ]
    return ["/usr/bin/tesseract", "/usr/local/bin/tesseract"]


def _find_tessdata(exe: str) -> str:
    """
    Locate the tessdata folder holding eng.traineddata for a given binary.

    Windows keeps tessdata next to the binary, while Homebrew/MacPorts and most
    Linux packages install it under ../share, so both layouts are searched.
    """
    root = os.path.dirname(exe)
    prefix = os.path.dirname(root)
    candidates = [
        os.environ.get("TESSDATA_PREFIX", ""),
        os.path.join(os.environ.get("TESSDATA_PREFIX", ""), "tessdata"),
        os.path.join(root, "tessdata"),
        os.path.join(prefix, "share", "tessdata"),
        os.path.join(prefix, "share", "tesseract-ocr", "tessdata"),
        "/opt/homebrew/share/tessdata",
        "/usr/local/share/tessdata",
        "/opt/local/share/tessdata",
        "/usr/share/tessdata",
        "/usr/share/tesseract-ocr/tessdata",
    ]
    # Homebrew nests tessdata under a version folder, e.g. share/tessdata/5.
    for base in ("/opt/homebrew/share", "/usr/local/share"):
        versioned = os.path.join(base, "tesseract-ocr")
        if os.path.isdir(versioned):
            for entry in sorted(os.listdir(versioned), reverse=True):
                candidates.append(os.path.join(versioned, entry, "tessdata"))

    for alt in candidates:
        if alt and os.path.isfile(os.path.join(alt, "eng.traineddata")):
            return alt
    return ""


def _find_tesseract() -> Tuple[str, str]:
    """
    Locate the tesseract binary and its tessdata folder.
    Returns (tesseract_exe, tessdata_dir) or ("", "").
    """
    import shutil

    candidates = []
    which = shutil.which("tesseract")
    if which:
        candidates.append(which)
    candidates.extend(_tesseract_candidates())

    first_exe = ""
    for exe in candidates:
        if not exe or not os.path.isfile(exe):
            continue
        if not first_exe:
            first_exe = exe
        tessdata = _find_tessdata(exe)
        if tessdata:
            return exe, tessdata
    # Binary present but no language data: report it so the caller can say so.
    return first_exe, ""


def _tesseract_install_command() -> str:
    if sys.platform == "darwin":
        return "brew install tesseract"
    if os.name == "nt":
        return "winget install --id UB-Mannheim.TesseractOCR"
    return "sudo apt install tesseract-ocr"


def _tesseract_install_help() -> str:
    if sys.platform == "darwin":
        source = (
            "If the brew command is not found, install Homebrew first\n"
            "from https://brew.sh and then run the command above.\n"
        )
    elif os.name == "nt":
        source = "Or download: https://github.com/UB-Mannheim/tesseract/wiki\n"
    else:
        source = "Or use your distribution's package manager.\n"
    return (
        "This PDF is scanned (image-only) — Tesseract OCR is required.\n\n"
        "Install once, then restart SSA PDF Studio:\n"
        f"  {_tesseract_install_command()}\n\n"
        + source
        + "After install, run PDF to Excel again."
    )


def _prepare_tesseract_env() -> str:
    """
    Put Tesseract on PATH and return tessdata directory for PyMuPDF OCR.
    Raises RuntimeError with install steps when Tesseract is missing.
    """
    exe, tessdata = _find_tesseract()
    if not exe:
        raise RuntimeError(_tesseract_install_help())

    tess_dir = os.path.dirname(exe)
    path_parts = os.environ.get("PATH", "").split(os.pathsep)
    if tess_dir not in path_parts:
        os.environ["PATH"] = tess_dir + os.pathsep + os.environ.get("PATH", "")

    if tessdata and os.path.isdir(tessdata) and os.path.isfile(
        os.path.join(tessdata, "eng.traineddata")
    ):
        # Tesseract expects TESSDATA_PREFIX = parent of the tessdata folder,
        # which is the bin dir on Windows but ../share under Homebrew.
        os.environ["TESSDATA_PREFIX"] = os.path.dirname(tessdata) + os.sep
        return tessdata

    raise RuntimeError(
        "Tesseract was found but eng.traineddata is missing.\n"
        "Reinstall Tesseract and enable the English language pack.\n\n"
        + _tesseract_install_help()
    )


def _extract_ocr_pymupdf(pdf_path: str, status_cb=None) -> List[List[List[str]]]:
    """
    OCR fallback for scanned / image-only PDFs (needs Tesseract installed).
    Returns word-clustered rows so columns stay as close to the page layout as possible.
    Fails fast if Tesseract is missing (does not loop every page for 10+ minutes).
    """
    try:
        import pymupdf
    except Exception:
        return []

    tessdata = _prepare_tesseract_env()
    if status_cb:
        status_cb("Tesseract found — starting OCR...")

    all_tables: List[List[List[str]]] = []
    doc = pymupdf.open(pdf_path)
    try:
        total = doc.page_count
        for i, page in enumerate(doc, start=1):
            if status_cb:
                status_cb(f"OCR (scanned PDF): page {i}/{total}...")
            try:
                tp = page.get_textpage_ocr(
                    flags=3,
                    language="eng",
                    dpi=200,
                    full=True,
                    tessdata=tessdata,
                )
                words = page.get_text("words", textpage=tp) or []
            except TypeError:
                # Older PyMuPDF without tessdata= kwarg
                try:
                    tp = page.get_textpage_ocr(flags=3, language="eng", dpi=200, full=True)
                    words = page.get_text("words", textpage=tp) or []
                except Exception as exc:
                    raise RuntimeError(
                        f"OCR failed on page {i}: {exc}\n\n{_tesseract_install_help()}"
                    ) from exc
            except Exception as exc:
                # Abort immediately — do not burn time on remaining pages
                raise RuntimeError(
                    f"OCR failed on page {i}: {exc}\n\n{_tesseract_install_help()}"
                ) from exc

            rows = _cluster_words_into_rows(words)
            if not rows:
                try:
                    text = page.get_text("text", textpage=tp) or ""
                except Exception:
                    text = ""
                rows = []
                for line in text.splitlines():
                    line = line.strip()
                    if not line:
                        continue
                    cells = [c.strip() for c in re.split(r"\s{2,}|\t", line) if c.strip()]
                    rows.append(cells or [line])
            if rows:
                exact = _normalize_table_exact(rows)
                if exact and any(any(c != "" for c in row) for row in exact):
                    all_tables.append(exact)
    finally:
        doc.close()
    return all_tables


def _flatten_tables_to_rows(page_tables: List[List[List[str]]]) -> List[List[str]]:
    """Concatenate all page blocks into one row list (exact cell text)."""
    out: List[List[str]] = []
    for block in page_tables:
        for row in block:
            if any((c or "").strip() for c in row):
                out.append(list(row))
    return out


def _extract_tables_pdfplumber_fast(pdf_path: str, status_cb=None) -> List[List[List[str]]]:
    """pdfplumber text-strategy only (never lines). Used as last resort."""
    try:
        import pdfplumber
    except Exception:
        return []

    page_tables: List[List[List[str]]] = []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages, start=1):
            if status_cb:
                status_cb(f"pdfplumber (fast): page {i}/{total}...")
            for table in _extract_page_tables(page):
                exact_rows = _normalize_table_exact(table)
                if exact_rows and any(any(c != "" for c in row) for row in exact_rows):
                    page_tables.append(exact_rows)
    return page_tables


def _parse_excel_number(value: str):
    """Turn a money cell like '6,46,301.00' / '-10,500.00' into a float. Else None."""
    text = (value or "").strip()
    if not text:
        return None
    t = text.replace(" ", "").replace("$", "").replace("₹", "")
    t = re.sub(r"^(?:Rs\.?|INR)", "", t, flags=re.I)
    t = re.sub(r"(?:Dr|Cr|DR|CR)$", "", t)
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg = True
        t = t[1:-1]
    if t.startswith("+"):
        t = t[1:]
    elif t.startswith("-"):
        neg = True
        t = t[1:]
    if not re.fullmatch(r"[\d,]+\.\d{2}", t):
        return None
    try:
        n = float(t.replace(",", ""))
    except ValueError:
        return None
    return -n if neg else n


def _header_is_amount_col(header: str) -> bool:
    low = (header or "").lower()
    if any(k in low for k in ("chq", "cheque", "check no", "ref no", "reference")):
        return False
    return any(
        k in low
        for k in (
            "withdraw",
            "deposit",
            "debit",
            "credit",
            "balance",
            "amount",
            "amt",
        )
    )


def _write_table_sheet(wb, sheet_name: str, table_rows: List[List[str]]):
    """Write extracted rows. Amount columns are real Excel numbers so SUM works."""
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet(title=sheet_name)
    header_fill = PatternFill("solid", fgColor="F3F4F6")
    header_font = Font(bold=True)
    header = table_rows[0] if table_rows else []
    amount_cols = set()
    if header and _looks_like_header_row(header):
        for i, h in enumerate(header):
            if _header_is_amount_col(h):
                amount_cols.add(i)

    for r_idx, row in enumerate(table_rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            text = "" if value is None else str(value)
            if r_idx == 1:
                cell = ws.cell(row=r_idx, column=c_idx, value=text)
                cell.number_format = "@"
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(wrap_text=True, vertical="center")
                continue
            num = None
            if (c_idx - 1) in amount_cols or _is_amount_text(text):
                num = _parse_excel_number(text)
            if num is not None:
                cell = ws.cell(row=r_idx, column=c_idx, value=num)
                cell.number_format = "#,##0.00"
            else:
                cell = ws.cell(row=r_idx, column=c_idx, value=text)
                cell.number_format = "@"

    if table_rows:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        col_count = len(table_rows[0])
        sample = table_rows[: min(30, len(table_rows))]
        for c_idx in range(1, col_count + 1):
            max_len = 0
            for row in sample:
                if c_idx - 1 < len(row):
                    max_len = max(max_len, len(str(row[c_idx - 1] or "")))
            width = 48 if c_idx == 2 else min(max(max_len + 2, 10), 36)
            ws.column_dimensions[get_column_letter(c_idx)].width = width


def _peek_pdf_text(pdf_path: str, pages: int = 2, max_chars: int = 12000) -> str:
    """First pages of selectable text — used only to detect which statement template to lock."""
    try:
        import pymupdf
    except Exception:
        return ""
    doc = pymupdf.open(pdf_path)
    try:
        chunks: List[str] = []
        for page in doc[: min(pages, doc.page_count)]:
            chunks.append(page.get_text("text") or "")
        return "\n".join(chunks)[:max_chars]
    finally:
        doc.close()


_BANK_PROFILE_LABEL = {
    "karnataka": "Karnataka Bank passbook",
    "kotak": "Kotak Mahindra Bank",
    "icici": "ICICI Bank",
    "axis": "Axis Bank",
    "hdfc": "HDFC Bank",
    "yes": "YES Bank",
    "sbi": "State Bank of India",
    "indusind": "IndusInd Bank",
    "statement": "bank statement",
    "unknown": "generic PDF",
}

# Known statement banks: lock one extractor. Never run passbook/bank-text contests.
_STATEMENT_LOCK_PROFILES = frozenset(
    {
        "karnataka",
        "kotak",
        "icici",
        "axis",
        "hdfc",
        "yes",
        "sbi",
        "indusind",
        "statement",
    }
)


def _normalize_spaced_letterhead(text: str) -> str:
    """Turn 'T H E   K A R N A T A K A   B A N K' into 'THE   KARNATAKA   BANK'."""
    # Only collapse a single space between one-letter tokens so word gaps stay.
    return re.sub(r"(?<=\b[A-Za-z]) (?=[A-Za-z]\b)", "", text or "")


def _detect_statement_profile(pdf_path: str) -> str:
    """
    Detect the bank/layout from the first-page letterhead and column titles.
    Ignore transaction narrations (UPI often contains other banks' names).
    """
    raw = _peek_pdf_text(pdf_path, pages=1, max_chars=8000)
    if not (raw or "").strip():
        raw = _peek_pdf_text(pdf_path, pages=2, max_chars=12000)
    letterhead = _normalize_spaced_letterhead(raw or "")
    # Cut before the first transaction date so UPI text cannot steal the bank.
    header_only = re.split(
        r"(?m)^\s*\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b",
        letterhead,
        maxsplit=1,
    )[0]
    head = re.sub(r"\s+", " ", header_only.lower())
    page = re.sub(r"\s+", " ", letterhead.lower())[:3500]
    if not head.strip() and not page.strip():
        return "unknown"

    # Layout fingerprints (more reliable than a bank name inside a UPI string)
    if "karnataka bank" in head or "karnataka bank" in page[:900]:
        return "karnataka"
    if "statement of axis account" in page or (
        "tran date" in page
        and "particulars" in page
        and "chq no" in page
        and ("init" in page or "opening balance" in page)
    ):
        return "axis"
    if "yesb0" in head or "yes bank" in head:
        return "yes"
    if (
        "your branch details" in page
        and "value date" in page
        and "description" in page
    ):
        return "yes"
    if "tran id" in page and "transaction remarks" in page and (
        "sl no" in page or "cheque no" in page
    ):
        return "icici"
    if "transaction remarks" in page and "cheque number" in page:
        return "icici"
    if "narration" in page and "closing balance" in page and (
        "chq./ref" in page or "chq/ref" in page or "withdrawal amt" in page
    ):
        return "hdfc"

    checks = (
        ("karnataka", ("karnataka bank", "the karnataka bank")),
        ("kotak", ("kotak mahindra", "kotak bank")),
        ("icici", ("icici bank", "icicibank.com")),
        ("axis", ("axis bank", "axisbank.com", "axis account")),
        ("hdfc", ("hdfc bank", "hdfcbank.com")),
        ("yes", ("yes bank", "yesbank.in")),
        ("sbi", ("state bank of india", "sbi.co.in", "sbi internet banking")),
        ("indusind", ("indusind bank",)),
    )
    for key, needles in checks:
        if any(n in head for n in needles):
            return key

    fname = os.path.basename(pdf_path or "").lower()
    file_hints = (
        ("karnataka", ("karnataka", "kbl")),
        ("kotak", ("kotak",)),
        ("icici", ("icici",)),
        ("axis", ("axis",)),
        ("hdfc", ("hdfc",)),
        ("yes", ("yes bank", "yesbank")),
        ("sbi", ("sbi", "state bank")),
        ("indusind", ("indusind",)),
    )
    for key, needles in file_hints:
        if any(n in fname for n in needles):
            return key

    grid = (
        "particulars",
        "narration",
        "transaction remarks",
        "tran date",
        "transaction date",
        "withdrawal amt",
        "withdrawal amount",
        "value date",
        "chq no",
        "cheque no",
        "chq./ref",
    )
    if "balance" in page and any(g in page for g in grid):
        return "statement"
    return "unknown"


def pdf_to_excel():
    try:
        from openpyxl import Workbook
    except Exception:
        raise RuntimeError("Please install: python3 -m pip install openpyxl")

    import time

    ensure_pdf_tools()
    pdf_path = ask_pdf_file("Select PDF to extract tables/text into Excel")

    reader = PdfReader(pdf_path)
    encrypted = False
    try:
        encrypted = bool(reader.is_encrypted)
    except Exception:
        encrypted = False

    output_folder = ask_output_folder()
    output_name = ask_text("Output file name", "Enter output Excel file name:", "pdf_data.xlsx")
    if not output_name.lower().endswith(".xlsx"):
        output_name += ".xlsx"
    out = os.path.join(output_folder, safe_filename(output_name))

    def status_cb(msg: str):
        try:
            print(msg, flush=True)
        except UnicodeEncodeError:
            enc = getattr(sys.stdout, "encoding", None) or "ascii"
            print(msg.encode(enc, errors="replace").decode(enc, errors="replace"), flush=True)
        root = getattr(tk, "_default_root", None)
        if root is not None:
            try:
                root.title(f"{APP_NAME} - {msg}")
                root.update_idletasks()
            except Exception:
                pass

    work_path, is_temp = _local_pdf_copy(pdf_path)
    if is_temp:
        status_cb("Copied PDF to local temp for speed...")
    if encrypted:
        status_cb("PDF is password-protected. Enter password to continue...")
        unlocked = _plain_copy_for_merge(work_path if is_temp else pdf_path)
        if is_temp:
            try:
                os.remove(work_path)
            except OSError:
                pass
        work_path = unlocked
        is_temp = True

    started = time.perf_counter()
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    used_names = set()
    combined_header: List[str] = []
    combined_rows: List[List[str]] = []
    raw_table_count = 0
    engine_used = "pymupdf"
    sheet_written = ""

    try:
        page_tables: List[List[List[str]]] = []
        text_chars = _pdf_extractable_char_count(work_path)
        is_scanned = text_chars < 20

        if is_scanned:
            # Skip empty table/text passes — they take minutes and find nothing on scans
            engine_used = "ocr"
            status_cb("PDF is scanned (no selectable text). Checking Tesseract OCR...")
            try:
                page_tables = _extract_ocr_pymupdf(work_path, status_cb=status_cb)
            except RuntimeError as exc:
                status_cb(str(exc).split("\n")[0])
                page_tables = []
                engine_used = "none"
        else:
            page_tables = []
            engine_used = "none"
            best_score = 0
            best_quality = -10_000

            def _consider(name: str, tables: List[List[List[str]]], label: str = ""):
                """Keep the candidate with the most real transaction rows (quality tie-break)."""
                nonlocal page_tables, engine_used, best_score, best_quality
                if not tables:
                    return
                # Never let Column_*/fragment/address-block garbage beat a real extract
                cand_garbage = _extract_is_garbage(tables)
                cur_garbage = _extract_is_garbage(page_tables) if page_tables else True
                if cand_garbage and not cur_garbage and best_score > 0:
                    return
                # Hard block: Column_* headers never win over a named-header extract
                cand_header = tables[0][0] if tables and tables[0] else []
                cur_header = page_tables[0][0] if page_tables and page_tables[0] else []
                if (
                    _header_is_column_numbered(cand_header)
                    and cur_header
                    and not _header_is_column_numbered(cur_header)
                    and _looks_like_header_row(cur_header)
                ):
                    return
                score = _txn_score(tables)
                # A real PDF table header (adaptive) must not be replaced by
                # passbook/bank-text remaps that split cells and invent columns.
                if (
                    engine_used in ("adaptive", "kotak", "column-anchored")
                    and name in ("passbook", "bank-text", "stacked-text")
                    and _looks_like_header_row(cur_header)
                    and not _header_is_column_numbered(cur_header)
                    and best_score >= 8
                    and score <= int(best_score * 1.15) + 2
                ):
                    return
                quality = _table_quality(tables)
                if cand_garbage or _header_is_column_numbered(cand_header):
                    quality = min(quality, -15_000)
                pref = _ENGINE_PREF.get(name, 0)
                better = False
                if score > best_score and not (cand_garbage and not cur_garbage):
                    # Structured adaptive/bank headers beat higher-score garbage always
                    if cand_garbage and best_quality > -5_000 and best_score >= 2:
                        better = False
                    elif (
                        engine_used in ("adaptive", "kotak", "column-anchored", "passbook")
                        and name in ("bank-text", "stacked-text")
                        and best_score >= 8
                        and quality < best_quality + 80
                        and score < int(best_score * 1.4) + 1
                    ):
                        # Stacked ICICI parser inflates txn count by splitting wrap
                        # lines and then guesses debit/credit — don't steal the grid.
                        better = False
                    else:
                        better = True
                elif score == best_score and score > 0:
                    if quality > best_quality:
                        better = True
                    elif quality == best_quality and pref > _ENGINE_PREF.get(engine_used, 0):
                        better = True
                elif best_score == 0 and not _tables_are_weak(tables):
                    better = True
                # Prefer non-garbage with slightly fewer rows over garbage with more
                if (
                    not better
                    and not cand_garbage
                    and cur_garbage
                    and score >= max(2, best_score // 2)
                ):
                    better = True
                # Prefer real statement header even if txn count is a bit lower
                if (
                    not better
                    and not cand_garbage
                    and not cur_garbage
                    and score >= 2
                    and quality > best_quality + 200
                    and score >= best_score * 0.6
                ):
                    better = True
                # Keep *something* rather than an empty Excel when this is the first hit
                if (
                    not better
                    and not page_tables
                    and name in (
                        "pymupdf",
                        "pymupdf-words",
                        "pymupdf-text",
                        "pdfplumber-text",
                        "ocr",
                    )
                ):
                    has_cells = any(
                        (c or "").strip()
                        for block in tables
                        for row in block
                        for c in row
                    )
                    if has_cells:
                        better = True
                if better:
                    page_tables = tables
                    engine_used = name
                    best_score = score
                    best_quality = quality
                    if label:
                        status_cb(label)

            profile = _detect_statement_profile(work_path)
            profile_label = _BANK_PROFILE_LABEL.get(profile, profile)
            status_cb(f"Detected: {profile_label}. Using that template only...")

            page_count = 0
            try:
                import pymupdf as _fitz_count

                _dcount = _fitz_count.open(work_path)
                page_count = _dcount.page_count
                _dcount.close()
            except Exception:
                page_count = 0

            # One bank → one extractor. Contesting parsers is what invented extra columns.
            # ICICI Detailed Statement has a drawn 10-col grid — use those cells, do not guess.
            ruled_cols = _peek_ruled_statement_ncols(work_path)
            if ruled_cols >= 8:
                status_cb(
                    f"PDF table grid ({ruled_cols} columns) — extracting cells as printed..."
                )
                _consider(
                    "pymupdf",
                    _extract_tables_pymupdf(work_path, status_cb=status_cb),
                )
                if best_score >= 8:
                    status_cb(
                        f"Locked grid extract ({best_score} transactions) — empty cells stay empty."
                    )

            if best_score < 8 and profile == "karnataka":
                status_cb("Karnataka Bank passbook template...")
                _consider(
                    "passbook",
                    _extract_passbook_statement(work_path, status_cb=status_cb),
                )
                if best_score < 8:
                    status_cb("Passbook weak — trying statement table...")
                    _consider(
                        "adaptive",
                        _extract_adaptive_statement(work_path, status_cb=status_cb),
                    )
            elif best_score < 8 and profile == "kotak":
                status_cb("Kotak statement template...")
                _consider("kotak", _extract_kotak_bank(work_path, status_cb=status_cb))
                if best_score < 8:
                    status_cb("Kotak template weak — trying header-learned table...")
                    _consider(
                        "adaptive",
                        _extract_adaptive_statement(work_path, status_cb=status_cb),
                    )
            elif best_score < 8:
                status_cb("Statement table (learn this PDF's header → all pages)...")
                _consider(
                    "adaptive",
                    _extract_adaptive_statement(work_path, status_cb=status_cb),
                )

            if engine_used != "none" and best_score > 0:
                status_cb(f"{profile_label}: {engine_used} ({best_score} transactions)")

            locked = best_score >= 8 and (
                engine_used == "pymupdf"
                or profile in _STATEMENT_LOCK_PROFILES
            )
            if locked:
                status_cb(
                    f"Locked {profile_label} extract — not mixing other parsers."
                )
            else:
                if best_score < 8 and page_count <= 30:
                    status_cb("Extracting tables (PyMuPDF)...")
                    _consider(
                        "pymupdf",
                        _extract_tables_pymupdf(work_path, status_cb=status_cb),
                    )
                elif best_score < 8:
                    status_cb("Skipping slow ruled-table scan on this long PDF...")

                keep_pdf_tables = (
                    engine_used in ("pymupdf", "adaptive")
                    and _usable_generic_table(page_tables)
                )
                if keep_pdf_tables:
                    status_cb("Keeping PDF tables as printed — not mixing other parsers.")
                elif profile not in _STATEMENT_LOCK_PROFILES:
                    status_cb("Unknown layout — comparing extra parsers...")
                    _consider(
                        "passbook",
                        _extract_passbook_statement(work_path, status_cb=status_cb),
                    )
                    _consider(
                        "bank-text",
                        _extract_bank_statement_pymupdf(work_path, status_cb=status_cb),
                    )
                    _consider(
                        "column-anchored",
                        _extract_column_anchored_bank(work_path, status_cb=status_cb),
                    )
                    _consider(
                        "kotak",
                        _extract_kotak_bank(work_path, status_cb=status_cb),
                    )
                elif best_score < 8:
                    # Known bank, but the locked template was weak — targeted fallbacks only.
                    status_cb("Template weak — trying column-anchored / Kotak...")
                    _consider(
                        "column-anchored",
                        _extract_column_anchored_bank(work_path, status_cb=status_cb),
                    )
                    _consider(
                        "kotak",
                        _extract_kotak_bank(work_path, status_cb=status_cb),
                    )

                if best_score > 0:
                    status_cb(f"Best engine: {engine_used} ({best_score} transactions)")

                # Stacked / word dumps only when still weak — never steal a good grid.
                if not keep_pdf_tables and (
                    _tables_are_weak(page_tables) or best_score == 0
                ):
                    status_cb("Trying stacked-text row rebuild...")
                    _consider(
                        "stacked-text",
                        _extract_stacked_statement(work_path, status_cb=status_cb),
                    )

                if not keep_pdf_tables and _tables_are_weak(page_tables):
                    status_cb("No strong tables — using word-position extract...")
                    _consider(
                        "pymupdf-words",
                        _extract_words_rows_pymupdf(work_path, status_cb=status_cb),
                    )

                if not keep_pdf_tables and _tables_are_weak(page_tables):
                    status_cb("Using fast text extract...")
                    _consider(
                        "pymupdf-text",
                        _extract_text_rows_pymupdf(work_path, status_cb=status_cb),
                    )

                if not keep_pdf_tables and _tables_are_weak(page_tables):
                    status_cb("Retrying stacked-text rebuild...")
                    _consider(
                        "stacked-text",
                        _extract_stacked_statement(work_path, status_cb=status_cb),
                    )

                if not keep_pdf_tables and _tables_are_weak(page_tables):
                    status_cb("Trying pdfplumber text mode...")
                    _consider(
                        "pdfplumber-text",
                        _extract_tables_pdfplumber_fast(work_path, status_cb=status_cb),
                    )

                if not keep_pdf_tables and _tables_are_weak(page_tables) and text_chars < 20:
                    status_cb("No text tables found — trying OCR...")
                    try:
                        _consider(
                            "ocr",
                            _extract_ocr_pymupdf(work_path, status_cb=status_cb),
                        )
                    except RuntimeError as exc:
                        status_cb(str(exc).split("\n")[0])

            if not page_tables:
                status_cb("No statement table — writing raw PDF text...")
                raw_text = _extract_text_rows_pymupdf(work_path, status_cb=status_cb)
                if raw_text:
                    page_tables = raw_text
                    engine_used = "raw-text"
                else:
                    raw_words = _extract_words_rows_pymupdf(work_path, status_cb=status_cb)
                    if raw_words:
                        page_tables = raw_words
                        engine_used = "raw-words"

        skip_combined = False
        if (
            engine_used == "pymupdf"
            and len(page_tables) >= 2
            and not _tables_share_statement_header(page_tables)
        ):
            status_cb("Writing each PDF table to its own Excel sheet...")
            n_sheets = 0
            for block in page_tables:
                exact = _normalize_table_exact(block)
                if not exact or not any(
                    (c or "").strip() for row in exact for c in row
                ):
                    continue
                n_sheets += 1
                _write_table_sheet(
                    wb,
                    _unique_sheet_name(used_names, f"Table {n_sheets}"),
                    exact,
                )
            if n_sheets:
                skip_combined = True
                sheet_written = f"{n_sheets} tables"
                raw_table_count = n_sheets

        if not skip_combined:
            status_cb(f"Merging {len(page_tables)} table block(s)...")
        raw_all_rows = _flatten_tables_to_rows(page_tables)
        # Adaptive/bank parsers already emit 1-row-per-txn — don't re-smash them
        skip_heavy_merge = engine_used in (
            "adaptive",
            "passbook",
            "kotak",
            "column-anchored",
            "bank-text",
            "stacked-text",
            "pymupdf",
        )

        for exact_rows in page_tables:
            raw_table_count += 1
            if skip_heavy_merge:
                cleaned = [
                    r
                    for r in exact_rows
                    if not _is_junk_or_legend_row(r)
                ]
            else:
                cleaned = _merge_multiline_transaction_rows(exact_rows)
            if not cleaned:
                continue

            if _looks_like_header_row(cleaned[0]):
                page_header, page_body = cleaned[0], cleaned[1:]
            else:
                page_header, page_body = [], cleaned

            if not combined_header and page_header:
                combined_header = list(page_header)
            elif not combined_header and page_body:
                # Prefer not inventing Column_* when a later pass may supply names;
                # only invent if we truly have no header at all.
                width = max(len(r) for r in page_body)
                combined_header = [f"Column_{i}" for i in range(1, width + 1)]

            # Never keep Column_* if this page brought a real header
            if (
                page_header
                and _looks_like_header_row(page_header)
                and not _header_is_column_numbered(page_header)
                and (
                    not combined_header
                    or _header_is_column_numbered(combined_header)
                )
            ):
                combined_header = list(page_header)

            for row in page_body:
                if page_header and _same_header(row, page_header):
                    continue
                if _looks_like_header_row(row):
                    continue
                width = len(combined_header) if combined_header else len(row)
                if len(row) < width:
                    row = row + [""] * (width - len(row))
                elif len(row) > width and combined_header:
                    overflow = [
                        (c or "").strip()
                        for c in row[width:]
                        if (c or "").strip()
                    ]
                    row = row[:width]
                    if overflow:
                        narr_i = -1
                        for i, h in enumerate(combined_header):
                            hl = (h or "").lower()
                            if any(
                                k in hl
                                for k in (
                                    "particular",
                                    "narrat",
                                    "remark",
                                    "desc",
                                    "detail",
                                )
                            ):
                                narr_i = i
                                break
                        extra = " ".join(overflow)
                        if narr_i >= 0:
                            row[narr_i] = ((row[narr_i] or "") + " " + extra).strip()
                        else:
                            row[-1] = ((row[-1] or "") + " " + extra).strip()
                combined_rows.append(row[:width] if len(row) >= width else row)

        # If bank-merge dropped everything but we still have exact raw rows, keep them
        if not combined_rows and raw_all_rows:
            status_cb("Using exact raw extracted rows (merge found no txn pattern)...")
            combined_rows = raw_all_rows
            if not combined_header:
                width = max(len(r) for r in combined_rows)
                combined_header = [f"Column_{i}" for i in range(1, width + 1)]

        status_cb("Writing Excel...")
        if skip_combined:
            pass
        elif combined_header or combined_rows:
            width = len(combined_header) if combined_header else (
                max((len(r) for r in combined_rows), default=1)
            )
            if not combined_header:
                combined_header = [f"Column_{i}" for i in range(1, width + 1)]
            # Don't rewrite US / raw extracts — cleanup is for Indian txn tables only
            if engine_used in (
                "adaptive",
                "passbook",
                "kotak",
                "column-anchored",
                "bank-text",
                "stacked-text",
            ):
                combined_header, combined_rows = _cleanup_statement_table(
                    combined_header, combined_rows
                )
            width = len(combined_header) if combined_header else 1
            final_table = [combined_header] + [
                (r + [""] * (width - len(r)))[:width] for r in combined_rows
            ]
            _write_table_sheet(wb, _unique_sheet_name(used_names, "Transactions"), final_table)
            sheet_written = "Transactions"
        else:
            tips = [
                ["Message"],
                ["No extractable text was found in this PDF."],
                ["Your PDF looks scanned/image-only."],
                ["Install Tesseract OCR, then restart this app and try again:"],
                [_tesseract_install_command()],
            ]
            _write_table_sheet(wb, _unique_sheet_name(used_names, "Note"), tips)
            sheet_written = "Note"

        out = _save_workbook_unlocked(wb, os.path.normpath(out))
    finally:
        if is_temp:
            try:
                os.remove(work_path)
            except OSError:
                pass
        root = getattr(tk, "_default_root", None)
        if root is not None:
            try:
                root.title(f"{APP_NAME} v{APP_VERSION}")
            except Exception:
                pass

    elapsed = time.perf_counter() - started
    if combined_rows:
        summary = (
            f"Excel file created:\n{out}\n"
            f"Engine: {engine_used}\n"
            f"Time: {elapsed:.1f} seconds\n"
            f"Raw tables seen: {raw_table_count}\n"
            f"Data rows: {len(combined_rows)}\n"
            f"Output: '{sheet_written}' sheet with exact extracted text."
        )
    else:
        if is_scanned:
            hint = (
                "This PDF is scanned (image only). Install Tesseract, restart the app, try again:\n"
                f"{_tesseract_install_command()}"
            )
        else:
            hint = (
                "This PDF has text, but no transaction table was detected.\n"
                "Try PDF to Word for a page copy, or install Tesseract if pages are photos:\n"
                f"{_tesseract_install_command()}"
            )
        summary = (
            f"Excel file created:\n{out}\n"
            f"Engine: {engine_used}\n"
            f"Time: {elapsed:.1f} seconds\n"
            f"Raw tables seen: {raw_table_count}\n"
            f"No data found.\n"
            f"{hint}"
        )
    return summary, output_folder


# --------------------------- GUI ---------------------------

class PDFStudioApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("1120x740")
        self.minsize(980, 640)
        self.configure(bg="#f6f7fb")

        self.last_output_folder = None
        self._configure_styles()
        self._build_ui()

    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#f6f7fb")
        style.configure("Header.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#ffffff", foreground="#111827", font=("Helvetica", 20, "bold"))
        style.configure("Subtitle.TLabel", background="#ffffff", foreground="#4b5563", font=("Helvetica", 11))
        style.configure("Card.TButton", font=("Helvetica", 11, "bold"), padding=14)
        style.configure("Footer.TLabel", background="#f6f7fb", foreground="#6b7280", font=("Helvetica", 9))

    def _build_ui(self):
        header = ttk.Frame(self, style="Header.TFrame")
        header.pack(fill="x")

        title_area = ttk.Frame(header, style="Header.TFrame")
        title_area.pack(fill="x", padx=26, pady=18)

        ttk.Label(title_area, text="SSA PDF Studio", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            title_area,
            text="Offline desktop PDF tools: merge, split, extract, convert, compress, watermark, protect and more.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        self.status_var = tk.StringVar(value="Ready. Select any tool below.")

        # Scrollable main content (mouse wheel + scrollbar)
        scroll_container = ttk.Frame(self)
        scroll_container.pack(fill="both", expand=True, padx=22, pady=18)

        self.canvas = tk.Canvas(scroll_container, bg="#f6f7fb", highlightthickness=0)
        self.v_scroll = ttk.Scrollbar(scroll_container, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set)

        self.v_scroll.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        main = ttk.Frame(self.canvas)
        self.main_window = self.canvas.create_window((0, 0), window=main, anchor="nw")

        main.bind("<Configure>", self._on_main_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # Mouse wheel / trackpad scroll (Windows, macOS, Linux)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_linux)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_linux)

        categories = [
            ("Organize PDF", [
                ("Merge PDF", "Combine PDFs in the order you set (Move up / Move down).", merge_pdfs),
                ("Split PDF", "Create separate files for each page range.", split_pdf),
                ("Extract Pages", "Create one PDF from selected pages/ranges.", extract_pages),
                ("Remove Pages", "Delete selected pages from a PDF.", remove_pages),
            ]),
            ("Optimize & Edit", [
                ("Compress PDF", "Optimize PDF size where possible.", compress_pdf),
                ("Rotate PDF", "Rotate all or selected pages.", rotate_pdf),
                ("Add Watermark", "Add text watermark to every page.", add_text_watermark),
                ("Add Page Numbers", "Add Page X of Y at bottom-right.", add_page_numbers),
            ]),
            ("Convert PDF", [
                ("PDF to Images", "Export every page as PNG image.", pdf_to_images),
                ("Images to PDF", "Convert JPG/PNG images into one PDF.", images_to_pdf),
                ("PDF to Word", "Convert PDF to editable Word (text, headings, tables, images).", pdf_to_word),
                ("PDF to Excel", "Extract exact PDF data to Excel (tables, text layout, OCR for scans).", pdf_to_excel),
            ]),
            ("PDF Security", [
                ("Protect PDF", "Encrypt PDF with a password.", protect_pdf),
                ("Unlock PDF", "Remove password if you know it.", unlock_pdf),
            ]),
        ]

        for col, (cat_name, tools) in enumerate(categories):
            frame = ttk.LabelFrame(main, text=cat_name, padding=14)
            frame.grid(row=0, column=col, padx=8, pady=8, sticky="nsew")
            main.columnconfigure(col, weight=1)
            for r, (name, desc, func) in enumerate(tools):
                self._add_tool_card(frame, name, desc, func, r)

        main.rowconfigure(0, weight=1)

        footer = ttk.Frame(self)
        footer.pack(fill="x", padx=24, pady=(0, 16))
        ttk.Label(footer, textvariable=self.status_var, style="Footer.TLabel").pack(side="left")
        ttk.Button(footer, text="Open Last Output Folder", command=self._open_last_folder).pack(side="right")

    def _on_main_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self.main_window, width=event.width)

    def _on_mousewheel(self, event):
        # Windows uses multiples of 120; macOS often sends smaller deltas
        if not event.delta:
            return
        if sys.platform == "darwin":
            self.canvas.yview_scroll(int(-1 * event.delta), "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event):
        # Linux: Button-4 = up, Button-5 = down
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")

    def _add_tool_card(self, parent, name: str, desc: str, func, row: int):
        card = tk.Frame(parent, bg="#ffffff", highlightbackground="#e5e7eb", highlightthickness=1)
        card.grid(row=row, column=0, sticky="ew", pady=8)
        parent.columnconfigure(0, weight=1)

        label = tk.Label(card, text=name, bg="#ffffff", fg="#111827", font=("Helvetica", 12, "bold"), anchor="w")
        label.pack(fill="x", padx=14, pady=(12, 0))
        desc_label = tk.Label(card, text=desc, bg="#ffffff", fg="#6b7280", font=("Helvetica", 9), anchor="w", wraplength=210, justify="left")
        desc_label.pack(fill="x", padx=14, pady=(3, 10))
        btn = ttk.Button(card, text="Open Tool", style="Card.TButton", command=lambda: self._run_tool(func))
        btn.pack(fill="x", padx=14, pady=(0, 12))

    def _run_tool(self, func):
        try:
            self.status_var.set("Working...")
            self.update_idletasks()
            message, folder = func()
            self.last_output_folder = folder
            self.status_var.set("Completed successfully.")
            messagebox.showinfo(APP_NAME, message)
        except Exception as exc:
            error_msg = str(exc)
            cancel_msgs = {
                "No PDF file selected.",
                "No PDF files selected.",
                "No output folder selected.",
                "No input provided.",
            }
            if error_msg in cancel_msgs:
                self.status_var.set("Cancelled.")
                return
            self.status_var.set("Error occurred.")
            if not error_msg:
                error_msg = "Unknown error occurred."
            print(traceback.format_exc())
            messagebox.showerror(APP_NAME, error_msg)

    def _open_last_folder(self):
        if self.last_output_folder and os.path.isdir(self.last_output_folder):
            open_folder(self.last_output_folder)
        else:
            messagebox.showinfo(APP_NAME, "No output folder available yet.")


if __name__ == "__main__":
    app = PDFStudioApp()
    app.mainloop()
