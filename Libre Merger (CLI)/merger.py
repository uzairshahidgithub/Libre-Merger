#!/usr/bin/env python3
"""
Document Merger – Merge PPTX, DOCX, and PDF files into one PDF.
Supports sorting by name or date, ascending/descending.
Uses LibreOffice for conversion of non-PDF files.
"""

import os
import sys
import glob
import tempfile
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

# Try importing required third‑party libraries
try:
    import fitz  # PyMuPDF
except ImportError:
    print("Error: PyMuPDF is not installed. Please install it using:")
    print("    pip install PyMuPDF")
    sys.exit(1)

# Try importing tkinter for the file save pdialog
try:
    import tkinter as tk
    from tkinter import filedialog
except ImportError:
    print("Warning: tkinter not found. The save dialog will fall back to console input.")
    tk = None

# ----------------------------------------------------------------------
# Helper functions
# ----------------------------------------------------------------------

def check_libreoffice():
    """Check if LibreOffice is installed and accessible."""
    try:
        subprocess.run(["soffice", "--version"], capture_output=True, check=True)
        return True
    except (subprocess.SubprocessError, FileNotFoundError):
        return False

def convert_to_pdf(input_path, output_dir):
    """
    Convert a file to PDF using LibreOffice.
    Returns the path to the generated PDF, or None on failure.
    """
    input_path = Path(input_path).resolve()
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    # Run soffice conversion
    cmd = [
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(input_path)
    ]
    try:
        subprocess.run(cmd, check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as e:
        print(f"Error converting {input_path.name}: {e.stderr}")
        return None

    # The output PDF is named after the input file, but with .pdf extension
    pdf_name = input_path.stem + ".pdf"
    pdf_path = output_dir / pdf_name
    if pdf_path.exists():
        return str(pdf_path)
    else:
        print(f"Conversion of {input_path.name} failed: output PDF not found.")
        return None

def get_file_list(folder, ext_choice):
    """
    Return a list of file paths matching the chosen extensions.
    ext_choice: 1 -> pptx, 2 -> pdf, 3 -> docx, 4 -> mix (all)
    """
    extensions = []
    if ext_choice == 1:
        extensions = ["pptx"]
    elif ext_choice == 2:
        extensions = ["pdf"]
    elif ext_choice == 3:
        extensions = ["docx"]
    elif ext_choice == 4:
        extensions = ["pptx", "pdf", "docx"]
    else:
        print("Invalid extension choice. Exiting.")
        sys.exit(1)

    files = []
    for ext in extensions:
        pattern = os.path.join(folder, f"*.{ext}")
        files.extend(glob.glob(pattern))
    # Remove duplicates (if any, though unlikely)
    files = list(set(files))
    return files

def sort_files(files, sort_choice):
    """
    Sort files according to the user's choice.
    sort_choice: 1 -> name ASC, 2 -> name DESC, 3 -> date ASC, 4 -> date DESC
    """
    if sort_choice == 1:
        files.sort(key=lambda x: os.path.basename(x).lower())
    elif sort_choice == 2:
        files.sort(key=lambda x: os.path.basename(x).lower(), reverse=True)
    elif sort_choice == 3:
        files.sort(key=lambda x: os.path.getmtime(x))
    elif sort_choice == 4:
        files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    else:
        print("Invalid sort choice. Exiting.")
        sys.exit(1)
    return files

def display_table(files):
    """Print a simple ASCII table with numbered files."""
    print("\n" + "=" * 70)
    print(f"{'No.':<5} {'Filename':<60}")
    print("=" * 70)
    for idx, path in enumerate(files, start=1):
        filename = os.path.basename(path)
        # Truncate long filenames
        if len(filename) > 60:
            filename = filename[:57] + "..."
        print(f"{idx:<5} {filename:<60}")
    print("=" * 70)

def save_dialog(default_filename="merged_document.pdf"):
    """Open a file save dialog using tkinter, fallback to console input."""
    if tk:
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=default_filename
        )
        root.destroy()
        return file_path if file_path else None
    else:
        # Fallback to console input
        print("\nNo GUI save dialog available. Please enter output file path.")
        return input("Output file path: ").strip()

def merge_pdfs(pdf_paths, output_path):
    """Merge a list of PDF files into one using PyMuPDF."""
    try:
        doc = fitz.open()  # empty document
        for path in pdf_paths:
            if not os.path.exists(path):
                print(f"Warning: {path} not found, skipping.")
                continue
            with fitz.open(path) as src:
                doc.insert_pdf(src)
        doc.save(output_path)
        doc.close()
        return True
    except Exception as e:
        print(f"Error during PDF merging: {e}")
        return False

def main():
    print("=" * 70)
    print("Document Merger – Combine PPTX, DOCX, PDF into one PDF")
    print("=" * 70)

    # Step 1: Folder path
    folder = input("Enter folder path: ").strip()
    if not os.path.isdir(folder):
        print("Error: Folder does not exist.")
        sys.exit(1)

    # Step 2: File type selection
    print("\nSelect file types to include:")
    print("1. pptx")
    print("2. pdf")
    print("3. docx")
    print("4. Mix (all three)")
    ext_choice = int(input("Enter choice (1-4): ").strip())

    # Step 3: Sorting selection
    print("\nSelect sorting order:")
    print("1. By Name (Ascending)")
    print("2. By Name (Descending)")
    print("3. By Date (Ascending)")
    print("4. By Date (Descending)")
    sort_choice = int(input("Enter choice (1-4): ").strip())

    # Step 4: Confirm settings
    print("\n" + "-" * 70)
    print(f"Folder: {folder}")
    ext_names = {1: "pptx", 2: "pdf", 3: "docx", 4: "pptx, pdf, docx"}
    print(f"Included types: {ext_names[ext_choice]}")
    sort_names = {1: "Name (A-Z)", 2: "Name (Z-A)", 3: "Date (Oldest first)", 4: "Date (Newest first)"}
    print(f"Sorting: {sort_names[sort_choice]}")
    print(f"Output format: PDF (default)")
    print("-" * 70)

    # Get file list
    files = get_file_list(folder, ext_choice)
    if not files:
        print("No matching files found. Exiting.")
        sys.exit(1)

    # Sort files
    files = sort_files(files, sort_choice)

    # Display table
    display_table(files)

    # Wait for confirmation
    input("\nPress Enter to start merging...")

    # Check for LibreOffice
    if not check_libreoffice():
        print("\nERROR: LibreOffice is not installed or not in PATH.")
        print("Please install LibreOffice (https://www.libreoffice.org/) and make sure 'soffice' is accessible.")
        print("Exiting.")
        sys.exit(1)

    # Create temporary directory for converted PDFs
    with tempfile.TemporaryDirectory() as tmpdir:
        print("\nConverting non-PDF files to PDF...")
        pdf_paths = []
        total = len(files)
        for idx, src_path in enumerate(files, 1):
            ext = os.path.splitext(src_path)[1].lower()
            print(f"[{idx}/{total}] Processing: {os.path.basename(src_path)}")
            if ext == ".pdf":
                pdf_paths.append(src_path)
            else:
                pdf_path = convert_to_pdf(src_path, tmpdir)
                if pdf_path:
                    pdf_paths.append(pdf_path)
                else:
                    print(f"  Failed to convert {src_path}. Skipping.")

        if not pdf_paths:
            print("No files could be converted/merged. Exiting.")
            sys.exit(1)

        # Merge all PDFs
        print("\nMerging PDFs...")
        # Show a simple progress indicator
        sys.stdout.write("Progress: ")
        sys.stdout.flush()
        # Use a single file for final output, but we need a temporary location
        merged_temp = os.path.join(tmpdir, "merged_temp.pdf")
        if merge_pdfs(pdf_paths, merged_temp):
            sys.stdout.write(" done.\n")
            # Ask where to save
            print("\nPlease select a location to save the merged PDF.")
            output_path = save_dialog("merged_document.pdf")
            if output_path:
                # Copy the merged file to the chosen location
                shutil.copy2(merged_temp, output_path)
                print(f"\nSuccess! Merged PDF saved to:\n{output_path}")
            else:
                print("\nSave cancelled. The merged PDF was not saved.")
        else:
            print(" failed.")
            sys.exit(1)

    print("\nAll done. Goodbye!")

if __name__ == "__main__":
    main()