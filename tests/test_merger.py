import os
import sys
import tempfile
from pathlib import Path
import pytest
from PIL import Image
import pymupdf as fitz

# Ensure root directory is on sys.path
sys.path.insert(0, str(Path(__file__).parent.parent))

from merger import (
    parse_page_ranges,
    is_image,
    is_libreoffice_supported,
    check_libreoffice,
    get_libreoffice_cmd,
    convert_image_to_pdf,
    convert_lo_to_pdf,
    apply_watermark,
    create_file_icon
)


class TestHelpers:
    def test_parse_page_ranges(self):
        # Empty or "all"
        assert parse_page_ranges("all", 10) is None
        assert parse_page_ranges("ALL", 10) is None
        assert parse_page_ranges("", 10) is None
        assert parse_page_ranges(None, 10) is None

        # Single page (1-indexed -> 0-indexed)
        assert parse_page_ranges("1", 5) == [0]
        assert parse_page_ranges("3", 5) == [2]

        # Ranges
        assert parse_page_ranges("1-3", 5) == [0, 1, 2]
        assert parse_page_ranges("2-4", 10) == [1, 2, 3]

        # Mixed comma-separated list
        assert parse_page_ranges("1, 3-4, 6", 10) == [0, 2, 3, 5]

        # Out-of-bounds handling
        assert parse_page_ranges("1-20", 5) == [0, 1, 2, 3, 4]
        assert parse_page_ranges("10", 5) is None

        # Fault-tolerant with invalid entries
        assert parse_page_ranges("abc, 1-2, xyz", 5) == [0, 1]

    def test_is_image(self):
        assert is_image("jpg") is True
        assert is_image("jpeg") is True
        assert is_image("png") is True
        assert is_image("bmp") is True
        assert is_image("tiff") is True
        assert is_image("pdf") is False
        assert is_image("docx") is False

    def test_is_libreoffice_supported(self):
        assert is_libreoffice_supported("docx") is True
        assert is_libreoffice_supported("doc") is True
        assert is_libreoffice_supported("pptx") is True
        assert is_libreoffice_supported("xlsx") is True
        assert is_libreoffice_supported("odt") is True
        assert is_libreoffice_supported("txt") is True
        assert is_libreoffice_supported("html") is True
        assert is_libreoffice_supported("pdf") is False
        assert is_libreoffice_supported("bin") is False

    def test_libreoffice_availability(self):
        cmd = get_libreoffice_cmd()
        assert cmd is not None, "LibreOffice executable should be found on this system"
        assert check_libreoffice() is True, "check_libreoffice() should return True"

    def test_create_file_icon(self):
        icon = create_file_icon("pdf")
        assert icon is not None
        icon_docx = create_file_icon("docx")
        assert icon_docx is not None


class TestDocumentProcessing:
    def test_convert_image_to_pdf(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            img_path = Path(tmpdir) / "test_sample.png"
            img = Image.new("RGB", (200, 200), color=(0, 128, 255))
            img.save(str(img_path))

            pdf_out = convert_image_to_pdf(str(img_path), tmpdir)
            assert pdf_out is not None
            assert os.path.exists(pdf_out)

            doc = fitz.open(pdf_out)
            assert doc.page_count == 1
            doc.close()

    def test_convert_lo_to_pdf(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            txt_path = Path(tmpdir) / "sample_doc.txt"
            txt_path.write_text("Libre Merger Automated Test Document\nLine 2 content.")

            pdf_out = convert_lo_to_pdf(str(txt_path), tmpdir)
            assert pdf_out is not None, "LibreOffice headless conversion failed"
            assert os.path.exists(pdf_out)

            doc = fitz.open(pdf_out)
            assert doc.page_count >= 1
            text_extracted = doc[0].get_text()
            assert "Libre Merger" in text_extracted
            doc.close()

    def test_merge_and_watermark(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create two test PDFs
            pdf1_path = Path(tmpdir) / "doc1.pdf"
            pdf2_path = Path(tmpdir) / "doc2.pdf"

            doc1 = fitz.open()
            p1 = doc1.new_page(width=400, height=600)
            p1.insert_text((50, 100), "First PDF Document", fontsize=20)
            doc1.save(str(pdf1_path))
            doc1.close()

            doc2 = fitz.open()
            p2 = doc2.new_page(width=400, height=600)
            p2.insert_text((50, 100), "Second PDF Document Page 1", fontsize=20)
            p3 = doc2.new_page(width=400, height=600)
            p3.insert_text((50, 100), "Second PDF Document Page 2", fontsize=20)
            doc2.save(str(pdf2_path))
            doc2.close()

            # Merge
            master_doc = fitz.open()
            src1 = fitz.open(str(pdf1_path))
            master_doc.insert_pdf(src1)
            src1.close()

            src2 = fitz.open(str(pdf2_path))
            # Slice only page 2 of doc2 (index 1)
            master_doc.insert_pdf(src2, from_page=1, to_page=1)
            src2.close()

            assert master_doc.page_count == 2

            # Watermark
            apply_watermark(master_doc, "CONFIDENTIAL")

            # Bookmarks / TOC
            toc = [
                [1, "Document 1", 1],
                [1, "Document 2 (Selected)", 2]
            ]
            master_doc.set_toc(toc)

            merged_output = Path(tmpdir) / "final_merged.pdf"
            master_doc.save(str(merged_output))
            master_doc.close()

            # Verify saved output
            saved_doc = fitz.open(str(merged_output))
            assert saved_doc.page_count == 2
            saved_toc = saved_doc.get_toc()
            assert len(saved_toc) == 2
            assert saved_toc[0][1] == "Document 1"
            assert saved_toc[1][1] == "Document 2 (Selected)"
            saved_doc.close()

    def test_pdf_encryption(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            test_pdf = Path(tmpdir) / "secret.pdf"
            doc = fitz.open()
            page = doc.new_page(width=300, height=300)
            page.insert_text((30, 50), "Classified Information", fontsize=14)

            # Save with encryption
            doc.save(
                str(test_pdf),
                encryption=fitz.PDF_ENCRYPT_AES_256,
                owner_pw="SuperSecret123",
                user_pw="SuperSecret123"
            )
            doc.close()

            # Verify that document requires password
            encrypted_doc = fitz.open(str(test_pdf))
            assert encrypted_doc.is_encrypted is True
            assert encrypted_doc.authenticate("WrongPassword") == 0
            assert encrypted_doc.authenticate("SuperSecret123") > 0
            assert "Classified" in encrypted_doc[0].get_text()
            encrypted_doc.close()

    def test_gui_initialization_and_queue(self):
        from merger import LibreMergerApp
        app = LibreMergerApp()
        try:
            app.update_idletasks()
            assert app.title() == "Libre Merger"
            assert len(app.files_to_merge) == 0
            # Test queueing files
            app.queue_files(["README.md", "Libre Merger.png"])
            assert len(app.files_to_merge) >= 1
            app.clear_files()
            assert len(app.files_to_merge) == 0
        finally:
            app.destroy()
