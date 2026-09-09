"""Regression tests for attachment conversion, EXIF handling, and filename sanitization.

Verifies:
1. Palette images with transparency (PNG/GIF) composite onto a pure white background instead
   of converting to black (0, 0, 0) or raising UserWarnings.
2. Images with EXIF orientation (0x0112) are transposed so pages in generated PDFs are
   not sideways/upside down.
3. sanitize_filename prevents directory traversal ("..", ".") and Windows reserved device names.
4. merge_pdf_with_body supports modern pypdf (PdfWriter) as well as legacy PyPDF2 (PdfMerger).
"""

import io
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock
from PIL import Image

# Ensure project root is in sys.path
sys.path.insert(0, str(Path(__file__).parent.parent))

# Headless mock for PyQt/PySide if needed
if "PySide6" not in sys.modules:
    _qt_mock = MagicMock()
    for mod in ["PySide6", "PySide6.QtWidgets", "PySide6.QtCore", "PySide6.QtGui"]:
        sys.modules[mod] = _qt_mock

import UniversalInvoiceMail as uim


class TestAttachmentConversionRegressions(unittest.TestCase):
    def test_palette_image_with_transparency_composites_on_white(self):
        """Palette PNG with transparent background must be composited on white (255, 255, 255)."""
        # Create a 32x32 image: transparent background, red block in center
        from PIL import ImageDraw
        im_rgba = Image.new("RGBA", (32, 32), (0, 0, 0, 0))
        draw = ImageDraw.Draw(im_rgba)
        draw.rectangle([10, 10, 22, 22], fill=(255, 0, 0, 255))
        im_p = im_rgba.convert("P")
        buf = io.BytesIO()
        im_p.save(buf, format="PNG")
        png_data = buf.getvalue()

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "transparent.pdf"
            success, msg = uim.convert_attachment_to_pdf(png_data, "logo.png", pdf_path)
            self.assertTrue(success, f"Conversion failed: {msg}")
            self.assertTrue(pdf_path.exists())

            # Verify through pypdf that the pixel in the PDF image is white
            from pypdf import PdfReader

            reader = PdfReader(str(pdf_path))
            self.assertEqual(len(reader.pages), 1)
            extracted_img = reader.pages[0].images[0].image
            px_bg = extracted_img.getpixel((0, 0))
            self.assertGreater(
                min(px_bg),
                220,
                f"Transparent pixel should be white, got {px_bg}",
            )
            px_fg = extracted_img.getpixel((16, 16))
            self.assertGreater(px_fg[0], 180, f"Foreground red channel should be high, got {px_fg}")
            self.assertLess(px_fg[1], 80, f"Foreground green channel should be low, got {px_fg}")

    def test_exif_orientation_is_transposed(self):
        """Images with EXIF orientation tags must be transposed to correct orientation."""
        im = Image.new("RGB", (100, 200), (255, 255, 255))
        exif = im.getexif()
        exif[0x0112] = 6  # 90 degrees clockwise
        buf = io.BytesIO()
        im.save(buf, format="JPEG", exif=exif)
        jpeg_data = buf.getvalue()

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "exif.pdf"
            success, msg = uim.convert_attachment_to_pdf(jpeg_data, "receipt.jpg", pdf_path)
            self.assertTrue(success, f"Conversion failed: {msg}")

            from pypdf import PdfReader

            reader = PdfReader(str(pdf_path))
            extracted_img = reader.pages[0].images[0].image
            # Transposed image should have swapped width and height
            self.assertEqual(extracted_img.size, (200, 100))

    def test_sanitize_filename_prevents_directory_traversal(self):
        """sanitize_filename must not allow relative traversal paths like '..' or '.'."""
        self.assertEqual(uim.sanitize_filename(".."), "unnamed")
        self.assertEqual(uim.sanitize_filename("."), "unnamed")
        self.assertEqual(uim.sanitize_filename("..."), "unnamed")
        self.assertEqual(uim.sanitize_filename("../"), "unnamed")
        self.assertEqual(uim.sanitize_filename("..\\"), "unnamed")

    def test_sanitize_filename_windows_reserved_and_trailing_dots(self):
        """sanitize_filename must strip trailing dots/spaces and prefix Windows reserved names."""
        self.assertEqual(uim.sanitize_filename("invoice.pdf."), "invoice.pdf")
        self.assertEqual(uim.sanitize_filename("invoice.pdf "), "invoice.pdf")
        self.assertEqual(uim.sanitize_filename("CON"), "_CON")
        self.assertEqual(uim.sanitize_filename("aux.pdf"), "_aux.pdf")
        self.assertEqual(uim.sanitize_filename("nul.txt"), "_nul.txt")

    def test_merge_pdf_with_body_uses_pypdf_writer(self):
        """merge_pdf_with_body must support pypdf.PdfWriter when PyPDF2 is absent."""
        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "attachment.pdf"
            output_path = Path(tmpdir) / "merged.pdf"

            # Create a simple valid PDF for the attachment using pypdf
            from pypdf import PdfWriter

            writer = PdfWriter()
            writer.add_blank_page(width=200, height=200)
            with open(pdf_path, "wb") as f:
                writer.write(f)
            writer.close()

            # Mock PyPDF2 as not installed to test the pypdf branch
            with patch("UniversalInvoiceMail.XHTML2PDF_AVAILABLE", True), \
                 patch.dict("sys.modules", {"PyPDF2": None}), \
                 patch("UniversalInvoiceMail.html_to_pdf", side_effect=lambda html, path, meta, mode: path.write_bytes(pdf_path.read_bytes()) or True):
                result = uim.merge_pdf_with_body(pdf_path, "<p>Mail content</p>", {}, output_path)

            self.assertTrue(result)
            self.assertTrue(output_path.exists())
            from pypdf import PdfReader

            reader = PdfReader(str(output_path))
            # Merged PDF must contain 2 pages: attachment + body
            self.assertEqual(len(reader.pages), 2)


if __name__ == "__main__":
    unittest.main()
