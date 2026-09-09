#!/usr/bin/env python3
"""Unit-Tests fuer UniversalInvoiceMail Helper-Funktionen."""

import sys
import os
import io
import tempfile
from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

# PyQt6 und optionale Abhaengigkeiten mocken fuer headless Test
for mod in ['PyQt6', 'PyQt6.QtWidgets', 'PyQt6.QtCore', 'PyQt6.QtGui',
            'xhtml2pdf', 'xhtml2pdf.pisa', 'pytesseract', 'pypdfium2',
            'pypdf', 'PIL', 'PIL.Image', 'selenium', 'selenium.webdriver',
            'selenium.webdriver.edge.options', 'selenium.webdriver.edge.service',
            'selenium.webdriver.chrome.options', 'selenium.webdriver.chrome.service',
            'webdriver_manager', 'webdriver_manager.microsoft', 'webdriver_manager.chrome',
            'googleapiclient', 'googleapiclient.discovery',
            'google_auth_oauthlib', 'google_auth_oauthlib.flow',
            'google.auth', 'google.auth.transport', 'google.auth.transport.requests',
            'google.oauth2', 'google.oauth2.credentials',
            'google.auth.exceptions', 'keyring',
            'reportlab', 'reportlab.pdfgen', 'reportlab.lib',
            'reportlab.lib.pagesizes', 'reportlab.lib.units', 'reportlab.lib.utils',
            'docx2pdf', 'win32com', 'win32com.client', 'pythoncom']:
    if mod not in sys.modules:
        if mod.startswith(("PyQt6", "selenium", "webdriver")):
            sys.modules[mod] = MagicMock()
        else:
            try:
                __import__(mod)
            except ImportError:
                sys.modules[mod] = MagicMock()

# Spezielle Mocks fuer Konstanten die beim Import gebraucht werden
sys.modules['PyQt6.QtCore'].Qt = MagicMock()
sys.modules['PyQt6.QtCore'].QThread = MagicMock()
sys.modules['PyQt6.QtCore'].Signal = MagicMock(return_value=MagicMock())
sys.modules['PyQt6.QtCore'].QUrl = MagicMock()
sys.modules['PyQt6.QtCore'].QSize = MagicMock()

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import unittest
from UniversalInvoiceMail import (
    sanitize_filename,
    format_imap_date,
    safe_b64decode,
    calculate_hash,
    calculate_file_hash,
    convert_attachment_to_pdf,
    decode_mail_header,
    get_attachment_conversion_type,
    MailAccount,
    InvoiceProfile,
    InvoiceWorker,
    OCRProcessor,
)
import UniversalInvoiceMail
UniversalInvoiceMail.XHTML2PDF_AVAILABLE = True
UniversalInvoiceMail.PILLOW_AVAILABLE = True
UniversalInvoiceMail.PYTHON_DOCX_AVAILABLE = True
UniversalInvoiceMail.OPENPYXL_AVAILABLE = True
UniversalInvoiceMail.WIN32COM_AVAILABLE = True
UniversalInvoiceMail.OCR_AVAILABLE = True
UniversalInvoiceMail.pythoncom = sys.modules['pythoncom']
UniversalInvoiceMail.win32com = sys.modules['win32com']
UniversalInvoiceMail.pisa = sys.modules['xhtml2pdf.pisa']


class TestSanitizeFilename(unittest.TestCase):

    def test_basic(self):
        self.assertEqual(sanitize_filename("invoice.pdf"), "invoice.pdf")

    def test_special_chars(self):
        result = sanitize_filename('Rechnung<>:"/\\|?*.pdf')
        self.assertNotIn("<", result)
        self.assertNotIn(">", result)
        self.assertNotIn(":", result)
        self.assertTrue(result.endswith(".pdf"))

    def test_only_invalid_chars_falls_back_to_unnamed(self):
        self.assertEqual(sanitize_filename('<>:"/\\|?*'), "unnamed")

    def test_spaces(self):
        result = sanitize_filename("my   file   name.pdf")
        self.assertNotIn("  ", result)

    def test_empty(self):
        self.assertEqual(sanitize_filename(""), "unnamed")
        self.assertEqual(sanitize_filename(None), "unnamed")

    def test_max_length(self):
        long_name = "x" * 200 + ".pdf"
        result = sanitize_filename(long_name)
        self.assertLessEqual(len(result), 120)


class TestFormatImapDate(unittest.TestCase):

    def test_basic(self):
        dt = datetime(2026, 3, 8)
        self.assertEqual(format_imap_date(dt), "08-Mar-2026")

    def test_january(self):
        dt = datetime(2026, 1, 1)
        self.assertEqual(format_imap_date(dt), "01-Jan-2026")

    def test_december(self):
        dt = datetime(2025, 12, 31)
        self.assertEqual(format_imap_date(dt), "31-Dec-2025")

    def test_english_month_names(self):
        """Stellt sicher dass Monatsnamen IMMER englisch sind (nicht locale-abhaengig)."""
        for month in range(1, 13):
            result = format_imap_date(datetime(2026, month, 15))
            month_part = result.split("-")[1]
            self.assertIn(month_part, ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                                        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])


class TestSafeB64Decode(unittest.TestCase):

    def test_empty(self):
        self.assertEqual(safe_b64decode(""), b"")
        self.assertEqual(safe_b64decode(None), b"")

    def test_valid(self):
        import base64
        original = b"Hello World"
        encoded = base64.urlsafe_b64encode(original).decode()
        self.assertEqual(safe_b64decode(encoded), original)

    def test_missing_padding(self):
        import base64
        original = b"Test Data"
        encoded = base64.urlsafe_b64encode(original).decode().rstrip("=")
        self.assertEqual(safe_b64decode(encoded), original)

    def test_whitespace_in_base64(self):
        import base64
        original = b"Test"
        encoded = base64.urlsafe_b64encode(original).decode()
        # Simuliere Newlines wie in E-Mails
        encoded_with_newlines = encoded[:2] + "\n" + encoded[2:]
        self.assertEqual(safe_b64decode(encoded_with_newlines), original)


class TestCalculateHash(unittest.TestCase):

    def test_basic(self):
        result = calculate_hash(b"test")
        self.assertEqual(len(result), 64)

    def test_deterministic(self):
        self.assertEqual(calculate_hash(b"abc"), calculate_hash(b"abc"))

    def test_different_input(self):
        self.assertNotEqual(calculate_hash(b"abc"), calculate_hash(b"def"))


class TestCalculateFileHash(unittest.TestCase):

    def test_existing_file(self):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as f:
            f.write(b"test content for hashing")
            path = f.name
        try:
            result = calculate_file_hash(path)
            self.assertIsNotNone(result)
            self.assertEqual(len(result), 64)
        finally:
            os.unlink(path)

    def test_nonexistent(self):
        self.assertIsNone(calculate_file_hash("/nonexistent/path.txt"))

    def test_deterministic(self):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as f:
            f.write(b"deterministic content")
            path = f.name
        try:
            self.assertEqual(calculate_file_hash(path), calculate_file_hash(path))
        finally:
            os.unlink(path)


class TestDecodeMailHeader(unittest.TestCase):

    def test_plain_text(self):
        self.assertEqual(decode_mail_header("Hello"), "Hello")

    def test_empty(self):
        self.assertEqual(decode_mail_header(""), "")
        self.assertEqual(decode_mail_header(None), "")


class TestMailAccount(unittest.TestCase):

    def test_roundtrip(self):
        acc = MailAccount(id="test1", name="Test", provider="IMAP",
                          host="imap.test.com", port=993, username="user@test.com")
        d = acc.to_dict()
        acc2 = MailAccount.from_dict(d)
        self.assertEqual(acc.id, acc2.id)
        self.assertEqual(acc.host, acc2.host)

    def test_defaults(self):
        acc = MailAccount(id="x", name="X", provider="IMAP")
        self.assertEqual(acc.port, 993)
        self.assertFalse(acc.use_gmail_api)


class TestInvoiceProfile(unittest.TestCase):

    def test_roundtrip(self):
        p = InvoiceProfile(id="p1", name="Amazon", account_id="a1",
                           sender_filter="amazon", subject_filter="Rechnung")
        d = p.to_dict()
        p2 = InvoiceProfile.from_dict(d)
        self.assertEqual(p.name, p2.name)
        self.assertEqual(p.sender_filter, p2.sender_filter)


class TestAttachmentConversion(unittest.TestCase):

    def test_attachment_type_detection(self):
        self.assertEqual(get_attachment_conversion_type("invoice.pdf"), "pdf")
        self.assertEqual(get_attachment_conversion_type("scan.PNG"), "image")
        self.assertEqual(get_attachment_conversion_type("report.docx"), "docx")
        self.assertEqual(get_attachment_conversion_type("table.xlsx"), "xlsx")
        self.assertEqual(get_attachment_conversion_type("legacy.xls"), "legacy_office")
        self.assertIsNone(get_attachment_conversion_type("archive.zip"))

    def test_convert_xlsx_attachment_uses_html_pipeline(self):
        try:
            from openpyxl import Workbook
        except ImportError:
            self.skipTest("openpyxl nicht installiert")

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Rechnungen"
        sheet.append(["Bestellung", "Betrag"])
        sheet.append(["A-100", 19.99])

        buffer = io.BytesIO()
        workbook.save(buffer)
        workbook.close()

        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "attachment.pdf"

            def fake_html_to_pdf(html_content, output_pdf, mail_meta, mode="fast"):
                self.assertIn("Bestellung", html_content)
                self.assertIn("A-100", html_content)
                output_pdf.write_bytes(b"%PDF-1.4 xlsx")
                return True

            with patch("UniversalInvoiceMail.html_to_pdf", side_effect=fake_html_to_pdf) as mocked_pdf:
                success, message = convert_attachment_to_pdf(
                    buffer.getvalue(), "rechnung.xlsx", output_path
                )

            self.assertTrue(success)
            self.assertIn("XLSX", message)
            self.assertTrue(output_path.exists())
            mocked_pdf.assert_called_once()

    def test_legacy_attachment_uses_com_converter(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "attachment.pdf"

            def fake_convert(source_path, temp_output):
                self.assertEqual(source_path.suffix, ".xls")
                temp_output.write_bytes(b"%PDF-1.4 legacy-com")
                return "Excel COM"

            with patch("UniversalInvoiceMail.convert_legacy_office_via_com", side_effect=fake_convert) as mocked_com:
                with patch("UniversalInvoiceMail.convert_legacy_office_via_libreoffice") as mocked_libreoffice:
                    success, message = convert_attachment_to_pdf(
                        b"legacy-data", "altbestand.xls", output_path
                    )

            self.assertTrue(success)
            self.assertIn("Excel COM", message)
            self.assertTrue(output_path.exists())
            mocked_com.assert_called_once()
            mocked_libreoffice.assert_not_called()

    def test_legacy_attachment_falls_back_to_libreoffice(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "attachment.pdf"

            def fake_libreoffice(source_path, temp_output):
                self.assertEqual(source_path.suffix, ".doc")
                temp_output.write_bytes(b"%PDF-1.4 legacy-libreoffice")
                return "LibreOffice"

            with patch(
                "UniversalInvoiceMail.convert_legacy_office_via_com",
                side_effect=RuntimeError("Word/Excel nicht installiert"),
            ) as mocked_com:
                with patch(
                    "UniversalInvoiceMail.convert_legacy_office_via_libreoffice",
                    side_effect=fake_libreoffice,
                ) as mocked_libreoffice:
                    success, message = convert_attachment_to_pdf(
                        b"legacy-data", "altbestand.doc", output_path
                    )

            self.assertTrue(success)
            self.assertIn("LibreOffice", message)
            self.assertTrue(output_path.exists())
            mocked_com.assert_called_once()
            mocked_libreoffice.assert_called_once()

    def test_legacy_attachment_reports_missing_converters(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / "attachment.pdf"

            with patch(
                "UniversalInvoiceMail.convert_legacy_office_via_com",
                side_effect=RuntimeError("COM nicht verfügbar"),
            ):
                with patch(
                    "UniversalInvoiceMail.convert_legacy_office_via_libreoffice",
                    side_effect=RuntimeError("LibreOffice nicht gefunden"),
                ):
                    success, message = convert_attachment_to_pdf(
                        b"legacy-data",
                        "altbestand.xls",
                        output_path,
                    )

        self.assertFalse(success)
        self.assertIn("Legacy-Format", message)
        self.assertIn("konnte nicht konvertiert werden", message)


class TestOcrContentEscaping(unittest.TestCase):
    """Regression-Test: OCR-Text in enhance_with_ocr muss HTML-escaped werden.

    Bug: pytesseract kann Text mit '<', '>', '&' zurueckgeben (z.B. 'Preis < 10 EUR').
    Dieser wurde direkt in ein <pre>-Tag eingefuegt, was das HTML-Dokument brach.
    Fix: escape() aus html-Modul wird auf ocr_content angewendet.
    """

    def test_ocr_text_with_special_chars_is_escaped_in_html(self):
        from unittest.mock import patch, MagicMock
        from UniversalInvoiceMail import OCRProcessor

        ocr_text_with_special = "Preis < 10 EUR & VAT > 0 fuer AT&T"
        captured = {}

        fake_img = MagicMock()

        def fake_image_to_string(img, lang=None):
            return ocr_text_with_special

        def fake_create_pdf(html_src, dest, encoding='utf-8'):
            captured['html'] = html_src
            dest.write(b"%PDF-1.4 fake-ocr")
            return MagicMock()

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "test.pdf"
            pdf_path.write_bytes(b"%PDF-1.4 original")
            ocr_page_path = pdf_path.with_suffix(".ocr_page.pdf")

            processor = OCRProcessor()

            with patch("UniversalInvoiceMail.OCR_AVAILABLE", True), \
                 patch("UniversalInvoiceMail.XHTML2PDF_AVAILABLE", True), \
                 patch.object(processor, "_pdf_to_images", return_value=[fake_img]), \
                 patch("UniversalInvoiceMail.pytesseract") as mock_tess, \
                 patch("UniversalInvoiceMail.pisa") as mock_pisa, \
                 patch("UniversalInvoiceMail.PdfReader") as mock_reader, \
                 patch("UniversalInvoiceMail.PdfWriter") as mock_writer:

                mock_tess.image_to_string.side_effect = fake_image_to_string
                mock_pisa.CreatePDF.side_effect = fake_create_pdf

                # PdfReader/PdfWriter stubs
                fake_reader_inst = MagicMock()
                fake_reader_inst.pages = []
                mock_reader.return_value = fake_reader_inst
                fake_writer_inst = MagicMock()
                mock_writer.return_value = fake_writer_inst

                # ocr_page.pdf muss existieren damit der Check besteht
                ocr_page_path.write_bytes(b"%PDF-1.4 fake-ocr-page")

                processor.enhance_with_ocr(pdf_path)

        html_src = captured.get('html', '')
        self.assertIn('&lt;', html_src,
                      "< in OCR text must be HTML-escaped to &lt;")
        self.assertIn('&gt;', html_src,
                      "> in OCR text must be HTML-escaped to &gt;")
        self.assertIn('&amp;', html_src,
                      "& in OCR text must be HTML-escaped to &amp;")
        self.assertNotIn('<10', html_src,
                         "Raw < must not appear as part of text in HTML")


class TestHtmlToPdfMailMetaEscaping(unittest.TestCase):
    """Regression-Tests: mail_meta-Werte muessen HTML-escaped werden.

    Bug: Sender wie 'Amazon <noreply@amazon.de>' enthielten rohe spitze Klammern,
    die das HTML-Dokument brachen und die E-Mail-Adresse im PDF unsichtbar machten.
    Fix: escape() aus html-Modul wird auf alle mail_meta-Werte angewendet.
    """

    def _call_html_to_pdf_capture_src(self, mail_meta):
        from unittest.mock import patch, MagicMock
        from UniversalInvoiceMail import html_to_pdf

        captured = {}

        def fake_create_pdf(src, dest, link_callback=None, encoding='utf-8'):
            captured['src'] = src
            dest.write(b"%PDF-1.4 fake")
            mock_result = MagicMock()
            mock_result.err = 0
            return mock_result

        with tempfile.TemporaryDirectory() as tmpdir:
            out = Path(tmpdir) / "out.pdf"
            with patch("UniversalInvoiceMail.pisa") as mock_pisa:
                mock_pisa.CreatePDF.side_effect = fake_create_pdf
                # XHTML2PDF_AVAILABLE auf True setzen
                with patch("UniversalInvoiceMail.XHTML2PDF_AVAILABLE", True):
                    html_to_pdf("<p>Inhalt</p>", out, mail_meta, mode="fast")
        return captured.get('src', '')

    def test_sender_with_angle_brackets_is_escaped(self):
        """Absender 'Amazon <noreply@amazon.de>' muss als &lt;...&gt; erscheinen."""
        mail_meta = {
            'sender': 'Amazon <noreply@amazon.de>',
            'subject': 'Ihre Bestellung',
            'date': '2026-01-15',
        }
        src = self._call_html_to_pdf_capture_src(mail_meta)
        self.assertIn('&lt;noreply@amazon.de&gt;', src,
                      "Angle brackets in sender must be HTML-escaped")
        self.assertNotIn('<noreply@amazon.de>', src,
                         "Raw angle brackets in sender must not appear in HTML")

    def test_subject_with_ampersand_is_escaped(self):
        """Betreff mit '&' muss als '&amp;' erscheinen."""
        mail_meta = {
            'sender': 'shop@example.com',
            'subject': 'Rechnung fuer Artikel A & B',
            'date': '2026-01-15',
        }
        src = self._call_html_to_pdf_capture_src(mail_meta)
        self.assertIn('&amp;', src,
                      "Ampersand in subject must be HTML-escaped")

    def test_subject_with_angle_brackets_is_escaped(self):
        """Betreff mit '<' und '>' muss escaped werden."""
        mail_meta = {
            'sender': 'no-reply@shop.de',
            'subject': 'Preis < 10 EUR > Aktionsware',
            'date': '2026-01-15',
        }
        src = self._call_html_to_pdf_capture_src(mail_meta)
        self.assertIn('&lt;', src,
                      "< in subject must be HTML-escaped")
        self.assertIn('&gt;', src,
                      "> in subject must be HTML-escaped")


class TestEmlMsgPlainTextEscaping(unittest.TestCase):
    """Regression-Tests: Unkodierter Plain-Text in EML/MSG-Fallback-Pfad muss escaped werden.

    Bug: _convert_eml_to_pdf und _convert_msg_to_pdf fügten Plain-Text mit
    '<', '>', '&' unescaped in <pre>-Tags ein, was xhtml2pdf-HTML korrumpierte.
    Fix: escape() wird auf den Plain-Text-Inhalt angewendet.
    """

    def _make_eml_bytes(self, body_text: str) -> bytes:
        import email.mime.text
        msg = email.mime.text.MIMEText(body_text, "plain", "utf-8")
        msg["Subject"] = "Test"
        msg["From"] = "sender@example.com"
        return msg.as_bytes()

    def test_eml_plain_text_special_chars_are_escaped(self):
        """EML-Plain-Text mit '<', '>' und '&' muss HTML-escaped an pisa uebergeben werden."""
        from UniversalInvoiceMail import MainWindow

        body_with_specials = "Preis < 10 EUR > Aktionsware & mehr"
        captured = {}

        def fake_create_pdf(src, dest, **kwargs):
            captured["html"] = src
            dest.write(b"%PDF-1.4 fake")
            return MagicMock()

        with tempfile.TemporaryDirectory() as tmpdir:
            eml_path = Path(tmpdir) / "test.eml"
            eml_path.write_bytes(self._make_eml_bytes(body_with_specials))

            win = MainWindow.__new__(MainWindow)

            with patch("UniversalInvoiceMail.XHTML2PDF_AVAILABLE", True), \
                 patch("UniversalInvoiceMail.pisa") as mock_pisa:
                mock_pisa.CreatePDF.side_effect = fake_create_pdf
                win._convert_eml_to_pdf(eml_path)

        html = captured.get("html", "")
        self.assertIn("&lt;", html, "< in EML plain text must be HTML-escaped")
        self.assertIn("&gt;", html, "> in EML plain text must be HTML-escaped")
        self.assertIn("&amp;", html, "& in EML plain text must be HTML-escaped")
        self.assertNotIn("< 10", html, "Raw < must not appear in EML HTML")

    def test_msg_plain_text_special_chars_are_escaped(self):
        """MSG-Plain-Text-Fallback mit '<', '>', '&' muss HTML-escaped an pisa uebergeben werden."""
        from UniversalInvoiceMail import MainWindow

        body_with_specials = "Betrag < 100 EUR & Steuer > 0"
        captured = {}

        def fake_create_pdf(src, dest, **kwargs):
            captured["html"] = src
            dest.write(b"%PDF-1.4 fake")
            return MagicMock()

        mock_msg = MagicMock()
        mock_msg.htmlBody = None
        mock_msg.body = body_with_specials

        mock_extract_msg = MagicMock()
        mock_extract_msg.Message.return_value = mock_msg

        with tempfile.TemporaryDirectory() as tmpdir:
            msg_path = Path(tmpdir) / "test.msg"
            msg_path.write_bytes(b"dummy msg content")

            win = MainWindow.__new__(MainWindow)

            with patch.dict("sys.modules", {"extract_msg": mock_extract_msg}), \
                 patch("UniversalInvoiceMail.XHTML2PDF_AVAILABLE", True), \
                 patch("UniversalInvoiceMail.pisa") as mock_pisa:
                mock_pisa.CreatePDF.side_effect = fake_create_pdf
                win._convert_msg_to_pdf(msg_path)

        html = captured.get("html", "")
        self.assertIn("&lt;", html, "< in MSG plain text must be HTML-escaped")
        self.assertIn("&gt;", html, "> in MSG plain text must be HTML-escaped")
        self.assertIn("&amp;", html, "& in MSG plain text must be HTML-escaped")
        self.assertNotIn("< 100", html, "Raw < must not appear in MSG HTML")


class TestGetMessageBodyEscaping(unittest.TestCase):
    """_get_message_body() muss plain-text mit html.escape() schuetzen."""

    def _make_payload(self, plain_text: str) -> dict:
        import base64
        data = base64.urlsafe_b64encode(plain_text.encode("utf-8")).decode()
        return {
            "mimeType": "multipart/alternative",
            "parts": [
                {
                    "mimeType": "text/plain",
                    "body": {"data": data},
                }
            ],
        }

    def test_plain_text_special_chars_are_escaped(self):
        class WorkerStub:
            _get_all_body_parts = InvoiceWorker._get_all_body_parts
            _get_message_body = InvoiceWorker._get_message_body

        stub = WorkerStub()
        payload = self._make_payload("Preis < 10 EUR & Steuer > 0")
        result = stub._get_message_body(payload)
        self.assertIn("&lt;", result)
        self.assertIn("&gt;", result)
        self.assertIn("&amp;", result)
        self.assertNotIn("< 10", result)

    def test_html_body_is_returned_unescaped(self):
        """HTML-Bodie sollen unverändert zurückgegeben werden."""
        import base64

        class WorkerStub:
            _get_all_body_parts = InvoiceWorker._get_all_body_parts
            _get_message_body = InvoiceWorker._get_message_body

        html_body = "<p>Hello <b>World</b></p>"
        data = base64.urlsafe_b64encode(html_body.encode("utf-8")).decode()
        payload = {
            "mimeType": "text/html",
            "body": {"data": data},
        }
        stub = WorkerStub()
        result = stub._get_message_body(payload)
        self.assertEqual(result, html_body)


class TestPdfToImagesClosesDocument(unittest.TestCase):
    """_pdf_to_images() muss PdfDocument auch im Fehlerfall schliessen."""

    def test_pdf_document_closed_on_render_error(self):
        from UniversalInvoiceMail import OCRProcessor

        mock_pdf = MagicMock()
        mock_pdf.__len__ = MagicMock(return_value=2)
        mock_page = MagicMock()
        mock_page.render.side_effect = RuntimeError("render crash")
        mock_pdf.__getitem__ = MagicMock(return_value=mock_page)

        with patch("UniversalInvoiceMail.pdfium") as mock_pdfium, \
             patch("UniversalInvoiceMail.OCR_AVAILABLE", True):
            mock_pdfium.PdfDocument.return_value = mock_pdf
            proc = OCRProcessor()
            result = proc._pdf_to_images(Path("dummy.pdf"))

        mock_pdf.close.assert_called_once()
        self.assertEqual(result, [])

    def test_pdf_document_closed_on_success(self):
        from UniversalInvoiceMail import OCRProcessor

        mock_bitmap = MagicMock()
        mock_bitmap.to_pil.return_value = MagicMock()
        mock_page = MagicMock()
        mock_page.render.return_value = mock_bitmap

        mock_pdf = MagicMock()
        mock_pdf.__len__ = MagicMock(return_value=1)
        mock_pdf.__getitem__ = MagicMock(return_value=mock_page)

        with patch("UniversalInvoiceMail.pdfium") as mock_pdfium, \
             patch("UniversalInvoiceMail.OCR_AVAILABLE", True):
            mock_pdfium.PdfDocument.return_value = mock_pdf
            proc = OCRProcessor()
            result = proc._pdf_to_images(Path("dummy.pdf"))

        mock_pdf.close.assert_called_once()
        self.assertEqual(len(result), 1)


class TestImapMergeBodyEscaping(unittest.TestCase):
    """_process_imap_message() muss plain-text Body vor dem Merge-Pfad escapen."""

    def _make_multipart_imap_msg(self, plain_text: str) -> object:
        import email.mime.multipart
        import email.mime.text
        import email.mime.application
        msg = email.mime.multipart.MIMEMultipart("mixed")
        msg["Subject"] = "Test Rechnung"
        msg["From"] = "sender@example.com"
        msg["Date"] = "Thu, 1 Jan 2026 12:00:00 +0000"
        msg.attach(email.mime.text.MIMEText(plain_text, "plain", "utf-8"))
        att = email.mime.application.MIMEApplication(b"%PDF-1.4 fake", _subtype="pdf")
        att.add_header("Content-Disposition", "attachment", filename="rechnung.pdf")
        msg.attach(att)
        return msg

    def test_plain_text_body_is_escaped_for_merge(self):
        from UniversalInvoiceMail import InvoiceWorker, AppSettings

        plain_text = "Betrag < 50 EUR & mehr > 0"
        msg = self._make_multipart_imap_msg(plain_text)

        settings = AppSettings()
        settings.merge_body_with_attachments = True
        settings.download_attachments = True
        settings.convert_body_to_pdf = False
        settings.enable_hash_check = False

        captured = {}

        class WorkerStub:
            _get_imap_message_body = InvoiceWorker._get_imap_message_body
            _check_message_filters = staticmethod(lambda _profile, _subject, _body: True)
            _compute_target_dir = staticmethod(lambda _profile: Path("/tmp"))

            def _save_attachment_invoice(self, **kwargs):
                captured.update(kwargs)
                return False

            log = MagicMock()

        stub = WorkerStub()
        stub.settings = settings

        profile = InvoiceProfile(id="prof1", name="Test", account_id="acc1")
        InvoiceWorker._process_imap_message(stub, msg, profile)

        body = captured.get("body_html", "")
        self.assertIn("&lt;", body, "< in IMAP plain body must be escaped for merge")
        self.assertIn("&amp;", body, "& in IMAP plain body must be escaped for merge")
        self.assertNotIn("< 50", body, "Raw < must not appear in IMAP merge body_html")


class TestMergePdfTempCleanup(unittest.TestCase):
    """merge_pdf_with_body() muss temp-Datei auch in Fehlerpfaden loeschen."""

    def test_temp_file_deleted_when_html_to_pdf_fails(self):
        from UniversalInvoiceMail import merge_pdf_with_body

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "source.pdf"
            pdf_path.write_bytes(b"%PDF-1.4 fake")
            output_path = Path(tmpdir) / "out.pdf"

            created_paths = []

            original_ntf = tempfile.NamedTemporaryFile

            def capturing_ntf(**kwargs):
                result = original_ntf(**kwargs)
                created_paths.append(Path(result.name))
                return result

            with patch("UniversalInvoiceMail.XHTML2PDF_AVAILABLE", True), \
                 patch.dict("sys.modules", {"PyPDF2": MagicMock()}), \
                 patch("UniversalInvoiceMail.html_to_pdf", return_value=False), \
                 patch("tempfile.NamedTemporaryFile", side_effect=capturing_ntf):
                merge_pdf_with_body(pdf_path, "<p>body</p>", {}, output_path)

        for p in created_paths:
            self.assertFalse(p.exists(), f"Temp-Datei nicht geloescht: {p}")

    def test_temp_file_deleted_on_merge_exception(self):
        from UniversalInvoiceMail import merge_pdf_with_body

        created_paths = []
        original_ntf = tempfile.NamedTemporaryFile

        def capturing_ntf(**kwargs):
            result = original_ntf(**kwargs)
            created_paths.append(Path(result.name))
            return result

        mock_pypdfs = MagicMock()
        mock_pypdfs.PdfMerger.return_value.append.side_effect = RuntimeError("merge crash")

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "source.pdf"
            pdf_path.write_bytes(b"%PDF-1.4 fake")
            output_path = Path(tmpdir) / "out.pdf"

            with patch("UniversalInvoiceMail.XHTML2PDF_AVAILABLE", True), \
                 patch.dict("sys.modules", {"PyPDF2": mock_pypdfs}), \
                 patch("UniversalInvoiceMail.html_to_pdf", return_value=True), \
                 patch("tempfile.NamedTemporaryFile", side_effect=capturing_ntf):
                merge_pdf_with_body(pdf_path, "<p>body</p>", {}, output_path)

        for p in created_paths:
            self.assertFalse(p.exists(), f"Temp-Datei nicht geloescht: {p}")


class TestMsgConvertClosesMessage(unittest.TestCase):
    """_convert_msg_to_pdf() muss extract_msg.Message auch im Fehlerfall schliessen."""

    def _make_mock_msg(self, plain_body: str = "Test"):
        mock_msg = MagicMock()
        mock_msg.htmlBody = None
        mock_msg.body = plain_body
        return mock_msg

    def test_msg_closed_on_pisa_exception(self):
        from UniversalInvoiceMail import MainWindow

        mock_msg = self._make_mock_msg()
        mock_extract_msg = MagicMock()
        mock_extract_msg.Message.return_value = mock_msg

        with tempfile.TemporaryDirectory() as tmpdir:
            msg_path = Path(tmpdir) / "test.msg"
            msg_path.write_bytes(b"dummy")

            win = MainWindow.__new__(MainWindow)

            with patch.dict("sys.modules", {"extract_msg": mock_extract_msg}), \
                 patch("UniversalInvoiceMail.XHTML2PDF_AVAILABLE", True), \
                 patch("UniversalInvoiceMail.pisa") as mock_pisa:
                mock_pisa.CreatePDF.side_effect = RuntimeError("pisa crash")
                result = win._convert_msg_to_pdf(msg_path)

        self.assertIsNone(result)
        mock_msg.close.assert_called_once()

    def test_msg_closed_on_success(self):
        from UniversalInvoiceMail import MainWindow

        mock_msg = self._make_mock_msg()
        mock_extract_msg = MagicMock()
        mock_extract_msg.Message.return_value = mock_msg

        with tempfile.TemporaryDirectory() as tmpdir:
            msg_path = Path(tmpdir) / "test.msg"
            msg_path.write_bytes(b"dummy")

            win = MainWindow.__new__(MainWindow)

            def fake_create_pdf(html, dest, **kwargs):
                dest.write(b"%PDF-1.4 fake")
                return MagicMock(err=0)

            with patch.dict("sys.modules", {"extract_msg": mock_extract_msg}), \
                 patch("UniversalInvoiceMail.XHTML2PDF_AVAILABLE", True), \
                 patch("UniversalInvoiceMail.pisa") as mock_pisa:
                mock_pisa.CreatePDF.side_effect = fake_create_pdf
                result = win._convert_msg_to_pdf(msg_path)

        mock_msg.close.assert_called_once()
        self.assertIsNotNone(result)


class TestBrowserRendererCloseOnAppExit(unittest.TestCase):
    """closeEvent() muss _browser_renderer schliessen wenn er initialisiert ist."""

    def test_browser_renderer_closed_on_close_event(self):
        import UniversalInvoiceMail as uim
        from UniversalInvoiceMail import MainWindow

        mock_renderer = MagicMock()
        original_renderer = uim._browser_renderer

        win = MainWindow.__new__(MainWindow)
        win.worker = None
        win.profiles = []
        win.accounts = []
        win.invoices = []
        win.settings = MagicMock()

        try:
            uim._browser_renderer = mock_renderer
            mock_event = MagicMock()

            with patch.object(win, 'save_config'):
                win.closeEvent(mock_event)

            mock_renderer.close.assert_called_once()
            self.assertIsNone(uim._browser_renderer)
        finally:
            uim._browser_renderer = original_renderer

    def test_close_event_safe_without_renderer(self):
        import UniversalInvoiceMail as uim
        from UniversalInvoiceMail import MainWindow

        original_renderer = uim._browser_renderer

        win = MainWindow.__new__(MainWindow)
        win.worker = None
        win.profiles = []
        win.accounts = []
        win.invoices = []
        win.settings = MagicMock()

        try:
            uim._browser_renderer = None
            mock_event = MagicMock()

            with patch.object(win, 'save_config'):
                win.closeEvent(mock_event)  # darf nicht werfen

            mock_event.accept.assert_called_once()
        finally:
            uim._browser_renderer = original_renderer


class TestImapConnectionShutdown(unittest.TestCase):
    """_process_imap() muss IMAP-Verbindung auch bei Exceptions schliessen."""

    def test_shutdown_called_on_login_error(self):
        from UniversalInvoiceMail import InvoiceWorker, MailAccount, AppSettings

        mock_mail = MagicMock()
        mock_mail.login.side_effect = Exception("login failed")

        account = MailAccount(
            id="test", name="Test", provider="IMAP",
            host="imap.test.com", port=993,
            username="user@test.com", use_gmail_api=False
        )
        settings = AppSettings()

        worker = InvoiceWorker.__new__(InvoiceWorker)
        worker.accounts = [account]
        worker.profiles = []
        worker.settings = settings
        worker.should_stop = False
        worker.found_count = 0
        worker.log = MagicMock()
        worker.log.emit = MagicMock()

        with patch("UniversalInvoiceMail.imaplib") as mock_imaplib, \
             patch("UniversalInvoiceMail.KEYRING_AVAILABLE", True), \
             patch("UniversalInvoiceMail.keyring") as mock_keyring:
            mock_keyring.get_password.return_value = "password"
            mock_imaplib.IMAP4_SSL.return_value = mock_mail
            mock_imaplib.IMAP4.error = Exception
            worker._process_imap(account, [])

        mock_mail.shutdown.assert_called_once()

    def test_shutdown_called_on_success(self):
        from UniversalInvoiceMail import InvoiceWorker, MailAccount, AppSettings

        mock_mail = MagicMock()
        mock_mail.login.return_value = ("OK", [])
        mock_mail.list.return_value = ("OK", [b'(\\HasNoChildren) "/" "INBOX"'])
        mock_mail.select.return_value = ("OK", [b"1"])
        mock_mail.search.return_value = ("OK", [b""])
        mock_mail.logout.return_value = ("BYE", [])

        account = MailAccount(
            id="test2", name="Test2", provider="IMAP",
            host="imap.test.com", port=993,
            username="user@test.com", use_gmail_api=False
        )
        settings = AppSettings()

        worker = InvoiceWorker.__new__(InvoiceWorker)
        worker.accounts = [account]
        worker.profiles = []
        worker.settings = settings
        worker.should_stop = False
        worker.found_count = 0
        worker.log = MagicMock()
        worker.log.emit = MagicMock()

        with patch("UniversalInvoiceMail.imaplib") as mock_imaplib, \
             patch("UniversalInvoiceMail.KEYRING_AVAILABLE", True), \
             patch("UniversalInvoiceMail.keyring") as mock_keyring:
            mock_keyring.get_password.return_value = "password"
            mock_imaplib.IMAP4_SSL.return_value = mock_mail
            mock_imaplib.IMAP4.error = Exception
            worker._process_imap(account, [])

        mock_mail.shutdown.assert_called_once()


class TestMergePdfMergerClosed(unittest.TestCase):
    """merge_pdf_with_body() muss PdfMerger auch bei Exceptions schliessen."""

    def test_merger_closed_on_append_exception(self):
        from UniversalInvoiceMail import merge_pdf_with_body

        mock_merger = MagicMock()
        mock_merger.append.side_effect = RuntimeError("append crash")
        mock_pypdfs = MagicMock()
        mock_pypdfs.PdfMerger.return_value = mock_merger

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "source.pdf"
            pdf_path.write_bytes(b"%PDF-1.4 fake")
            output_path = Path(tmpdir) / "out.pdf"

            with patch("UniversalInvoiceMail.XHTML2PDF_AVAILABLE", True), \
                 patch.dict("sys.modules", {"PyPDF2": mock_pypdfs}), \
                 patch("UniversalInvoiceMail.html_to_pdf", return_value=True):
                result = merge_pdf_with_body(pdf_path, "<p>body</p>", {}, output_path)

        mock_merger.close.assert_called_once()
        # Fallback-Kopie sollte trotzdem stattgefunden haben
        self.assertTrue(result)

    def test_merger_closed_on_success(self):
        from UniversalInvoiceMail import merge_pdf_with_body

        mock_merger = MagicMock()
        mock_pypdfs = MagicMock()
        mock_pypdfs.PdfMerger.return_value = mock_merger

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "source.pdf"
            pdf_path.write_bytes(b"%PDF-1.4 fake")
            output_path = Path(tmpdir) / "out.pdf"
            output_path.write_bytes(b"%PDF-1.4 merged")

            with patch("UniversalInvoiceMail.XHTML2PDF_AVAILABLE", True), \
                 patch.dict("sys.modules", {"PyPDF2": mock_pypdfs}), \
                 patch("UniversalInvoiceMail.html_to_pdf", return_value=True):
                merge_pdf_with_body(pdf_path, "<p>body</p>", {}, output_path)

        mock_merger.close.assert_called_once()


class TestBrowserRendererHeaderInjection(unittest.TestCase):
    """render_html_to_pdf() darf Backslashes in mail_meta nicht als Regex-Backrefs auswerten."""

    def test_backslash_in_sender_does_not_duplicate_body_tag(self):
        from UniversalInvoiceMail import BrowserPDFRenderer
        import re

        renderer = BrowserPDFRenderer.__new__(BrowserPDFRenderer)
        renderer.log = lambda _: None
        renderer.driver = None
        renderer._initialized = False

        html = "<html><body><p>Test</p></body></html>"
        mail_meta = {
            "sender": r"CORP\1administrator",
            "subject": "Rechnung",
            "date": "2026-01-01",
        }

        # Wir testen nur den Header-Einfüge-Pfad, nicht den vollen Render
        from html import escape
        header_html = (
            f'<div>{escape(mail_meta["sender"])}</div>'
        )
        if "<body" in html.lower():
            _hdr = header_html
            result = re.sub(
                r"(<body[^>]*>)",
                lambda m: m.group(1) + _hdr,
                html,
                count=1,
                flags=re.IGNORECASE,
            )
        else:
            result = f"<html><body>{header_html}{html}</body></html>"

        # <body> darf nur einmal vorkommen
        self.assertEqual(result.lower().count("<body>"), 1)
        # Sender-Text muss korrekt drin sein
        self.assertIn("CORP\\1administrator", result)


class TestOcrMsgVarShadowing(unittest.TestCase):
    """OCR-Rückgabewert darf den msg-Parameter nicht überschreiben."""

    def _check_no_msg_tuple_assign(self, method):
        import ast
        import inspect
        import textwrap
        src = textwrap.dedent(inspect.getsource(method))
        tree = ast.parse(src)
        for node in ast.walk(tree):
            if isinstance(node, ast.Assign):
                for target in node.targets:
                    if isinstance(target, ast.Tuple):
                        names = [
                            elt.id for elt in target.elts
                            if isinstance(elt, ast.Name)
                        ]
                        self.assertNotIn(
                            "msg", names,
                            f"OCR-Rückgabe darf 'msg'-Parameter nicht überschreiben in {method.__name__}"
                        )

    def test_gmail_ocr_result_uses_separate_variable(self):
        from UniversalInvoiceMail import InvoiceWorker
        self._check_no_msg_tuple_assign(InvoiceWorker._process_gmail_message)

    def test_imap_ocr_result_uses_separate_variable(self):
        from UniversalInvoiceMail import InvoiceWorker
        self._check_no_msg_tuple_assign(InvoiceWorker._process_imap_message)


class TestAddTextLayerTempCleanup(unittest.TestCase):
    """add_text_layer() muss die Temp-Datei auch bei Exceptions bereinigen."""

    def test_temp_file_cleaned_on_move_exception(self):
        proc = OCRProcessor.__new__(OCRProcessor)

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "source.pdf"
            pdf_path.write_bytes(b"%PDF-1.4 fake")
            temp_path = pdf_path.with_suffix(".ocr_temp.pdf")

            sys.modules['pytesseract'].image_to_pdf_or_hocr = MagicMock(
                return_value=b"%PDF-1.4 fake"
            )

            with patch("UniversalInvoiceMail.OCR_AVAILABLE", True), \
                 patch.object(proc, "_pdf_to_images", return_value=[MagicMock()]), \
                 patch("UniversalInvoiceMail.shutil") as mock_shutil:

                mock_shutil.move.side_effect = RuntimeError("move failed")
                success, msg = proc.add_text_layer(pdf_path)

        self.assertFalse(success)
        self.assertFalse(
            temp_path.exists(),
            "Temp-Datei muss nach Exception bereinigt werden"
        )

    def test_no_temp_file_on_early_return(self):
        proc = OCRProcessor.__new__(OCRProcessor)

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "source.pdf"
            pdf_path.write_bytes(b"%PDF-1.4 fake")
            temp_path = pdf_path.with_suffix(".ocr_temp.pdf")

            with patch("UniversalInvoiceMail.OCR_AVAILABLE", True), \
                 patch.object(proc, "_pdf_to_images", return_value=[]):

                success, msg = proc.add_text_layer(pdf_path)

        self.assertFalse(success)
        self.assertFalse(temp_path.exists())


class TestGmailDateFilterMonthsFallback(unittest.TestCase):
    """_build_gmail_search_query() darf date_filter_months nur ohne date_from UND date_to auslösen.

    Bug #20: Wenn nur date_to gesetzt war, wurde fälschlicherweise ein after:-Filter
    aus date_filter_months hinzugefügt — dies schnitt Mails vor dem Fallback-Datum ab.
    Fix: Bedingung auf 'not date_from and not date_to' erweitert.
    """

    def _make_worker(self, date_from="", date_to="", months=12):
        from UniversalInvoiceMail import InvoiceWorker, AppSettings
        settings = AppSettings(date_from=date_from, date_to=date_to,
                               date_filter_months=months)
        worker = InvoiceWorker.__new__(InvoiceWorker)
        worker.settings = settings
        worker.log = MagicMock()
        return worker

    def _make_profile(self):
        from UniversalInvoiceMail import InvoiceProfile
        return InvoiceProfile(id="p1", name="Test", account_id="a1",
                              sender_filter="", subject_filter="")

    def test_only_date_to_set_no_fallback_start(self):
        """Wenn nur date_to gesetzt, darf kein 'after:' aus date_filter_months erscheinen."""
        worker = self._make_worker(date_to="2025-12-31", months=12)
        query = worker._build_gmail_search_query(self._make_profile())
        self.assertIn("before:2025/12/31", query)
        self.assertNotIn("after:", query)

    def test_both_empty_fallback_triggers(self):
        """Wenn weder date_from noch date_to gesetzt, muss Fallback 'after:...' erscheinen."""
        worker = self._make_worker(months=12)
        query = worker._build_gmail_search_query(self._make_profile())
        self.assertIn("after:", query)

    def test_only_date_from_set_no_fallback(self):
        """Wenn date_from gesetzt, darf kein zusätzliches 'after:' aus Fallback erscheinen."""
        worker = self._make_worker(date_from="2025-01-01", months=12)
        query = worker._build_gmail_search_query(self._make_profile())
        self.assertIn("after:2025/01/01", query)
        # Nur ein 'after:' erlaubt (das aus date_from), nicht zwei
        self.assertEqual(query.count("after:"), 1)


class TestEnhanceWithOcrFileHandleRelease(unittest.TestCase):
    """enhance_with_ocr() muss PdfReader-Handles vor unlink() freigeben (Bug #19).

    Auf Windows sperrt ein offener PdfReader die Datei; unlink() würde mit
    PermissionError fehlschlagen und OCR still falsch zurückgeben.
    Fix: del ocr_reader / del original_reader nach den Pages-Schleifen.
    """

    def test_ocr_page_temp_file_is_deleted_on_success(self):
        """Die .ocr_page.pdf Temp-Datei muss nach erfolgreichem Merge gelöscht werden."""
        from UniversalInvoiceMail import OCRProcessor

        proc = OCRProcessor.__new__(OCRProcessor)

        fake_page = MagicMock()
        mock_writer = MagicMock()
        mock_writer.pages = [fake_page]
        mock_reader_instance = MagicMock()
        mock_reader_instance.pages = [fake_page]

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / "source.pdf"
            pdf_path.write_bytes(b"%PDF-1.4 fake")
            ocr_page_path = pdf_path.with_suffix(".ocr_page.pdf")
            ocr_page_path.write_bytes(b"%PDF-1.4 ocr")

            with patch("UniversalInvoiceMail.OCR_AVAILABLE", True), \
                 patch("UniversalInvoiceMail.XHTML2PDF_AVAILABLE", True), \
                 patch.object(proc, "_pdf_to_images", return_value=[MagicMock()]), \
                 patch("UniversalInvoiceMail.pytesseract") as mock_tess, \
                 patch("UniversalInvoiceMail.pisa") as mock_pisa, \
                 patch("UniversalInvoiceMail.PdfWriter") as mock_writer_cls, \
                 patch("UniversalInvoiceMail.PdfReader") as mock_reader_cls, \
                 patch("UniversalInvoiceMail.shutil"):

                mock_tess.image_to_string.return_value = "Rechnungsbetrag 99,00 EUR"
                mock_pisa.CreatePDF.side_effect = lambda html, dest, **kw: dest.write(b"%PDF-1.4 ocr")
                mock_writer_cls.return_value = mock_writer
                mock_reader_cls.return_value = mock_reader_instance

                proc.enhance_with_ocr(pdf_path)

            self.assertEqual(mock_reader_cls.call_count, 2)


class TestLoadConfigTypeErrorRobustness(unittest.TestCase):
    """load_config() darf bei korrupten JSON-Einträgen (fehlende Pflichtfelder) nicht crashen."""

    def _make_window(self):
        from UniversalInvoiceMail import MainWindow
        win = MainWindow.__new__(MainWindow)
        win.accounts = []
        win.profiles = []
        win.invoices = []
        win.settings = __import__('UniversalInvoiceMail').AppSettings()
        win.datev_config = None
        return win

    def test_load_config_survives_corrupt_invoice_db(self):
        """Invoices-DB mit fehlendem Pflichtfeld 'id' darf keinen Startup-Crash verursachen."""
        corrupt_invoices = [{"profile_name": "Test", "filename": "f.pdf",
                              "date": "2026-01-01", "path": "/tmp/f.pdf"}]
        # 'id' fehlt → Invoice(**filtered) wirft TypeError ohne den Fix

        win = self._make_window()

        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = Path(tmpdir) / "invoices.json"
            db_path.write_text(__import__('json').dumps(corrupt_invoices), encoding='utf-8')

            with patch('UniversalInvoiceMail.INVOICES_DB', db_path), \
                 patch('UniversalInvoiceMail.CONFIG_FILE', Path(tmpdir) / "cfg.json"):
                win.load_config()  # darf nicht werfen

        self.assertEqual(win.invoices, [])

    def test_load_config_survives_corrupt_profile_config(self):
        """Config mit fehlendem Pflichtfeld in einem Profil darf keinen Startup-Crash verursachen."""
        import json

        corrupt_config = {
            "settings": {},
            "accounts": [],
            "profiles": [{"name": "Ohne ID", "account_id": "acc1"}]
            # 'id' fehlt → InvoiceProfile(**filtered) wirft TypeError
        }

        win = self._make_window()

        with tempfile.TemporaryDirectory() as tmpdir:
            cfg_path = Path(tmpdir) / "config.json"
            cfg_path.write_text(json.dumps(corrupt_config), encoding='utf-8')

            with patch('UniversalInvoiceMail.CONFIG_FILE', cfg_path), \
                 patch('UniversalInvoiceMail.INVOICES_DB', Path(tmpdir) / "inv.json"):
                win.load_config()  # darf nicht werfen

        self.assertEqual(win.profiles, [])


if __name__ == "__main__":
    unittest.main()
