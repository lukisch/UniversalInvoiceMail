"""Integration tests for UniversalInvoiceMail IMAP and Gmail API flows.

Uses unittest.mock to avoid real IMAP/Gmail API connections.
Run with: python -m pytest tests/test_integration.py
"""

import sys
import base64
import email
import unittest
from email.message import EmailMessage
from pathlib import Path
from unittest.mock import MagicMock, patch

# Add project root to sys.path so we can import UniversalInvoiceMail
sys.path.insert(0, str(Path(__file__).parent.parent))

# Mock PyQt6 before importing to allow headless execution
_qt_mock = MagicMock()
sys.modules['PyQt6'] = _qt_mock
sys.modules['PyQt6.QtWidgets'] = _qt_mock
sys.modules['PyQt6.QtCore'] = _qt_mock
sys.modules['PyQt6.QtGui'] = _qt_mock

# Mock optional dependencies that might not be installed
for mod in ['xhtml2pdf', 'xhtml2pdf.pisa', 'pytesseract', 'pypdfium2',
            'pypdf', 'PIL', 'PIL.Image', 'selenium', 'webdriver_manager',
            'webdriver_manager.microsoft', 'webdriver_manager.chrome',
            'googleapiclient', 'googleapiclient.discovery',
            'google_auth_oauthlib', 'google_auth_oauthlib.flow',
            'google.auth', 'google.auth.transport', 'google.auth.transport.requests',
            'google.auth.exceptions', 'google.oauth2', 'google.oauth2.credentials',
            'keyring']:
    if mod not in sys.modules:
        if mod.startswith(("selenium", "webdriver")):
            sys.modules[mod] = MagicMock()
        else:
            try:
                __import__(mod)
            except ImportError:
                sys.modules[mod] = MagicMock()

import UniversalInvoiceMail as uim


class TestHtmlPdfSanitizer(unittest.TestCase):
    """Regression tests for HTML filtering before PDF conversion."""

    def test_basic_sanitizer_removes_script_end_tags_with_spaces(self):
        html = "<p>OK</p><script>alert(1)</script ><style>.x{}</style><p>Done</p>"

        cleaned = uim.sanitize_html_for_pdf(html)

        self.assertIn("<p>OK</p>", cleaned)
        self.assertIn("<p>Done</p>", cleaned)
        self.assertNotIn("alert(1)", cleaned)
        self.assertNotIn("<script", cleaned.lower())
        self.assertNotIn("<style", cleaned.lower())

    def test_full_sanitizer_removes_script_end_tags_with_spaces(self):
        html = "<p>OK</p><script type=\"text/javascript\">alert(1)</script ><img src=\"cid:test\">"

        cleaned = uim.sanitize_html_for_pdf_full(html)

        self.assertIn("<p>OK</p>", cleaned)
        self.assertIn("cid:test", cleaned)
        self.assertNotIn("alert(1)", cleaned)
        self.assertNotIn("<script", cleaned.lower())


class TestMailAccountFromDict(unittest.TestCase):
    """Regression tests for MailAccount.from_dict robustness."""

    def test_from_dict_ignores_unknown_keys(self):
        """from_dict must not raise TypeError on future/unknown config keys."""
        d = {
            "id": "a1", "name": "Test", "provider": "IMAP",
            "host": "imap.example.com", "port": 993, "username": "u@x.de",
            "use_gmail_api": False,
            "unknown_future_field": "some_value",
        }
        account = uim.MailAccount.from_dict(d)
        self.assertEqual(account.host, "imap.example.com")

    def test_from_dict_uses_defaults_for_missing_keys(self):
        """from_dict must use dataclass defaults for keys absent in the saved dict."""
        d = {"id": "a1", "name": "Test", "provider": "IMAP"}
        account = uim.MailAccount.from_dict(d)
        self.assertFalse(account.use_gmail_api)
        self.assertEqual(account.port, 993)


class TestImapConnect(unittest.TestCase):
    """Tests for IMAP connection and authentication flow."""

    def _make_account(self, use_gmail_api=False):
        return uim.MailAccount(
            id="acc1",
            name="Test Account",
            provider="IMAP",
            host="imap.example.com",
            port=993,
            username="user@example.com",
            use_gmail_api=use_gmail_api,
        )

    def _make_profile(self):
        return uim.InvoiceProfile(
            id="prof1",
            name="TestShop",
            account_id="acc1",
            sender_filter="shop@example.com",
            subject_filter="Rechnung",
            enabled=True,
        )

    def _make_settings(self, tmp_path=None):
        s = uim.AppSettings()
        if tmp_path:
            s.download_path = str(tmp_path)
        return s

    def test_imap_connect_success(self):
        """IMAP connection succeeds with valid credentials."""
        account = self._make_account()
        profile = self._make_profile()
        settings = self._make_settings()
        worker = uim.InvoiceWorker([account], [profile], settings, set())

        mock_mail = MagicMock()
        mock_mail.list.return_value = (None, [b'(\\HasNoChildren) "/" INBOX'])

        with patch('imaplib.IMAP4_SSL', return_value=mock_mail):
            with patch('UniversalInvoiceMail.KEYRING_AVAILABLE', True):
                with patch('keyring.get_password', return_value='secret'):
                    mock_mail.uid.return_value = (None, [b''])
                    worker._process_imap(account, [profile])

        mock_mail.login.assert_called_once_with('user@example.com', 'secret')

    def test_imap_connect_no_password(self):
        """IMAP connection aborts when no password is stored."""
        account = self._make_account()
        profile = self._make_profile()
        settings = self._make_settings()
        worker = uim.InvoiceWorker([account], [profile], settings, set())
        log_messages = []
        worker.log = MagicMock()
        worker.log.emit = lambda msg: log_messages.append(msg)

        with patch('UniversalInvoiceMail.KEYRING_AVAILABLE', True):
            with patch('keyring.get_password', return_value=None):
                worker._process_imap(account, [profile])

        self.assertTrue(any("Kein Passwort" in m for m in log_messages))

    def test_imap_connect_auth_error(self):
        """IMAP connection handles authentication error gracefully."""
        import imaplib
        account = self._make_account()
        profile = self._make_profile()
        settings = self._make_settings()
        worker = uim.InvoiceWorker([account], [profile], settings, set())
        log_messages = []
        worker.log = MagicMock()
        worker.log.emit = lambda msg: log_messages.append(msg)

        mock_mail = MagicMock()
        mock_mail.login.side_effect = imaplib.IMAP4.error("LOGIN failed")

        with patch('imaplib.IMAP4_SSL', return_value=mock_mail):
            with patch('UniversalInvoiceMail.KEYRING_AVAILABLE', True):
                with patch('keyring.get_password', return_value='wrongpwd'):
                    worker._process_imap(account, [profile])

        self.assertTrue(any("IMAP Fehler" in m for m in log_messages))

    def test_uid_called_not_msn(self):
        """_search_imap muss uid() statt search()/fetch() verwenden (MSN/UID-Bug).

        MSN-Nummern sind instabil sobald andere Clients Mails verschieben oder
        löschen. uid() adressiert Mails über stabile UIDs (RFC 3501 §2.3.1.1).
        """
        account = self._make_account()
        profile = self._make_profile()
        settings = self._make_settings()
        worker = uim.InvoiceWorker([account], [profile], settings, set())

        mock_mail = MagicMock()
        # uid('search', ...) gibt leere ID-Liste zurück → keine weiteren Fetch-Calls
        mock_mail.uid.return_value = (None, [b''])

        worker._search_imap(mock_mail, profile)

        # search() darf NICHT aufgerufen worden sein
        mock_mail.search.assert_not_called()
        # uid() muss mit 'search' aufgerufen worden sein
        calls = [str(c) for c in mock_mail.uid.call_args_list]
        self.assertTrue(
            any('search' in c.lower() for c in calls),
            f"uid() wurde nicht mit 'search' aufgerufen. Calls: {calls}",
        )

    def test_uid_nil_response_guard(self):
        """NIL-Antwort von uid('fetch') führt zu 'continue', nicht zu einem Absturz.

        Ein IMAP-Server kann für eine nicht mehr existierende UID eine leere
        Antwortstruktur zurückgeben. Der Guard verhindert AttributeError auf
        msg_data[0][1].
        """
        account = self._make_account()
        profile = self._make_profile()
        settings = self._make_settings()
        worker = uim.InvoiceWorker([account], [profile], settings, set())
        log_messages = []
        worker.log = MagicMock()
        worker.log.emit = lambda msg: log_messages.append(msg)
        worker.progress = MagicMock()

        mock_mail = MagicMock()
        # uid('search') liefert eine UID zurück, uid('fetch') liefert NIL
        mock_mail.uid.side_effect = [
            (None, [b'42']),   # search-Aufruf: 1 UID gefunden
            (None, [None]),    # fetch-Aufruf: NIL-Antwort (z. B. Mail inzwischen gelöscht)
        ]

        # Kein Exception-Crash erwartet
        worker._search_imap(mock_mail, profile)

        # Es muss eine Warnung geloggt worden sein
        self.assertTrue(
            any("UID" in m or "Keine Daten" in m or "Leere" in m for m in log_messages),
            f"Erwartete NIL-Warnung nicht gefunden. Log: {log_messages}",
        )

    def test_imap_message_body_charset(self):
        """_get_imap_message_body dekodiert non-UTF-8 Charsets korrekt.

        Viele ältere oder deutsche Mailer versenden text/plain in ISO-8859-1 oder
        windows-1252. Früher wurde blind UTF-8 angenommen, was zu '?' für Umlaute führte.
        """
        account = self._make_account()
        profile = self._make_profile()
        settings = self._make_settings()
        worker = uim.InvoiceWorker([account], [profile], settings, set())

        # Erstellt eine E-Mail mit ISO-8859-1-Body (enthält echte Umlaute)
        raw_body = "Rechnung für Müller & Söhne GmbH".encode('iso-8859-1')
        msg = email.message.Message()
        msg['Content-Type'] = 'text/plain; charset="iso-8859-1"'
        msg['Content-Transfer-Encoding'] = 'quoted-printable'
        msg.set_payload(raw_body)
        # get_payload(decode=True) gibt raw bytes zurück; wir mocken es direkt
        msg.set_payload(raw_body.decode('iso-8859-1'), charset='iso-8859-1')

        # Direkter Byte-Test über ein synthetisches email.message.Message-Objekt
        from email import message as email_message
        part = email_message.Message()
        part['Content-Type'] = 'text/plain; charset="iso-8859-1"'
        part._payload = raw_body  # interne Repräsentation als bytes

        class FakeMsg:
            def walk(self_inner):
                return [part]

        with patch.object(type(part), 'get_payload', return_value=raw_body):
            result = worker._get_imap_message_body(FakeMsg())

        # Alle Umlaute müssen korrekt dekodiert sein
        self.assertIn("Müller", result)
        self.assertIn("Söhne", result)


class TestGmailAuth(unittest.TestCase):
    """Tests for Gmail API authentication flow."""

    def test_gmail_auth_missing_credentials_file(self):
        """Gmail auth aborts when credentials.json is absent and cannot be found."""
        account = uim.MailAccount(
            id="acc2", name="Gmail", provider="Gmail API",
            use_gmail_api=True
        )
        profile = uim.InvoiceProfile(
            id="p2", name="Amazon", account_id="acc2", enabled=True
        )
        settings = uim.AppSettings()
        worker = uim.InvoiceWorker([account], [profile], settings, set())
        log_messages = []
        worker.log = MagicMock()
        worker.log.emit = lambda msg: log_messages.append(msg)

        with patch('UniversalInvoiceMail.GMAIL_API_AVAILABLE', True):
            with patch.object(Path, 'exists', return_value=False):
                creds = worker._get_gmail_credentials()

        self.assertIsNone(creds)

    def test_gmail_api_unavailable(self):
        """_process_gmail_api logs an error when Gmail API is not installed."""
        account = uim.MailAccount(
            id="acc3", name="Gmail2", provider="Gmail API", use_gmail_api=True
        )
        profile = uim.InvoiceProfile(
            id="p3", name="Google", account_id="acc3", enabled=True
        )
        settings = uim.AppSettings()
        worker = uim.InvoiceWorker([account], [profile], settings, set())
        log_messages = []
        worker.log = MagicMock()
        worker.log.emit = lambda msg: log_messages.append(msg)

        with patch('UniversalInvoiceMail.GMAIL_API_AVAILABLE', False):
            worker._process_gmail_api(account, [profile])

        self.assertTrue(any("nicht verfügbar" in m or "nicht verf" in m
                            for m in log_messages))
    def test_invoice_profile_defaults_gmail_query_for_old_configs(self):
        """Older profile configs load with an empty gmail_query field."""
        profile = uim.InvoiceProfile.from_dict({
            "id": "p-old",
            "name": "Amazon",
            "account_id": "acc-old",
            "subject_filter": "Rechnung",
        })
        self.assertTrue(hasattr(profile, "gmail_query"))
        self.assertEqual(profile.gmail_query, "")

    def test_search_gmail_includes_saved_gmail_query_and_filters(self):
        """Saved Gmail raw queries are combined with the regular Gmail API filters."""
        account = uim.MailAccount(
            id="acc4", name="Gmail3", provider="Gmail API", use_gmail_api=True
        )
        profile = uim.InvoiceProfile(
            id="p4",
            name="Amazon",
            account_id="acc4",
            sender_filter="billing@example.com, invoices@example.com",
            subject_filter="Rechnung,Invoice",
            gmail_query="label:finance has:attachment",
            enabled=True,
        )
        settings = uim.AppSettings(date_from="2026-05-01", date_to="2026-05-31")
        worker = uim.InvoiceWorker([account], [profile], settings, set())
        worker.log = MagicMock()
        worker.log.emit = MagicMock()
        seen = {}

        class _ListCall:
            def execute(self):
                return {"messages": []}

        class _Messages:
            def list(self, **kwargs):
                seen["query"] = kwargs["q"]
                seen["maxResults"] = kwargs["maxResults"]
                return _ListCall()

        class _Users:
            def messages(self):
                return _Messages()

        class _Service:
            def users(self):
                return _Users()

        worker._search_gmail(_Service(), profile)

        self.assertIn("label:finance has:attachment", seen["query"])
        self.assertIn("from:billing@example.com", seen["query"])
        self.assertIn('subject:"Rechnung"', seen["query"])
        self.assertIn("after:2026/05/01", seen["query"])
        self.assertIn("before:2026/05/31", seen["query"])


class TestDownloadAttachment(unittest.TestCase):
    """Tests for attachment detection and download."""

    def test_filter_messages_blacklist(self):
        """Messages matching the blacklist are rejected by _check_message_filters."""
        account = uim.MailAccount(id="a", name="A", provider="IMAP")
        profile = uim.InvoiceProfile(
            id="p", name="Shop", account_id="a",
            blacklist="newsletter,werbung",
            enabled=True,
        )
        settings = uim.AppSettings()
        worker = uim.InvoiceWorker([account], [profile], settings, set())
        worker.log = MagicMock()
        worker.log.emit = MagicMock()

        # Message with blacklisted word in subject
        result = worker._check_message_filters(profile, "Dein Newsletter", "")
        self.assertFalse(result)

        # Message with blacklisted word in body
        result = worker._check_message_filters(profile, "Rechnung", "tolle werbung hier")
        self.assertFalse(result)

    def test_filter_messages_passes(self):
        """Messages with no blacklist match pass _check_message_filters."""
        account = uim.MailAccount(id="a", name="A", provider="IMAP")
        profile = uim.InvoiceProfile(
            id="p", name="Shop", account_id="a",
            blacklist="newsletter",
            body_must_contain="rechnung",
            enabled=True,
        )
        settings = uim.AppSettings()
        worker = uim.InvoiceWorker([account], [profile], settings, set())
        worker.log = MagicMock()
        worker.log.emit = MagicMock()

        result = worker._check_message_filters(
            profile, "Ihre Rechnung bei Shop", "ihre rechnung ist beigefuegt"
        )
        self.assertTrue(result)

    def test_filter_messages_body_must_contain_fails(self):
        """Messages missing required body terms are rejected."""
        account = uim.MailAccount(id="a", name="A", provider="IMAP")
        profile = uim.InvoiceProfile(
            id="p", name="Shop", account_id="a",
            body_must_contain="rechnung,invoice",
            enabled=True,
        )
        settings = uim.AppSettings()
        worker = uim.InvoiceWorker([account], [profile], settings, set())
        worker.log = MagicMock()
        worker.log.emit = MagicMock()

        result = worker._check_message_filters(
            profile, "Hallo, toller Inhalt", "keine relevanten inhalte hier"
        )
        self.assertFalse(result)

    def test_compute_target_dir(self, tmp_path=None):
        """_compute_target_dir creates the expected subdirectory."""
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            account = uim.MailAccount(id="a", name="A", provider="IMAP")
            profile = uim.InvoiceProfile(
                id="p", name="TestShop", account_id="a", enabled=True
            )
            settings = uim.AppSettings()
            settings.download_path = tmp
            worker = uim.InvoiceWorker([account], [profile], settings, set())

            target = worker._compute_target_dir(profile)
            self.assertTrue(target.exists())
            self.assertTrue(target.name.startswith("TestShop"))

    def test_compute_target_dir_with_subfolder(self):
        """_compute_target_dir uses target_subfolder when specified."""
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            account = uim.MailAccount(id="a", name="A", provider="IMAP")
            profile = uim.InvoiceProfile(
                id="p", name="Shop", account_id="a",
                target_subfolder="Rechnungen2024",
                enabled=True,
            )
            settings = uim.AppSettings()
            settings.download_path = tmp
            worker = uim.InvoiceWorker([account], [profile], settings, set())

            target = worker._compute_target_dir(profile)
            self.assertTrue(target.exists())
            self.assertIn("Rechnungen2024", str(target))

    def test_build_attachment_output_path_sanitizes_profile_name(self):
        """Output filenames must stay inside the target dir even if the profile name contains separators."""
        import tempfile
        with tempfile.TemporaryDirectory() as tmp:
            account = uim.MailAccount(id="a", name="A", provider="IMAP")
            profile = uim.InvoiceProfile(
                id="p", name="Bad/Name", account_id="a", enabled=True
            )
            settings = uim.AppSettings()
            settings.download_path = tmp
            worker = uim.InvoiceWorker([account], [profile], settings, set())

            target = worker._compute_target_dir(profile)
            output_path, safe_name = worker._build_attachment_output_path(
                target, profile, "2026-05-20", "Subject", "seed"
            )

            self.assertEqual(output_path.parent, target)
            self.assertTrue(safe_name.startswith("Bad_Name_"))
            self.assertNotIn("/", safe_name)
            self.assertNotIn("\\", safe_name)

    def test_imap_png_attachment_is_converted_and_emitted(self):
        """IMAP processing accepts supported non-PDF attachments via the shared converter."""
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            account = uim.MailAccount(id="a", name="A", provider="IMAP")
            profile = uim.InvoiceProfile(
                id="p", name="TestShop", account_id="a",
                subject_filter="Rechnung",
                enabled=True,
            )
            settings = uim.AppSettings()
            settings.download_path = tmp
            worker = uim.InvoiceWorker([account], [profile], settings, set())
            worker.log = MagicMock()
            worker.log.emit = MagicMock()
            worker.invoice_found = MagicMock()
            worker.invoice_found.emit = MagicMock()

            message = EmailMessage()
            message["From"] = "shop@example.com"
            message["Subject"] = "Ihre Rechnung"
            message["Date"] = "Mon, 01 Apr 2026 10:00:00 +0000"
            message.set_content("Im Anhang befindet sich die Rechnung.")
            message.add_attachment(
                b"fake-png",
                maintype="image",
                subtype="png",
                filename="rechnung.png",
            )

            def fake_convert(file_data, source_name, output_path):
                self.assertEqual(file_data, b"fake-png")
                self.assertEqual(source_name, "rechnung.png")
                output_path.write_bytes(b"%PDF-1.4 fake")
                return True, "Bild-Anhang konvertiert"

            with patch("UniversalInvoiceMail.convert_attachment_to_pdf", side_effect=fake_convert) as mocked_convert:
                worker._process_imap_message(message, profile)

            mocked_convert.assert_called_once()
            worker.invoice_found.emit.assert_called_once()
            invoice = worker.invoice_found.emit.call_args.args[0]
            self.assertTrue(invoice.filename.endswith(".pdf"))
            self.assertTrue(Path(invoice.path).exists())

    def test_imap_body_pdf_sanitizes_profile_name(self):
        """Body-only exports must not turn profile names into nested paths."""
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            account = uim.MailAccount(id="a", name="A", provider="IMAP")
            profile = uim.InvoiceProfile(
                id="p", name="Bad/Name", account_id="a",
                enabled=True,
            )
            settings = uim.AppSettings()
            settings.download_path = tmp
            settings.download_attachments = False
            settings.convert_body_to_pdf = True
            settings.merge_body_with_attachments = False
            worker = uim.InvoiceWorker([account], [profile], settings, set())
            worker.log = MagicMock()
            worker.log.emit = MagicMock()
            worker.invoice_found = MagicMock()
            worker.invoice_found.emit = MagicMock()

            message = EmailMessage()
            message["From"] = "shop@example.com"
            message["Subject"] = "Ihre Rechnung"
            message["Date"] = "Mon, 01 Apr 2026 10:00:00 +0000"
            message.set_content(
                "<html><body>" + ("Rechnung " * 30) + "</body></html>",
                subtype="html",
            )

            def fake_html_to_pdf(html_content, output_path, mail_meta, mode="fast"):
                self.assertIn("Rechnung", html_content)
                output_path.write_bytes(b"%PDF-1.4 fake body")
                return True

            with patch("UniversalInvoiceMail.html_to_pdf", side_effect=fake_html_to_pdf):
                worker._process_imap_message(message, profile)

            worker.invoice_found.emit.assert_called_once()
            invoice = worker.invoice_found.emit.call_args.args[0]
            self.assertEqual(Path(invoice.path).parent, Path(tmp) / "Bad_Name")
            self.assertTrue(Path(invoice.path).exists())

    def test_gmail_png_attachment_is_converted_and_emitted(self):
        """Gmail payload processing accepts supported non-PDF attachments."""
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            account = uim.MailAccount(id="a", name="A", provider="Gmail API", use_gmail_api=True)
            profile = uim.InvoiceProfile(
                id="p", name="TestShop", account_id="a",
                enabled=True,
            )
            settings = uim.AppSettings()
            settings.download_path = tmp
            worker = uim.InvoiceWorker([account], [profile], settings, set())
            worker.log = MagicMock()
            worker.log.emit = MagicMock()
            worker.invoice_found = MagicMock()
            worker.invoice_found.emit = MagicMock()

            attachment_bytes = b"gmail-image"
            encoded_attachment = base64.urlsafe_b64encode(attachment_bytes).decode().rstrip("=")
            gmail_message = {
                "id": "msg-123",
                "payload": {
                    "headers": [
                        {"name": "From", "value": "shop@example.com"},
                        {"name": "Subject", "value": "Ihre Rechnung"},
                        {"name": "Date", "value": "Mon, 01 Apr 2026 10:00:00 +0000"},
                    ],
                    "parts": [
                        {
                            "filename": "rechnung.png",
                            "mimeType": "image/png",
                            "body": {"data": encoded_attachment},
                        }
                    ],
                },
            }

            def fake_convert(file_data, source_name, output_path):
                self.assertEqual(file_data, attachment_bytes)
                self.assertEqual(source_name, "rechnung.png")
                output_path.write_bytes(b"%PDF-1.4 fake")
                return True, "Bild-Anhang konvertiert"

            with patch("UniversalInvoiceMail.convert_attachment_to_pdf", side_effect=fake_convert) as mocked_convert:
                worker._process_gmail_message(MagicMock(), gmail_message, profile)

            mocked_convert.assert_called_once()
            worker.invoice_found.emit.assert_called_once()
            invoice = worker.invoice_found.emit.call_args.args[0]
            self.assertTrue(invoice.filename.endswith(".pdf"))
            self.assertTrue(Path(invoice.path).exists())

    def test_search_imap_uses_gmail_raw_when_supported(self):
        """Profiles with gmail_query should use X-GM-RAW on Gmail-capable IMAP servers."""
        account = uim.MailAccount(id="a", name="A", provider="IMAP")
        profile = uim.InvoiceProfile(
            id="p", name="TestShop", account_id="a",
            sender_filter="shop@example.com",
            gmail_query="label:finance has:attachment",
            enabled=True,
        )
        settings = uim.AppSettings(date_from="2026-05-20")
        worker = uim.InvoiceWorker([account], [profile], settings, set())
        worker.log = MagicMock()
        worker.log.emit = MagicMock()
        worker._process_imap_message = MagicMock()

        class FakeMail:
            capabilities = ("IMAP4REV1", "X-GM-EXT-1")

            def __init__(self):
                self.uid_args = None

            def uid(self, command, *args):
                if command.upper() == 'SEARCH':
                    self.uid_args = args
                    return "OK", [b"1"]
                # FETCH: eine vollständige Test-Mail zurückgeben
                msg = EmailMessage()
                msg["From"] = "shop@example.com"
                msg["Subject"] = "Ihre Rechnung"
                msg["Date"] = "Tue, 20 May 2026 10:00:00 +0000"
                msg.set_content("Test")
                return "OK", [(None, msg.as_bytes())]

        mail = FakeMail()
        worker._search_imap(mail, profile)

        # uid_args enthält die Argumente nach dem Command-Parameter
        self.assertEqual(mail.uid_args[0:2], (None, "X-GM-RAW"))
        self.assertIn("label:finance has:attachment", mail.uid_args[2])
        self.assertIn("from:shop@example.com", mail.uid_args[2])
        self.assertIn("after:2026/05/20", mail.uid_args[2])
        worker._process_imap_message.assert_called_once()

    def test_search_imap_falls_back_without_gmail_extension(self):
        """Servers without X-GM-RAW must fall back to normal IMAP search criteria."""
        account = uim.MailAccount(id="a", name="A", provider="IMAP")
        profile = uim.InvoiceProfile(
            id="p", name="TestShop", account_id="a",
            sender_filter="shop@example.com",
            subject_filter="Invoice",
            gmail_query="label:finance",
            enabled=True,
        )
        settings = uim.AppSettings(date_from="2026-05-20")
        worker = uim.InvoiceWorker([account], [profile], settings, set())
        worker.log = MagicMock()
        worker.log.emit = MagicMock()
        worker._process_imap_message = MagicMock()

        class FakeMail:
            capabilities = ("IMAP4REV1",)

            def __init__(self):
                self.uid_args = None

            def uid(self, command, *args):
                if command.upper() == 'SEARCH':
                    self.uid_args = args
                    return "OK", [b"1"]
                # FETCH: eine vollständige Test-Mail zurückgeben
                msg = EmailMessage()
                msg["From"] = "shop@example.com"
                msg["Subject"] = "Invoice"
                msg["Date"] = "Tue, 20 May 2026 10:00:00 +0000"
                msg.set_content("Test")
                return "OK", [(None, msg.as_bytes())]

        mail = FakeMail()
        worker._search_imap(mail, profile)

        # uid_args enthält die Argumente nach dem Command-Parameter
        self.assertEqual(
            mail.uid_args,
            (None, "SINCE", "20-May-2026", "FROM", '"shop@example.com"', "SUBJECT", '"Invoice"')
        )
        worker._process_imap_message.assert_called_once()


class TestImapSearchArgsBugs(unittest.TestCase):
    """Regression tests for bugs fixed in bugsweep 2026-06-11."""

    def _make_worker(self, sender_filter="", subject_filter="", date_from="", date_to=""):
        account = uim.MailAccount(id="a", name="A", provider="IMAP")
        profile = uim.InvoiceProfile(
            id="p", name="Shop", account_id="a",
            sender_filter=sender_filter,
            subject_filter=subject_filter,
            enabled=True,
        )
        settings = uim.AppSettings(date_from=date_from, date_to=date_to)
        # Disable legacy date filter so search_args contains only sender/subject tokens —
        # otherwise date_filter_months=12 prepends a dynamic SINCE that breaks exact asserts.
        settings.date_filter_months = 0
        worker = uim.InvoiceWorker([account], [profile], settings, set())
        worker.log = MagicMock()
        worker.log.emit = MagicMock()
        return worker, profile

    def test_build_imap_search_args_single_sender(self):
        """Single sender produces exactly [FROM, <quoted>] — no OR."""
        worker, profile = self._make_worker(sender_filter="amazon.de")
        args = worker._build_imap_search_args(profile)
        self.assertEqual(args, ["FROM", '"amazon.de"'], f"Unexpected args: {args}")

    def test_build_imap_search_args_two_senders_uses_imap_or(self):
        """Two senders produce exact IMAP OR token sequence: OR FROM a FROM b."""
        worker, profile = self._make_worker(sender_filter="amazon.de, amazon.com")
        args = worker._build_imap_search_args(profile)
        self.assertEqual(
            args,
            ["OR", "FROM", '"amazon.de"', "FROM", '"amazon.com"'],
            f"Unexpected args: {args}",
        )

    def test_build_imap_search_args_three_senders_uses_imap_or(self):
        """Three senders produce exact nested IMAP OR: OR FROM a OR FROM b FROM c."""
        worker, profile = self._make_worker(sender_filter="a.de, b.de, c.de")
        args = worker._build_imap_search_args(profile)
        self.assertEqual(
            args,
            ["OR", "FROM", '"a.de"', "OR", "FROM", '"b.de"', "FROM", '"c.de"'],
            f"Unexpected args: {args}",
        )

    def test_build_imap_search_args_single_subject(self):
        """Single subject filter produces exactly [SUBJECT, <quoted>]."""
        worker, profile = self._make_worker(subject_filter="Rechnung")
        args = worker._build_imap_search_args(profile)
        self.assertEqual(args, ["SUBJECT", '"Rechnung"'], f"Unexpected args: {args}")

    def test_build_imap_search_args_two_subjects_uses_imap_or(self):
        """Two subjects produce exact IMAP OR token sequence: OR SUBJECT a SUBJECT b."""
        worker, profile = self._make_worker(subject_filter="Rechnung, Invoice")
        args = worker._build_imap_search_args(profile)
        self.assertEqual(
            args,
            ["OR", "SUBJECT", '"Rechnung"', "SUBJECT", '"Invoice"'],
            f"Unexpected args: {args}",
        )

    def test_build_imap_search_args_multi_subject_no_drop(self):
        """With 3 subjects the previous silent-drop must not occur: all 3 must appear."""
        worker, profile = self._make_worker(subject_filter="Rechnung, Invoice, Beleg")
        args = worker._build_imap_search_args(profile)
        self.assertEqual(args.count("OR"), 2, f"Expected 2 OR tokens, got: {args}")
        self.assertIn('"Rechnung"', args)
        self.assertIn('"Invoice"', args)
        self.assertIn('"Beleg"', args)

    def test_build_imap_search_args_sender_and_subject_combined(self):
        """Single sender + single subject: exact token order FROM … SUBJECT …"""
        worker, profile = self._make_worker(
            sender_filter="shop@example.com", subject_filter="Rechnung"
        )
        args = worker._build_imap_search_args(profile)
        self.assertEqual(
            args,
            ["FROM", '"shop@example.com"', "SUBJECT", '"Rechnung"'],
            f"Unexpected args: {args}",
        )

    def test_build_imap_search_args_multi_sender_and_multi_subject(self):
        """Multi sender + multi subject: OR-blocks must not interfere with each other."""
        worker, profile = self._make_worker(
            sender_filter="a.de, b.de", subject_filter="R, I"
        )
        args = worker._build_imap_search_args(profile)
        self.assertEqual(
            args,
            [
                "OR", "FROM", '"a.de"', "FROM", '"b.de"',
                "OR", "SUBJECT", '"R"', "SUBJECT", '"I"',
            ],
            f"Unexpected args: {args}",
        )


class TestOnInvoiceFoundPersistence(unittest.TestCase):
    """Regression test: on_invoice_found must call save_invoices_db, not save_config."""

    def test_on_invoice_found_calls_save_invoices_db_not_save_config(self):
        """Each found invoice must trigger save_invoices_db only, not the heavier save_config."""
        save_config_calls = []
        save_db_calls = []

        window = MagicMock()
        window.invoices = []
        window.save_config = lambda: save_config_calls.append(1)
        window.save_invoices_db = lambda: save_db_calls.append(1)

        invoice = uim.Invoice(
            id="inv1", profile_name="Shop", filename="test.pdf",
            date="2026-06-11", path="/tmp/test.pdf",
            sender="shop@example.com", subject="Rechnung", hash="abc123",
        )

        uim.MainWindow.on_invoice_found(window, invoice)

        self.assertEqual(len(save_db_calls), 1, "save_invoices_db must be called once")
        self.assertEqual(len(save_config_calls), 0, "save_config must NOT be called")
        self.assertIn(invoice, window.invoices)


if __name__ == "__main__":
    unittest.main()
