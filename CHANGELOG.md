# Changelog - UniversalInvoiceMail

Alle wichtigen Änderungen an diesem Projekt werden in dieser Datei dokumentiert.

## [2.4.1] - 2026-07-27 - Technical Hygiene & Maintenance

### Changed
- `llms.txt`: Header auf `Last-checked: 2026-07-27` und 125 verifizierte Tests (115 Pytest + 10 Node Web Companion PWA) aktualisiert.
- Technischen Hygiene- & Maintenance-Check (Pfad A) in den internen Wartungsprotokollen registriert.
- Test-Verifikation: 115/115 Pytest-Tests 100% grün (0 Fehler, 24.08s execution time), `py_compile` fehlerfrei.

## [2.4.0] - 2026-07-26

### Added
- Standardisierte `pyproject.toml` (PEP 621) mit Metadaten, Keywords, Klassifikatoren und `[tool.pytest.ini_options]` (`pythonpath = ["."]`).
- Shields.io-Badges für Pytest (110 passed), Web Companion (10 passed), Lizenz (MIT), Local-First-Datenschutz und LLM-Ready-Kontext.
- Mermaid-Systemarchitekturdiagramm & Datenfluss-Visualisierung in `README.md` und `README-DE.md`.
- GFM-KI-Agenten-Hinweis (`> [!NOTE]`) für LLM-Verarbeitbarkeit und sicheres Auffinden von `llms.txt`.
- Sprachwechsler-Leiste (`[English](README.md) | [Deutsch](README-DE.md)`).

### Changed
- `llms.txt` Header-Datum auf `Last-checked: 2026-07-26` aktualisiert und Verifikationsteststand auf 120 grüne Tests (110 Pytest + 10 Node Web Companion) angeglichen.

## [Unreleased]

### DATEV mapping contract and roadmap readback (2026-08-26)
- Prevented empty mapping keys, non-numeric account values, empty adviser/client
  numbers, and case-insensitive duplicate sender/keyword keys from being silently
  discarded or replaced before `DATEVSettingsDialog` validation.
- Kept the mapping table as a user-configured technical aid; professional account
  assignment remains with accounting or tax advisers, and the 93-column export
  contract is unchanged.
- Synchronized `ROADMAP.txt` with the verified v2.3.0/Companion baseline of
  154 Pytest and 10 Node tests and identified TASKPLAN as the canonical task status.

### UX: Accurate DATEV validation guidance (2026-08-25)
- Corrected the German and English user-facing DATEV documentation: the settings dialog
  validates its technical inputs before saving and reports errors; it does not silently fall
  back to default account values. The guide now distinguishes this check from an accountant's
  professional review of the account assignment.

### Bugfix & Hardening: DATEV CSV Quoting, Multi-Format Date Parsing & Dynamic Fiscal Year (2026-08-25)
- **DATEV EXTF CSV Escaping & Delimiter Injection Protection (`datev_exporter.py` / `DATEVBuchung`)**:
  - Replaced naive string concatenation with standard `csv.writer(..., delimiter=";", quoting=csv.QUOTE_MINIMAL)` in `DATEVExporter.export()`, preventing CSV column shifts and parser failures when invoice filenames (`belegfeld1`), provider names, or descriptions (`buchungstext`) contain semicolons `;` or quotation marks `"`.
  - Added whitespace and control character sanitation (stripping `\r`, `\n`, `\t` and boundary whitespace) in `DATEVBuchung.to_row()`.
  - Guaranteed invariant: Exported booking rows always parse to exactly 93 EXTF columns under RFC 4180 / DATEV CSV specifications.
- **Robust Multi-Format Date Parsing (`datev_exporter.py` / `parse_datev_datetime`)**:
  - Implemented unified `parse_datev_datetime()` supporting ISO dates, timestamps (`YYYY-MM-DDTHH:MM:SS`, `YYYY-MM-DD HH:MM:SS`), German dates (`DD.MM.YYYY`, `DD-MM-YYYY`), dotted dates (`YYYY.MM.DD`), and compact formats (`YYYYMMDD`).
  - Unified date parsing across `validate_invoices_for_export()` and `DATEVExporter` to prevent false invalid date fallback warnings.
  - Enabled dynamic `wj_beginn` fiscal year alignment in `DATEVExporter._build_header()` when exporting historical invoices across multiple tax years.
- **Provider Resolution Fallback (`UniversalInvoiceMail.py`)**:
  - Added fallback from `inv.profile_name` to `inv.sender` in `_export_datev()` to preserve provider-based SKR03/SKR04 account mapping when profile names are generic.
- **Test Coverage Expansion**:
  - Added comprehensive unit and regression tests in `tests/test_datev.py` and `tests/test_datev_validation.py` (3 new tests, 153/153 Pytest passed, 10/10 Node Web Companion passed, 100% green).

### UX & Accessibility: Manual Amount Validation (2026-08-23)
- Ungültige manuelle Betragseingaben werden nicht mehr stillschweigend verworfen: Die Tabelle stellt den zuletzt gespeicherten Betrag wieder her und erläutert den Fehler in einem zugänglichen Warnhinweis sowie im Aktivitätsprotokoll.
- Neuer Offscreen-Regressionstest stellt sicher, dass ungültige Eingaben keinen gespeicherten Betrag überschreiben.

### Bugfix & Hardening: Amount Normalization & Companion Exchange Bundle (2026-08-23)
- **Amount Normalization & Parsing Hardening (`invoice_bundle.py` / `UniversalInvoiceMail.py`)**:
  - Hardened `_normalize_amount()` against formatted European numbers with thousand-separators (`1.234,56`), US format (`1,234.56`), currency symbols (`€`, `$`, `£`, `¥`, `₹`), ISO currency codes (`EUR`, `USD`, `CHF`, `GBP`), whitespace-only strings, and negative credit amounts (`-15,50`).
  - Switched `build_invoice_bundle()` to use `_normalize_amount()` instead of raw `float(amount)` to prevent unhandled `ValueError` crashes during bundle export.
  - Added safe integer fallback parsing `_safe_int()` for Sachkontenlänge in DATEV bundle configurations.
  - Preserved fallback `profile_id` on invoices when profile collections do not contain matching profile name/id entries.
  - Added support for `collections.abc.MutableMapping` (e.g. `UserDict`) in `apply_invoice_bundle_changes()`.
  - Hardened GUI amount input handler `_on_invoice_amount_changed()` in `UniversalInvoiceMail.py` using `_normalize_amount()`.
  - Added comprehensive regression test suite in `tests/test_invoice_bundle.py` (3 new tests, 149/149 passed).

### DATEV Validation Hardening & Multi-Language I18N System (2026-08-21)
- **DATEV Account Validation & Mapping Hardening (TW-UIM-04 / TASKPLAN #1156)**:
  - Added `validate_account_number()` verifying numerical validity, positive values, and standard 4-to-8 digit ranges for SKR03/SKR04 accounts.
  - Added `validate_datev_config()` validating Beraternummer (1-7 digits), Mandantennummer (1-5 digits), Sachkontenlänge (4-8), and non-empty mappings.
  - Added `validate_invoices_for_export()` and `DATEVExporter.validate()` producing structured `DATEVValidationReport` batch diagnostics (valid counts, skipped zero-amount items, unparseable dates).
  - Enhanced `DATEVSettingsDialog` in `UniversalInvoiceMail.py` with pre-save input validation in `accept()`, providing clear user-facing guidance on configuration errors.
  - Added test suite `tests/test_datev_validation.py` with 10 contract and unit tests (100% green).
- **I18N / Multi-Language System (Policy P-006 / 6-Languages Support)**:
  - Implemented `translator.py` (`TranslationSystem`) supporting 6 standard languages (`de`, `en`, `es`, `zh`, `ja`, `ru`) with fallback chain (current -> en -> de -> key).
  - Created `locales/translations.json` with complete translations for core actions, statuses, table headers, DATEV dialogs, validation messages, and mail filters.
  - Added test suite `tests/test_i18n.py` with 16 tests verifying completeness, token consistency, formatting, and language switching across all 6 languages.
- **Metadata, Linting & Parity**:
  - Updated `tests/test_metadata.py` with `translator.py` and `locales/translations.json`.
  - Pytest test suite expanded to 146 tests (146 passed in 2.74s, 100% green).
  - Node Web Companion test suite maintained at 10 passed (156 total passed tests).
  - Clean `ruff check .` (0 errors) and `python -m compileall .` (0 errors).

### Maintainer verification, Hygiene & Discoverability (2026-08-16)
- **Ruff Linting**: Resolved all 34 pre-existing Ruff linting errors (`F401` unused imports, `E402` module-level imports, `F841` unused variables) across `test_helpers.py`, `tests/test_datev.py`, and `tests/test_integration.py`. Added `[tool.ruff]` and `[tool.ruff.lint]` configuration in `pyproject.toml`.
- **Automated Metadata Parity Tests**: Added `tests/test_metadata.py` verifying version parity, documentation file presence, `llms.txt` integrity, web companion PWA assets and UTF-8 encoding.
- **Discoverability & Badges**: Synchronized test badges to 120 Pytest passed (100% green) and 10 Node Web Companion passed across `README.md` and `README-DE.md`. Added doc-bricks and open-bricks ecosystem sibling tool grid.
- **Testsuite Status**: 120/120 Pytest tests passing, 10/10 Node Web Companion tests passing, source-platform smoke passing, `ruff check .` 100% clean.

### Repo hygiene (2026-08-15)
- Internal `TASKPLAN_STATUS_*.md` readbacks are now ignored and no longer intended for public Git tracking because they can contain local Plan-D evidence paths.

### TASKPLAN Steuerdokumente und DATEV-Readback (2026-08-12)
- ROADMAP, AUFGABEN, README/README-DE, User Guide und `llms.txt` gegen den v2.3.0-/Companion-Stand sowie den frischen Plan-D-Readback abgeglichen.
- `115/115` Pytest-Tests, Source-Platform-Smoke und `compileall` liefen grün. Die getrackte Node-Baseline bleibt `10/10`; eine fremde, uncommittete Manifest-Variante liefert `9/10` und wurde nicht übernommen.
- Die DATEV-Mapping-UI (Tabelle, Hinzufügen/Entfernen, Standard-Wiederherstellung, Laden/Speichern und Accessibility-Metadaten) ist dokumentiert. Kontenbereichs- und Duplikat-/Konfliktregeln bleiben bis zur Accounting-Entscheidung offen; der 93-Spalten-Exportvertrag wurde nicht verändert.
- Keine kanonische OneDrive-Projektion: aktiver Cloud-Lock und fremde Arbeitsbaumänderungen wurden nicht überschrieben.

### Maintainer verification (2026-08-10)
- Fresh local readback: 114/114 Pytest, source-platform smoke, `py_compile` and
  JavaScript syntax checks passed. Ruff reports 34 existing test/import findings.
- The Node Web Companion suite is 9/10 only because the pre-existing uncommitted
  `web_companion/manifest.webmanifest` change points to lowercase/root-relative
  icons while the tracked assets use `./icons/Icon-*.png`. The foreign manifest
  change was preserved and not staged or committed; the public 10/10 baseline was
  not rewritten from this dirty working tree.

### Maintainer verification & Discoverability (2026-08-04)
- Discoverability, README-Design & SEO Check (Pfad B) für `doc-bricks/UniversalInvoiceMail` durchgeführt.
- Shields.io-Badges für doc-bricks Organisation, open-bricks Ökosystem, Pytest (114 passed), Web Companion (10 passed), Python 3.10+ und MIT-Lizenz in `README.md` & `README-DE.md` integriert.
- Interaktive Mermaid-Systemarchitekturdiagramme für Datenfluss (IMAP/Gmail API -> Conversion/OCR -> Archive/DATEV/Web Companion) in deutscher und englischer Dokumentation hinterlegt.
- GFM-KI-Agenten-Callout-Box (`> [!NOTE]`) für `llms.txt` Discovery Index eingebunden; `llms.txt` Timestamp auf `2026-08-04` und 124 passed Tests (114 Pytest + 10 Node Web Companion) aktualisiert.
- PWA-Manifest `manifest.webmanifest` Icon-Pfade für case-sensitive und offline PWA Installationen gehärtet (`./icons/Icon-192.png`, etc.); Node test suite (10/10 passed) und Pytest test suite (114/114 passed) 100% grün.

### Maintainer verification (2026-08-01)
- 114/114 Pytest-Tests, 10/10 Node-Web-Companion-PWA-Tests, `py_compile` und der
  Source-Platform-Smoke erfolgreich verifiziert. Der Plattform-Smoke meldete nur
  einen nicht-fatalen Qt-Font-Hinweis; der echte Android-/iOS-Geräte-Signoff bleibt offen.

### Fixed
- PWA-Manifest und Companion-HTML referenzieren wieder ausschließlich die versionierten Icons unter `web_companion/icons/`; dadurch bleibt die Offline-Installation auf case-sensitiven Hosts funktionsfähig und der mobile PWA-Smoke-Test grün.
- IMAP MSN→UID (kritisch): `_search_imap` verwendet jetzt `uid('search')` und `uid('fetch')` statt `search()`/`fetch()`. MSN-Nummern sind instabil wenn andere Clients gleichzeitig Mails verschieben/löschen; UIDs sind stabile Kennungen gemäß RFC 3501 §2.3.1.1.
- IMAP NIL-Guard: `uid('fetch')` kann bei nicht mehr existierenden UIDs eine leere/fehlerhafte Antwortstruktur zurückgeben; Guard verhindert AttributeError auf `msg_data[0][1]`.
- MIME-Charset: `_get_imap_message_body` liest den Charset aus dem Content-Type-Header (`get_content_charset`) statt blind UTF-8 anzunehmen; verhindert Mojibake bei ISO-8859-1/windows-1252-Mails.
- PDF-HTML-Sanitizer entfernen `script`-/`style`-Blöcke jetzt parserbasiert,
  sodass auch Varianten wie `</script >` zuverlässig gefiltert werden.
- IMAP multi-subject OR: Wenn 2+ Betreff-Filter konfiguriert waren, wurden Betreff-Einträge nach dem ersten stillschweigend verworfen; es wurde kein OR-Ausdruck aufgebaut, sodass nur Nachrichten mit dem ersten Betreff gefunden wurden.
- AccountDialog: `use_gmail_api` wurde beim Bearbeiten eines Gmail-Kontos (use_gmail_api=False) durch `on_provider_changed()` auf True zurückgesetzt; der gespeicherte Wert wird jetzt nach dem Provider-Lookup wiederhergestellt.
- MailAccount.from_dict: Unbekannte Schlüssel wurden stillschweigend verworfen; jetzt werden neue Felder toleriert (Vorwärtskompatibilität).
- IMAP multi-sender OR: Für 2+ Absender-Filter wurde die korrekte verschachtelte OR-FROM-Kette aufgebaut; zuvor wurden Absender mit AND verknüpft, sodass keine Nachricht passte.
- on_invoice_found nutzt jetzt save_invoices_db statt save_config (kein vollständiges Rewrite der Config bei jeder gefundenen Rechnung).
- start_grabbing: redundanter log_output.clear()-Aufruf entfernt, der Sync-Status-Meldungen löschte, bevor der Worker-Thread startete.

### Changed
- Der DATEV-Einstellungsdialog erläutert seine Konten-Mapping-Felder, Tabellenaktionen und Speichern-/Abbrechen-Aktionen jetzt zusätzlich per Accessible Description und Tooltip; das kompakte Layout bleibt unverändert.
- README, README-DE und `llms.txt` mit Startpunkten, local-first Invoice-Archive-/Gmail-/IMAP-/DATEV-Suchkontext und klarer Privacy-Abgrenzung geschärft.
- Interne Wartungsdateien (`CHECKS-LOG*.txt`, `LOCK*.txt`) sind jetzt gitignored; das Repo führt stattdessen nur veröffentlichbare Projektdateien.
- `EXPORTFORMAT.md` und `AUFGABEN.txt` auf den realen Bundle-Export/-Import-Stand gehoben; Companion-Rückfluss ist jetzt klar auf Betrag, Prüfflag und Notiz begrenzt.
- Die kompakte Rechnungs-Aktionsleiste exponiert jetzt klare Accessible Names, Descriptions und Tooltips für Auswahl-, Export-, Bundle- und DATEV-Aktionen, ohne die UI sichtbar aufzublähen.

### Added
- macOS/Linux platform smoke `tests/source_platform_smoke.py` (renamed from `tests/linux_platform_smoke.py` via `git mv`, history preserved) for offscreen PySide6 start, missing-keyring fallback, LibreOffice SOFFICE_PATH env-override detection and CSV export.
- GitHub Actions workflow `.github/workflows/source-platform-smoke.yml` on `ubuntu-latest` + `macos-latest`; installs PySide6 only (avoids pywin32/google-auth build failures on non-Windows).
- Neues Hilfsmodul `invoice_bundle.py` für redigierten Bundle-Export/-Import samt UI-Aktionen `Bundle Export` und `Bundle Import`.
- Neue Regressionstests `tests/test_invoice_bundle.py` für Exportvertrag, Hash-Konflikte und UI-Roundtrip.
- Committebare Web-Companion-PWA-Ressourcen (`favicon.ico`, `favicon.png`, `apple-touch-icon-180.png`) für lokale Installierbarkeit ohne tote Repo-Referenzen.

### CI
- Source-platform smoke workflow now uses verified `actions/checkout@v6` and `actions/setup-python@v6`, matching the main test workflow, and forces UTF-8 Python output.

### Fixed
- HTML-Injection in PDF-Covern: Mail-Metadaten (Datum, Betreff, Absender) werden nun mit `html.escape()` gesichert, bevor sie in xhtml2pdf/Selenium-HTML eingebettet werden.
- HTML-Injection bei OCR-Ergebnissen: OCR-Text in `<pre>`-Tags wird mit `html.escape()` gesichert.
- HTML-Injection in EML/MSG-Fallback: Plain-Text-Körper aus EML- und MSG-Dateien werden vor dem Einbetten in `<pre>`-Tags escaped.
- HTML-Injection in Gmail-Body: `_get_message_body()` escaped Plain-Text-Fallback jetzt mit `html.escape()`.
- HTML-Injection in IMAP-Merge-Pfad: `_process_imap_message()` escaped den Body beim Zusammenführen mit PDF-Anhängen.
- Ressourcen-Leak in `_pdf_to_images()`: `pdfium.PdfDocument.close()` wird jetzt per `try/finally` auch bei Rendering-Exceptions aufgerufen.
- Ressourcen-Leak in `_convert_msg_to_pdf()`: `extract_msg.Message.close()` wird jetzt per `try/finally` auch bei pisa-Exceptions aufgerufen.
- Temp-Datei in `add_text_layer()` wird bei Fehlern bereinigt: `temp_path` wird jetzt vor dem `try`-Block deklariert, damit der `except`-Handler sie per `unlink(missing_ok=True)` löschen kann.
- Variablen-Shadowing in `_process_gmail_message()` und `_process_imap_message()`: `success, msg = ocr.enhance_with_ocr(...)` überschrieb den `msg`-Parameter (E-Mail-Objekt); umbenannt zu `ocr_msg`.
- Regex-Backreference-Bug in `BrowserPDFRenderer.render_html_to_pdf()`: Ein Absendername mit `\1` (z. B. `CORP\1user`) wurde von `re.sub()` als Backreferenz interpretiert und duplizierte den `<body>`-Tag; Ersatz durch Lambda-Funktion behoben.
- Import-Crash ohne Gmail-Pakete: Die Rückgabe-Annotation `-> Optional[Credentials]` in `_get_gmail_credentials()` wurde eager ausgewertet; wenn Gmail-Pakete fehlen, ist `Credentials` undefiniert und das gesamte Modul schlägt beim Import fehl. Annotation auf `-> "Optional[Credentials]"` (String, lazy) umgestellt.
- Gmail-Datumsfilter-Inkonsistenz in `_build_gmail_search_query()`: Der `date_filter_months`-Fallback wurde ausgelöst wenn nur `date_to` gesetzt war (ohne `date_from`), sodass fälschlicherweise eine `after:`-Schranke eingefügt wurde; IMAP-Pendant prüft korrekt `if not search_args`. Bedingung auf `not date_from and not date_to` korrigiert.
- Windows-File-Lock in `enhance_with_ocr()`: `PdfReader(ocr_pdf_path)` hielt die `ocr_page.pdf`-Datei nach der Pages-Schleife offen; `unlink()` schlug auf Windows mit `PermissionError` fehl, OCR gab `False` zurück und hinterließ Temp-Dateien. `del ocr_reader` (und `del original_reader`) nach den jeweiligen Pages-Schleifen hinzugefügt, damit CPythons Refcounting die File-Handles sofort freigibt.
- Startup-Crash bei korrupter Konfiguration: `load_config()` fing `TypeError` nicht ab; fehlende Pflichtfelder in `Invoice` oder `InvoiceProfile` (z. B. nach Sync-Fehler oder manueller Bearbeitung der JSON-Dateien) ließen `cls(**filtered)` mit `TypeError` fehlschlagen, der aus dem Konstruktor propagierte. `TypeError` zu beiden `except`-Klauseln in `load_config()` hinzugefügt.

### Changed
- Porting status: macOS and Linux source smoke unified under `source_platform_smoke.py`; both platforms covered by CI.
- DATEV-Header nutzt jetzt dieselbe Datumslogik wie die Buchungszeilen, damit auch `TT/MM/JJJJ` das korrekte Exportintervall setzt.
### Added
- Gmail Query Builder im Profil-Dialog ergänzt; optionale Raw Queries können jetzt ohne manuelle Syntaxpflege vorbereitet werden
- GitHub-Actions-Testworkflow für Python 3.10, 3.11 und 3.12 ergänzt
- `llms.txt` als maschinenlesbarer Projektkontext ergänzt

### Changed
- Gmail-Suchen kombinieren gespeicherte Raw Queries jetzt mit Sender-, Betreff- und Datumsfiltern
- IMAP nutzt bei Gmail-kompatiblen Servern `X-GM-RAW` und fällt sonst sauber auf normale IMAP-Kriterien zurück

### Verified
- DATEV-Export als bereits vorhandene Migration gegen Code, Dialog, Doku und Regressionstests nachgezogen; `AUFGABEN.txt` entsprechend korrigiert
- Lokaler Teststand auf 104 grüne Tests aktualisiert

## [2.3.0] - 2026-05-02
### Added
- DATEV-Export: Rechnungen als DATEV-Buchungsstapel (CSV, cp1252) exportieren
- Invoice.amount Feld: Rechnungsbetrag direkt in der Tabelle editierbar
- DATEVSettingsDialog: Berater-Nr. und Mandant-Nr. konfigurierbar
