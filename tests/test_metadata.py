# -*- coding: utf-8 -*-
"""Metadata, Manifest, and Documentation Parity Tests for UniversalInvoiceMail."""
import json
import re
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent


def test_version_parity():
    """Verify that all version strings across package definitions match exactly."""
    pyproject_path = REPO_ROOT / "pyproject.toml"
    assert pyproject_path.exists(), "pyproject.toml missing"
    pyproject_content = pyproject_path.read_text(encoding="utf-8")
    pyproject_ver_match = re.search(r'version\s*=\s*"([^"]+)"', pyproject_content)
    assert pyproject_ver_match, "version not found in pyproject.toml"
    version = pyproject_ver_match.group(1)

    # Check UniversalInvoiceMail.py header docstring
    app_path = REPO_ROOT / "UniversalInvoiceMail.py"
    assert app_path.exists(), "UniversalInvoiceMail.py missing"
    app_content = app_path.read_text(encoding="utf-8")
    assert f"UniversalInvoiceMail V{version}" in app_content or f"UniversalInvoiceMail v{version}" in app_content, (
        f"Version V{version} missing from UniversalInvoiceMail.py docstring"
    )

    # Check CHANGELOG.md
    changelog_path = REPO_ROOT / "CHANGELOG.md"
    assert changelog_path.exists(), "CHANGELOG.md missing"
    changelog_content = changelog_path.read_text(encoding="utf-8")
    assert f"[{version}]" in changelog_content, (
        f"Release [{version}] header missing from CHANGELOG.md"
    )


def test_core_documentation_files():
    """Verify presence and non-emptiness of core documentation and policy files."""
    required_files = [
        "README.md",
        "README-DE.md",
        "llms.txt",
        "LICENSE",
        "CHANGELOG.md",
        "ROADMAP.txt",
        "pyproject.toml",
        "EXPORTFORMAT.md",
        "USER_GUIDE.md",
        "translator.py",
        "locales/translations.json",
        "web_companion/package.json",
        "web_companion/index.html",
        "web_companion/manifest.webmanifest",
    ]
    for rel_path in required_files:
        file_path = REPO_ROOT / rel_path
        assert file_path.exists(), f"Required file {rel_path} does not exist"
        assert file_path.stat().st_size > 0, f"Required file {rel_path} is empty"


def test_llms_txt_structure():
    """Verify that llms.txt provides complete RAG context and valid metadata."""
    llms_path = REPO_ROOT / "llms.txt"
    content = llms_path.read_text(encoding="utf-8")

    assert "# UniversalInvoiceMail" in content
    assert "Last-checked:" in content
    assert "Verification:" in content
    assert "Safety boundary:" in content
    assert "DATEV" in content
    assert "https://github.com/doc-bricks/UniversalInvoiceMail" in content


def test_web_companion_pwa_assets():
    """Verify that web companion files and icons are complete and valid JSON/HTML."""
    web_dir = REPO_ROOT / "web_companion"
    pkg_path = web_dir / "package.json"
    manifest_path = web_dir / "manifest.webmanifest"

    pkg_data = json.loads(pkg_path.read_text(encoding="utf-8"))
    assert "name" in pkg_data
    assert "scripts" in pkg_data and "test" in pkg_data["scripts"]

    manifest_data = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest_data.get("name") == "UniversalInvoiceMail Companion"
    assert "icons" in manifest_data
    assert len(manifest_data["icons"]) >= 2

    # Check icons exist
    for icon_entry in manifest_data["icons"]:
        src = icon_entry["src"].lstrip("./")
        icon_file = web_dir / src
        assert icon_file.exists(), f"Companion icon missing: {src}"


def test_utf8_encoding_and_no_mojibake():
    """Verify that all markdown and python source files are clean UTF-8 without mojibake."""
    suspect_patterns = ["\u00c3\u00a4", "\u00c3\u00bc", "\u00c3\u00b6", "\u00c3\u009f", "\ufffd"]
    check_exts = {".py", ".md", ".txt", ".json", ".toml"}

    for p in REPO_ROOT.rglob("*"):
        if any(part.startswith(".") or part in ("node_modules", "__pycache__") for part in p.parts):
            continue
        if p.is_file() and p.suffix in check_exts:
            try:
                content = p.read_text(encoding="utf-8")
            except UnicodeDecodeError:
                assert False, f"File {p} failed UTF-8 decoding"
            for pat in suspect_patterns:
                assert pat not in content, f"Mojibake pattern '{pat}' found in {p.relative_to(REPO_ROOT)}"


def test_datev_validation_guidance_matches_the_dialog_contract():
    """User-facing DATEV guidance must not describe the pre-save validation as deferred."""
    german_readme = (REPO_ROOT / "README-DE.md").read_text(encoding="utf-8")
    english_readme = (REPO_ROOT / "README.md").read_text(encoding="utf-8")
    user_guide = (REPO_ROOT / "USER_GUIDE.md").read_text(encoding="utf-8")
    roadmap = (REPO_ROOT / "ROADMAP.txt").read_text(encoding="utf-8")

    assert "Der Dialog prüft vor dem Speichern" in german_readme
    assert "The settings dialog validates" in english_readme
    assert "Der Dialog prüft beim Speichern" in user_guide
    assert "case-insensitive uniqueness" in english_readme
    assert "Groß-/Kleinschreibung" in german_readme
    assert "Rand-Leerzeichen" in user_guide
    assert "Formale Kontenbereichs- sowie Duplikat-/Konfliktregeln bleiben" not in german_readme
    assert "Formal account-range and duplicate/conflict validation is intentionally deferred" not in english_readme
    assert "Die formale Prüfung erlaubter Kontenbereiche" not in user_guide
    assert "154/154" in roadmap
    assert "10/10" in roadmap
    assert "Mapping-Tabelle bleibt als technische" in roadmap
    assert "Automatische oder fachlich verbindliche Kontierung" in roadmap
    assert "93-Spalten-Exportvertrag bleibt" in roadmap
    assert "TASKPLAN ist die kanonische" in roadmap
