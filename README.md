# Vyaprix — Desktop Invoicing and Business Records

Vyaprix is a Python desktop application from EllowLabs intended to help a small business manage invoices and related records from a local desktop interface.

## Main capabilities

Based on the current application source, workflows include:

- Invoice creation and management
- Client and catalog-item records
- Payment tracking and invoice balance/status calculations
- GST-oriented invoice calculations and printable PDF output
- Local SQLite persistence
- Database backup and import workflows
- Custom business identity, contact details, and invoice branding
- Optional barcode, image, and document utilities

Confirm the available behavior in the current app build before relying on it for financial or tax compliance. Vyaprix is a record-keeping tool, not accounting or legal advice.

## Technology

- Python desktop GUI using Tkinter and ttkbootstrap
- SQLite for local application data
- ReportLab for PDF generation
- PyMuPDF for PDF/document support
- Pillow for image handling
- Optional barcode and print integrations

The source currently uses Python's native `lazy import` syntax and is marked for Python 3.15. Earlier Python interpreters may not be able to parse the file. Use a compatible Python version and verify dependency support before installation.

## Setup

1. Install a Python release that supports the syntax used by the current source.
2. Install the required third-party packages in an isolated environment.
3. Run the application from a writable folder so it can create its local database and configuration.

Example environment setup (adapt the interpreter command to your installed Python version):

```powershell
py -3.15 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install ttkbootstrap reportlab PyMuPDF Pillow
python Vyaprix.py
```

Python 3.15 may not be installed on every system. If the command is unavailable, consult the Python release available for your platform and the source compatibility requirements before proceeding. Barcode-related features may require the optional `python-barcode` package and its image support.

## Data safety

- The application stores business records locally; back up its data folder before upgrading, importing a database, or changing machines.
- Test database import on a copy before replacing your working records.
- Use synthetic clients and invoices for demonstrations and screenshots.
- Keep invoice PDFs, exported databases, and business configuration files private.
- Verify tax rates, GSTINs, business identity, and totals before issuing invoices.
- Review bundled fonts, logos, and other assets against their own licensing terms.

## Troubleshooting

- If startup fails with a syntax error around `lazy import`, the interpreter is too old for the current source syntax.
- If PDF generation fails, verify ReportLab and PyMuPDF installation and the writable output location.
- If a local database cannot be opened, close other instances and check folder permissions before attempting recovery.
- Keep a separate backup before running any database migration or import.

## Contributions and licensing

This repository currently does not grant an open-source license for Vyaprix or its bundled branding/assets. See [LICENSE](LICENSE) before copying or redistributing it. Third-party libraries remain under their respective licenses. To propose a bug fix or feature, contact the maintainer before submitting a substantial patch.
