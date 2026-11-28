
# Vyaprix — Desktop Invoicing & Business Management

**Vyaprix by EllowLabs** is a Python-based desktop application designed to simplify invoicing, customer management, payment tracking, and business record-keeping.

Built with **PySide6**, Vyaprix provides a modern desktop experience with local SQLite storage, PDF generation, and essential business management tools.

## ✨ Key Features

- **Invoice Management** — Create and manage professional invoices.
- **Customer Management** — Maintain customer profiles and records.
- **Product & Service Catalog** — Organize items for invoicing.
- **Payment Tracking** — Monitor payments and outstanding balances.
- **GST Support** — Calculate GST and tax breakdowns.
- **PDF Generation** — Generate, print, and manage invoice PDFs.
- **Business Customization** — Configure business details, logos, and branding.
- **Local Database** — Store records using SQLite.
- **Backup & Restore** — Manage database backups and restoration.
- **Additional Utilities** — Barcode generation, PDF merging, and document tools.

## 🛠️ Technology Stack

| Component | Technology |
| --- | --- |
| Language | Python |
| GUI | PySide6 (Qt) |
| Database | SQLite3 |
| PDF Generation | ReportLab |
| PDF Processing | pypdf |
| Barcode | python-barcode |
| Image Processing | Pillow |
| Data Formats | JSON, CSV |

## 💻 Requirements

- Windows, Linux, or macOS (subject to compatibility).
- Compatible Python installation.
- Required dependencies and appropriate file permissions.

## 🚀 Installation

**1. Clone the repository**

```bash
git clone YOUR_REPOSITORY_URL
cd Vyaprix
```

**2. Create and activate a virtual environment**

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

**3. Install dependencies**

```bash
python -m pip install --upgrade pip
pip install PySide6 reportlab pypdf python-barcode Pillow
```

**4. Run the application**

```bash
python Vyaprix.py
```

Replace `Vyaprix.py` with the actual entry-point filename if different.

## 🔐 Data & Security

- Back up the SQLite database regularly.
- Keep copies of important invoices and business records.
- Protect customer information and configuration files.
- Verify invoice details, GST calculations, and tax information before issuing invoices.

## 📦 Project Status

**PySide6 Edition — Preview Build**

Vyaprix has migrated from Tkinter to PySide6. Converted screens, dialogs, and interactive components are included. Final cleanup and verification are still in progress.

Some features and integrations may require further testing.

## ⚠️ Disclaimer

Vyaprix is an invoicing and business record-keeping utility, not a replacement for professional accounting, taxation, or legal advice. Users are responsible for verifying invoice accuracy and applicable tax compliance.

## 📄 Licensing & Credits

**Developed by EllowLabs.**

Third-party dependencies retain their respective licenses. Refer to the repository's LICENSE file, if available, for applicable project licensing terms.

---

**Vyaprix by EllowLabs**

*Simplifying Business. Streamlining Invoicing. Organizing Records.*
