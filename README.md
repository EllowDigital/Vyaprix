# Vyaprix — Desktop Invoicing & Business Management System

**Vyaprix by EllowLabs** is a Python-based desktop application designed to simplify invoice generation, customer management, payment tracking, and everyday business record-keeping.

Built with a modern desktop interface, Vyaprix aims to provide small businesses with a practical, reliable, and locally operated solution for managing essential business activities without depending on a constant internet connection.

## ✨ Key Features

Vyaprix brings essential business operations together in one desktop application.

- **Invoice Management** — Create, manage, and maintain professional customer invoices.
- **Customer Management** — Maintain customer profiles, contact information, and related records.
- **Product & Service Catalog** — Organize catalog items for convenient invoice creation.
- **Payment Tracking** — Record payments and monitor outstanding balances.
- **Invoice Status Management** — Calculate invoice totals, paid amounts, and pending balances.
- **GST Calculations** — Support GST-oriented invoice calculations and tax breakdowns.
- **PDF Generation** — Generate professionally formatted, printable PDF invoices.
- **Business Customization** — Configure business identity, contact details, logos, and invoice branding.
- **Local Database** — Store application records using SQLite for local data management.
- **Database Backup & Import** — Support database backup and restoration workflows.
- **Barcode Utilities** — Optional barcode generation and image-based barcode support.
- **Document Utilities** — PDF merging and other document-related operations.
- **Desktop Experience** — Native desktop interface with dialogs, tables, shortcuts, and interactive components.

## 🛠️ Technology Stack

Vyaprix uses Python and established libraries to provide its desktop interface, local data management, and document-generation capabilities.

| Component | Technology |
| --- | --- |
| Programming Language | Python |
| Desktop GUI Framework | PySide6 (Qt) |
| Database | SQLite3 |
| PDF Generation | ReportLab |
| PDF Merging | pypdf |
| Barcode Generation | python-barcode |
| Image Processing | Pillow |
| Data Serialization | JSON, CSV |
| Application Logging | Python Logging |
| File & System Operations | Python Standard Library |

### Core Dependencies

- PySide6
- ReportLab
- pypdf
- python-barcode
- Pillow

Additional functionality may depend on optional integrations and utilities.

## 💻 System Requirements

- Windows, Linux, or macOS (subject to platform compatibility).
- A compatible Python installation.
- Sufficient disk space for application data, generated invoices, and backups.
- Write permissions for the application data and configuration directories.
- Required Python dependencies installed in the active environment.

## 🚀 Installation & Setup

### 1. Clone the Repository

```bash
git clone YOUR_REPOSITORY_URL
cd Vyaprix
```

### 2. Create a Virtual Environment

```powershell
python -m venv .venv
```

### 3. Activate the Environment

```powershell
.venv\Scripts\Activate.ps1
```

### 4. Upgrade pip

```bash
python -m pip install --upgrade pip
```

### 5. Install Dependencies

```bash
python -m pip install PySide6 reportlab pypdf python-barcode Pillow
```

### 6. Launch Vyaprix

```bash
python Vyaprix.py
```

**Note:** Replace the entry-point filename with the actual application startup file if it differs. Use a Python interpreter compatible with the source code and the supported dependency versions.

## 🔐 Data Management & Security

Vyaprix is designed around local business data management. Responsible handling of business records is important.

- Back up the SQLite database regularly.
- Maintain separate copies of important invoice PDFs and exported records.
- Test database imports and restoration procedures using backup copies.
- Protect customer information and business configuration files.
- Use appropriate file-system permissions for sensitive business records.
- Verify customer details, GSTINs, tax rates, and invoice calculations before issuing invoices.

## ⚠️ Important Disclaimer

Vyaprix is a desktop business record-keeping and invoicing utility. It does not replace professional accounting, taxation, or legal services.

Users are responsible for verifying invoice information, GST calculations, applicable tax requirements, and statutory compliance before using generated invoices for official purposes.

## 🔧 Troubleshooting

**Application does not start**

- Verify that the correct Python environment is activated.
- Confirm that all required dependencies are installed.
- Check the terminal output for missing modules or compatibility errors.

**PDF generation issues**

- Verify the ReportLab installation.
- Check the output directory permissions.
- Confirm that required fonts and image assets are accessible.

**Database access issues**

- Ensure the application has permission to access its database directory.
- Close other running application instances if database locks occur.
- Restore from a verified backup if database corruption is suspected.

## 📦 Project Status

**PySide6 Edition — Preview Build**

The graphical interface has been migrated from the earlier Tkinter implementation to PySide6. The preview build includes converted application screens, dialogs, and interactive components. Final cleanup and verification are still in progress.

Features and integrations should be considered subject to verification in the current build.

## 📄 Licensing & Credits

**Developed by EllowLabs**

Vyaprix and its associated branding, source code, and bundled assets are subject to their applicable ownership and licensing terms.

Third-party dependencies retain their respective licenses. Consult the repository's LICENSE file, if provided, before copying, modifying, or redistributing any project materials.

---

**Vyaprix by EllowLabs**

*Simplifying Business. Streamlining Invoicing. Organizing Records.*
