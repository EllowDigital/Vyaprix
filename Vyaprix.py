# =============================================================================
#  Vyaprix By EllowLabs  -  Professional Edition   (PySide6 Edition)
#  Python 3.15  -  Native Lazy Imports (PEP 810)
#
#  Single-file desktop invoicing application: PySide6 (Qt 6) user interface,
#  SQLite storage, ReportLab PDF generation, Qt's own PDF viewer and printing.
#
#  Version: the version, the build number and the release notes all come from
#           the CHANGELOG below (About -> What's new shows the same list).
#  Needs:   Python 3.15+, PySide6 (with QtPdf + QtPrintSupport), reportlab, pypdf
#           python-barcode is optional (invoices simply print without a barcode)
#           NotoSans-Regular.ttf / NotoSans-Bold.ttf in a "fonts" folder and the
#           logo images in an "assets" folder, next to this file (as before)
#  Run:     python Vyaprix.py
# =============================================================================

# --- Standard Library: Eager Imports ---
# (logging and re are already used while this file is still loading -- the log file is opened and the
#  validation patterns are compiled -- so making them lazy would gain nothing.)
import logging
import re
import sys

from contextlib import closing
from dataclasses import dataclass, replace
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP
from enum import Enum
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple, Type
from xml.sax.saxutils import escape as xml_escape

# --- Standard Library: Lazy Imports ---
lazy import csv
lazy import datetime
lazy import io
lazy import json
lazy import os
lazy import queue
lazy import shutil
lazy import sqlite3
lazy import subprocess
lazy import tempfile
lazy import threading
lazy import time
lazy import urllib.parse
lazy import webbrowser

# --- GUI Dependencies (PySide6 / Qt) ---
from PySide6.QtCore import (
    Qt, QCoreApplication, QDate, QEvent, QItemSelectionModel, QLocale, QObject, QPoint, QPointF,
    QRect, QRectF, QRegularExpression, QSize, QStandardPaths, QStringListModel, QTimer, QUrl,
    Signal,
)
from PySide6.QtGui import (
    QBrush, QColor, QDesktopServices, QFont, QFontDatabase, QFontMetrics, QGuiApplication, QIcon,
    QImage, QKeySequence, QPainter, QPainterPath, QPalette, QPen, QPixmap,
    QRegularExpressionValidator, QShortcut, QValidator,
)
from PySide6.QtWidgets import (
    QAbstractItemView, QAbstractScrollArea, QAbstractSpinBox, QApplication, QButtonGroup, QCheckBox,
    QColorDialog, QComboBox, QCompleter, QDateEdit, QDialog, QFileDialog, QFormLayout, QFrame, QGridLayout,
    QHBoxLayout, QHeaderView, QLabel, QLayout, QLineEdit, QMainWindow, QMenu, QMessageBox, QPlainTextEdit,
    QProgressBar, QPushButton, QRadioButton, QScrollArea, QSizePolicy, QStackedWidget, QTabWidget,
    QTableWidget, QTableWidgetItem, QTableWidgetSelectionRange, QToolButton, QVBoxLayout, QWidget,
)

# --- ReportLab: Lazy Imports ---
lazy from reportlab.lib import colors
lazy from reportlab.lib.enums import TA_CENTER, TA_RIGHT
lazy from reportlab.lib.fonts import addMapping
lazy from reportlab.lib.pagesizes import A4
lazy from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
lazy from reportlab.lib.units import cm
lazy from reportlab.lib.utils import simpleSplit
lazy from reportlab.pdfbase import pdfmetrics
lazy from reportlab.pdfbase.ttfonts import TTFont
lazy from reportlab.pdfgen.canvas import Canvas
lazy from reportlab.platypus import (
    BaseDocTemplate,
    CondPageBreak,
    Flowable,
    Frame,
    Image,
    KeepTogether,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

# --- pypdf: needed for bulk printing (merging PDFs) ---
# Qt's own QtPdf shows the PDFs; pypdf only merges them. A missing pypdf disables bulk printing only.
try:
    from pypdf import PdfWriter
    PYPDF_AVAILABLE = True
except ImportError:
    PdfWriter = None
    PYPDF_AVAILABLE = False
    print("WARNING: pypdf not installed. Bulk-print PDF merging will be skipped (pip install pypdf).")

# --- Python-Barcode: Optional Dependency ---
try:
    import barcode
    from barcode.writer import ImageWriter

    BARCODE_AVAILABLE = True
except ImportError:
    barcode = None
    ImageWriter = None
    BARCODE_AVAILABLE = False
    print("WARNING: python-barcode not installed. Barcodes will be skipped.")


# =============================================================================
#  1. APPLICATION METADATA
# =============================================================================
@dataclass(frozen=True)
class Release:
    """One release of Vyaprix: its version, the day it was handed over, a headline and every change made in it."""
    version: str
    date: str
    title: str
    changes: Tuple[str, ...]


# The file this project started from: Vyaprix v1.5, build 1.5.2 (the Tkinter build of 22 October 2026).
BASE_VERSION: str       = "v1.5"
BASE_BUILD_VERSION: str = "1.5.2"

# Newest first. EVERY change made to the application adds a line below -- and a new Release whenever a new build is handed
# over. APP_VERSION, the build number, the release date and the "What's new" window all follow from this list, so they
# can never disagree: the build number is "<version>.<number of changes made since the uploaded v1.5 / build 1.5.2>".
CHANGELOG: Tuple[Release, ...] = (
    Release("v2.3", "12 November 2026", "One GST line, professional A4 invoices, sharp and responsive on every screen", (
        "CGST and SGST (or IGST) appear once on the invoice form and on the PDF -- one combined line each, however many "
        "GST rates the items use. Hover the line on the form to see the split by rate.",
        "A \"GST summary\" on the PDF shows the taxable value and the tax at each rate whenever an invoice uses more than "
        "one rate.",
        "The invoice PDF is redesigned for A4: header, bill-to and invoice details, items, totals, payment box and "
        "signature on even margins, in all four templates.",
        "Long invoices run on to as many A4 pages as they need: the item headings repeat on every page, no line is "
        "ever cut in two, and each page carries the invoice number and \"Page x of y\".",
        "Auto-fit: an invoice that runs only a few lines onto an extra page is fitted onto fewer pages -- the gaps "
        "tighten a little, the type stays the same size.",
        "The item columns size themselves to what is in them -- quantity, unit, rate, HSN / SAC, GST % and amounts never "
        "wrap or get cut, and long descriptions wrap inside their cell.",
        "Very large amounts and very long names, addresses and notes still fit inside the page.",
        "The template miniatures in Settings are redrawn to match the new layouts.",
        "The application is DPI aware: a Windows scale of 125 % or 150 % is followed exactly, the logo, icons and template "
        "miniatures stay sharp, and the invoice preview and Settings re-sharpen when you move them to another monitor.",
        "The main window opens at a size that fits your screen, cannot be made too small to use, and fits itself again "
        "when you move it to a smaller monitor.",
        "The sidebar shrinks to icons in a narrow window and drops the logo in a short one, so every screen stays usable "
        "as you resize.",
        "The version, the build number and the \"What's new\" list all come from one change log: the build number counts "
        "every change made since v1.5 (build 1.5.2).",
    )),
    Release("v2.2", "10 November 2026", "Invoice templates", (
        "Four invoice templates -- Classic, Modern, Minimal and Corporate -- picked in Settings > Payment & PDF, with a "
        "miniature of each and a \"Preview sample invoice\" button.",
        "A brand colour for Modern, Minimal and Corporate: ten ready colours, or any colour of your own.",
        "A template can be chosen for a single invoice (\"PDF Template\" on the invoice form) and is remembered with it.",
    )),
    Release("v2.1", "9 November 2026", "Professional invoicing", (
        "Every invoice line has a quantity, a unit and a rate (250 Nos x Rs. 50); the line amount is worked out for you.",
        "One catalog item can be sold in several units, each with its own rate, and Vyaprix remembers the rate per unit.",
        "GST for the whole invoice or a different GST rate for every item, chosen per invoice.",
        "HSN / SAC code on every item, saved with the catalog item.",
        "Item discounts as an amount or a percentage, besides the overall discount.",
        "The place of supply decides CGST + SGST or IGST (a customer's GSTIN fills it in).",
        "Due date (on receipt, 7, 15, 30 ... days, or a date you pick), customer PO / reference number and a "
        "reverse-charge flag.",
        "Optional round off of the invoice total to the nearest rupee.",
        "Notes for the customer and Terms & Conditions printed on the invoice, with default terms set in Settings.",
        "The invoice total in words (Rupees ... Only), switched on in Settings.",
        "The catalog shows and edits unit, HSN / SAC, GST % and the other units an item is sold in.",
        "The PDF prints quantity, unit, rate, discount and GST columns whenever an invoice uses them; older invoices "
        "print as before.",
        "Databases and backups made by earlier versions upgrade themselves; saved invoices keep exactly their totals.",
        "Faster entry: Enter in the last field of the last line adds a new line.",
    )),
    Release("v2.0", "6 November 2026", "Rebuilt on PySide6", (
        "The whole interface is rebuilt on PySide6 (Qt 6); Tkinter, ttkbootstrap and PyMuPDF are no longer needed.",
        "Runs on Python 3.15 and uses PEP 810 lazy imports, so the application starts faster.",
        "Invoice preview uses Qt's own PDF viewer (zoom, page navigation, print, share).",
        "Bulk PDF merging uses pypdf instead of PyMuPDF.",
        "The eight colour themes, including the dark ones, are drawn from one palette each.",
        "Every screen of the Tk build is kept: Dashboard, invoice form, Clients, Catalog, Settings, About, Keyboard "
        "shortcuts and the splash screen.",
        "Windows and dialogs are limited to the size of the screen they open on.",
    )),
)
CHANGE_COUNT: int = sum(len(release.changes) for release in CHANGELOG)

APP_NAME: str          = "Vyaprix"
APP_VERSION: str       = CHANGELOG[0].version
APP_BUILD_VERSION: str = f"{APP_VERSION.lstrip('v')}.{CHANGE_COUNT}"
APP_RELEASE_DATE: str  = CHANGELOG[0].date
APP_TAGLINE: str       = "Business Made Smarter."

APP_ATTRIBUTION: str = (
    "Designed & Developed by Sarwan Yadav\n"
    "at EllowLabs\n"
    "Published by EllowDigital\n"
    "© 2026 EllowDigital. All Rights Reserved."
)


# =============================================================================
#  2. SYSTEM PATHS & DIRECTORIES
# =============================================================================
def _resource_root() -> Path:
    """Return the read-only root containing resources bundled by PyInstaller."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


# Read-Only paths (inside the compiled EXE)
RESOURCE_DIR: Path        = _resource_root()
RESOURCE_ASSETS_DIR: Path = RESOURCE_DIR / "assets"
FONTS_DIR: Path           = RESOURCE_DIR / "fonts"


def _program_data_root() -> Path:
    """Windows: the shared Public folder, exactly where the Tkinter build kept its data. Elsewhere
    (e.g. developing on macOS or Linux) a folder in the user's home, instead of a stray 'C:\\Users\\Public'
    directory inside the current folder."""
    if public := os.environ.get("PUBLIC"):
        return Path(public)
    if sys.platform == "win32":
        return Path(r"C:\Users\Public")
    return Path.home() / ".local" / "share"


# Mutable application state mapped to Windows ProgramData
PROGRAM_DATA_DIR: Path = _program_data_root()
DATA_DIR: Path         = PROGRAM_DATA_DIR / "EllowDigital" / "Vyaprix"

# Writable subdirectories
ASSETS_DIR: Path       = DATA_DIR / "assets"
DATA_FONTS_DIR: Path   = DATA_DIR / "fonts"
INVOICES_DIR: Path     = DATA_DIR / "invoices"
BATCH_DIR: Path        = INVOICES_DIR / "_batches"
PRINT_BATCH_SIZE: int  = 100  # bulk print merges PDFs in groups of this size -- see DashboardFrame._bulk_print_work
LOGS_DIR: Path         = DATA_DIR / "logs"

# Core Application Files
DB_NAME: Path          = DATA_DIR / "Vyaprix.db"
CONFIG_FILE: Path      = DATA_DIR / "config.json"
LOG_FILE: Path         = LOGS_DIR / "Vyaprix.log"

# UI Specific Assets (Read directly from EXE for speed)
APP_LOGO_PNG: Path     = RESOURCE_ASSETS_DIR / "app_logo.png"
APP_LOGO_ICO: Path     = RESOURCE_ASSETS_DIR / "app_logo.ico"


# =============================================================================
#  3. ENVIRONMENT INITIALIZATION
# =============================================================================
def _copy_bundled_files(src_dir: Path, dest_dir: Path, allowed_files: Tuple[str, ...] = None) -> None:
    """Helper method to securely copy bundled files to the writable directory."""
    if not src_dir.exists():
        return
        
    for item in src_dir.iterdir():
        if item.is_file():
            # Strict whitelist to prevent local dev clutter (DBs, JSONs) from extracting
            if allowed_files and item.name not in allowed_files:
                continue
                
            dest = dest_dir / item.name
            if not dest.exists():
                try:
                    shutil.copy2(item, dest)
                except Exception as e:
                    print(f"Failed to copy {item.name}: {e}")


def _initialize_environment() -> None:
    """Create mutable directories and copy default assets/fonts on first launch."""
    directories_to_create = (
        DATA_DIR, ASSETS_DIR, DATA_FONTS_DIR, 
        INVOICES_DIR, BATCH_DIR, LOGS_DIR
    )
    for directory in directories_to_create:
        directory.mkdir(parents=True, exist_ok=True)

    # Extract defaults from EXE to Writable paths using strict whitelists
    # FIXED: Added "logo.png" so the default invoice logo extracts successfully
    expected_assets = ("app_logo.png", "app_logo.ico", "logo.png", "sign.png", "upi_qr.png")
    _copy_bundled_files(RESOURCE_ASSETS_DIR, ASSETS_DIR, allowed_files=expected_assets)
    
    expected_fonts = ("NotoSans-Regular.ttf", "NotoSans-Bold.ttf")
    _copy_bundled_files(FONTS_DIR, DATA_FONTS_DIR, allowed_files=expected_fonts)


# Trigger initialization immediately on module load
_initialize_environment()


# =============================================================================
#  4. LOGGING
# =============================================================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        # forever on a machine that runs Vyaprix for years.
        RotatingFileHandler(LOG_FILE, mode='a', maxBytes=5 * 1024 * 1024,
                            backupCount=5, encoding='utf-8'),
        logging.StreamHandler(),
    ],
)


logger = logging.getLogger(__name__)


# =============================================================================
#  5. PDF FONTS
# =============================================================================
def register_fonts() -> str:
    for font_dir in (DATA_FONTS_DIR, FONTS_DIR):
        reg  = font_dir / "NotoSans-Regular.ttf"
        bold = font_dir / "NotoSans-Bold.ttf"
        if not (reg.exists() and bold.exists()):
            continue
        try:
            # Remember to keep the str() fix!
            pdfmetrics.registerFont(TTFont("NotoSans",      str(reg)))
            pdfmetrics.registerFont(TTFont("NotoSans-Bold", str(bold)))
            addMapping("NotoSans", 0, 0, "NotoSans")
            addMapping("NotoSans", 1, 0, "NotoSans-Bold")
            return "NotoSans"
        except Exception as exc:
            logger.warning(f"Could not register custom fonts from {font_dir}: {exc}")
    logger.error(
        "NotoSans font files were not found in %s or %s -- falling back to "
        "Helvetica. PDFs will NOT render the currency symbol or non-Latin "
        "characters correctly until this is fixed.", DATA_FONTS_DIR, FONTS_DIR,
    )
    return "Helvetica"


PDF_FONT = register_fonts()


# =============================================================================
#  6. UI CONSTANTS & FORMATTING
# =============================================================================
FONT_FAMILY: str         = "Segoe UI"      # Qt silently falls back to the platform default if this font is absent
DB_DATE_FORMAT: str      = "%Y-%m-%d"
DISPLAY_DATE_FORMAT: str = "%d %b %Y"
PAYMENT_MODES: List[str] = ['Cash', 'UPI', 'Bank Transfer', 'Card', 'Cheque', 'Other']

# Regex to safely sanitize invoice numbers before converting them to PDF filenames.
# Prevents path traversal and Windows filesystem crashes. Unchanged.
INVALID_FILENAME_CHARS: re.Pattern = re.compile(r'[\\/:*?"<>|]')

# The Dashboard table colours each row's text from STATUS_BADGE (a hex foreground per status). The invoice
# form's status badge takes a theme *tone* instead (STATUS_TONES below -- the Tk build's ttkbootstrap
# STATUS_BOOTSTYLE), so that it follows the active theme. The PDF watermark keeps its own palette in
# PDFGenerator._apply_watermark (the same hues: a printed page never follows the UI theme).
STATUS_BADGE: Dict[str, Dict[str, str]] = {
    "PAID":          {"foreground": "#10B981", "background": ""},  # Soft Green
    "DUE":           {"foreground": "#EF4444", "background": ""},  # Soft Red
    "PARTIAL":       {"foreground": "#F59E0B", "background": ""},  # Orange/Amber
    "OVERPAID":      {"foreground": "#8B5CF6", "background": ""},  # Violet -- distinct from "Paid" green so an overpayment stands out
    "COMPLIMENTARY": {"foreground": "#3B82F6", "background": ""},  # Blue
}

STATUS_DISPLAY_LABELS: Dict[str, str] = {
    "Due": "UNPAID", "Partial": "PARTIALLY PAID", "Paid": "PAID",
    "Overpaid": "OVERPAID", "Complimentary": "COMPLIMENTARY",
}

#: ttkbootstrap "bootstyle" of the status badge  ->  theme tone (the Tk build's STATUS_BOOTSTYLE)
STATUS_TONES: Dict[str, str] = {
    "Paid": "success", "Due": "danger", "Partial": "warning", "Overpaid": "info", "Complimentary": "info",
}


class InvoiceStatus(Enum):
    PAID          = "Paid"
    DUE           = "Due"
    PARTIAL       = "Partial"
    OVERPAID      = "Overpaid"
    COMPLIMENTARY = "Complimentary"


# =============================================================================
#  7. CONFIGURATION
# =============================================================================
# Sample values earlier builds pre-filled as *defaults*. They are not real business details, so they
# must never reach a customer's invoice: new installs start with these fields empty, old configs are
# cleaned on load, and the PDF generator ignores them even if one slips through.
_SAMPLE_GSTIN: str = "22AAAAA0000A1Z5"
_SAMPLE_UPI_ID: str = "yourupi@bank"


def _real_value(value: Any, sample: str) -> str:
    """`value` trimmed -- or "" when it is empty or just the sample text an older build pre-filled."""
    text = str(value or "").strip()
    return "" if text.casefold() == sample.casefold() else text


def get_default_config() -> Dict[str, Any]:
    return {
        "theme": "litera",
        "user_details": {
            "name": "", "phone": "", "email": "",
            "website": "", "address": "", "city": "",
            "state": "", "pin_code": "",
            "contact_info": "",
        },
        "company_details": {"name": APP_NAME, "description": APP_TAGLINE, "gstin": ""},
        "payment_details": {
            "upi_id": "", "payment_terms": "Due upon receipt",
            "bank_name": "", "account_holder_name": "", "account_number": "",
            "ifsc_code": "", "payment_notes": "",
        },
        # TARGET user invoice logo specifically to 'logo.png'
        "asset_paths":     {"logo": "assets/logo.png", "signature": "assets/sign.png", "upi_qr_code": "assets/upi_qr.png"},
        "pdf_settings":    {"footer_message": "Generated with Vyaprix • Business Made Smarter.", "always_show_qr_code": False,
                            "amount_in_words": False,           # print "Rupees ... Only" under the invoice total
                            "default_terms": "",                # terms & conditions every new invoice starts with
                            "template": "classic",              # the invoice look (see PDF_TEMPLATES)
                            "accent_color": ""},                # brand colour for the templates that take one ("" = the template's own)
        "currency_symbol": "₹",
    }


def _migrate_legacy_business_identity(config: Dict[str, Any]) -> bool:
    """Upgrades an older config in place; True if anything changed (so the caller saves it)."""
    ud   = config.get("user_details", {})
    comp = config.setdefault("company_details", {})
    pay  = config.setdefault("payment_details", {})
    changed = False

    # The sample GSTIN / UPI ID older builds pre-filled were never the user's own details.
    if comp.get("gstin") and not _real_value(comp.get("gstin"), _SAMPLE_GSTIN):
        comp["gstin"] = ""
        changed = True
    if pay.get("upi_id") and not _real_value(pay.get("upi_id"), _SAMPLE_UPI_ID):
        pay["upi_id"] = ""
        changed = True

    legacy_firm = (ud.get("firm_name") or "").strip()
    if legacy_firm and (not comp.get("name") or comp.get("name") == APP_NAME):
        comp["name"] = legacy_firm
        changed = True

    legacy_gst = (ud.get("gst_number") or "").strip()
    if legacy_gst and not comp.get("gstin"):
        comp["gstin"] = legacy_gst
        changed = True

    return changed


def load_config() -> Dict[str, Any]:
    if not CONFIG_FILE.exists():
        _cfg = get_default_config()
        _save_config_raw(_cfg)
        return _cfg
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
            data = json.load(fh)
            defaults = get_default_config()
            for k, v in defaults.items():
                if k not in data:
                    data[k] = v
                elif isinstance(v, dict):
                    for sk, sv in v.items():
                        data[k].setdefault(sk, sv)
            if _migrate_legacy_business_identity(data):
                try:
                    _save_config_raw(data)
                except OSError as exc:      # the upgrade is re-applied on the next start; never discard the loaded settings
                    logger.warning("Could not save the upgraded settings: %s", exc)
            return data
    except Exception:
        return get_default_config()


def _save_config_raw(data: Dict[str, Any]):
    """Writes config.json through a temporary file and os.replace(), so a crash or power cut part-way
    through a save can never leave a truncated (unreadable) settings file behind."""
    tmp = CONFIG_FILE.with_name(CONFIG_FILE.name + ".tmp")
    try:
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=4)
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, CONFIG_FILE)
        return
    except OSError:
        pass                                # e.g. the file is briefly locked by a virus scan: plain write below
    finally:
        try:
            tmp.unlink(missing_ok=True)     # nothing left to remove after a successful replace
        except OSError:
            pass
    with open(CONFIG_FILE, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=4)


def save_config(data: Dict[str, Any]):
    _save_config_raw(data)
    global CONFIG, CURRENCY_SYMBOL
    CONFIG          = data
    CURRENCY_SYMBOL = data.get("currency_symbol", "₹")


CONFIG          = load_config()
CURRENCY_SYMBOL = CONFIG.get("currency_symbol", "₹")


# =============================================================================
#  8. FORMATTING & VALIDATION HELPERS
# =============================================================================
def parse_date(date_str: Optional[str], in_fmt: str, out_fmt: str) -> str:
    if not date_str:
        return "N/A"
    try:
        return datetime.datetime.strptime(date_str, in_fmt).strftime(out_fmt)
    except (ValueError, TypeError):
        return date_str


def fmt_money(value: Any) -> str:
    try:
        amount = float(value or 0)
    except (TypeError, ValueError):
        amount = 0.0
    if abs(amount) < 0.005:            # would round to 0.00 either way -- never display "-0.00"
        amount = 0.0
    return f"{CURRENCY_SYMBOL}{amount:,.2f}"


def like_pattern(text: str) -> str:
    """'%text%' for a SQL `LIKE ? ESCAPE '\\'` search in which what the user typed is matched
    literally: a typed %, _ or \\ is an ordinary character, not a wildcard (searching "50%"
    used to match every row that merely contained "50")."""
    escaped = str(text).replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
    return f"%{escaped}%"


_CENTS = Decimal("0.01")


def _round2(value: Any) -> Decimal:
    if isinstance(value, Decimal):
        return value.quantize(_CENTS, rounding=ROUND_HALF_UP)
    return Decimal(str(value or 0)).quantize(_CENTS, rounding=ROUND_HALF_UP)


def safe_float(value: Any, fallback: float = 0.0) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return fallback


def _allocate_2dp(total: Decimal, weights: List[Decimal]) -> List[Decimal]:
    """Splits `total` (a 2dp amount) over lines in proportion to `weights`. Every share is a 2dp amount, the shares
    add up to exactly `total` (the paise lost to rounding go to the lines with the largest remainders) and no
    share is larger than its weight -- so an overall discount spread over the items never leaves a stray paisa."""
    weight_sum = sum(weights, Decimal("0"))
    if total <= 0 or weight_sum <= 0:
        return [Decimal("0.00")] * len(weights)
    exact  = [total * w / weight_sum for w in weights]
    shares = [e.quantize(_CENTS, rounding=ROUND_DOWN) for e in exact]
    missing = int(((total - sum(shares, Decimal("0"))) / _CENTS).to_integral_value())
    for i in sorted(range(len(shares)), key=lambda k: (-(exact[k] - shares[k]), k))[:missing]:
        shares[i] += _CENTS
    return shares


# --- Item lines (quantity / unit / rate), GST rates, places of supply and amounts in words ---------------------
UNIT_PRESETS: List[str] = [
    "Nos", "Pcs", "Set", "Unit", "Day", "Hour", "Night", "Month", "Event", "Person", "Plate", "Trip", "Lot",
    "Box", "Pack", "Dozen", "Kg", "Gm", "Ltr", "Mtr", "Sq.ft", "Sq.mtr",
]
# GST slabs since the 22 Sep 2025 reform (0 / 5 / 18 / 40, plus the special 0.25 and 3 for gems and gold). Any other
# rate -- an older invoice's 12 or 28, say -- can still be typed in.
GST_RATE_PRESETS: List[str] = ["0", "0.25", "3", "5", "18", "40"]
TAX_MODE_OVERALL, TAX_MODE_PER_ITEM = "overall", "per_item"
_DUE_CUSTOM = -1        # the due-date list's "pick a date" choice
#: the due-date list: (caption, days after the invoice date) -- None: no due date, _DUE_CUSTOM: a date picked by hand
DUE_CHOICES: List[Tuple[str, Optional[int]]] = [("None", None), ("On receipt", 0), ("7 days", 7), ("15 days", 15),
                                                ("30 days", 30), ("45 days", 45), ("60 days", 60), ("Custom", _DUE_CUSTOM)]

#: GST state / union-territory codes (the first two digits of a GSTIN) -- the "place of supply" list
INDIAN_STATES: List[Tuple[str, str]] = [
    ("01", "Jammu and Kashmir"), ("02", "Himachal Pradesh"), ("03", "Punjab"), ("04", "Chandigarh"),
    ("05", "Uttarakhand"), ("06", "Haryana"), ("07", "Delhi"), ("08", "Rajasthan"), ("09", "Uttar Pradesh"),
    ("10", "Bihar"), ("11", "Sikkim"), ("12", "Arunachal Pradesh"), ("13", "Nagaland"), ("14", "Manipur"),
    ("15", "Mizoram"), ("16", "Tripura"), ("17", "Meghalaya"), ("18", "Assam"), ("19", "West Bengal"),
    ("20", "Jharkhand"), ("21", "Odisha"), ("22", "Chhattisgarh"), ("23", "Madhya Pradesh"), ("24", "Gujarat"),
    ("26", "Dadra and Nagar Haveli and Daman and Diu"), ("27", "Maharashtra"), ("29", "Karnataka"), ("30", "Goa"),
    ("31", "Lakshadweep"), ("32", "Kerala"), ("33", "Tamil Nadu"), ("34", "Puducherry"),
    ("35", "Andaman and Nicobar Islands"), ("36", "Telangana"), ("37", "Andhra Pradesh"), ("38", "Ladakh"),
    ("97", "Other Territory"),
]
_STATE_NAME_BY_CODE: Dict[str, str] = dict(INDIAN_STATES)


def _state_key(text: Any) -> str:
    """A state name reduced to lower-case words ("Jammu & Kashmir" == "jammu and kashmir")."""
    return " ".join(re.sub(r"[^a-z0-9 ]+", " ", str(text or "").lower().replace("&", " and ")).split())


_STATE_CODE_BY_NAME: Dict[str, str] = {_state_key(n): c for c, n in INDIAN_STATES}
_STATE_CODE_BY_NAME.update({"orissa": "21", "pondicherry": "34", "new delhi": "07", "nct of delhi": "07",
                            "uttaranchal": "05", "telengana": "36", "daman and diu": "26",
                            "dadra and nagar haveli": "26", "andaman nicobar islands": "35"})


def state_code_of(text: Any) -> str:
    """The two-digit GST code of a place written as "27-Maharashtra", "27", or just "Maharashtra" ("" if unknown)."""
    text = str(text or "").strip()
    head = re.match(r"^(\d{2})(?!\d)", text)
    if head and head.group(1) in _STATE_NAME_BY_CODE:
        return head.group(1)
    return _STATE_CODE_BY_NAME.get(_state_key(text), "")


def state_label(code: str) -> str:
    """"27-Maharashtra" -- how a place of supply is written on an invoice."""
    return f"{code}-{_STATE_NAME_BY_CODE[code]}" if code in _STATE_NAME_BY_CODE else ""


def seller_state_code(config: Dict[str, Any]) -> str:
    """The state the business is registered in: the first two digits of its GSTIN, else the state typed in Settings."""
    gstin = _real_value((config.get("company_details") or {}).get("gstin"), _SAMPLE_GSTIN).upper()
    if is_valid_gstin(gstin) and gstin[:2] in _STATE_NAME_BY_CODE:
        return gstin[:2]
    return state_code_of((config.get("user_details") or {}).get("state"))


def gst_type_for_place(config: Dict[str, Any], place_of_supply: Any) -> Optional[str]:
    """"IGST" when the place of supply is in another state than the seller's, "CGST/SGST" when it is the same one,
    None while either is unknown (then the tax split is left for the user to choose)."""
    seller, place = seller_state_code(config), state_code_of(place_of_supply)
    if not seller or not place:
        return None
    return "CGST/SGST" if seller == place else "IGST"


_WORDS_BELOW_20 = ("Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven",
                   "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
_WORDS_TENS = ("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")


def _words_below_1000(n: int) -> str:
    parts: List[str] = []
    if n >= 100:
        parts.append(f"{_WORDS_BELOW_20[n // 100]} Hundred")
        n %= 100
    if n >= 20:
        parts.append(_WORDS_TENS[n // 10] + (f"-{_WORDS_BELOW_20[n % 10]}" if n % 10 else ""))
    elif n > 0:
        parts.append(_WORDS_BELOW_20[n])
    return " ".join(parts)


def _words_indian(n: int) -> str:
    """A whole number in words with the Indian grouping (thousand, lakh, crore)."""
    if n == 0:
        return "Zero"
    parts: List[str] = []
    for size, name in ((10 ** 7, "Crore"), (10 ** 5, "Lakh"), (1000, "Thousand")):
        if n >= size:
            head, n = divmod(n, size)
            parts.append(f"{_words_indian(head)} {name}")
    if n:
        parts.append(_words_below_1000(n))
    return " ".join(parts)


def amount_in_words(amount: Any, symbol: str = "₹") -> str:
    """"Rupees One Thousand Two Hundred Thirty-Four and Fifty Paise Only" -- rupee amounts only ("" for any other currency)."""
    if symbol != "₹":
        return ""
    value  = _round2(abs(Decimal(str(amount or 0))))
    rupees = int(value)
    paise  = int((value - rupees) * 100)
    text   = f"Rupees {_words_indian(rupees)}"
    if paise:
        text += f" and {_words_indian(paise)} Paise"
    return text + " Only"


def fmt_qty(value: Any) -> str:
    """A quantity as plain text: "2", "2.5", "0.125" (three decimals at most, never a trailing zero)."""
    text = f"{safe_float(value, 1.0):.3f}".rstrip("0").rstrip(".")
    return text or "0"


def item_qty(item: Dict[str, Any]) -> float:
    """The quantity of a stored item (1 for the plain amount-only lines every older invoice has)."""
    raw = item.get("qty")
    return 1.0 if raw is None or raw == "" else safe_float(raw, 1.0)


def item_is_extended(item: Dict[str, Any]) -> bool:
    """True when a line carries more than a title and an amount -- a quantity other than 1, a unit, an HSN/SAC code
    or its own GST rate. Invoices made only of plain lines keep the compact two-column PDF table of earlier builds."""
    return bool(item_qty(item) != 1.0 or str(item.get("unit") or "").strip() or str(item.get("hsn") or "").strip()
                or item.get("tax_rate") not in (None, ""))


_UNSET: Any = object()        # "argument not passed" (None is a real value for some of them)


def catalog_units(entry: Dict[str, Any]) -> List[Tuple[str, float]]:
    """Every unit a catalog item is sold in, with its rate: the main unit first (when it has one), then the other
    units saved with it."""
    units: List[Tuple[str, float]] = []
    main = str(entry.get("unit") or "").strip()
    if main:
        units.append((main, safe_float(entry.get("amount"))))
    try:
        others = json.loads(entry.get("other_units") or "[]")
    except (TypeError, ValueError):
        others = []
    for other in others if isinstance(others, list) else []:
        name = str(other.get("unit") or "").strip() if isinstance(other, dict) else ""
        if name and name.casefold() not in {u.casefold() for u, _ in units}:
            units.append((name, safe_float(other.get("rate"))))
    return units


EMAIL_FORMAT_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]{2,}$")
GSTIN_FORMAT_PATTERN = re.compile(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$")
PHONE_FORMAT_PATTERN = re.compile(r"^[+]?[\d\s().-]+$")


def is_valid_email(value: str) -> bool:
    return bool(EMAIL_FORMAT_PATTERN.match(value.strip()))


def is_valid_gstin(value: str) -> bool:
    return bool(GSTIN_FORMAT_PATTERN.match(value.strip().upper()))


def is_valid_phone(value: str) -> bool:
    value = value.strip()
    if not PHONE_FORMAT_PATTERN.match(value):
        return False
    digit_count = sum(ch.isdigit() for ch in value)
    return 7 <= digit_count <= 15


def safe_markup(value: Any) -> str:
    """Escape user-supplied text before it's embedded in ReportLab Paragraph
    strings, which parse a small XML/HTML-like markup subset. Unchanged --
    this is PDF generation logic and has nothing to do with the GUI toolkit."""
    return xml_escape(str(value or ""))


def compose_your_name(config: Dict[str, Any]) -> str:
    """Turns Settings into the single bold headline used at the top of an
    invoice's FROM block. The business name (Settings > General >
    Company Details) is the single source of truth for this -- it's the
    exact same name shown in the invoice letterhead, so the two can
    never drift into showing two different company names on one
    document. Falls back to the personal "Your Name" field only if no
    business name has been configured at all."""
    company_name  = (config.get("company_details", {}).get("name") or "").strip()
    personal_name = (config.get("user_details", {}).get("name") or "").strip()
    return company_name or personal_name or "Your Name/Company"


def compose_your_contact(config: Dict[str, Any]) -> str:
    """Turns Settings > Your Details into the multi-line contact block
    shown under the FROM headline on an invoice. Only fields that are
    actually filled in are included, so an incomplete profile doesn't
    leave blank lines on the PDF.

    Deliberately does NOT include the GSTIN: it already has its own
    dedicated line elsewhere on the same invoice (see _build_parties() /
    company_details.gstin), and printing the same GSTIN a second time
    here is confusing, not helpful -- that duplication was a real bug in
    an earlier build.

    Falls back to the legacy free-text contact_info field when none of
    the structured fields have been filled in yet -- this only matters
    right after upgrading from an older build that only had that one
    field, so existing config doesn't suddenly show blank contact
    details until the user revisits Settings and re-enters them here."""
    company_name  = (config.get("company_details", {}).get("name") or "").strip()
    user_details  = config.get("user_details", {})
    personal_name = (user_details.get("name") or "").strip()
    lines = []
    if company_name and personal_name:
        lines.append(personal_name)  # personal name already used as the headline when there's no business name
    for key in ("phone", "email", "website"):
        if val := (user_details.get(key) or "").strip():
            lines.append(val)
    if address := (user_details.get("address") or "").strip():
        lines.append(address)
    city  = (user_details.get("city") or "").strip()
    state = (user_details.get("state") or "").strip()
    pin   = (user_details.get("pin_code") or "").strip()
    city_state = ", ".join(p for p in (city, state) if p)
    if city_state and pin:
        lines.append(f"{city_state} - {pin}")
    elif city_state or pin:
        lines.append(city_state or pin)
    if lines:
        return "\n".join(lines)
    return (user_details.get("contact_info") or "").strip()


# =============================================================================
#  9. DATABASE LAYER
# =============================================================================
class DatabaseManager:
    """
    Data-access layer.

    Performance notes:
      * Every invoice row now carries denormalized `grand_total`, `total_paid`,
        and `balance_due` columns, recomputed once whenever an invoice or its
        payments change. The dashboard and invoice list therefore never need
        to re-parse item JSON or recompute tax math just to render a table or
        show summary stats -- they read plain numbers straight out of SQLite.
      * Indexes are created on the columns used for search / filter / sort.
      * WAL journal mode + NORMAL sync give much snappier writes on desktop
        hardware without risking corruption.
    """

    def __init__(self, db_path: Path):
        self.db_path = db_path
        self._init_db()

    def _conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path, timeout=10)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    # ------------------------------------------------------------------ #
    #  Schema setup & migration
    # ------------------------------------------------------------------ #
    def _init_db(self):
        with closing(self._conn()) as c:
            try:
                c.execute("PRAGMA journal_mode = WAL")
                c.execute("PRAGMA synchronous = NORMAL")
            except sqlite3.Error:
                pass
            with c:
                # ONE explicit transaction for the whole set-up. Python's sqlite3 module opens a
                # transaction only before INSERT/UPDATE/DELETE -- never before DDL -- so without this
                # every ALTER / CREATE / RENAME below committed on its own, and a crash half-way
                # through a table migration could leave `clients_legacy` behind with the real data
                # stranded in it. SQLite DDL is transactional: this makes the set-up all-or-nothing.
                c.execute("BEGIN IMMEDIATE")
                c.execute("""
                    CREATE TABLE IF NOT EXISTS invoices (
                        invoice_number        TEXT PRIMARY KEY,
                        bill_to               TEXT NOT NULL,
                        client_contact        TEXT,
                        client_gstin          TEXT,
                        your_name             TEXT NOT NULL,
                        your_contact          TEXT,
                        items                 TEXT NOT NULL,
                        amount_paid           REAL    DEFAULT 0,
                        tax_rate              REAL    DEFAULT 0,
                        invoice_date          TEXT,
                        payment_date          TEXT,
                        status                TEXT    DEFAULT "Due",
                        payment_terms         TEXT,
                        payment_mode          TEXT,
                        show_qr_code          INTEGER DEFAULT 1,
                        is_complimentary      INTEGER DEFAULT 0,
                        is_gst_invoice        INTEGER DEFAULT 1,
                        global_discount_value REAL    DEFAULT 0,
                        global_discount_type  TEXT    DEFAULT 'amount'
                    )
                """)
                c.execute("""
                    CREATE TABLE IF NOT EXISTS clients (
                        name TEXT PRIMARY KEY,
                        contact TEXT,
                        gstin TEXT
                    )
                """)
                c.execute("""
                    CREATE TABLE IF NOT EXISTS catalog (
                        title TEXT PRIMARY KEY,
                        description TEXT,
                        amount REAL
                    )
                """)

                # --- invoices: add newer columns as needed -----------------
                existing = {r["name"] for r in c.execute("PRAGMA table_info(invoices)").fetchall()}
                for col, ctype in {
                    "advance_amount":       "REAL DEFAULT 0",
                    "advance_date":         "TEXT",
                    "gst_type":             "TEXT DEFAULT 'CGST/SGST'",
                    "notes":                "TEXT DEFAULT ''",
                    "grand_total":          "REAL DEFAULT 0",
                    "total_paid":           "REAL DEFAULT 0",
                    "balance_due":          "REAL DEFAULT 0",
                    "created_ts":           "REAL DEFAULT 0",
                    # --- professional invoicing (items with quantity / unit / own GST rate, round off, ...) ---
                    "tax_mode":             "TEXT DEFAULT 'overall'",     # 'overall' (one GST rate) | 'per_item'
                    "round_off":            "INTEGER DEFAULT 0",          # round the grand total to the nearest rupee
                    "place_of_supply":      "TEXT DEFAULT ''",
                    "po_number":            "TEXT DEFAULT ''",            # customer's PO / reference number
                    "due_date":             "TEXT",
                    "reverse_charge":       "INTEGER DEFAULT 0",
                    "customer_notes":       "TEXT DEFAULT ''",            # printed on the invoice (`notes` is internal)
                    "terms":                "TEXT DEFAULT ''",            # terms & conditions, printed
                    "pdf_template":         "TEXT DEFAULT ''",            # this invoice's own look ('' = the Settings default)
                }.items():
                    if col not in existing:
                        c.execute(f"ALTER TABLE invoices ADD COLUMN {col} {ctype}")

                # --- clients / catalog: migrate to id-keyed schema so a
                #     dedicated manager screen can rename records safely ----
                self._migrate_to_id_schema(
                    c, "clients", "name",
                    "contact TEXT, gstin TEXT", ["contact", "gstin"],
                )
                self._migrate_to_id_schema(
                    c, "catalog", "title",
                    "description TEXT, amount REAL", ["description", "amount"],
                )
                # --- catalog: unit, HSN/SAC, GST rate and the other units an item is sold in ------------
                existing_catalog = {r["name"] for r in c.execute("PRAGMA table_info(catalog)").fetchall()}
                for col, ctype in {
                    "unit":        "TEXT DEFAULT ''",     # the item's main unit; `amount` is the rate per that unit
                    "hsn":         "TEXT DEFAULT ''",
                    "tax_rate":    "REAL",                # NULL = no rate saved with the item
                    "other_units": "TEXT DEFAULT ''",     # JSON [{"unit": "Dozen", "rate": 100.0}, ...]
                }.items():
                    if col not in existing_catalog:
                        c.execute(f"ALTER TABLE catalog ADD COLUMN {col} {ctype}")

                # --- payments table (multiple partial payments per invoice) -
                c.execute("""
                    CREATE TABLE IF NOT EXISTS payments (
                        id             INTEGER PRIMARY KEY AUTOINCREMENT,
                        invoice_number TEXT NOT NULL,
                        amount         REAL NOT NULL DEFAULT 0,
                        payment_date   TEXT,
                        mode           TEXT,
                        note           TEXT,
                        FOREIGN KEY (invoice_number) REFERENCES invoices(invoice_number)
                            ON DELETE CASCADE ON UPDATE CASCADE
                    )
                """)

                # --- indexes for search / sort / filter perf ----------------
                # One per column DashboardFrame can sort by server-side
                # (see DatabaseManager.get_invoices_page) or that get_next_
                # invoice_number()/backfill scan by, so none of those ever
                # fall back to a full table scan even at 100,000+ rows.
                c.execute("CREATE INDEX IF NOT EXISTS idx_invoices_status  ON invoices(status)")
                c.execute("CREATE INDEX IF NOT EXISTS idx_invoices_date    ON invoices(invoice_date)")
                c.execute("CREATE INDEX IF NOT EXISTS idx_invoices_billto  ON invoices(bill_to)")
                c.execute("CREATE INDEX IF NOT EXISTS idx_invoices_total   ON invoices(grand_total)")
                c.execute("CREATE INDEX IF NOT EXISTS idx_invoices_paid    ON invoices(total_paid)")
                c.execute("CREATE INDEX IF NOT EXISTS idx_invoices_balance ON invoices(balance_due)")
                c.execute("CREATE INDEX IF NOT EXISTS idx_invoices_created ON invoices(created_ts)")
                c.execute("CREATE INDEX IF NOT EXISTS idx_payments_invnum  ON payments(invoice_number)")

                self._migrate_legacy_payments(c)
                self._backfill_financials(c)

    @staticmethod
    def _migrate_to_id_schema(c: sqlite3.Connection, table: str, key_col: str,
                               extra_cols_sql: str, extra_col_names: List[str]):
        """Give `clients`/`catalog` a stable integer id so records can be renamed."""
        cols = {r["name"] for r in c.execute(f"PRAGMA table_info({table})").fetchall()}
        if not cols or "id" in cols:
            return
        c.execute(f"ALTER TABLE {table} RENAME TO {table}_legacy")
        c.execute(f"""
            CREATE TABLE {table} (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {key_col} TEXT UNIQUE NOT NULL,
                {extra_cols_sql}
            )
        """)
        col_list = ", ".join([key_col] + extra_col_names)
        c.execute(f"INSERT OR IGNORE INTO {table} ({col_list}) SELECT {col_list} FROM {table}_legacy")
        c.execute(f"DROP TABLE {table}_legacy")

    @staticmethod
    def _migrate_legacy_payments(c: sqlite3.Connection):
        """One-time migration: fold legacy advance_amount/amount_paid into the
        new `payments` table so existing invoices keep their payment history."""
        already = c.execute("SELECT COUNT(*) AS n FROM payments").fetchone()["n"]
        if already:
            return
        rows = c.execute("""
            SELECT invoice_number, advance_amount, advance_date, amount_paid,
                   payment_date, payment_mode, invoice_date
            FROM invoices
        """).fetchall()
        for r in rows:
            adv = float(r["advance_amount"] or 0)
            if adv > 0:
                c.execute(
                    "INSERT INTO payments (invoice_number, amount, payment_date, mode, note) "
                    "VALUES (?,?,?,?,?)",
                    (r["invoice_number"], adv, r["advance_date"] or r["invoice_date"],
                     r["payment_mode"], "Advance payment"),
                )
            paid = float(r["amount_paid"] or 0)
            if paid > 0:
                c.execute(
                    "INSERT INTO payments (invoice_number, amount, payment_date, mode, note) "
                    "VALUES (?,?,?,?,?)",
                    (r["invoice_number"], paid, r["payment_date"] or r["invoice_date"],
                     r["payment_mode"], ""),
                )

    def _backfill_financials(self, c: sqlite3.Connection):
        """Populate grand_total/total_paid/balance_due for any row missing them
        (fresh migration, or rows saved by an older build)."""
        rows = c.execute(
            "SELECT invoice_number FROM invoices WHERE grand_total = 0 AND is_complimentary = 0"
        ).fetchall()
        for r in rows:
            self._recompute_one(c, r["invoice_number"])

    # ------------------------------------------------------------------ #
    #  Totals math (pure function -- also used by the PDF generator & UI)
    # ------------------------------------------------------------------ #
    @staticmethod
    def calculate_totals(items: List[Dict], tax_rate: float,
                         global_disc_val: float = 0.0,
                         global_disc_type: str = "amount",
                         tax_mode: str = TAX_MODE_OVERALL,
                         gst_type: str = "CGST/SGST",
                         round_off: bool = False) -> Dict[str, Any]:
        """Computes every figure on the invoice summary, in this fixed
        order: Subtotal -> Discount -> Taxable Amount -> CGST/SGST ->
        (Round off ->) Grand Total. All arithmetic happens in Decimal with
        ROUND_HALF_UP, rounding to the nearest paisa/cent at each
        intermediate step rather than only at the very end -- that's
        what guarantees CGST + SGST always sum to exactly the combined
        tax shown elsewhere, instead of occasionally being a paisa off
        from ordinary floating-point drift.

        Two ways to charge GST:
          * tax_mode "overall"  -- one rate (`tax_rate`) on the whole taxable amount. This is the only mode older
            invoices know, and its arithmetic is untouched, so every stored invoice keeps its exact totals.
          * tax_mode "per_item" -- every line carries its own `tax_rate` (a line without one takes `tax_rate`, the
            default). Each line's discount is taken off first, the overall discount is spread over the lines in
            proportion to what they are worth, the taxable values are grouped by rate and CGST / SGST are computed
            once per rate -- the way a GST invoice's tax summary is laid out. With the same rate on every line this
            gives exactly the "overall" figures.

        `round_off` rounds the grand total to the nearest rupee (half up); the difference is returned as
        "round_off" (negative when rounding down). `gst_type` only decides what "igst" holds -- cgst and sgst are
        always the two equal halves, as before, and IGST is their sum.

        Returns plain floats (already rounded to 2dp), not Decimal objects -- every existing caller (SQLite writes,
        fmt_money(), the PDF generator) keeps working exactly as before. Besides the original keys it returns
        "tax" (all tax), "igst", "round_off" and "tax_groups": one {"rate", "taxable", "cgst", "sgst", "igst",
        "tax"} dict per GST rate, lowest rate first (empty when no tax is charged).
        """
        def d(x: Any) -> Decimal:
            return Decimal(str(x or 0))

        zero = Decimal("0")
        per_item = tax_mode == TAX_MODE_PER_ITEM
        if per_item:
            # every line is rounded on its own, and a line's discount can never be larger than the line itself
            gross_list = [_round2(d(i.get("amount", 0))) for i in items]
            disc_list  = [max(zero, min(_round2(d(i.get("discount", 0))), g)) for i, g in zip(items, gross_list)]
            subtotal   = sum(gross_list, zero)
            item_disc  = sum(disc_list, zero)
        else:
            subtotal   = _round2(sum((d(i.get("amount", 0)) for i in items), zero))
            item_disc  = _round2(sum((d(i.get("discount", 0)) for i in items), zero))
        after_item = max(zero, subtotal - item_disc)

        if global_disc_type == "%":
            global_disc = _round2(after_item * d(global_disc_val) / Decimal("100"))
        else:
            global_disc = _round2(d(global_disc_val))
        global_disc = max(zero, min(global_disc, after_item))

        total_disc = item_disc + global_disc
        taxable    = max(zero, subtotal - total_disc)

        # (rate, taxable value at that rate) pairs, lowest rate first
        if per_item:
            nets   = [g - dsc for g, dsc in zip(gross_list, disc_list)]
            shares = _allocate_2dp(global_disc, nets)
            bases: Dict[Decimal, Decimal] = {}
            for item, net, share in zip(items, nets, shares):
                raw  = item.get("tax_rate")
                try:
                    rate = d(tax_rate) if raw is None or raw == "" else d(raw)
                except ArithmeticError:                      # an unreadable rate counts as "no rate": the default
                    rate = d(tax_rate)
                rate = max(zero, min(rate, Decimal("100")))
                bases[rate] = bases.get(rate, zero) + (net - share)
            groups = sorted(bases.items())
            taxable = sum(bases.values(), zero)
        else:
            groups = [(d(tax_rate), taxable)] if d(tax_rate) > 0 else []

        cgst = sgst = zero
        tax_groups: List[Dict[str, float]] = []
        for rate, base in groups:
            half = _round2(base * (rate / Decimal("2")) / Decimal("100"))
            cgst += half
            sgst += half
            tax_groups.append({"rate": float(rate), "taxable": float(base), "cgst": float(half), "sgst": float(half),
                               "igst": float(half + half), "tax": float(half + half)})

        before_round = taxable + cgst + sgst
        if round_off:
            grand_total = before_round.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
        else:
            grand_total = before_round
        tax = cgst + sgst

        return {
            "subtotal": float(subtotal), "item_discount": float(item_disc),
            "global_discount": float(global_disc), "total_discount": float(total_disc),
            "taxable": float(taxable), "cgst": float(cgst), "sgst": float(sgst),
            "grand_total": float(grand_total),
            "tax": float(tax), "igst": float(tax) if gst_type == "IGST" else 0.0,
            "round_off": float(grand_total - before_round), "tax_groups": tax_groups,
        }

    @classmethod
    def totals_for(cls, data: Dict[str, Any], items: Optional[List[Dict]] = None) -> Dict[str, Any]:
        """The totals of one stored invoice -- `data` is its database row as a dict -- read from every column that
        matters (GST on/off, the rate or per-item mode, the tax split, the overall discount, round off). The database
        layer (_recompute_one), the PDF and the live invoice form all call this one function, so what is saved, printed
        and shown while typing can never disagree. `items` may be passed to avoid parsing the JSON again."""
        if items is None:
            try:
                items = json.loads(data.get("items") or "[]")
            except (TypeError, ValueError):
                items = []
        is_gst = data.get("is_gst_invoice", 1)
        return cls.calculate_totals(
            items,
            safe_float(data.get("tax_rate")) if is_gst else 0.0,
            safe_float(data.get("global_discount_value")),
            data.get("global_discount_type") or "amount",
            tax_mode=(data.get("tax_mode") or TAX_MODE_OVERALL) if is_gst else TAX_MODE_OVERALL,
            gst_type=data.get("gst_type") or "CGST/SGST",
            round_off=bool(data.get("round_off")),
        )

    @staticmethod
    def compute_payment_status(grand_total: float, paid: float) -> Tuple[str, float, float]:
        """Given a grand total and amount paid, returns
        (status, balance_due, advance_credit) -- the single shared
        source of truth for "is this invoice unpaid, partially paid,
        paid, or overpaid", used by the database layer (_recompute_one),
        the live invoice-form preview, and the PDF, so the three can
        never disagree with each other.

        balance_due is never negative: a customer who paid more than
        the grand total isn't owed a "negative bill" -- the excess
        comes back separately as advance_credit, which callers display
        as "Advance / Credit" instead of a confusing negative balance.
        `status` is one of InvoiceStatus's own stored values, so this
        can be written straight to the database.
        """
        gt, p = _round2(grand_total), _round2(paid)
        diff = gt - p                                    # > 0 still owed, < 0 overpaid
        if diff > 0:
            status = InvoiceStatus.DUE.value if p <= 0 else InvoiceStatus.PARTIAL.value
            return status, float(diff), 0.0
        if diff < 0:
            return InvoiceStatus.OVERPAID.value, 0.0, float(-diff)
        return (InvoiceStatus.PAID.value if gt > 0 else InvoiceStatus.DUE.value), 0.0, 0.0

    # ------------------------------------------------------------------ #
    #  Internal: recompute + persist grand_total/total_paid/balance/status
    # ------------------------------------------------------------------ #
    def _recompute_one(self, c: sqlite3.Connection, invoice_number: str):
        row = c.execute("SELECT * FROM invoices WHERE invoice_number=?", (invoice_number,)).fetchone()
        if not row:
            return
        d = dict(row)
        if d.get("is_complimentary"):
            c.execute(
                "UPDATE invoices SET grand_total=0, total_paid=0, balance_due=0, status=? "
                "WHERE invoice_number=?",
                (InvoiceStatus.COMPLIMENTARY.value, invoice_number),
            )
            return
        totals = self.totals_for(d)
        total_paid = c.execute(
            "SELECT COALESCE(SUM(amount),0) AS s FROM payments WHERE invoice_number=?",
            (invoice_number,),
        ).fetchone()["s"]
        grand_total = totals["grand_total"]
        status, balance, _credit = self.compute_payment_status(grand_total, total_paid)
        # balance_due is stored floored at zero (never negative) -- it
        # feeds a SUM() for the dashboard's "total balance due" stat, and
        # a negative value from one overpaid invoice would silently
        # (and wrongly) offset what other invoices are genuinely still
        # owed. The overpayment itself (_credit) is a per-invoice figure,
        # computed on demand wherever it's shown rather than persisted.
        c.execute(
            "UPDATE invoices SET grand_total=?, total_paid=?, balance_due=?, status=? "
            "WHERE invoice_number=?",
            (grand_total, total_paid, balance, status, invoice_number),
        )

    def recompute_invoice_financials(self, invoice_number: str):
        with closing(self._conn()) as c:
            with c:
                self._recompute_one(c, invoice_number)

    # ------------------------------------------------------------------ #
    #  Clients
    # ------------------------------------------------------------------ #
    def get_all_clients(self, search: str = "") -> List[Dict]:
        with closing(self._conn()) as c:
            if search:
                rows = c.execute(
                    "SELECT * FROM clients WHERE name LIKE ? ESCAPE '\\' ORDER BY name",
                    (like_pattern(search),),
                ).fetchall()
            else:
                rows = c.execute("SELECT * FROM clients ORDER BY name").fetchall()
            return [dict(r) for r in rows]

    def get_client_by_id(self, client_id: int) -> Optional[Dict]:
        with closing(self._conn()) as c:
            row = c.execute("SELECT * FROM clients WHERE id=?", (client_id,)).fetchone()
            return dict(row) if row else None

    def save_client(self, name: str, contact: str, gstin: str):
        """Upsert-by-name -- used for the quick auto-save-from-invoice-form path."""
        if not name:
            return
        with closing(self._conn()) as c:
            with c:
                c.execute("""
                    INSERT INTO clients (name, contact, gstin)
                    VALUES (?, ?, ?)
                    ON CONFLICT(name) DO UPDATE SET
                    contact=excluded.contact, gstin=excluded.gstin
                """, (name, contact, gstin))

    def upsert_client_by_id(self, client_id: Optional[int], name: str, contact: str, gstin: str) -> Tuple[bool, Optional[str]]:
        """Used by the dedicated Clients manager screen -- supports renames."""
        if not name.strip():
            return False, "Client name is required."
        try:
            with closing(self._conn()) as c:
                with c:
                    dupe = c.execute(
                        "SELECT id FROM clients WHERE name=? AND id IS NOT ?",
                        (name, client_id if client_id is not None else -1),
                    ).fetchone()
                    if dupe:
                        return False, f"A client named '{name}' already exists."
                    if client_id is None:
                        c.execute("INSERT INTO clients (name, contact, gstin) VALUES (?,?,?)",
                                  (name, contact, gstin))
                    else:
                        c.execute("UPDATE clients SET name=?, contact=?, gstin=? WHERE id=?",
                                  (name, contact, gstin, client_id))
            return True, None
        except sqlite3.Error as exc:
            return False, str(exc)

    def delete_client(self, client_id: int):
        with closing(self._conn()) as c:
            with c:
                c.execute("DELETE FROM clients WHERE id=?", (client_id,))

    def delete_clients(self, client_ids: List[int]):
        with closing(self._conn()) as c:
            with c:
                c.executemany("DELETE FROM clients WHERE id=?", [(i,) for i in client_ids])

    # ------------------------------------------------------------------ #
    #  Catalog
    # ------------------------------------------------------------------ #
    def get_all_catalog_items(self, search: str = "") -> List[Dict]:
        with closing(self._conn()) as c:
            if search:
                rows = c.execute(
                    "SELECT * FROM catalog WHERE title LIKE ? ESCAPE '\\' ORDER BY title",
                    (like_pattern(search),),
                ).fetchall()
            else:
                rows = c.execute("SELECT * FROM catalog ORDER BY title").fetchall()
            return [dict(r) for r in rows]

    def get_catalog_item_by_id(self, item_id: int) -> Optional[Dict]:
        with closing(self._conn()) as c:
            row = c.execute("SELECT * FROM catalog WHERE id=?", (item_id,)).fetchone()
            return dict(row) if row else None

    def save_catalog_item(self, title: str, description: str, amount: float,
                          unit: str = "", hsn: str = "", tax_rate: Optional[float] = None):
        """Upsert by title. The unit, HSN/SAC and GST rate are only changed when given (an empty one keeps what the
        item already has)."""
        if not title:
            return
        with closing(self._conn()) as c:
            with c:
                c.execute("""
                    INSERT INTO catalog (title, description, amount, unit, hsn, tax_rate)
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(title) DO UPDATE SET
                    description=excluded.description, amount=excluded.amount,
                    unit=CASE WHEN excluded.unit != '' THEN excluded.unit ELSE catalog.unit END,
                    hsn=CASE WHEN excluded.hsn != '' THEN excluded.hsn ELSE catalog.hsn END,
                    tax_rate=COALESCE(excluded.tax_rate, catalog.tax_rate)
                """, (title, description, amount, unit or "", hsn or "", tax_rate))

    def remember_catalog_item(self, title: str, description: str, rate: float, unit: str = "", hsn: str = "",
                              tax_rate: Optional[float] = None):
        """What saving an invoice teaches the catalog about one of its lines: the description, HSN/SAC and GST rate
        are refreshed, and `rate` is stored for the unit the line was sold in -- as the item's main rate, or as the
        rate of one of its other units (a unit it has not been sold in before is added to them). That is how one
        item ends up with a price per unit ("Tent": Nos 500, Day 150) without any set-up."""
        if not title:
            return
        unit = (unit or "").strip()
        with closing(self._conn()) as c:
            with c:
                row = c.execute("SELECT * FROM catalog WHERE title=?", (title,)).fetchone()
                if row is None:
                    c.execute("INSERT INTO catalog (title, description, amount, unit, hsn, tax_rate) VALUES (?,?,?,?,?,?)",
                              (title, description, rate, unit, hsn or "", tax_rate))
                    return
                entry  = dict(row)
                main   = str(entry.get("unit") or "").strip()
                others = [{"unit": u, "rate": r} for u, r in catalog_units(entry)[1 if main else 0:]]
                amount, new_main = entry.get("amount"), main
                if not unit or not main or unit.casefold() == main.casefold():
                    amount, new_main = rate, main or unit
                else:
                    for o in others:
                        if o["unit"].casefold() == unit.casefold():
                            o["rate"] = rate
                            break
                    else:
                        others.append({"unit": unit, "rate": rate})
                c.execute(
                    "UPDATE catalog SET description=?, amount=?, unit=?, hsn=?, tax_rate=?, other_units=? WHERE id=?",
                    (description, amount, new_main, hsn or entry.get("hsn") or "",
                     tax_rate if tax_rate is not None else entry.get("tax_rate"),
                     json.dumps(others) if others else "", entry["id"]))

    def upsert_catalog_item_by_id(self, item_id: Optional[int], title: str, description: str, amount: float,
                                  unit: Any = _UNSET, hsn: Any = _UNSET, tax_rate: Any = _UNSET,
                                  other_units: Any = _UNSET) -> Tuple[bool, Optional[str]]:
        """Used by the Catalog manager. `unit`, `hsn`, `tax_rate` (None = no rate) and `other_units` (a list of
        (unit, rate) pairs) are left as they are when not passed."""
        if not title.strip():
            return False, "Item title is required."
        extra = {"unit": unit, "hsn": hsn, "tax_rate": tax_rate,
                 "other_units": _UNSET if other_units is _UNSET else
                 (json.dumps([{"unit": u, "rate": r} for u, r in other_units]) if other_units else "")}
        extra = {k: v for k, v in extra.items() if v is not _UNSET}
        try:
            with closing(self._conn()) as c:
                with c:
                    dupe = c.execute(
                        "SELECT id FROM catalog WHERE title=? AND id IS NOT ?",
                        (title, item_id if item_id is not None else -1),
                    ).fetchone()
                    if dupe:
                        return False, f"An item named '{title}' already exists."
                    if item_id is None:
                        cols = ["title", "description", "amount", *extra]
                        c.execute(f"INSERT INTO catalog ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                                  (title, description, amount, *extra.values()))
                    else:
                        sets = ", ".join(f"{k}=?" for k in ("title", "description", "amount", *extra))
                        c.execute(f"UPDATE catalog SET {sets} WHERE id=?",
                                  (title, description, amount, *extra.values(), item_id))
            return True, None
        except sqlite3.Error as exc:
            return False, str(exc)

    def delete_catalog_item(self, item_id: int):
        with closing(self._conn()) as c:
            with c:
                c.execute("DELETE FROM catalog WHERE id=?", (item_id,))

    def delete_catalog_items(self, item_ids: List[int]):
        with closing(self._conn()) as c:
            with c:
                c.executemany("DELETE FROM catalog WHERE id=?", [(i,) for i in item_ids])

    # ------------------------------------------------------------------ #
    #  Payments
    # ------------------------------------------------------------------ #
    def get_payments(self, invoice_number: str) -> List[Dict]:
        with closing(self._conn()) as c:
            rows = c.execute(
                "SELECT * FROM payments WHERE invoice_number=? ORDER BY payment_date, id",
                (invoice_number,),
            ).fetchall()
            return [dict(r) for r in rows]

    def replace_payments(self, invoice_number: str, payments: List[Dict]):
        """Replace the full payment history for one invoice in a single
        transaction, then recompute its denormalized totals/status."""
        with closing(self._conn()) as c:
            with c:
                c.execute("DELETE FROM payments WHERE invoice_number=?", (invoice_number,))
                for p in payments:
                    c.execute(
                        "INSERT INTO payments (invoice_number, amount, payment_date, mode, note) "
                        "VALUES (?,?,?,?,?)",
                        (invoice_number, p.get("amount", 0), p.get("payment_date"),
                         p.get("mode"), p.get("note", "")),
                    )
                self._recompute_one(c, invoice_number)

    # ------------------------------------------------------------------ #
    #  Invoices
    # ------------------------------------------------------------------ #
    def get_all_invoices(self, search_query: str = "", status_filter: str = "All") -> List[sqlite3.Row]:
        clauses, params = [], []
        if search_query:
            clauses.append("(invoice_number LIKE ? ESCAPE '\\' OR bill_to LIKE ? ESCAPE '\\' "
                           "OR notes LIKE ? ESCAPE '\\')")
            q = like_pattern(search_query)
            params += [q, q, q]
        if status_filter and status_filter != "All":
            clauses.append("UPPER(status) = ?")
            params.append(status_filter.upper())
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""

        with closing(self._conn()) as c:
            return c.execute(
                f"SELECT * FROM invoices {where} ORDER BY invoice_date DESC, invoice_number DESC",
                params,
            ).fetchall()

    # Maps the Dashboard's UI sort-column ids to a real, indexed SQL
    # column (see the indexes created in _init_db()) -- "note" isn't a
    # real column, so it's expressed as a boolean CASE instead.
    _SORT_COLUMN_SQL = {
        "id": "invoice_number", "client": "bill_to", "date": "invoice_date",
        "status": "status", "total": "grand_total", "paid": "total_paid",
        "balance": "balance_due",
        "note": "(CASE WHEN notes IS NOT NULL AND TRIM(notes) != '' THEN 1 ELSE 0 END)",
    }

    def get_invoices_page(self, search_query: str = "", status_filter: str = "All",
                           sort_col: str = "date", sort_desc: bool = True,
                           page: int = 1, page_size: int = 50) -> Tuple[List[sqlite3.Row], int]:
        """Search + filter + sort + pagination, all done in SQL. This is
        the core of keeping the Dashboard fast at any invoice count: it
        never pulls more than one page's worth of rows into Python, no
        matter how large the table is (tested to 100,000+ rows) -- unlike
        get_all_invoices() above, which is only safe for genuinely
        whole-table operations like a full CSV export.

        Returns (rows_for_this_page, total_matching_rows) -- the total is
        needed for the pagination bar's "Page X of Y" and comes from a
        second, cheap COUNT(*) query using the same WHERE clause.
        """
        clauses, params = [], []
        if search_query:
            clauses.append("(invoice_number LIKE ? ESCAPE '\\' OR bill_to LIKE ? ESCAPE '\\' "
                           "OR notes LIKE ? ESCAPE '\\')")
            q = like_pattern(search_query)
            params += [q, q, q]
        if status_filter and status_filter != "All":
            clauses.append("UPPER(status) = ?")
            params.append(status_filter.upper())
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""

        sql_col = self._SORT_COLUMN_SQL.get(sort_col, "invoice_date")
        direction = "DESC" if sort_desc else "ASC"
        # Tie-break on invoice_number so the order is stable across pages
        # even when many rows share the same sort value (e.g. the same
        # date) -- without a stable tie-break, rows can shuffle between
        # pages, or repeat/get skipped, as the user pages through.
        order_by = f"ORDER BY {sql_col} {direction}, invoice_number {direction}"

        page = max(1, page)
        page_size = max(1, page_size)
        offset = (page - 1) * page_size

        with closing(self._conn()) as c:
            total = c.execute(f"SELECT COUNT(*) FROM invoices {where}", params).fetchone()[0]
            rows = c.execute(
                f"SELECT * FROM invoices {where} {order_by} LIMIT ? OFFSET ?",
                params + [page_size, offset],
            ).fetchall()
        return rows, total

    def get_invoices_by_ids(self, inv_ids: List[str]) -> List[sqlite3.Row]:
        if not inv_ids:
            return []
        with closing(self._conn()) as c:
            q_marks = ",".join("?" * len(inv_ids))
            return c.execute(
                f"SELECT * FROM invoices WHERE invoice_number IN ({q_marks}) "
                "ORDER BY invoice_date DESC, invoice_number DESC",
                inv_ids,
            ).fetchall()

    def iter_all_invoices_chunked(self, chunk_size: int = 1000):
        """Yields the whole invoices table in fixed-size chunks (ordered
        by rowid, the cheapest possible stable order) instead of loading
        every row into memory at once via fetchall(). Used by the
        streaming CSV export path so exporting 100,000+ invoices costs a
        bounded, small amount of memory -- roughly one chunk's worth --
        no matter how large the table actually is."""
        with closing(self._conn()) as c:
            last_rowid = 0
            while True:
                rows = c.execute(
                    "SELECT rowid AS _rowid, * FROM invoices WHERE rowid > ? ORDER BY rowid LIMIT ?",
                    (last_rowid, chunk_size),
                ).fetchall()
                if not rows:
                    return
                yield rows
                last_rowid = rows[-1]["_rowid"]

    def get_invoice_by_id(self, inv_id: str) -> Optional[sqlite3.Row]:
        with closing(self._conn()) as c:
            return c.execute("SELECT * FROM invoices WHERE invoice_number=?", (inv_id,)).fetchone()

    def update_notes(self, invoice_number: str, notes: str):
        with closing(self._conn()) as c:
            with c:
                c.execute("UPDATE invoices SET notes=? WHERE invoice_number=?", (notes, invoice_number))

    def save_invoice(self, data: Dict[str, Any], is_update: bool,
                     original_id: Optional[str] = None,
                     payments: Optional[List[Dict]] = None) -> Tuple[bool, Optional[str]]:
        try:
            with closing(self._conn()) as c:
                with c:
                    data = dict(data)
                    data.setdefault("created_ts", time.time())
                    if is_update:
                        nid = data["invoice_number"]
                        if nid != original_id and c.execute(
                            "SELECT 1 FROM invoices WHERE invoice_number=?", (nid,)
                        ).fetchone():
                            return False, f"Invoice number '{nid}' already exists."
                        existing = c.execute(
                            "SELECT created_ts FROM invoices WHERE invoice_number=?", (original_id,)
                        ).fetchone()
                        if existing and existing["created_ts"]:
                            data["created_ts"] = existing["created_ts"]
                        qd = {**data, "original_id": original_id}
                        c.execute(
                            f"UPDATE invoices SET {', '.join(f'{k}=:{k}' for k in data)} WHERE invoice_number=:original_id",
                            qd,
                        )
                        if payments is not None and nid != original_id:
                            c.execute("UPDATE payments SET invoice_number=? WHERE invoice_number=?",
                                      (nid, original_id))
                    else:
                        if c.execute(
                            "SELECT 1 FROM invoices WHERE invoice_number=?", (data["invoice_number"],)
                        ).fetchone():
                            return False, f"Invoice number '{data['invoice_number']}' already exists."
                        c.execute(
                            f"INSERT INTO invoices ({', '.join(data)}) VALUES (:{', :'.join(data)})",
                            data,
                        )
                    final_id = data["invoice_number"]
                    if payments is not None:
                        c.execute("DELETE FROM payments WHERE invoice_number=?", (final_id,))
                        for p in payments:
                            c.execute(
                                "INSERT INTO payments (invoice_number, amount, payment_date, mode, note) "
                                "VALUES (?,?,?,?,?)",
                                (final_id, p.get("amount", 0), p.get("payment_date"),
                                 p.get("mode"), p.get("note", "")),
                            )
                    self._recompute_one(c, final_id)
            return True, None
        except sqlite3.Error as exc:
            return False, str(exc)

    def delete_invoice(self, inv_id: str) -> bool:
        with closing(self._conn()) as c:
            with c:
                c.execute("DELETE FROM invoices WHERE invoice_number=?", (inv_id,))
        return True

    def delete_invoices(self, inv_ids: List[str], chunk_size: int = 200,
                         progress_cb: Optional[Callable[[int, int], None]] = None) -> int:
        """Deletes the given invoices in fixed-size chunks, each its own
        short transaction, rather than one single executemany() covering
        everything at once. Two benefits at scale: any one write-lock is
        only held briefly (friendlier to concurrent readers, e.g. the
        Dashboard refreshing in the background), and a caller can report
        incremental progress via progress_cb(deleted_so_far, total) --
        used by the bulk-delete progress dialog. Returns the number of
        invoices actually deleted."""
        total = len(inv_ids)
        deleted = 0
        for i in range(0, total, chunk_size):
            chunk = inv_ids[i:i + chunk_size]
            with closing(self._conn()) as c:
                with c:
                    c.executemany("DELETE FROM invoices WHERE invoice_number=?", [(x,) for x in chunk])
            deleted += len(chunk)
            if progress_cb:
                progress_cb(deleted, total)
        return deleted

    def get_next_invoice_number(self) -> str:
        """Continue the numbering scheme of the most recently *created* invoice
        (by creation time, falling back to insertion order), NOT the highest
        numeric suffix anywhere in the table.

        BUG FIX: the old version scanned every invoice_number in the table and
        kept whichever had the single highest numeric suffix, regardless of
        its prefix or how long ago it was created. A one-off invoice number in
        a different format (e.g. a number copied from an old system, or a
        client-requested series like "PROJECT-2025-999") would then hijack the
        suggestion for every future invoice forever -- even years later --
        since nothing in the old logic let the normal INV-xxx sequence "win
        back" the suggestion. Anchoring on recency instead means a single
        manually-typed correction immediately restores the right sequence.
        """
        with closing(self._conn()) as c:
            rows = c.execute(
                "SELECT invoice_number FROM invoices ORDER BY created_ts DESC, rowid DESC"
            ).fetchall()
        for r in rows:
            m = re.match(r"^(.*-)(\d+)$", r["invoice_number"])
            if m:
                prefix, num_str = m.group(1), m.group(2)
                return f"{prefix}{int(num_str) + 1:0{len(num_str)}d}"
        return "INV-001"

    def export_to_csv(self, file_path: str, invoice_ids: Optional[List[str]] = None) -> Tuple[bool, Optional[str]]:
        try:
            if invoice_ids:
                with closing(self._conn()) as c:
                    q_marks = ",".join("?" * len(invoice_ids))
                    rows = c.execute(
                        f"SELECT * FROM invoices WHERE invoice_number IN ({q_marks}) "
                        "ORDER BY invoice_date DESC, invoice_number DESC",
                        invoice_ids,
                    ).fetchall()
            else:
                rows = self.get_all_invoices()
            if not rows:
                return False, "No invoices to export."
            with open(file_path, "w", newline="", encoding="utf-8") as fh:
                writer = csv.writer(fh)
                writer.writerow(rows[0].keys())
                writer.writerows(rows)
            return True, None
        except Exception as exc:
            return False, str(exc)

    def get_dashboard_stats(self) -> Dict[str, Any]:
        """Single aggregate query -- O(1) round trip regardless of invoice count."""
        with closing(self._conn()) as c:
            row = c.execute("""
                SELECT
                    COUNT(*)                                                     AS total,
                    COALESCE(SUM(CASE WHEN UPPER(status)='PAID'    THEN 1 ELSE 0 END), 0) AS paid,
                    COALESCE(SUM(CASE WHEN UPPER(status)='DUE'     THEN 1 ELSE 0 END), 0) AS due,
                    COALESCE(SUM(CASE WHEN UPPER(status)='PARTIAL' THEN 1 ELSE 0 END), 0) AS partial,
                    COALESCE(SUM(CASE WHEN is_complimentary=0 THEN grand_total ELSE 0 END), 0) AS total_billed,
                    COALESCE(SUM(CASE WHEN is_complimentary=0 THEN total_paid  ELSE 0 END), 0) AS total_paid,
                    COALESCE(SUM(CASE WHEN is_complimentary=0 THEN balance_due ELSE 0 END), 0) AS balance_due
                FROM invoices
            """).fetchone()
            return dict(row)


# -----------------------------------------------------------------------------
#  Backup, restore and import
# -----------------------------------------------------------------------------
# Every copy of the database goes through SQLite's own online-backup API instead of a raw file copy. A raw copy of a
# WAL-mode database can silently miss recent changes that still sit in the -wal file; the backup API always produces
# a complete, consistent file, and it works while the application has the database open.
_FILE_DIALOG_STATE: Dict[str, Optional[str]] = {"backup_dir": None, "import_dir": None, "asset_dir": None}


def _start_dir(key: str) -> Path:
    """Where a file dialog opens: the folder used last time, else the user's Documents folder."""
    remembered = _FILE_DIALOG_STATE.get(key)
    if remembered and Path(remembered).is_dir():
        return Path(remembered)
    return documents_dir()


def _remember_dir(key: str, chosen_file: Any) -> None:
    _FILE_DIALOG_STATE[key] = str(Path(str(chosen_file)).parent)


def _same_file(a: Path, b: Path) -> bool:
    try:
        return os.path.samefile(a, b)
    except OSError:
        return False


def _overwrite_database(source: Path, dest: Path) -> None:
    """Makes the SQLite database at `dest` an exact copy of the one at `source`."""
    with closing(sqlite3.connect(str(source), timeout=10)) as src, closing(sqlite3.connect(str(dest), timeout=10)) as dst:
        try:
            # SQLite refuses a backup into a WAL-mode database when the two page sizes differ.
            dst.execute("PRAGMA journal_mode = DELETE")
        except sqlite3.Error:
            pass
        src.backup(dst)


def _write_database_copy(source: Path, dest: Path) -> None:
    """Writes a complete single-file copy of the database `source` to `dest`, replacing whatever is there. The copy
    is built under a temporary name beside `dest` first, so a failure half-way never leaves a damaged file behind."""
    source, dest = Path(source), Path(dest)
    if not source.is_file():
        raise FileNotFoundError(f"The database file was not found:\n{source}")
    if dest.exists() and _same_file(source, dest):
        raise ValueError("That is the database Vyaprix is using right now.\nPlease choose a different file.")
    temp = dest.with_name(dest.name + ".part")
    sidecars = ("-wal", "-shm", "-journal")
    try:
        _overwrite_database(source, temp)
        try:
            with closing(sqlite3.connect(str(temp), timeout=10)) as conn:
                conn.execute("PRAGMA journal_mode = DELETE")        # a self-contained file: nothing left in a -wal
        except sqlite3.Error:
            pass
        for suffix in sidecars:                                      # a stale -wal beside the old file must not be replayed
            try:
                dest.with_name(dest.name + suffix).unlink()
            except OSError:
                pass
        os.replace(temp, dest)
    finally:
        for leftover in [temp] + [temp.with_name(temp.name + s) for s in sidecars]:
            try:
                leftover.unlink()
            except OSError:
                pass


def backup_database_to(dest: Path) -> None:
    """Saves a complete copy of the live database to `dest`."""
    _write_database_copy(Path(DB_NAME), Path(dest))


# Columns every Vyaprix database has had since the first release (newer ones are added by DatabaseManager).
_CORE_INVOICE_COLUMNS = frozenset({"invoice_number", "bill_to", "items", "status", "invoice_date"})


def validate_import_database(path: Path) -> Optional[str]:
    """None when `path` is a usable Vyaprix database, otherwise a short human-readable reason why it is not."""
    path = Path(path)
    if not path.is_file():
        return "The file could not be found."
    try:
        with closing(sqlite3.connect(str(path), timeout=10)) as conn:
            tables = {row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type = 'table'")}
            if "invoices" not in tables:
                return "It is a database, but not a Vyaprix one (there is no invoices table in it)."
            missing = sorted(_CORE_INVOICE_COLUMNS - {row[1] for row in conn.execute("PRAGMA table_info(invoices)")})
            if missing:
                return f"It is a database, but not a Vyaprix one (its invoices table has no {', '.join(missing)} column)."
            verdict = conn.execute("PRAGMA quick_check(1)").fetchone()
            if not verdict or verdict[0] != "ok":
                return f"The file failed SQLite's integrity check and may be damaged ({verdict[0] if verdict else 'no result'})."
    except sqlite3.Error as exc:
        return str(exc)
    return None


class DatabaseImportError(Exception):
    """Raised by import_database_from(). `stage` is "backup" (the safety copy could not be made, nothing was touched)
    or "import" (the import failed; `restored` says whether the previous database could be put back)."""

    def __init__(self, message: str, stage: str, restored: bool = True, backup_path: Optional[Path] = None):
        super().__init__(message)
        self.stage = stage
        self.restored = restored
        self.backup_path = backup_path


def import_database_from(source: Path) -> Path:
    """Replaces everything in the live database with the contents of `source`. The current database is copied to a
    timestamped safety file first (its path is returned) and put back if anything goes wrong."""
    source = Path(source)
    live = Path(DB_NAME)
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = DATA_DIR / f"Vyaprix-pre-import-backup-{stamp}.db"
    counter = 2
    while backup_path.exists():         # never overwrite an earlier safety copy (it may even be the file being imported)
        backup_path = DATA_DIR / f"Vyaprix-pre-import-backup-{stamp}-{counter}.db"
        counter += 1
    had_live = live.exists()
    try:
        if had_live:
            _write_database_copy(live, backup_path)
    except Exception as exc:
        raise DatabaseImportError(str(exc), "backup", backup_path=backup_path) from exc
    try:
        _overwrite_database(source, live)
        DatabaseManager(live)       # adds anything an older / smaller file is missing; fails on a file that is not ours
    except Exception as exc:
        restored = False
        try:
            if had_live:
                _overwrite_database(backup_path, live)
                restored = True
        except Exception:
            logger.exception("Could not restore the previous database after a failed import")
        raise DatabaseImportError(str(exc), "import", restored, backup_path) from exc
    return backup_path


def prompt_and_backup_database(parent: Optional[QWidget] = None) -> None:
    """Asks where to save a copy of the live database, then writes it there. Shared by the Settings dialog's
    "Save a Backup Copy" button and the Ctrl+B shortcut, so there is one definition of "back up the database"."""
    default_name = f"Vyaprix-backup-{datetime.date.today().isoformat()}.db"
    path, _ = QFileDialog.getSaveFileName(parent, "Save Database Backup", str(_start_dir("backup_dir") / default_name),
                                          "SQLite Database (*.db)")
    if not path:
        return
    target = Path(path)
    if not target.suffix:
        target = target.with_suffix(".db")
    _remember_dir("backup_dir", target)
    try:
        backup_database_to(target)
    except Exception as exc:
        Messagebox.show_error("Backup Failed", str(exc), parent=parent)
        return
    Messagebox.show_info("Backup Saved", f"Database backed up to:\n{target}", parent=parent)


# =============================================================================
#  10. PDF GENERATION
# =============================================================================
_HEX_COLOR = re.compile(r"#[0-9A-Fa-f]{6}")


def gst_breakdown_text(taxed: List[Dict[str, float]], key: str, rate_share: float) -> str:
    """One line per GST rate for a tooltip: "9% on 1,000.00 = 90.00". `key` is the tax figure to print ("cgst",
    "sgst" or "tax") and `rate_share` the part of the group's rate that figure stands for (half for CGST / SGST)."""
    return "\n".join(f"{g['rate'] * rate_share:g}% on {fmt_money(g['taxable'])} = {fmt_money(g[key])}" for g in taxed)


def normalize_hex_color(value: Any) -> str:
    """'#2563eb' -> '#2563EB'; anything that is not a #RRGGBB colour -> ''."""
    text = str(value or "").strip()
    return text.upper() if _HEX_COLOR.fullmatch(text) else ""


def _pdf_rgb(color: str) -> Tuple[int, int, int]:
    return int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16)


def _pdf_mix(color: str, other: str, share: float) -> str:
    """`color` blended with `other`: a share of 0 is all `color`, 1 is all `other`."""
    return "#%02X%02X%02X" % tuple(round(a + (b - a) * share) for a, b in zip(_pdf_rgb(color), _pdf_rgb(other)))


def _pdf_tint(color: str, share: float) -> str:
    """`color` lightened towards white."""
    return _pdf_mix(color, "#FFFFFF", share)


def _pdf_readable_on(color: str) -> str:
    """The text colour to print on `color`: white on a dark fill, near-black on a light one."""
    def linear(channel: int) -> float:
        v = channel / 255.0
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (linear(c) for c in _pdf_rgb(color))
    return "#111827" if 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.18 else "#FFFFFF"


@dataclass(frozen=True)
class PDFTemplate:
    """The look of a printed invoice: colours, line styles and the shape of the letterhead -- what is printed (and
    where) is the same for every template. Every field defaults to the CLASSIC look, the one invoices have always
    had, so a template lists only what it changes. A template with a `recipe` takes a brand colour: with_accent()
    recomputes its colours around it (the recipes are below)."""
    key: str
    label: str
    blurb: str = ""
    # --- colours --------------------------------------------------------------------------------------------
    name_color: str    = "#1F2937"      # company name
    tagline_color: str = "#6B7280"      # the line under the company name
    title_color: str   = "#111827"      # INVOICE / RECEIPT
    label_color: str   = "#6B7280"      # FROM: / BILL TO:
    rule: str          = "#E5E7EB"      # line under the letterhead
    meta_bg: str       = "#F9FAFB"      # the invoice-number / date box ("" = no fill)
    meta_border: str   = "#E5E7EB"      # its outline ("" = none)
    meta_lines: str    = ""             # hairlines above and below it ("" = none)
    head_bg: str       = "#1F2937"      # item table header ("" = no fill)
    head_fg: str       = "#FFFFFF"
    head_rule: str     = ""             # line under the item header ("" = none)
    row_line: str      = "#F3F4F6"      # line between item rows
    end_line: str      = "#D1D5DB"      # line under the last item row ...
    end_width: float   = 2              # ... and its weight
    zebra: str         = ""             # fill of every second item row ("" = none)
    grid: bool         = False          # a full grid around the item cells ...
    grid_line: str     = "#D1D5DB"      # ... in this colour
    pay_head_bg: str   = "#374151"      # payment history header
    pay_head_fg: str   = "#FFFFFF"
    total_bg: str      = ""             # fill behind the "Invoice Total" row ("" = none)
    total_fg: str      = "#FFFFFF"      # ... and its text
    total_rule: str    = ""             # line above that row ("" = none)
    # --- shape ----------------------------------------------------------------------------------------------
    header: str          = "split"      # "split": name left, title right, a rule below  |  "band": one coloured band
    band_bg: str         = ""           # fill of the band
    bar: str             = ""           # colour of a strip across the top edge of every page ("" = none)
    name_size: float     = 16
    title_size: float    = 18
    title_leading: float = 12
    rule_width: float    = 1
    # --- brand colour ---------------------------------------------------------------------------------------
    accent: str = ""                    # the template's accent colour ("" = it takes no brand colour)
    recipe: str = ""                    # how its colours follow the accent: "modern" | "minimal" | "corporate"

    @property
    def brandable(self) -> bool:
        return self.recipe in _RECIPES and bool(self.accent)

    def with_accent(self, accent: Any = "") -> "PDFTemplate":
        """This template in the brand colour `accent` (#RRGGBB); without a valid colour it keeps its own accent.
        A template with no recipe takes no brand colour and is returned unchanged."""
        recipe = _RECIPES.get(self.recipe)
        if recipe is None:
            return self
        color = normalize_hex_color(accent) or self.accent
        return replace(self, accent=color, **recipe(color))


def _recipe_modern(a: str) -> Dict[str, Any]:
    on = _pdf_readable_on(a)
    return dict(title_color=a, label_color=a, rule=a, bar=a, meta_bg=_pdf_tint(a, 0.94), meta_border=_pdf_tint(a, 0.80),
                head_bg=a, head_fg=on, row_line=_pdf_tint(a, 0.86), end_line=a, zebra=_pdf_tint(a, 0.965),
                pay_head_bg=a, pay_head_fg=on, total_bg=a, total_fg=on)


def _recipe_minimal(a: str) -> Dict[str, Any]:
    return dict(title_color=a, rule=a, meta_lines=_pdf_tint(a, 0.70), head_rule=a, end_line=a, total_rule=a)


def _recipe_corporate(a: str) -> Dict[str, Any]:
    on = _pdf_readable_on(a)
    return dict(band_bg=a, name_color=on, title_color=on, tagline_color=_pdf_mix(on, a, 0.25), label_color=a,
                meta_bg=_pdf_tint(a, 0.95), meta_border=_pdf_tint(a, 0.65), head_bg=a, head_fg=on, grid_line=_pdf_tint(a, 0.55),
                end_line=a, pay_head_bg=a, pay_head_fg=on, total_bg=a, total_fg=on)


_RECIPES: Dict[str, Callable[[str], Dict[str, Any]]] = {
    "modern": _recipe_modern, "minimal": _recipe_minimal, "corporate": _recipe_corporate,
}

PDF_TEMPLATE_CLASSIC = PDFTemplate(
    "classic", "Classic", "The look Vyaprix invoices have always had: a plain letterhead and a dark table header.")
PDF_TEMPLATE_MODERN = PDFTemplate(
    "modern", "Modern", "A bold accent colour: a colour bar on every page, a coloured table header, tinted rows and a "
                        "coloured total.",
    name_color="#111827", title_size=22, title_leading=26, rule_width=2,
    accent="#2563EB", recipe="modern").with_accent()
PDF_TEMPLATE_MINIMAL = PDFTemplate(
    "minimal", "Minimal", "Clean lines and plenty of white space. Light on ink, and it photocopies well.",
    name_color="#111827", name_size=15, title_size=14, title_leading=18, rule_width=0.6, meta_bg="", meta_border="",
    head_bg="", head_fg="#111827", row_line="#E5E7EB", end_width=0.8, pay_head_bg="#F3F4F6", pay_head_fg="#111827",
    accent="#111827", recipe="minimal").with_accent()
PDF_TEMPLATE_CORPORATE = PDFTemplate(
    "corporate", "Corporate", "A formal look: a coloured header band, a ruled table and a boxed total.",
    header="band", name_size=17, title_size=20, title_leading=24, grid=True,
    accent="#1E3A5F", recipe="corporate").with_accent()

PDF_TEMPLATES: Dict[str, PDFTemplate] = {t.key: t for t in (
    PDF_TEMPLATE_CLASSIC, PDF_TEMPLATE_MODERN, PDF_TEMPLATE_MINIMAL, PDF_TEMPLATE_CORPORATE)}


def pdf_template_for(config: Dict[str, Any], data: Optional[Dict[str, Any]] = None) -> PDFTemplate:
    """The template an invoice prints in: its own choice, else the one picked in Settings, else Classic -- in the
    brand colour from Settings when the template takes one."""
    settings = config.get("pdf_settings") or {}
    chosen = PDF_TEMPLATE_CLASSIC
    for key in ((data or {}).get("pdf_template"), settings.get("template")):
        if key and key in PDF_TEMPLATES:
            chosen = PDF_TEMPLATES[key]
            break
    return chosen.with_accent(settings.get("accent_color"))


class PDFGenerator:
    """Draws one invoice as an A4 PDF.

    The page: a letterhead, the two parties, a strip of reference details (number, dates, terms, PO, place of supply),
    the item table and then the totals -- with the amount in words and, when GST is charged at several rates, a GST
    summary by rate beside them -- followed by the payment history, notes, terms and the payment / signature block.

    Every width comes from the printable width of the page, so the layout always fills A4 exactly. A long invoice flows
    onto more pages: the item table repeats its header, a very long item description may split between pages, the
    totals are kept in one piece, a running header names the invoice on every page after the first, and every page
    says "Page x of y". What the page looks like (colours, letterhead, table style) comes from the template."""

    MARGIN_X, MARGIN_TOP, MARGIN_BOTTOM = 1.5, 1.8, 2.3           # centimetres
    SUMMARY_SHARE = 0.48                                          # the totals box: a share of the printable width
    # Auto-fit: an invoice that runs only a little onto a second page is printed again with tighter gaps and row padding
    # (never smaller type) and kept when that saves a page. "A little" = the last page is filled less than this share.
    TIGHT_SCALE = 0.55
    TIGHT_WHEN_LAST_PAGE_BELOW = 0.30

    def __init__(self, data: Dict[str, Any], config: Dict[str, Any],
                 payments: Optional[List[Dict]] = None, out_dir: Optional[Path] = None):
        self.data     = data
        self.config   = config
        self.payments = payments or []
        self.tpl      = pdf_template_for(config, data)
        self.font     = PDF_FONT
        self.bold     = f"{PDF_FONT}-Bold" if PDF_FONT != "Helvetica" else "Helvetica-Bold"

        # --- Smart File Naming to bypass Windows file locks ---
        # Defensive re-sanitize: invoice_number is validated at save time now,
        # but older rows (saved before that check existed) could still carry
        # unsafe characters, so never trust it blindly as a path component.
        raw_number    = str(data.get('invoice_number') or 'NA').strip()
        safe_number   = INVALID_FILENAME_CHARS.sub('_', raw_number).strip('. ') or "NA"
        base_filename = f"Invoice-{safe_number}"
        target_dir    = out_dir or INVOICES_DIR
        self.filename = target_dir / f"{base_filename}.pdf"

        counter = 1
        while self.filename.exists():
            try:
                # Test if we can write to the existing file without altering it yet
                with open(self.filename, 'a'):
                    pass
                break  # If successful, no PermissionError is thrown; we are clear to overwrite.
            except PermissionError:
                # If Windows says the file is locked, append a counter and try the new name.
                self.filename = target_dir / f"{base_filename} ({counter}).pdf"
                counter += 1
        # ------------------------------------------------------

        self.is_gst   = bool(data.get("is_gst_invoice", 1))
        self.tax_mode = (data.get("tax_mode") or TAX_MODE_OVERALL) if self.is_gst else TAX_MODE_OVERALL
        self.styles   = self._setup_styles()
        self.page_height = A4[1] - (self.MARGIN_TOP + self.MARGIN_BOTTOM) * cm      # what a page can hold, in points
        self.tight      = False                  # laying out with tighter gaps (the auto-fit pass)
        self.tightened  = False                  # the tight layout is the one that was printed
        self.page_count = 0                      # pages of the last layout
        self._room_left = 0.0                    # room left under the last flowable of the last layout, in points

    def _g(self, points: float) -> float:
        """A gap or padding: as designed, or scaled down in the auto-fit pass."""
        return points * self.TIGHT_SCALE if self.tight else points

    def _asset(self, key: str) -> Optional[Path]:
        configured = self.config.get("asset_paths", {}).get(key)
        if not configured:
            return None

        path = Path(configured)
        if not path.is_absolute():
            path = DATA_DIR / path

        try:
            path = path.resolve()
            path.relative_to(DATA_DIR.resolve())
        except (OSError, ValueError):
            logger.warning("Ignoring asset outside Vyaprix's ProgramData folder: %s", path)
            return None

        if path.is_file():
            return path
        return None

    def _setup_styles(self) -> Dict[str, ParagraphStyle]:
        ss, base, bold, t = getSampleStyleSheet(), self.font, self.bold, self.tpl
        ss.add(ParagraphStyle("NC",          fontName=base, fontSize=9,  parent=ss["Normal"],   leading=11))
        ss.add(ParagraphStyle("Right",       parent=ss["NC"], alignment=TA_RIGHT))
        ss.add(ParagraphStyle("Center",      parent=ss["NC"], alignment=TA_CENTER))
        ss.add(ParagraphStyle("Small",       parent=ss["NC"], fontSize=7.5, textColor=colors.HexColor("#6B7280")))
        ss.add(ParagraphStyle("SmallRight",  parent=ss["Small"], alignment=TA_RIGHT))
        ss.add(ParagraphStyle("BillTo",      parent=ss["NC"], leading=11.5))
        ss.add(ParagraphStyle("PartyLabel",  parent=ss["NC"], fontName=bold, fontSize=7, leading=9,
                              textColor=colors.HexColor(t.label_color)))
        ss.add(ParagraphStyle("PartyName",   parent=ss["NC"], fontName=bold, fontSize=10.5, leading=13))
        ss.add(ParagraphStyle("ItemDesc",    parent=ss["NC"], textColor=colors.HexColor("#6B7280"), fontSize=8, leading=10))
        ss.add(ParagraphStyle("CompanyName", fontName=bold,   fontSize=t.name_size, textColor=colors.HexColor(t.name_color),
                              leading=t.name_size + 2))
        ss.add(ParagraphStyle("Tagline",     parent=ss["NC"], fontSize=8.5, textColor=colors.HexColor(t.tagline_color)))
        ss.add(ParagraphStyle("InvoiceTitle",fontName=bold,   fontSize=t.title_size, alignment=TA_RIGHT,
                              textColor=colors.HexColor(t.title_color), leading=t.title_leading))
        ss.add(ParagraphStyle("NoteBox",     parent=ss["NC"], fontSize=8, leading=10.5, textColor=colors.HexColor("#374151")))
        ss.add(ParagraphStyle("NoteText",    parent=ss["NC"], fontSize=8, leading=10.5, textColor=colors.HexColor("#4B5563")))
        ss.add(ParagraphStyle("Caption",     parent=ss["NC"], fontName=bold, fontSize=7.5, leading=9.5,
                              textColor=colors.HexColor(t.label_color), keepWithNext=1))
        ss.add(ParagraphStyle("PayLabel",    parent=ss["NC"], fontSize=8, leading=10, textColor=colors.HexColor("#6B7280")))
        ss.add(ParagraphStyle("PayValue",    parent=ss["NC"], fontSize=8.5, leading=10.5))
        return ss

    def _text_width(self, text: str, size: float, bold: bool = False) -> float:
        """How wide `text` prints, in points -- used to size the table columns from what they hold."""
        return pdfmetrics.stringWidth(text, self.bold if bold else self.font, size)

    def _head_styles(self, size: float = 8.0) -> Tuple[ParagraphStyle, ParagraphStyle, ParagraphStyle]:
        """Table-header text: left, right and centred, in the template's header colour."""
        left = ParagraphStyle("HS", parent=self.styles["NC"], fontSize=size, leading=size + 2,
                              textColor=colors.HexColor(self.tpl.head_fg))
        return left, ParagraphStyle("HSR", parent=left, alignment=TA_RIGHT), ParagraphStyle("HSC", parent=left, alignment=TA_CENTER)

    # ------------------------------------------------------------------ #
    #  Letterhead, parties, reference details
    # ------------------------------------------------------------------ #
    def _build_header(self, width: float) -> Table:
        comp  = self.config.get("company_details", {})
        left  = [Paragraph(safe_markup(compose_your_name(self.config)), self.styles["CompanyName"])]
        if d := comp.get("description"):
            left += [Spacer(1, 2), Paragraph(safe_markup(d), self.styles["Tagline"])]

        status    = self.data.get("status")
        doc_title = ("RECEIPT"       if status in (InvoiceStatus.PAID.value, InvoiceStatus.OVERPAID.value) else
                     "COMPLIMENTARY" if status == InvoiceStatus.COMPLIMENTARY.value else "INVOICE")

        # --- BARCODE GENERATION ---
        # Code128 only supports a limited ASCII character set, and invoice
        # numbers are user-editable free text -- so this must never be able
        # to take down the whole PDF. If it fails for any reason, we simply
        # skip the barcode instead of aborting invoice generation.
        title_block_rows = []
        if BARCODE_AVAILABLE:
            try:
                code128 = barcode.get('code128', str(self.data.get('invoice_number') or '000'), writer=ImageWriter())
                buffer = io.BytesIO()
                code128.write(buffer, options={"write_text": False, "module_width": 0.2, "module_height": 5.0})
                buffer.seek(0)
                title_block_rows.append([Image(buffer, width=3.5*cm, height=0.8*cm)])
            except Exception as exc:
                logger.warning(f"Skipping barcode for invoice "
                            f"{self.data.get('invoice_number')!r}: {exc}")
        # --------------------------

        # Stack the barcode ABOVE the status title (barcode omitted if it failed)
        title_block_rows.append([Paragraph(doc_title, self.styles["InvoiceTitle"])])
        title_block = Table(title_block_rows, style=[("ALIGN", (0,0), (-1,-1), "RIGHT")])

        t = self.tpl
        band = t.header == "band"
        half = width / 2.0
        if logo := self._asset("logo"):
            try:
                img = Image(str(logo), 2.0*cm, 2.0*cm, kind="proportional")
                logo_style = [("VALIGN",(0,0),(-1,-1),"MIDDLE"), ("LEFTPADDING",(0,0),(-1,-1),0)]
                if band:        # on the coloured band the logo sits on a white tile, so any logo stays readable
                    logo_style += [("BACKGROUND",(0,0),(0,0),colors.white), ("ALIGN",(0,0),(0,0),"CENTER"),
                                   ("LEFTPADDING",(1,0),(1,0),10)]
                left_side = Table([[img, left]], colWidths=[2.5*cm, half - 2.5*cm - (24 if band else 0)], style=logo_style)
            except Exception:
                left_side = left
        else:
            left_side = left

        if band:
            return Table([[left_side, title_block]], colWidths=[half, half], style=[
                ("BACKGROUND",(0,0),(-1,-1),colors.HexColor(t.band_bg)),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"), ("ALIGN",(1,0),(1,-1),"RIGHT"),
                ("LEFTPADDING",(0,0),(-1,-1),12),   ("RIGHTPADDING",(0,0),(-1,-1),12),
                ("TOPPADDING",(0,0),(-1,-1),self._g(12)),    ("BOTTOMPADDING",(0,0),(-1,-1),self._g(12)),
            ])
        return Table([[left_side, title_block]], colWidths=[half, half], style=[
            ("VALIGN",(0,0),(-1,-1),"TOP"), ("ALIGN",(1,0),(1,-1),"RIGHT"),
            ("LEFTPADDING",(0,0),(-1,-1),0),   ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("LINEBELOW",(0,0),(-1,-1),t.rule_width,colors.HexColor(t.rule)),
            ("BOTTOMPADDING",(0,0),(-1,-1),self._g(10)),
        ])

    def _build_parties(self, width: float) -> Table:
        """BILL TO and FROM, side by side."""
        cg = yg = ""
        if self.is_gst:
            if cgstin := self.data.get("client_gstin"):
                cg = f"GSTIN: {safe_markup(cgstin)}"
            if ygstin := _real_value(self.config.get("company_details", {}).get("gstin"), _SAMPLE_GSTIN):
                yg = f"GSTIN: {safe_markup(ygstin)}"
        # Escape first, THEN turn newlines into <br/> -- safe_markup() only
        # touches &/</>, so it won't disturb the newlines we're about to replace.
        your_c   = safe_markup(self.data.get("your_contact")).replace("\n", "<br/>")
        client_c = safe_markup(self.data.get("client_contact")).replace("\n", "<br/>")

        def block(label: str, name: str, *lines: str) -> List[Any]:
            cell = [Paragraph(label, self.styles["PartyLabel"]), Paragraph(name, self.styles["PartyName"])]
            body = "<br/>".join(line for line in lines if line)
            if body:
                cell.append(Paragraph(body, self.styles["BillTo"]))
            return cell

        bill = block("BILL TO", safe_markup(self.data.get("bill_to")), client_c, cg)
        mine = block("FROM", safe_markup(self.data.get("your_name")), your_c, yg)
        return Table([[bill, mine]], colWidths=[width / 2.0, width / 2.0], style=[
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(0,0),0),  ("RIGHTPADDING",(0,0),(0,0),14),
            ("LEFTPADDING",(1,0),(1,0),14), ("RIGHTPADDING",(1,0),(1,0),0),
        ])

    def _build_meta(self, width: float) -> Table:
        """The reference details as a strip of labelled cells: invoice number and date, due date, payment terms and,
        when they are used, the customer's PO number, the place of supply and the reverse-charge flag."""
        date_str = parse_date(self.data.get("invoice_date"), DB_DATE_FORMAT, DISPLAY_DATE_FORMAT)
        free     = not self.data.get("is_complimentary")
        cells    = [("Invoice #", safe_markup(self.data.get("invoice_number"))), ("Invoice Date", date_str)]
        if free and self.data.get("due_date"):
            cells.append(("Due Date", parse_date(self.data.get("due_date"), DB_DATE_FORMAT, DISPLAY_DATE_FORMAT)))
        if free:
            cells.append(("Payment Terms", safe_markup(self.data.get("payment_terms") or "N/A")))
            if po := str(self.data.get("po_number") or "").strip():
                cells.append(("PO / Ref #", safe_markup(po)))
            if self.is_gst and (place := str(self.data.get("place_of_supply") or "").strip()):
                cells.append(("Place of Supply", safe_markup(place)))
            if self.is_gst and self.data.get("reverse_charge"):
                cells.append(("Reverse Charge", "Yes"))

        t = self.tpl
        count = len(cells)
        per_row = count if count <= 4 else (3 if count <= 6 else 4)        # 5-6 cells: two even rows of three
        rows: List[List[Any]] = []
        for start in range(0, count, per_row):
            chunk = cells[start:start + per_row]
            row = [Paragraph(f"<font size='7' color='{t.label_color}'>{label}</font><br/><b>{value}</b>", self.styles["NC"])
                   for label, value in chunk]
            rows.append(row + [""] * (per_row - len(row)))
        style: List[Tuple] = []
        if t.meta_border:
            style.append(("BOX",(0,0),(-1,-1),1,colors.HexColor(t.meta_border)))
        if t.meta_bg:
            style.append(("BACKGROUND",(0,0),(-1,-1),colors.HexColor(t.meta_bg)))
        if t.meta_lines:
            style += [("LINEABOVE",(0,0),(-1,0),0.6,colors.HexColor(t.meta_lines)),
                      ("LINEBELOW",(0,-1),(-1,-1),0.6,colors.HexColor(t.meta_lines))]
        pad = 8 if (t.meta_border or t.meta_bg) else 0          # an unboxed strip lines up with the page margin
        style += [("VALIGN",(0,0),(-1,-1),"TOP"), ("TOPPADDING",(0,0),(-1,-1),self._g(6)), ("BOTTOMPADDING",(0,0),(-1,-1),self._g(6)),
                  ("LEFTPADDING",(0,0),(-1,-1),pad), ("RIGHTPADDING",(0,0),(-1,-1),4)]
        if not pad:
            style.append(("LEFTPADDING",(1,0),(-1,-1),8))
        return Table(rows, colWidths=[width / per_row] * per_row, style=TableStyle(style))

    # ------------------------------------------------------------------ #
    #  Items
    # ------------------------------------------------------------------ #
    def _items(self) -> List[Dict[str, Any]]:
        try:
            items = json.loads(self.data.get("items") or "[]")
        except (TypeError, ValueError):
            items = []
        return items if isinstance(items, list) else []

    def _items_style(self, last_row: int) -> List[Tuple]:
        """The item table's look (the same for the compact and the full table), from the template."""
        t = self.tpl
        style: List[Tuple] = []
        if t.head_bg:
            style.append(("BACKGROUND",(0,0),(-1,0),colors.HexColor(t.head_bg)))
        style += [
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LINEBELOW",(0,1),(-1,-2),1,colors.HexColor(t.row_line)),
            ("LINEBELOW",(0,-1),(-1,-1),t.end_width,colors.HexColor(t.end_line)),
            ("TOPPADDING",(0,0),(-1,-1),self._g(5)),   ("BOTTOMPADDING",(0,0),(-1,-1),self._g(5)),
            ("LEFTPADDING",(0,0),(-1,-1),6),  ("RIGHTPADDING",(0,0),(-1,-1),6),
        ]
        if t.head_rule:
            style.append(("LINEBELOW",(0,0),(-1,0),1.2,colors.HexColor(t.head_rule)))
        if t.zebra:
            style += [("BACKGROUND",(0,r),(-1,r),colors.HexColor(t.zebra)) for r in range(2, last_row + 1, 2)]
        if t.grid:
            style.append(("GRID",(0,0),(-1,-1),0.5,colors.HexColor(t.grid_line)))
        return style

    def _items_table(self, rows: List[List[Any]], widths: List[float], style: List[Tuple]) -> Table:
        """The item table, its header repeated on every page. A row is never cut between two pages -- unless it is
        taller than a whole page (an enormous description), when it has to be."""
        table = Table(rows, colWidths=widths, repeatRows=1, style=TableStyle(style))
        table.wrap(sum(widths), self.page_height)
        if max(getattr(table, "_rowHeights", None) or [0]) > self.page_height - 2 * cm:
            table = Table(rows, colWidths=widths, repeatRows=1, splitInRow=1, style=TableStyle(style))
        return table

    def _build_items(self, width: float) -> Table:
        items = self._items()
        # Invoices made only of plain "title + amount" lines keep the compact two-column table they have always had;
        # as soon as a line has a quantity, unit, HSN/SAC code or its own GST rate the full table is printed.
        if any(item_is_extended(i) for i in items):
            return self._build_items_full(items, width)
        hs, hsr, _hsc = self._head_styles()
        rows = [[Paragraph("<b>ITEM / DESCRIPTION</b>", hs), Paragraph("<b>AMOUNT</b>", hsr)]]
        for item in items:
            title  = item.get("title")       or ""
            desc   = item.get("description") or ""
            amount = safe_float(item.get("amount"))
            if not title.strip() and not desc.strip() and amount == 0:
                rows.append([[Paragraph("<br/>", self.styles["NC"])], Paragraph("", self.styles["Right"])])
            else:
                cell = [Paragraph(f"<b>{safe_markup(title)}</b>", self.styles["NC"])]
                if desc.strip():
                    cell.append(Paragraph(safe_markup(desc), self.styles["ItemDesc"]))
                rows.append([cell, Paragraph(f"{CURRENCY_SYMBOL}{amount:,.2f}", self.styles["Right"])])
        # the amount column is as wide as the biggest figure needs (never less than a fifth of the page)
        figure = max([self._text_width(f"{CURRENCY_SYMBOL}{safe_float(i.get('amount')):,.2f}", 9) for i in items] + [0.0])
        amount_w = min(max(width * 0.20, figure + 16), width * 0.5)
        return self._items_table(rows, [width - amount_w, amount_w],
                                 [*self._items_style(len(rows) - 1), ("ALIGN",(1,0),(1,-1),"RIGHT")])

    # the share of the page each number column of the full item table asks for, in per cent -- a column takes more
    # when its figures need it (see _build_items_full)
    _ITEM_COLUMN_SHARES = {"no": 4.5, "hsn": 10.5, "qty": 10.5, "rate": 12.5, "disc": 12.0, "gst": 6.5, "amount": 14.5}

    def _build_items_full(self, items: List[Dict[str, Any]], width: float) -> Table:
        """# | item (+ description) | HSN/SAC | quantity + unit | rate | discount | GST % | amount. The HSN/SAC column
        only appears when a line has a code, the discount column when a line has a discount and the GST column when GST
        is charged per item. The number columns are as wide as their longest figure needs (measured, never guessed) and
        the item column takes the rest, so no figure is ever cut and the table always spans the page exactly."""
        per_item = self.is_gst and self.tax_mode == TAX_MODE_PER_ITEM
        default_rate = safe_float(self.data.get("tax_rate"))
        show_hsn  = any(str(i.get("hsn") or "").strip() for i in items)
        show_disc = any(safe_float(i.get("discount")) > 0 for i in items)
        cols: List[Tuple[str, str]] = [("no", "#"), ("item", "ITEM / DESCRIPTION")]
        if show_hsn:
            cols.append(("hsn", "HSN/SAC"))
        cols += [("qty", "QTY"), ("rate", "RATE")]
        if show_disc:
            cols.append(("disc", "DISC."))
        if per_item:
            cols.append(("gst", "GST"))
        cols.append(("amount", "AMOUNT"))

        # --- every line as plain figures first, so the columns can be sized from what they hold -----------------
        lines: List[Optional[Dict[str, Any]]] = []
        number = 0
        for item in items:
            title = str(item.get("title") or "")
            desc  = str(item.get("description") or "")
            gross = safe_float(item.get("amount"))
            if not title.strip() and not desc.strip() and gross == 0:
                lines.append(None)                                        # a blank line stays a blank row
                continue
            number += 1
            qty  = item_qty(item)
            unit = str(item.get("unit") or "").strip()
            raw_rate = item.get("rate")
            rate = safe_float(raw_rate) if raw_rate not in (None, "") else (gross / qty if qty else gross)
            line: Dict[str, Any] = {
                "title": title, "desc": desc, "no": str(number),
                "hsn": str(item.get("hsn") or "").strip() or "-",
                "qty": fmt_qty(qty) + (f" {unit}" if unit else ""),
                "rate": f"{CURRENCY_SYMBOL}{rate:,.2f}", "amount": f"{CURRENCY_SYMBOL}{gross:,.2f}", "disc_pct": "",
            }
            disc = safe_float(item.get("discount"))
            line["disc"] = f"{CURRENCY_SYMBOL}{disc:,.2f}" if disc > 0 else "-"
            if disc > 0 and item.get("disc_type") == "percent":
                line["disc_pct"] = f"({safe_float(item.get('disc_value')):g}%)"
            raw_tax = item.get("tax_rate")
            line["gst"] = f"{(default_rate if raw_tax in (None, '') else safe_float(raw_tax, default_rate)):g}%"
            lines.append(line)

        # --- column widths ---------------------------------------------------------------------------------------
        size, pad = 9.0, 6.0
        fixed: Dict[str, float] = {}
        for attempt in (0, 1):
            fixed = {}
            for key, header in cols:
                if key == "item":
                    continue
                longest = self._text_width(header, size - 1, bold=True)
                for line in lines:
                    if line:
                        longest = max(longest, self._text_width(line[key], size),
                                      self._text_width(line["disc_pct"], size - 1) if key == "disc" else 0.0)
                share = self._ITEM_COLUMN_SHARES[key] / 100.0 * width * (1.0 if attempt == 0 else 0.92)
                fixed[key] = max(share, longest + 2 * pad + 2)
            item_w = width - sum(fixed.values())
            if item_w >= max(width * 0.30, 140.0):
                break
            size, pad = 8.0, 4.0                    # the figures need so much room that the item column would be too
                                                    # narrow: print them a size smaller, with less padding
        if item_w < 60.0:                           # extreme figures even at the smaller size: squeeze them all
            scale = (width - 60.0) / sum(fixed.values())
            fixed = {k: v * scale for k, v in fixed.items()}
            item_w = 60.0
        widths = [item_w if key == "item" else fixed[key] for key, _ in cols]

        # --- the cells -------------------------------------------------------------------------------------------
        hs, hsr, hsc = self._head_styles()
        head_style = {"no": hsc, "item": hs, "hsn": hsc}
        num  = ParagraphStyle("ItemNum", parent=self.styles["NC"], fontSize=size, leading=size + 2, alignment=TA_RIGHT)
        mid  = ParagraphStyle("ItemMid", parent=num, alignment=TA_CENTER)
        rows = [[Paragraph(f"<b>{header}</b>", head_style.get(key, hsr)) for key, header in cols]]
        for line in lines:
            if line is None:
                rows.append([[Paragraph("<br/>", self.styles["NC"])] if key == "item" else Paragraph("", num) for key, _ in cols])
                continue
            cell = [Paragraph(f"<b>{safe_markup(line['title'])}</b>", self.styles["NC"])]
            if line["desc"].strip():
                cell.append(Paragraph(safe_markup(line["desc"]), self.styles["ItemDesc"]))
            disc_cell = safe_markup(line["disc"])
            if line["disc_pct"]:
                disc_cell += f"<br/><font size='{size - 1:g}' color='#6B7280'>{safe_markup(line['disc_pct'])}</font>"
            values: Dict[str, Any] = {
                "no": Paragraph(line["no"], mid), "item": cell, "hsn": Paragraph(safe_markup(line["hsn"]), mid),
                "qty": Paragraph(safe_markup(line["qty"]), num), "rate": Paragraph(safe_markup(line["rate"]), num),
                "disc": Paragraph(disc_cell, num), "gst": Paragraph(safe_markup(line["gst"]), num),
                "amount": Paragraph(safe_markup(line["amount"]), num),
            }
            rows.append([values[key] for key, _ in cols])
        style = [*self._items_style(len(rows) - 1)]
        keys = [key for key, _ in cols]
        first_number = keys.index("qty")
        style += [("ALIGN",(first_number,0),(-1,-1),"RIGHT"), ("ALIGN",(0,0),(0,-1),"CENTER"),
                  ("LEFTPADDING",(2,0),(-1,-1),pad), ("RIGHTPADDING",(2,0),(-1,-1),pad),
                  ("LEFTPADDING",(0,0),(0,-1),pad), ("RIGHTPADDING",(0,0),(0,-1),pad)]
        if show_hsn:
            style.append(("ALIGN",(2,0),(2,-1),"CENTER"))
        return self._items_table(rows, widths, style)

    # ------------------------------------------------------------------ #
    #  Totals
    # ------------------------------------------------------------------ #
    def _build_totals(self, width: float, totals: Dict[str, Any]) -> Table:
        """The totals box. GST is always one CGST and one SGST line (or one IGST line) -- however many rates the
        invoice has; the rate-wise split goes into the GST summary beside it."""
        right = self.styles["Right"]
        total_paid = sum(safe_float(p.get("amount")) for p in self.payments)
        _status, balance, credit = DatabaseManager.compute_payment_status(totals["grand_total"], total_paid)

        rows = [[Paragraph("Subtotal", right), Paragraph(fmt_money(totals['subtotal']), right)]]
        if totals["total_discount"] > 0:
            rows.append([Paragraph("Discount", right), Paragraph(f"- {fmt_money(totals['total_discount'])}", right)])

        # --- IGST vs CGST/SGST PDF Rendering Logic ---
        taxed = [g for g in totals["tax_groups"] if g["rate"] > 0]
        if self.is_gst and taxed:
            rows.append([Paragraph("Taxable Amount", right), Paragraph(fmt_money(totals['taxable']), right)])
            several = len(taxed) > 1
            if self.data.get("gst_type") == "IGST":
                rows.append([Paragraph("IGST" if several else f"IGST ({taxed[0]['rate']:g}%)", right),
                             Paragraph(fmt_money(totals["tax"]), right)])
            else:
                half_rate = taxed[0]["rate"] / 2.0
                rows += [
                    [Paragraph("CGST" if several else f"CGST ({half_rate:g}%)", right),
                     Paragraph(fmt_money(totals["cgst"]), right)],
                    [Paragraph("SGST" if several else f"SGST ({half_rate:g}%)", right),
                     Paragraph(fmt_money(totals["sgst"]), right)],
                ]
        # ---------------------------------------------
        if totals["round_off"]:
            sign = "+" if totals["round_off"] > 0 else "-"
            rows.append([Paragraph("Round Off", right), Paragraph(f"{sign} {fmt_money(abs(totals['round_off']))}", right)])

        t = self.tpl
        total_row = len(rows)
        if t.total_bg:                                       # the total on a coloured bar
            rows.append([Paragraph(f"<font color='{t.total_fg}'><b>Invoice Total</b></font>", right),
                         Paragraph(f"<font color='{t.total_fg}'><b>{fmt_money(totals['grand_total'])}</b></font>", right)])
        else:
            rows.append([Paragraph("<b>Invoice Total</b>", right),
                         Paragraph(f"<b>{fmt_money(totals['grand_total'])}</b>", right)])
        if total_paid > 0:
            rows.append([Paragraph("Total Paid", right), Paragraph(f"- {fmt_money(total_paid)}", right)])

        # Balance Due is never negative -- an overpayment prints as
        # "Advance / Credit" instead, in green, the same way the
        # invoice-form badge shows it on screen.
        if credit > 0:
            bc = "#059669"
            rows.append([Paragraph("<b>Advance / Credit</b>", right),
                         Paragraph(f"<font color='{bc}'><b>{fmt_money(credit)}</b></font>", right)])
        else:
            bc = "#059669" if balance <= 0 else "#DC2626"
            rows.append([Paragraph("<b>Balance Due</b>", right),
                         Paragraph(f"<font color='{bc}'><b>{fmt_money(balance)}</b></font>", right)])
        style: List[Tuple] = [
            ("ALIGN",(0,0),(-1,-1),"RIGHT"), ("LEFTPADDING",(0,0),(-1,-1),10),
            ("TOPPADDING",(0,0),(-1,-1),self._g(3)),  ("BOTTOMPADDING",(0,0),(-1,-1),self._g(3)),
            ("BACKGROUND",(0,-1),(1,-1),colors.HexColor("#FEF2F2" if (balance > 0 and credit <= 0) else "#ECFDF5")),
            ("GRID",(0,-1),(1,-1),1,colors.HexColor("#E5E7EB")),
        ]
        if t.total_bg:
            style.append(("BACKGROUND",(0,total_row),(1,total_row),colors.HexColor(t.total_bg)))
        if t.total_rule:
            style.append(("LINEABOVE",(0,total_row),(1,total_row),0.8,colors.HexColor(t.total_rule)))
        return Table(rows, colWidths=[width * 0.58, width * 0.42], hAlign="RIGHT", style=TableStyle(style))

    def _build_gst_summary(self, width: float, totals: Dict[str, Any]) -> Optional[Table]:
        """Tax by rate -- the taxable value and the CGST + SGST (or IGST) at each rate. Printed only when GST is charged
        at more than one rate, which is why the totals box can show a single CGST and a single SGST line."""
        taxed = [g for g in totals["tax_groups"] if g["rate"] > 0]
        if not (self.is_gst and not self.data.get("is_complimentary") and len(taxed) > 1):
            return None
        igst = self.data.get("gst_type") == "IGST"
        t = self.tpl
        head_bg, head_fg = (t.head_bg, t.head_fg) if t.head_bg else (t.pay_head_bg, t.pay_head_fg)
        small, small_right = self.styles["Small"], self.styles["SmallRight"]
        ink = ParagraphStyle("GSTHead", parent=small, textColor=colors.HexColor(head_fg), fontName=self.bold)
        ink_right = ParagraphStyle("GSTHeadR", parent=ink, alignment=TA_RIGHT)
        bold_right = ParagraphStyle("GSTTotalR", parent=small_right, fontName=self.bold, textColor=colors.HexColor("#111827"))
        bold_left = ParagraphStyle("GSTTotalL", parent=small, fontName=self.bold, textColor=colors.HexColor("#111827"))

        def money(value: float) -> str:
            return f"{value:,.2f}"

        head = [Paragraph("RATE", ink), Paragraph(f"TAXABLE ({CURRENCY_SYMBOL})", ink_right)]
        head += [Paragraph("IGST", ink_right)] if igst else [Paragraph("CGST", ink_right), Paragraph("SGST", ink_right)]
        rows = [head]
        for g in taxed:
            row = [Paragraph(f"{g['rate']:g}%", small), Paragraph(money(g["taxable"]), small_right)]
            row += [Paragraph(money(g["tax"]), small_right)] if igst else [Paragraph(money(g["cgst"]), small_right),
                                                                              Paragraph(money(g["sgst"]), small_right)]
            rows.append(row)
        total = [Paragraph("Total", bold_left), Paragraph(money(sum(g["taxable"] for g in taxed)), bold_right)]
        total += [Paragraph(money(totals["tax"]), bold_right)] if igst else [Paragraph(money(totals["cgst"]), bold_right),
                                                                              Paragraph(money(totals["sgst"]), bold_right)]
        rows.append(total)
        shares = [0.18, 0.40, 0.42] if igst else [0.15, 0.31, 0.27, 0.27]
        style: List[Tuple] = [
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor(head_bg)),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LINEBELOW",(0,1),(-1,-2),0.5,colors.HexColor(t.row_line)),
            ("LINEABOVE",(0,-1),(-1,-1),0.8,colors.HexColor(t.end_line)),
            ("LINEBELOW",(0,-1),(-1,-1),0.8,colors.HexColor(t.end_line)),
            ("TOPPADDING",(0,0),(-1,-1),self._g(3)), ("BOTTOMPADDING",(0,0),(-1,-1),self._g(3)),
            ("LEFTPADDING",(0,0),(-1,-1),5), ("RIGHTPADDING",(0,0),(-1,-1),5),
        ]
        return Table(rows, colWidths=[width * s_ for s_ in shares], style=TableStyle(style))

    def _build_amount_in_words(self, grand_total: float) -> Optional[List[Any]]:
        """"Amount in words: Rupees ... Only" under the totals -- when Settings turns it on, for rupee invoices."""
        if not self.config.get("pdf_settings", {}).get("amount_in_words", False) or self.data.get("is_complimentary"):
            return None
        words = amount_in_words(grand_total, CURRENCY_SYMBOL)
        if not words:
            return None
        return [Paragraph("Amount in words", self.styles["Caption"]), Spacer(1, 2),
                Paragraph(safe_markup(words), self.styles["NoteBox"])]

    def _build_summary(self, width: float) -> List[Any]:
        """The totals box on the right; the amount in words and the GST summary, when there are any, on the left."""
        right_w = width * self.SUMMARY_SHARE
        if self.data.get("is_complimentary"):
            right = self.styles["Right"]
            box = Table([[Paragraph("<b>Total</b>", right), Paragraph("<b>No Charge</b>", right)]],
                        colWidths=[right_w * 0.58, right_w * 0.42], hAlign="RIGHT",
                        style=TableStyle([("ALIGN",(0,0),(-1,-1),"RIGHT"), ("LEFTPADDING",(0,0),(-1,-1),10)]))
            return [KeepTogether([box])]
        # Same calculation engine the invoice form uses live (and the database stores), so the
        # printed PDF can never show different numbers than what the
        # person saw on screen when they saved.
        totals = DatabaseManager.totals_for(self.data, self._items())
        box = self._build_totals(right_w, totals)
        gap = 0.6 * cm
        left_w = width - right_w - gap
        left: List[Any] = []
        if words := self._build_amount_in_words(totals["grand_total"]):
            left += words
        if summary := self._build_gst_summary(left_w, totals):
            left += ([Spacer(1, self._g(0.3 * cm))] if left else []) + [Paragraph("GST Summary", self.styles["Caption"]),
                                                               Spacer(1, 2), summary]
        if not left:
            return [KeepTogether([box])]
        outer = Table([[left, box]], colWidths=[left_w + gap, right_w], style=[
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(0,0),gap), ("RIGHTPADDING",(1,0),(1,0),0),
            ("TOPPADDING",(0,0),(-1,-1),0),  ("BOTTOMPADDING",(0,0),(-1,-1),0),
        ])
        return [KeepTogether([outer])]

    # ------------------------------------------------------------------ #
    #  Payments, notes, terms, signature
    # ------------------------------------------------------------------ #
    def _build_payment_history(self, width: float) -> Optional[Table]:
        """Small table listing each individual payment -- omitted entirely if
        there is no payment history to show (keeps simple invoices simple)."""
        if self.data.get("is_complimentary") or not self.payments:
            return None
        hs = ParagraphStyle("PHS", parent=self.styles["Small"], textColor=colors.HexColor(self.tpl.pay_head_fg))
        rows = [[Paragraph("<b>DATE</b>", hs), Paragraph("<b>MODE</b>", hs),
                 Paragraph("<b>NOTE</b>", hs), Paragraph("<b>AMOUNT</b>", ParagraphStyle("PHSR", parent=hs, alignment=TA_RIGHT))]]
        for p in sorted(self.payments, key=lambda x: x.get("payment_date") or ""):
            date_txt = parse_date(p.get("payment_date"), DB_DATE_FORMAT, DISPLAY_DATE_FORMAT)
            rows.append([
                Paragraph(date_txt, self.styles["Small"]),
                Paragraph(safe_markup(p.get("mode") or "-"), self.styles["Small"]),
                Paragraph(safe_markup(p.get("note") or "-"), self.styles["Small"]),
                Paragraph(f"{CURRENCY_SYMBOL}{safe_float(p.get('amount')):,.2f}", self.styles["SmallRight"]),
            ])
        return Table(rows, colWidths=[width * 0.22, width * 0.20, width * 0.33, width * 0.25], repeatRows=1,
                     style=TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor(self.tpl.pay_head_bg)),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LINEBELOW",(0,0),(-1,-2),0.5,colors.HexColor("#F3F4F6")),
            ("LINEBELOW",(0,-1),(-1,-1),0.8,colors.HexColor("#E5E7EB")),
            ("TOPPADDING",(0,0),(-1,-1),self._g(3)),   ("BOTTOMPADDING",(0,0),(-1,-1),self._g(3)),
            ("LEFTPADDING",(0,0),(-1,-1),6),  ("RIGHTPADDING",(0,0),(-1,-1),6),
        ]))

    def _invoice_needs_payment(self) -> bool:
        """Whether there's still money owed on this invoice -- the shared
        "nothing left to pay" gate for every payment-collection element
        below (QR code, bank/UPI details). True for Unpaid/Partially Paid
        invoices, or for anything once Settings > Images & Assets' "Always
        show payment details" override is turned on."""
        if self.config.get("pdf_settings", {}).get("always_show_qr_code", False):
            return True
        return self.data.get("status") in (InvoiceStatus.DUE.value, InvoiceStatus.PARTIAL.value)

    def _should_show_qr(self) -> bool:
        """Whether the QR code prints on this invoice.

        Always requires the per-invoice "Show QR Code" toggle to be on
        and the invoice not to be complimentary -- both pre-existing
        controls, unchanged. On top of those, by default, the QR only
        appears on invoices that still owe money (Unpaid/Partially
        Paid): there's nothing left to pay on a fully settled invoice,
        so a "scan to pay" code there is just noise. Settings > Images
        & Assets has an "Always show payment details" override for anyone
        who wants it printed regardless of payment status.
        """
        if not self.data.get("show_qr_code") or self.data.get("is_complimentary"):
            return False
        return self._invoice_needs_payment()

    def _payment_rows(self) -> List[Tuple[str, str]]:
        """The configured bank / UPI / payment-notes details (Settings > Payment & PDF) as (label, markup) pairs --
        empty when nothing is configured. Independent of the per-invoice QR toggle/asset -- a business that only
        takes bank transfers (no QR/UPI) still sees its bank details printed. Still shares _invoice_needs_payment()
        with the QR check, though, so a fully paid invoice doesn't keep printing "how to pay us" instructions for
        money that's already been collected. Skipped entirely for complimentary invoices, since there's nothing to
        pay on those. The payment notes come last, with no label."""
        if self.data.get("is_complimentary") or not self._invoice_needs_payment():
            return []
        pd = self.config.get("payment_details", {})
        rows: List[Tuple[str, str]] = []
        for label, key in (("Bank", "bank_name"), ("A/C Holder", "account_holder_name"),
                           ("A/C No", "account_number"), ("IFSC", "ifsc_code"),
                           ("UPI ID", "upi_id")):
            val = _real_value(pd.get(key), _SAMPLE_UPI_ID) if key == "upi_id" else pd.get(key)
            if val:
                rows.append((label, safe_markup(val)))
        if notes := pd.get("payment_notes"):
            rows.append(("", safe_markup(notes).replace("\n", "<br/>")))
        return rows

    def _build_payment_box(self, width: float) -> Optional[Table]:
        """PAYMENT DETAILS: bank / UPI lines, with the UPI QR code beside them when it prints."""
        rows = self._payment_rows()
        qr = self._asset("upi_qr_code") if self._should_show_qr() else None
        if not rows and not qr:
            return None
        if not rows:
            rows = [("", "Scan to pay.")]
        qr_w = 2.6 * cm if qr else 0.0
        inner_w = width - qr_w - 20                                  # the box's own padding
        label_w = 1.9 * cm
        # a row without a label (the payment notes) is one cell across both columns -- a span shows its FIRST cell
        lines = Table([[Paragraph(label, self.styles["PayLabel"]), Paragraph(value, self.styles["PayValue"])] if label else
                       [Paragraph(value, self.styles["PayValue"]), ""] for label, value in rows],
                      colWidths=[label_w, inner_w - label_w],
                      style=TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"), ("LEFTPADDING",(0,0),(-1,-1),0),
                                        ("RIGHTPADDING",(0,0),(-1,-1),4), ("TOPPADDING",(0,0),(-1,-1),1.5),
                                        ("BOTTOMPADDING",(0,0),(-1,-1),1.5)] +
                                       [("SPAN",(0,i),(1,i)) for i, (label, _v) in enumerate(rows) if not label]))
        cells: List[Any] = [[Paragraph("PAYMENT DETAILS", self.styles["Caption"]), Spacer(1, 3), lines]]
        col_widths = [inner_w + 20 - qr_w]
        if qr:
            cells.append(Image(str(qr), 2.2*cm, 2.2*cm, kind="proportional"))
            col_widths.append(qr_w)
        t = self.tpl
        style: List[Tuple] = [("VALIGN",(0,0),(-1,-1),"TOP"), ("LEFTPADDING",(0,0),(-1,-1),10), ("RIGHTPADDING",(0,0),(-1,-1),10),
                              ("TOPPADDING",(0,0),(-1,-1),self._g(8)), ("BOTTOMPADDING",(0,0),(-1,-1),self._g(8)),
                              ("BOX",(0,0),(-1,-1),0.8,colors.HexColor(t.meta_border or "#E5E7EB"))]
        if qr:
            style += [("ALIGN",(1,0),(1,0),"CENTER"), ("LEFTPADDING",(1,0),(1,0),0)]
        if t.meta_bg:
            style.append(("BACKGROUND",(0,0),(-1,-1),colors.HexColor(t.meta_bg)))
        return Table([cells], colWidths=col_widths, style=TableStyle(style))

    def _build_signature(self) -> Optional[Table]:
        """The signature image over a line and the name, right-aligned."""
        if not (sig := self._asset("signature")):
            return None
        name = safe_markup(self.data.get("your_name"))
        return Table([[Image(str(sig), 4.0*cm, 2.0*cm, kind="proportional")],
                      [Paragraph(f"<b>{name}</b>", self.styles["Center"])]], colWidths=[5.0*cm], hAlign="RIGHT",
                     style=[("LINEABOVE",(0,1),(0,1),1,colors.black), ("TOPPADDING",(0,1),(0,1),3),
                            ("ALIGN",(0,0),(-1,-1),"CENTER"), ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(-1,-1),0)])

    def _build_footer_section(self, width: float) -> Optional[Table]:
        """Payment details (with the QR code) on the left, the signature on the right -- None when there is neither."""
        box = self._build_payment_box(width * 0.62)
        sig = self._build_signature()
        if box is None and sig is None:
            return None
        return Table([[box if box is not None else "", sig if sig is not None else ""]],
                     colWidths=[width * 0.62, width * 0.38],
                     style=[("VALIGN",(0,0),(-1,-1),"BOTTOM"), ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(-1,-1),0),
                            ("TOPPADDING",(0,0),(-1,-1),0), ("BOTTOMPADDING",(0,0),(-1,-1),0)])

    def _build_notes_terms(self) -> List[Any]:
        """The customer-facing notes and the terms & conditions (the internal notes are never printed). A heading always
        stays with the first lines under it; a long text flows over to the next page."""
        out: List[Any] = []
        for label, key in (("Notes", "customer_notes"), ("Terms &amp; Conditions", "terms")):
            text = str(self.data.get(key) or "").strip()
            if text:
                out += [CondPageBreak(1.6*cm), Paragraph(label, self.styles["Caption"]), Spacer(1, self._g(0.1*cm)),
                        Paragraph(safe_markup(text).replace("\n", "<br/>"), self.styles["NoteText"]), Spacer(1, self._g(0.3*cm))]
        return out

    # ------------------------------------------------------------------ #
    #  Every page: top bar, running header, watermark, "Page x of y"
    # ------------------------------------------------------------------ #
    def _draw_top_bar(self, canvas):
        """The colour strip some templates draw across the top edge of every page."""
        if not self.tpl.bar:
            return
        canvas.saveState()
        canvas.setFillColor(colors.HexColor(self.tpl.bar))
        canvas.rect(0, A4[1] - 0.6*cm, A4[0], 0.6*cm, stroke=0, fill=1)
        canvas.restoreState()

    def _fit_text(self, canvas, text: str, max_width: float, size: float, font: Optional[str] = None) -> str:
        """`text`, cut with an ellipsis when it is wider than `max_width`."""
        font = font or self.font
        if canvas.stringWidth(text, font, size) <= max_width:
            return text
        while text and canvas.stringWidth(text + "…", font, size) > max_width:
            text = text[:-1]
        return text + "…"

    def _draw_running_header(self, canvas):
        """Pages after the first: whose invoice this is, so a loose sheet can be identified and filed."""
        width, height = A4
        left, right = self.MARGIN_X * cm, width - self.MARGIN_X * cm
        canvas.saveState()
        canvas.setFont(self.bold, 9)
        canvas.setFillColor(colors.HexColor("#111827"))
        canvas.drawString(left, height - 1.2*cm, self._fit_text(canvas, compose_your_name(self.config), 9.0*cm, 9, self.bold))
        canvas.setFont(self.font, 8.5)
        canvas.setFillColor(colors.HexColor("#6B7280"))
        number = str(self.data.get("invoice_number") or "")
        canvas.drawRightString(right, height - 1.2*cm, self._fit_text(canvas, f"Invoice {number}  ·  continued", 8.0*cm, 8.5))
        canvas.setStrokeColor(colors.HexColor(self.tpl.rule))
        canvas.setLineWidth(0.6)
        canvas.line(left, height - 1.5*cm, right, height - 1.5*cm)
        canvas.restoreState()

    def _apply_watermark(self, canvas, doc):
        wm_map = {
            InvoiceStatus.PAID.value:          ("PAID",            "#10B981", 100),
            InvoiceStatus.DUE.value:           ("BALANCE DUE",     "#EF4444",  50),
            InvoiceStatus.PARTIAL.value:       ("PARTIAL PAYMENT", "#F59E0B",  50),
            InvoiceStatus.OVERPAID.value:      ("ADVANCE / CREDIT","#8B5CF6",  50),
            InvoiceStatus.COMPLIMENTARY.value: ("COMPLIMENTARY",   "#3B82F6",  40),
        }
        if (entry := wm_map.get(self.data.get("status"))):
            txt, col, fsz = entry
            canvas.saveState()
            canvas.setFont("Helvetica-Bold", fsz)
            canvas.setFillColor(colors.HexColor(col), alpha=0.08)
            canvas.translate(A4[0]/2, A4[1]/2)
            canvas.rotate(35)
            canvas.drawCentredString(0, 0, txt)
            canvas.restoreState()

    def _draw_footer(self, canvas, page: int, total: int):
        """A hairline, the invoice number, the footer note from Settings (up to two lines) and "Page x of y"."""
        width, _height = A4
        left, right = self.MARGIN_X * cm, width - self.MARGIN_X * cm
        first = 1.45 * cm                                   # baseline of the first line of text
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#E5E7EB"))
        canvas.setLineWidth(0.6)
        canvas.line(left, 1.95*cm, right, 1.95*cm)
        canvas.setFont(self.font, 8)
        canvas.setFillColor(colors.HexColor("#6B7280"))
        canvas.drawString(left, first, self._fit_text(canvas, str(self.data.get("invoice_number") or ""), 4.0*cm, 8))
        canvas.drawRightString(right, first, f"Page {page} of {total}")
        note = str(self.config.get("pdf_settings", {}).get("footer_message", "Thank you!") or "")
        lines = [line for part in note.splitlines() for line in (simpleSplit(part, self.font, 8, 9.5*cm) or [""])]
        lines = [line for line in lines if line.strip()]
        shown = lines[:2]
        if len(lines) > 2:                                  # more text than two lines can hold
            shown[1] = self._fit_text(canvas, shown[1] + "…", 9.5*cm, 8)
        for n, line in enumerate(shown):
            canvas.drawCentredString(width / 2.0, first - n * 0.4*cm, line)
        canvas.restoreState()

    def _on_page(self, canvas, doc):
        """Called as every page begins: the colour bar, the running header (not on the first page) and the watermark
        -- all behind the content. The footer is drawn at the end, when the number of pages is known."""
        self._draw_top_bar(canvas)
        if doc.page > 1:
            self._draw_running_header(canvas)
        self._apply_watermark(canvas, doc)

    def _numbered_canvas(self):
        """A canvas that draws the footer once the number of pages is known."""
        generator = self

        class NumberedCanvas(Canvas):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self._page_states: List[Dict[str, Any]] = []

            def showPage(self):
                self._page_states.append(dict(self.__dict__))
                self._startPage()

            def save(self):
                total = len(self._page_states)
                generator.page_count = total
                for state in self._page_states:
                    self.__dict__.update(state)
                    generator._draw_footer(self, self._pageNumber, total)
                    super().showPage()
                super().save()

        return NumberedCanvas

    def _tail_marker(self):
        """A flowable with no size that goes last: when the layout reaches it, it is told how much room is left on the
        last page (the auto-fit pass decides from that whether the last page is nearly empty)."""
        generator = self

        class TailMarker(Flowable):
            def wrap(self, avail_width, avail_height):
                generator._room_left = avail_height
                return 0, 0

            def draw(self):
                pass

        return TailMarker()

    def _layout(self, target) -> None:
        """Lays the whole invoice out on A4 pages and writes the PDF to `target` (a path or a file-like object). Leaves
        the number of pages in self.page_count and the room left on the last page in self._room_left."""
        number = str(self.data.get("invoice_number") or "")
        doc = BaseDocTemplate(target, pagesize=A4,
                              rightMargin=self.MARGIN_X*cm, leftMargin=self.MARGIN_X*cm,
                              topMargin=self.MARGIN_TOP*cm, bottomMargin=self.MARGIN_BOTTOM*cm,
                              title=f"Invoice {number}", author=compose_your_name(self.config),
                              subject=f"Invoice {number} for {self.data.get('bill_to') or ''}".strip(),
                              creator=f"{APP_NAME} {APP_VERSION} (build {APP_BUILD_VERSION})")
        # A frame without padding: every table is exactly as wide as the printable area, so the page edges line up.
        frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="page",
                      leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        doc.addPageTemplates([PageTemplate(id="invoice", frames=[frame], onPage=self._on_page, pagesize=A4)])
        width, gap = doc.width, self._g
        story: List[Any] = [
            self._build_header(width),         Spacer(1, gap(0.45*cm)),
            self._build_parties(width),        Spacer(1, gap(0.4*cm)),
            self._build_meta(width),           Spacer(1, gap(0.45*cm)),
            self._build_items(width),          Spacer(1, gap(0.4*cm)),
            *self._build_summary(width),
        ]
        if history := self._build_payment_history(width):
            story += [Spacer(1, gap(0.35*cm)), CondPageBreak(2.5*cm), Paragraph("Payment History", self.styles["Caption"]),
                      Spacer(1, gap(0.12*cm)), history]
        if footer := self._build_footer_section(width):
            story += [Spacer(1, gap(0.4*cm)), KeepTogether([footer])]
        # Internal notes (self.data["notes"]) are intentionally never printed
        # on the customer-facing PDF -- they're for Vyaprix's eyes only. The customer notes and the terms are.
        notes = self._build_notes_terms()
        if notes:
            story += [Spacer(1, gap(0.4*cm)), *notes]
        story.append(self._tail_marker())
        doc.build(story, canvasmaker=self._numbered_canvas())

    def generate(self) -> Path:
        """Writes the invoice PDF and returns its path. An invoice that runs only a little onto a second (or later) page
        is laid out a second time with tighter gaps and row padding, and that version is printed when it saves a page."""
        self.tight = self.tightened = False
        plain = io.BytesIO()
        self._layout(plain)
        result, pages = plain.getvalue(), self.page_count
        if pages > 1 and self._room_left >= (1.0 - self.TIGHT_WHEN_LAST_PAGE_BELOW) * self.page_height:
            self.tight = True
            try:
                snug = io.BytesIO()
                self._layout(snug)
                if self.page_count < pages:
                    result, pages, self.tightened = snug.getvalue(), self.page_count, True
            except Exception:
                logger.warning("The tighter layout of invoice %r failed; printing it as designed.",
                               self.data.get("invoice_number"), exc_info=True)
            finally:
                self.tight = False
                self.page_count = pages
        self.filename.write_bytes(result)
        return self.filename


def sample_invoice_data(config: Dict[str, Any]) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
    """A made-up invoice with one payment, for previewing a template: printed under the user's own business details
    (name, logo, bank details), with the kinds of lines a real invoice has -- quantities, units, HSN codes and a GST
    rate per item -- so the whole table shows."""
    today = datetime.date.today()
    items = [
        {"title": "Stage and sound setup", "description": "Truss stage 20 x 16 ft with PA system", "amount": 90000.0,
         "discount": 9000.0, "qty": 2.0, "unit": "Event", "rate": 45000.0, "hsn": "998596", "tax_rate": 18.0,
         "disc_type": "percent", "disc_value": 10.0},
        {"title": "Banquet chairs", "description": "With covers", "amount": 12500.0, "discount": 0.0, "qty": 250.0,
         "unit": "Nos", "rate": 50.0, "hsn": "997212", "tax_rate": 18.0},
        {"title": "Printed brochures", "description": "A5 glossy, 4 colour", "amount": 8000.0, "discount": 0.0,
         "qty": 1000.0, "unit": "Pcs", "rate": 8.0, "hsn": "4911", "tax_rate": 5.0},
    ]
    data = {
        "invoice_number": "INV-0001", "bill_to": "Sample Customer Pvt Ltd",
        "client_contact": "12 Park Street\nKolkata - 700016\naccounts@example.com", "client_gstin": "",
        "your_name": compose_your_name(config), "your_contact": compose_your_contact(config),
        "items": json.dumps(items), "amount_paid": 0.0, "tax_rate": 18.0,
        "invoice_date": today.strftime(DB_DATE_FORMAT), "payment_date": None,
        "status": InvoiceStatus.PARTIAL.value, "payment_terms": "Net 15", "payment_mode": None,
        "show_qr_code": 1, "is_complimentary": 0, "is_gst_invoice": 1,
        "global_discount_value": 0.0, "global_discount_type": "amount", "advance_amount": 0.0, "advance_date": None,
        "gst_type": "CGST/SGST", "notes": "", "tax_mode": TAX_MODE_PER_ITEM, "round_off": 1,
        "place_of_supply": "", "po_number": "",
        "due_date": (today + datetime.timedelta(days=15)).strftime(DB_DATE_FORMAT), "reverse_charge": 0,
        "customer_notes": "Thank you for your business!",
        "terms": "1. Payment within 15 days of the invoice date.\n2. Goods once sold are not returnable.",
    }
    payments = [{"amount": 25000.0, "payment_date": today.strftime(DB_DATE_FORMAT), "mode": "UPI", "note": "Advance"}]
    return data, payments


def merge_pdfs(paths: List[Path], out_path: Path) -> Path:
    """Combine multiple single-invoice PDFs into one file for batch printing.

    CHANGED FROM THE TKINTER BUILD: this used `fitz.open()` /
    `merged.insert_pdf(src)` (PyMuPDF). Behaviour is identical -- every page
    of every input PDF, in order, into one output file -- just implemented
    with pypdf's PdfWriter instead, so PyMuPDF is no longer a dependency of
    this app at all once the PDF viewer (next stage) moves to Qt's own
    PySide6.QtPdf module."""
    if not PYPDF_AVAILABLE:
        raise RuntimeError("pypdf is required for bulk printing (pip install pypdf).")
    writer = PdfWriter()
    try:
        for p in paths:
            writer.append(str(p))
        with open(out_path, "wb") as fh:
            writer.write(fh)
    finally:
        writer.close()
    return out_path


# =============================================================================
#  11. THEME ENGINE
#  Eight named themes. Each is one palette; the Qt palette, the global stylesheet and the generated icons derive from it.
# =============================================================================
THEME_NAMES: List[str] = ["litera", "cosmo", "flatly", "darkly", "cyborg", "superhero", "morph", "pulse"]
DEFAULT_THEME: str = "litera"

#  bg      window background          surface  cards / inputs / menus
#  alt     zebra rows, headers, hover  border   hairlines
#  fg      body text                   muted    secondary text
THEMES: Dict[str, Dict[str, Any]] = {
    "litera": dict(dark=False, bg="#F4F6FA", surface="#FFFFFF", alt="#F1F3F7", border="#DDE1E7",
                   fg="#343A40", muted="#6C757D",
                   primary="#4582EC", secondary="#ADB5BD", success="#02B875",
                   info="#17A2B8", warning="#F0AD4E", danger="#D9534F"),
    "cosmo":  dict(dark=False, bg="#F5F7FA", surface="#FFFFFF", alt="#EFF2F6", border="#D9DEE4",
                   fg="#373A3C", muted="#6B7075",
                   primary="#2780E3", secondary="#7E8081", success="#3FB618",
                   info="#9954BB", warning="#FF7518", danger="#FF0039"),
    "flatly": dict(dark=False, bg="#F3F5F6", surface="#FFFFFF", alt="#ECF0F1", border="#D5DBDE",
                   fg="#2C3E50", muted="#7B8A8B",
                   primary="#2C3E50", secondary="#95A5A6", success="#18BC9C",
                   info="#3498DB", warning="#F39C12", danger="#E74C3C"),
    "darkly": dict(dark=True, bg="#222222", surface="#2D2D2D", alt="#363636", border="#454545",
                   fg="#F1F1F1", muted="#ADB5BD",
                   primary="#375A7F", secondary="#444444", success="#00BC8C",
                   info="#3498DB", warning="#F39C12", danger="#E74C3C"),
    "cyborg": dict(dark=True, bg="#060606", surface="#121212", alt="#1C1C1C", border="#2E2E2E",
                   fg="#E4E4E4", muted="#9A9A9A",
                   primary="#2A9FD6", secondary="#555555", success="#77B300",
                   info="#9933CC", warning="#FF8800", danger="#CC0000"),
    "superhero": dict(dark=True, bg="#2B3E50", surface="#34495E", alt="#3E5367", border="#4E6377",
                      fg="#EBEBEB", muted="#A9B8C6",
                      primary="#4C9BE8", secondary="#4E5D6C", success="#5CB85C",
                      info="#5BC0DE", warning="#F0AD4E", danger="#D9534F"),
    "morph":  dict(dark=False, bg="#DCE5F1", surface="#EDF2F9", alt="#E2EAF4", border="#C3D1E3",
                   fg="#44546A", muted="#6F7F94",
                   primary="#378DFC", secondary="#AAAAAA", success="#43CC29",
                   info="#5B62F4", warning="#FFC107", danger="#E52527"),
    "pulse":  dict(dark=False, bg="#F6F4FA", surface="#FFFFFF", alt="#F0ECF6", border="#DAD3E8",
                   fg="#3D3D4A", muted="#7A7589",
                   primary="#593196", secondary="#69676E", success="#13B955",
                   info="#009CDC", warning="#EFA31D", danger="#FC3939"),
}

_LEGACY_THEME_NAMES = {"light": "litera", "dark": "darkly"}   # written by early PySide6 previews
TONES: Tuple[str, ...] = ("primary", "secondary", "success", "info", "warning", "danger")
_ACTIVE_THEME_NAME: str = DEFAULT_THEME
_ACTIVE_PALETTE: Dict[str, str] = {}


def normalize_theme_name(name: Optional[str]) -> str:
    name = _LEGACY_THEME_NAMES.get((name or "").strip().lower(), (name or "").strip().lower())
    return name if name in THEMES else DEFAULT_THEME


# ---- tiny colour maths -------------------------------------------------------
def _rgb(h: str) -> Tuple[int, int, int]:
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _hex(r: float, g: float, b: float) -> str:
    return "#{:02X}{:02X}{:02X}".format(*(max(0, min(255, int(round(v)))) for v in (r, g, b)))


def _mix(a: str, b: str, t: float) -> str:
    """t=0 -> a, t=1 -> b."""
    ra, ga, ba = _rgb(a)
    rb, gb, bb = _rgb(b)
    return _hex(ra + (rb - ra) * t, ga + (gb - ga) * t, ba + (bb - ba) * t)


def _lum(h: str) -> float:
    def f(c):
        c /= 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = _rgb(h)
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def _contrast(a: str, b: str) -> float:
    la, lb = _lum(a), _lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def _on(bg: str) -> str:
    """Text colour for a filled background. White -- exactly what ttkbootstrap's `selectfg` gives every tone in all
    eight themes -- except on the rare pale fill where white would be close to invisible (Morph's yellow warning)."""
    return "#FFFFFF" if _contrast(bg, "#FFFFFF") >= 1.8 else "#1F2937"


def _readable(color: str, on: str, minimum: float = 3.2) -> str:
    """Nudge `color` toward black/white until it reads on `on` (for text & outlines)."""
    target = "#FFFFFF" if _lum(on) < 0.4 else "#000000"
    c = color
    for step in range(12):
        if _contrast(c, on) >= minimum:
            break
        c = _mix(c, target, 0.12)
    return c


def resolve_palette(theme: Dict[str, Any]) -> Dict[str, str]:
    """Theme dict -> flat dict of every token the QSS / QPalette need."""
    p: Dict[str, str] = {k: v for k, v in theme.items() if isinstance(v, str)}
    dark = bool(theme["dark"])
    p["is_dark"] = "1" if dark else "0"
    p["input_bg"] = p["surface"]
    p["border_strong"] = _mix(p["border"], p["fg"], 0.35)
    p["sel"] = _mix(p["surface"], p["primary"], 0.28 if dark else 0.20)
    p["sel_fg"] = p["fg"]
    p["disabled_bg"] = _mix(p["bg"], p["fg"], 0.06)
    p["disabled_fg"] = _mix(p["surface"], p["fg"], 0.42)
    p["placeholder"] = _mix(p["surface"], p["fg"], 0.48)
    p["shadow"] = "#000000"
    for tone in TONES:
        base = p[tone]
        p[f"{tone}_on"] = _on(base)
        p[f"{tone}_hover"] = _mix(base, "#FFFFFF" if dark else "#000000", 0.14)
        p[f"{tone}_pressed"] = _mix(base, "#FFFFFF" if dark else "#000000", 0.24)
        p[f"{tone}_soft"] = _mix(p["surface"], base, 0.18 if dark else 0.12)
        p[f"{tone}_text"] = _readable(base, p["surface"])
    p["muted_text"] = _readable(p["muted"], p["surface"], 3.0)
    p["hero_bg"] = p["primary"]
    p["hero_fg"] = _on(p["primary"])
    p["hero_sub"] = _mix(p["primary"], p["hero_fg"], 0.78)
    p["backdrop"] = _mix(p["bg"], "#000000", 0.18 if not dark else 0.30)    # PDF viewer / canvas surround
    return p


def theme_color(key: str) -> str:
    """Current theme's colour for a palette token (e.g. 'border', 'primary_text')."""
    if not _ACTIVE_PALETTE:
        _ACTIVE_PALETTE.update(resolve_palette(THEMES[DEFAULT_THEME]))
    return _ACTIVE_PALETTE[key]


def active_theme_name() -> str:
    return _ACTIVE_THEME_NAME


def theme_is_dark() -> bool:
    return theme_color("is_dark") == "1"


def _icon_dir() -> Path:
    for base in (DATA_DIR, Path(tempfile.gettempdir())):
        try:
            d = base / "ui_cache"
            d.mkdir(parents=True, exist_ok=True)
            return d
        except OSError:
            continue
    return Path(tempfile.gettempdir())


def _paint_icon(kind: str, color: str, px: int = 24) -> QImage:
    img = QImage(px, px, QImage.Format.Format_ARGB32)
    img.fill(0)
    qp = QPainter(img)
    qp.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    pen = QPen(QColor(color))
    pen.setWidthF(px * 0.11)
    pen.setCapStyle(Qt.PenCapStyle.RoundCap)
    pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    s = px / 24.0
    if kind in ("down", "up"):
        path = QPainterPath()
        y1, y2 = (8.5, 15.0) if kind == "down" else (15.0, 8.5)
        path.moveTo(6.0 * s, y1 * s)
        path.lineTo(12.0 * s, y2 * s)
        path.lineTo(18.0 * s, y1 * s)
        qp.setPen(pen)
        qp.drawPath(path)
    elif kind in ("left", "right"):
        path = QPainterPath()
        x1, x2 = (14.5, 9.0) if kind == "left" else (9.5, 15.0)
        path.moveTo(x1 * s, 6.0 * s)
        path.lineTo(x2 * s, 12.0 * s)
        path.lineTo(x1 * s, 18.0 * s)
        qp.setPen(pen)
        qp.drawPath(path)
    elif kind == "check":
        pen.setWidthF(px * 0.14)
        path = QPainterPath()
        path.moveTo(5.5 * s, 12.5 * s)
        path.lineTo(10.0 * s, 17.0 * s)
        path.lineTo(18.5 * s, 7.5 * s)
        qp.setPen(pen)
        qp.drawPath(path)
    elif kind == "dot":
        qp.setPen(Qt.PenStyle.NoPen)
        qp.setBrush(QBrush(QColor(color)))
        qp.drawEllipse(QRectF(7.5 * s, 7.5 * s, 9.0 * s, 9.0 * s))
    elif kind in ("tri_up", "tri_down"):
        qp.setPen(Qt.PenStyle.NoPen)
        qp.setBrush(QBrush(QColor(color)))
        path = QPainterPath()
        if kind == "tri_up":
            path.moveTo(12.0 * s, 8.0 * s); path.lineTo(18.0 * s, 16.0 * s); path.lineTo(6.0 * s, 16.0 * s)
        else:
            path.moveTo(12.0 * s, 16.0 * s); path.lineTo(18.0 * s, 8.0 * s); path.lineTo(6.0 * s, 8.0 * s)
        path.closeSubpath()
        qp.drawPath(path)
    qp.end()
    return img


def _icon_url(kind: str, color: str) -> str:
    path = _icon_dir() / f"{kind}_{color.lstrip('#').lower()}.png"
    if not path.exists():
        _paint_icon(kind, color).save(str(path), "PNG")
    return 'url("' + path.as_posix() + '")'


_QSS_BASE = """
* { outline: 0; }
QWidget { font-family: "@font_family@", "Noto Sans", "Helvetica Neue", Arial, sans-serif; font-size: 13px; color: @fg@; }
QMainWindow, QDialog { background: @bg@; }
QToolTip { background: @surface@; color: @fg@; border: 1px solid @border_strong@; padding: 4px 8px; border-radius: 4px; }

/* ---------- frames ---------- */
QFrame[role="card"]    { background: @surface@; border: 1px solid @border@; border-radius: 10px; }
QFrame[role="sidebar"] { background: @surface@; border: none; }          /* the 1 px divider next to it is its own vline() */
QFrame[role="toolbar"] { background: @surface@; border: 1px solid @border@; border-radius: 8px; }
QFrame[role="sep"]     { background: @border@; border: none; }
QFrame[role="hero"]    { background: @hero_bg@; border: none; border-radius: 0px; }
QFrame[role="hero-round"] { background: @hero_bg@; border: none; border-radius: 8px; }
QFrame[role="stat"]    { background: @surface@; border: 1px solid @border@; border-radius: 10px; }
QFrame[role="stat"][tone="primary"]   { border-left: 4px solid @primary@; }
QFrame[role="stat"][tone="success"]   { border-left: 4px solid @success@; }
QFrame[role="stat"][tone="danger"]    { border-left: 4px solid @danger@; }
QFrame[role="stat"][tone="warning"]   { border-left: 4px solid @warning@; }
QFrame[role="stat"][tone="info"]      { border-left: 4px solid @info@; }
QFrame[role="stat"][tone="secondary"] { border-left: 4px solid @secondary@; }
QFrame[role="stat-solid"] { border: none; border-radius: 10px; }   /* Dashboard cards: solid tone block, as in the Tk build */
QFrame[role="splash"]  { background: @surface@; border: none; }
QFrame[role="splash-border"] { background: @primary@; border: none; }      /* square, like the Tk splash window */
QFrame[role="thumb"]   { background: @alt@; border: 1px solid @border@; border-radius: 6px; }
QFrame[role="tpl"]     { background: @surface@; border: 2px solid @border@; border-radius: 8px; }      /* a template in Settings */
QFrame[role="tpl"][selected="true"] { border: 2px solid @primary@; }
QFrame[role="key"]     { background: @alt@; border: 1px solid @border_strong@; border-radius: 4px; }

/* ---------- labels ---------- */
QLabel { background: transparent; }
QLabel[role="h1"]      { font-size: 22px; font-weight: 700; }
QLabel[role="h2"]      { font-size: 15px; font-weight: 600; }
QLabel[role="h3"]      { font-size: 13px; font-weight: 700; }
QLabel[role="muted"]   { color: @muted_text@; }
QLabel[role="small"]   { color: @muted_text@; font-size: 11px; }
QLabel[role="section"] { font-size: 11px; font-weight: 700; color: @muted_text@; }
QLabel[role="caption"] { font-size: 12px; font-weight: 600; }          /* field captions on forms */
QLabel[role="display"] { font-size: 34px; font-weight: 700; }            /* About box version number */
QLabel[role="stat-value"] { font-size: 22px; font-weight: 700; }
QLabel[role="stat-title"] { font-size: 11px; font-weight: 700; color: @muted_text@; }
QLabel[role="stat-solid-title"] { font-size: 12px; }
QLabel[role="stat-solid-value"] { font-size: 24px; font-weight: 700; }
QLabel[role="stat-solid-value"][compact="true"] { font-size: 18px; }
QLabel[role="hero-title"] { color: @hero_fg@; font-size: 20px; font-weight: 700; }
QLabel[role="hero-sub"]   { color: @hero_sub@; font-size: 11px; }
QLabel[role="key"]     { font-weight: 700; font-size: 12px; background: @alt@; border: 1px solid @border_strong@; border-radius: 4px; padding: 4px 8px; }
QLabel[role="link"]    { color: @primary_text@; }
QLabel[role="badge"]   { font-size: 11px; font-weight: 700; border-radius: 9px; padding: 3px 10px; }
@LABEL_TONES@

/* ---------- inputs ---------- */
QLineEdit, QTextEdit, QPlainTextEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit, QTimeEdit {
    background: @input_bg@; color: @fg@; border: 1px solid @border_strong@; border-radius: 6px;
    padding: 5px 8px; selection-background-color: @sel@; selection-color: @sel_fg@;
}
QLineEdit:hover, QTextEdit:hover, QPlainTextEdit:hover, QComboBox:hover, QSpinBox:hover, QDoubleSpinBox:hover, QDateEdit:hover { border-color: @primary@; }
QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus, QDateEdit:focus,
QComboBox:on { border: 1px solid @primary@; }
QLineEdit:disabled, QTextEdit:disabled, QPlainTextEdit:disabled, QComboBox:disabled, QSpinBox:disabled, QDateEdit:disabled {
    background: @disabled_bg@; color: @disabled_fg@; border-color: @border@; }
QLineEdit[invalid="true"], QTextEdit[invalid="true"], QComboBox[invalid="true"] { border: 1px solid @danger@; }
QLineEdit[readOnly="true"] { background: @alt@; }
QLineEdit[flat="true"], QComboBox QLineEdit, QAbstractSpinBox QLineEdit, QDateEdit QLineEdit { border: none; background: transparent; padding: 0px; }

QComboBox { padding-right: 26px; }
QComboBox::drop-down, QDateEdit::drop-down { subcontrol-origin: padding; subcontrol-position: center right; width: 24px; border: none; }
QComboBox::down-arrow, QDateEdit::down-arrow { image: @ICON_DOWN_FG@; width: 12px; height: 12px; }
QComboBox::down-arrow:disabled, QDateEdit::down-arrow:disabled { image: @ICON_DOWN_DIS@; }
QComboBox QAbstractItemView { background: @surface@; color: @fg@; border: 1px solid @border_strong@; border-radius: 6px;
    selection-background-color: @primary@; selection-color: @primary_on@; padding: 3px; }
QComboBox QAbstractItemView::item { min-height: 24px; padding: 2px 6px; border-radius: 4px; }

QAbstractSpinBox::up-button, QAbstractSpinBox::down-button { width: 18px; border: none; background: transparent; }
QAbstractSpinBox::up-arrow { image: @ICON_UP_FG@; width: 10px; height: 10px; }
QAbstractSpinBox::down-arrow { image: @ICON_DOWN_FG@; width: 10px; height: 10px; }

QCheckBox, QRadioButton { spacing: 8px; background: transparent; }
QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid @border_strong@; border-radius: 4px; background: @input_bg@; }
QCheckBox::indicator:hover, QRadioButton::indicator:hover { border-color: @primary@; }
QCheckBox::indicator:checked { background: @primary@; border-color: @primary@; image: @ICON_CHECK@; }
QCheckBox::indicator:indeterminate { background: @primary@; border-color: @primary@; }
QCheckBox::indicator:disabled, QRadioButton::indicator:disabled { background: @disabled_bg@; border-color: @border@; }
QRadioButton::indicator { width: 16px; height: 16px; border: 1px solid @border_strong@; border-radius: 9px; background: @input_bg@; }
QRadioButton::indicator:checked { background: @primary@; border-color: @primary@; image: @ICON_DOT@; }

/* ---------- buttons ---------- */
QPushButton { background: @surface@; color: @fg@; border: 1px solid @border_strong@; border-radius: 6px; padding: 7px 14px; }
QPushButton:hover { background: @alt@; border-color: @muted@; }
QPushButton:pressed { background: @border@; }
QPushButton:focus { border-color: @primary@; }
QPushButton:disabled { color: @disabled_fg@; background: @disabled_bg@; border-color: @border@; }
QPushButton:checked { background: @primary@; color: @primary_on@; border-color: @primary@; }
@BUTTON_ROLES@
QPushButton[role="link"] { background: transparent; border: none; color: @primary_text@; padding: 4px 6px; }
QPushButton[role="link"]:hover { background: @primary_soft@; }
QPushButton[role="tool"] { padding: 4px 8px; min-width: 22px; font-size: 14px; background: @surface@; }
QPushButton[glyph="true"] { padding: 7px 6px; }                                 /* one-symbol buttons ("✕"): no wide side padding */
QPushButton[align="left"] { text-align: left; padding: 11px 15px; }        /* sidebar navigation / system buttons */
QPushButton[align="icon"] { text-align: center; padding: 11px 0px; }       /* ... in the compact (icons only) sidebar */
QToolButton { background: transparent; border: 1px solid transparent; border-radius: 6px; padding: 4px; }
QToolButton:hover { background: @alt@; border-color: @border@; }

/* ---------- item views ---------- */
QAbstractItemView { background: @surface@; alternate-background-color: @alt@; color: @fg@; border: 1px solid @border@;
    border-radius: 8px; gridline-color: @border@; selection-background-color: @sel@; selection-color: @sel_fg@; }
QTableView, QTreeView, QListView { show-decoration-selected: 1; }
QTableView::item, QTreeView::item { padding: 5px 8px; border: none; }
QTableView::item:selected, QTreeView::item:selected, QListView::item:selected { background: @sel@; color: @sel_fg@; }
QTableView::item:hover, QTreeView::item:hover, QListView::item:hover { background: @alt@; }
QTableView::item:selected:hover { background: @sel@; }
QListView::item { padding: 6px 8px; border-radius: 4px; }
QHeaderView { background: @alt@; border: none; }
QHeaderView::section { background: @alt@; color: @muted_text@; font-weight: 700; padding: 8px 10px; border: none;
    border-bottom: 1px solid @border_strong@; border-right: 1px solid @border@; }
QHeaderView::section:hover { background: @border@; color: @fg@; }
QHeaderView::section:first { border-top-left-radius: 8px; }
QHeaderView::section:last { border-top-right-radius: 8px; border-right: none; }
QHeaderView::up-arrow   { image: @ICON_TRI_UP@;   width: 12px; height: 12px; subcontrol-position: center right; }
QHeaderView::down-arrow { image: @ICON_TRI_DOWN@; width: 12px; height: 12px; subcontrol-position: center right; }
QTableCornerButton::section { background: @alt@; border: none; }

/* ---------- menus ---------- */
QMenu { background: @surface@; color: @fg@; border: 1px solid @border_strong@; border-radius: 8px; padding: 5px; }
QMenu::item { padding: 7px 26px 7px 14px; border-radius: 5px; }
QMenu::item:selected { background: @primary@; color: @primary_on@; }
QMenu::item:disabled { color: @disabled_fg@; }
QMenu::separator { height: 1px; background: @border@; margin: 5px 8px; }

/* ---------- scrolling ---------- */
QScrollArea { border: none; background: transparent; }
QScrollBar:vertical   { width: 12px; background: transparent; margin: 2px; }
QScrollBar:horizontal { height: 12px; background: transparent; margin: 2px; }
QScrollBar::handle:vertical   { background: @border_strong@; border-radius: 4px; min-height: 28px; margin: 0px 2px; }
QScrollBar::handle:horizontal { background: @border_strong@; border-radius: 4px; min-width: 28px; margin: 2px 0px; }
QScrollBar::handle:hover { background: @muted@; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0px; height: 0px; background: none; border: none; }
QScrollBar::add-page, QScrollBar::sub-page { background: transparent; }

/* ---------- containers ---------- */
QTabWidget::pane { border: 1px solid @border@; border-radius: 8px; top: -1px; background: transparent; }
QTabBar::tab { padding: 9px 18px; background: transparent; color: @muted_text@; border: none; border-bottom: 2px solid transparent; margin-right: 2px; }
QTabBar::tab:hover { color: @fg@; }
QTabBar::tab:selected { color: @primary_text@; font-weight: 600; border-bottom: 2px solid @primary@; }
QGroupBox { border: 1px solid @border@; border-radius: 8px; margin-top: 14px; padding: 14px 10px 10px 10px; background: transparent; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; left: 12px; padding: 0 6px; color: @muted_text@; font-weight: 700; }
QSplitter::handle { background: @border@; }
QProgressBar { border: none; border-radius: 4px; background: @border@; text-align: center; min-height: 8px; max-height: 8px; color: @fg@; }
QProgressBar::chunk { background: @primary@; border-radius: 4px; }
QProgressBar[tall="true"] { min-height: 18px; max-height: 18px; border-radius: 9px; }
QProgressBar[tall="true"]::chunk { border-radius: 9px; }
QMessageBox { background: @bg@; }
QMessageBox QLabel { color: @fg@; }

/* ---------- calendar popup (QDateEdit) ---------- */
QCalendarWidget QWidget { alternate-background-color: @alt@; }
QCalendarWidget QToolButton { background: transparent; color: @fg@; border: none; border-radius: 6px; padding: 4px 8px; font-weight: 600; }
QCalendarWidget QToolButton:hover { background: @alt@; }
QCalendarWidget QToolButton::menu-indicator { image: none; }
QCalendarWidget #qt_calendar_prevmonth { qproperty-icon: none; image: @ICON_LEFT@; width: 16px; height: 16px; }
QCalendarWidget #qt_calendar_nextmonth { qproperty-icon: none; image: @ICON_RIGHT@; width: 16px; height: 16px; }
QCalendarWidget QMenu { background: @surface@; }
QCalendarWidget QSpinBox { background: @surface@; border: 1px solid @border_strong@; border-radius: 4px; padding: 2px 4px; }
QCalendarWidget QAbstractItemView:enabled { background: @surface@; color: @fg@; border: none; border-radius: 0px;
    selection-background-color: @primary@; selection-color: @primary_on@; }
QCalendarWidget QAbstractItemView:disabled { color: @disabled_fg@; }
#qt_calendar_navigationbar { background: @alt@; border-bottom: 1px solid @border@; }
"""


def _label_tone_rules() -> str:
    out = []
    for tone in TONES:
        out.append(f'QLabel[tone="{tone}"] {{ color: @{tone}_text@; }}')
        out.append(f'QLabel[role="badge"][tone="{tone}"] {{ background: @{tone}@; color: @{tone}_on@; }}')
        out.append(f'QLabel[role="badge-soft"][tone="{tone}"] {{ background: @{tone}_soft@; color: @{tone}_text@; font-weight: 700; '
                   f'border-radius: 9px; padding: 3px 10px; font-size: 11px; }}')
        out.append(f'QFrame[role="stat-solid"][tone="{tone}"] {{ background: @{tone}@; }}')
        out.append(f'QLabel[role="stat-solid-title"][tone="{tone}"], QLabel[role="stat-solid-value"][tone="{tone}"] '
                   f'{{ color: @{tone}_on@; }}')
    out.append('QLabel[tone="muted"] { color: @muted_text@; }')
    return "\n".join(out)


def _button_role_rules() -> str:
    out = []
    for tone in TONES:
        out.append(
            f'QPushButton[role="{tone}"] {{ background: @{tone}@; color: @{tone}_on@; border: 1px solid @{tone}@; }}\n'
            f'QPushButton[role="{tone}"]:hover {{ background: @{tone}_hover@; border-color: @{tone}_hover@; }}\n'
            f'QPushButton[role="{tone}"]:pressed {{ background: @{tone}_pressed@; border-color: @{tone}_pressed@; }}\n'
            f'QPushButton[role="{tone}"]:disabled {{ background: @disabled_bg@; color: @disabled_fg@; border-color: @border@; }}\n'
            f'QPushButton[role="{tone}-outline"] {{ background: transparent; color: @{tone}_text@; border: 1px solid @{tone}_text@; }}\n'
            f'QPushButton[role="{tone}-outline"]:hover {{ background: @{tone}@; color: @{tone}_on@; border-color: @{tone}@; }}\n'
            f'QPushButton[role="{tone}-outline"]:pressed {{ background: @{tone}_pressed@; color: @{tone}_on@; }}\n'
            f'QPushButton[role="{tone}-outline"]:disabled {{ background: transparent; color: @disabled_fg@; border-color: @border@; }}'
        )
    return "\n".join(out)


def build_qss(p: Dict[str, str]) -> str:
    tokens = dict(p)
    tokens["font_family"] = FONT_FAMILY
    tokens["LABEL_TONES"] = _label_tone_rules()
    tokens["BUTTON_ROLES"] = _button_role_rules()
    # Icons are generated lazily and keyed by colour.
    tokens["ICON_DOWN_FG"] = _icon_url("down", p["muted_text"])
    tokens["ICON_DOWN_DIS"] = _icon_url("down", p["disabled_fg"])
    tokens["ICON_UP_FG"] = _icon_url("up", p["muted_text"])
    tokens["ICON_LEFT"] = _icon_url("left", p["fg"])
    tokens["ICON_RIGHT"] = _icon_url("right", p["fg"])
    tokens["ICON_CHECK"] = _icon_url("check", p["primary_on"])
    tokens["ICON_DOT"] = _icon_url("dot", p["primary_on"])
    tokens["ICON_TRI_UP"] = _icon_url("tri_up", p["muted_text"])
    tokens["ICON_TRI_DOWN"] = _icon_url("tri_down", p["muted_text"])
    qss = _QSS_BASE
    # two passes: expand the generated rule blocks first, then every @token@
    for key in ("LABEL_TONES", "BUTTON_ROLES"):
        qss = qss.replace(f"@{key}@", tokens[key])
    return re.sub(r"@(\w+)@", lambda m: tokens[m.group(1)], qss)


def build_qpalette(p: Dict[str, str]) -> QPalette:
    pal = QPalette()
    R = QPalette.ColorRole
    G = QPalette.ColorGroup
    pairs = {
        R.Window: p["bg"], R.WindowText: p["fg"], R.Base: p["input_bg"], R.AlternateBase: p["alt"],
        R.Text: p["fg"], R.Button: p["surface"], R.ButtonText: p["fg"], R.ToolTipBase: p["surface"],
        R.ToolTipText: p["fg"], R.Highlight: p["primary"], R.HighlightedText: p["primary_on"],
        R.PlaceholderText: p["placeholder"], R.Link: p["primary_text"], R.LinkVisited: p["info_text"],
        R.BrightText: "#FFFFFF", R.Light: p["surface"], R.Midlight: p["alt"], R.Mid: p["border"],
        R.Dark: p["border_strong"], R.Shadow: "#000000",
    }
    for role, color in pairs.items():
        pal.setColor(role, QColor(color))
    for role in (R.WindowText, R.Text, R.ButtonText):
        pal.setColor(G.Disabled, role, QColor(p["disabled_fg"]))
    return pal


def apply_app_style(app: QApplication, theme_name: Optional[str] = None) -> None:
    """Install the named theme application-wide (palette + stylesheet). Live: every
    open window re-polishes instantly, no restart and no per-widget fix-ups."""
    global _ACTIVE_THEME_NAME
    name = normalize_theme_name(theme_name)
    pal = resolve_palette(THEMES[name])
    _ACTIVE_THEME_NAME = name
    _ACTIVE_PALETTE.clear()
    _ACTIVE_PALETTE.update(pal)
    app.setStyle("Fusion")
    app.setPalette(build_qpalette(pal))
    app.setStyleSheet(build_qss(pal))


def _repolish(widget: QWidget) -> QWidget:
    st = widget.style()
    st.unpolish(widget)
    st.polish(widget)
    widget.update()
    return widget


def set_role(widget: QWidget, role: str, tone: Optional[str] = None) -> QWidget:
    """Sets the QSS-visible `role` (and optional `tone`) property -- the Qt
    equivalent of ttkbootstrap's bootstyle= -- and refreshes the widget so the
    change shows immediately even if it is already on screen."""
    widget.setProperty("role", role)
    if tone is not None:
        widget.setProperty("tone", tone)
    return _repolish(widget)


def set_flag(widget: QWidget, name: str, on: bool) -> QWidget:
    """Sets a boolean dynamic property that QSS can match as [name="true"] (e.g. a `compact` variant of a
    role) and refreshes the widget. No-op when the value is unchanged."""
    value = "true" if on else "false"
    if widget.property(name) == value:
        return widget
    widget.setProperty(name, value)
    return _repolish(widget)


def set_tone(widget: QWidget, tone: Optional[str]) -> QWidget:
    widget.setProperty("tone", tone or "")
    return _repolish(widget)


def hline() -> QFrame:
    f = QFrame()
    f.setProperty("role", "sep")
    f.setFixedHeight(1)
    f.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    return f


def vline() -> QFrame:
    f = QFrame()
    f.setProperty("role", "sep")
    f.setFixedWidth(1)
    f.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
    return f


def make_label(text: str = "", role: Optional[str] = None, tone: Optional[str] = None,
               wrap: bool = False, align: Optional[Qt.AlignmentFlag] = None) -> QLabel:
    lbl = QLabel(text)
    if role:
        lbl.setProperty("role", role)
    if tone:
        lbl.setProperty("tone", tone)
    if wrap:
        lbl.setWordWrap(True)
    if align is not None:
        lbl.setAlignment(align)
    return lbl


def make_button(text: str, role: Optional[str] = None, slot: Optional[Callable] = None,
                tooltip: str = "", min_height: int = 0) -> QPushButton:
    btn = QPushButton(text)
    if role:
        btn.setProperty("role", role)
    if tooltip:
        btn.setToolTip(tooltip)
    if min_height:
        btn.setMinimumHeight(min_height)
    if slot is not None:
        btn.clicked.connect(lambda _checked=False: slot())
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    return btn


# =============================================================================
#  12. SHARED UI HELPERS
# =============================================================================
class Messagebox:
    """Static helpers matching the exact call shape this app already uses
    everywhere -- title first, then message, an optional parent= kwarg,
    and real booleans back from confirmation dialogs -- so every existing
    call site (`Messagebox.show_error("Oops", "...", parent=self)`) keeps
    working completely unchanged. Only the implementation is new: QMessageBox
    instead of ttkbootstrap's dialogs."""

    @staticmethod
    def show_info(title, message, *args, parent=None, **kwargs):
        QMessageBox.information(parent, title, message)

    @staticmethod
    def show_warning(title, message, *args, parent=None, **kwargs):
        QMessageBox.warning(parent, title, message)

    @staticmethod
    def show_error(title, message, *args, parent=None, **kwargs):
        QMessageBox.critical(parent, title, message)

    @staticmethod
    def show_question(title, message, *args, parent=None, **kwargs):
        return QMessageBox.question(parent, title, message)

    @staticmethod
    def ok(title, message, *args, parent=None, **kwargs):
        QMessageBox.information(parent, title, message)

    @staticmethod
    def okcancel(title, message, *args, parent=None, **kwargs) -> bool:
        """True only if the user actually clicked OK."""
        result = QMessageBox.question(
            parent, title, message,
            QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel,
        )
        return result == QMessageBox.StandardButton.Ok

    @staticmethod
    def yesno(title, message, *args, parent=None, **kwargs) -> bool:
        """True only if the user actually clicked Yes."""
        result = QMessageBox.question(
            parent, title, message,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return result == QMessageBox.StandardButton.Yes


# -----------------------------------------------------------------------------
#  Number fields
# -----------------------------------------------------------------------------
class _DecimalValidator(QValidator):
    """Allows only a non-negative decimal number with at most
    `max_decimal_places` digits after the point, live as the user types.
    Applied via QLineEdit.setValidator() -- Qt's equivalent of the Tkinter
    build's `entry.configure(validate="key", validatecommand=...)`."""
    def __init__(self, max_decimal_places: int = 2, parent=None):
        super().__init__(parent)
        self._pattern = re.compile(r"^\d*\.?\d*$")
        self._max_places = max_decimal_places

    def validate(self, text: str, pos: int):
        if text == "":
            return QValidator.State.Acceptable, text, pos
        if not self._pattern.match(text) or len(text.partition(".")[2]) > self._max_places:
            return QValidator.State.Invalid, text, pos
        return QValidator.State.Acceptable, text, pos


def restrict_to_decimal(line_edit: QLineEdit, max_decimal_places: int = 2) -> QLineEdit:
    line_edit.setValidator(_DecimalValidator(max_decimal_places, line_edit))
    return line_edit


def restrict_to_digits(line_edit: QLineEdit, max_length: Optional[int] = None) -> QLineEdit:
    if max_length:
        line_edit.setMaxLength(max_length)
    line_edit.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d*$"), line_edit))
    return line_edit


class _BlurFormatFilter(QObject):
    """Reformats a field to a fixed number of decimal places the moment it
    loses focus (e.g. "5" -> "5.00"). QLineEdit has no del <FocusOut> event
    the way a Tk Entry does, so this uses an event filter instead -- one
    small object per field, parented to the field so Qt cleans it up
    automatically when the field is destroyed."""
    def __init__(self, places: int, parent=None):
        super().__init__(parent)
        self.places = places

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.FocusOut:
            current = obj.text()
            formatted = f"{safe_float(current):.{self.places}f}"
            if formatted != current:
                obj.setText(formatted)
        return False


def format_as_decimal_on_blur(line_edit: QLineEdit, places: int = 2) -> QLineEdit:
    line_edit.installEventFilter(_BlurFormatFilter(places, line_edit))
    return line_edit


# -----------------------------------------------------------------------------
#  Pictures, tooltips and small widgets
# -----------------------------------------------------------------------------
def load_app_logo_pixmap(max_width: int = 160) -> Optional[QPixmap]:
    """Loads the bundled app logo (sidebar + splash screen), scaled to at
    most `max_width` px wide. SIMPLER THAN THE TKINTER BUILD: Qt's QPixmap
    already does smooth, high-quality scaling to any size on its own -- the
    Pillow-with-a-Tk-PhotoImage-fallback dance the old load_app_logo_image()
    needed (Tk's native PhotoImage.subsample() only scales by whole-number
    factors) simply isn't necessary here."""
    if not APP_LOGO_PNG.exists():
        return None
    try:
        pixmap = QPixmap(str(APP_LOGO_PNG))
        if pixmap.isNull():
            return None
        if pixmap.width() > max_width:
            pixmap = pixmap.scaledToWidth(max_width, Qt.TransformationMode.SmoothTransformation)
        return pixmap
    except Exception as exc:
        logger.warning(f"Could not load app logo image: {exc}")
        return None


def load_thumbnail_pixmap(path: Optional[Path], max_size: int = 44) -> Optional[QPixmap]:
    """Same idea as load_app_logo_pixmap(), generalised to any path and
    target size -- used for the asset-upload preview rows in Settings."""
    if not path or not path.exists():
        return None
    try:
        pixmap = QPixmap(str(path))
        if pixmap.isNull():
            return None
        return pixmap.scaled(max_size, max_size, Qt.AspectRatioMode.KeepAspectRatio,
                              Qt.TransformationMode.SmoothTransformation)
    except Exception as exc:
        logger.warning(f"Could not load thumbnail for {path}: {exc}")
        return None


class ToolTip:
    """Qt widgets already have a built-in hover tooltip
    (QWidget.setToolTip()), so unlike the Tkinter build there is no custom
    popup Toplevel to build and manage here -- this class exists only so
    every `ToolTip(widget, "explanation")` call site elsewhere in the file
    keeps working unchanged."""
    def __init__(self, widget: QWidget, text: str, delay: int = 500):
        widget.setToolTip(text)


class Debouncer(QObject):
    """Coalesces rapid repeated calls (e.g. keystrokes) into one after
    `delay` ms. Same contract as the Tkinter build: construct once, then
    call the instance itself on every keystroke; `.cancel()` stops a
    pending call. Backed by QTimer instead of Tk's `.after()`."""
    def __init__(self, parent: QObject, delay: int, func: Callable):
        super().__init__(parent)
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(func)
        self._delay = delay

    def __call__(self, *args):
        self._timer.start(self._delay)

    def cancel(self):
        self._timer.stop()


class PlaceholderEntry(QLineEdit):
    """Single-line hint-text field, replacing the Tkinter build's
    PlaceholderEntry. Qt's native placeholderText already has the exact
    property that build's PlaceholderMixin had to hand-build: text() reads
    back "" while the hint is showing, so nothing can ever mistake the hint
    for real input. get()/insert()/delete()/focus_set() aliases are kept
    so the rest of this file's Tkinter-flavoured call sites need minimal
    changes; new code can just use the native Qt methods directly."""
    def __init__(self, parent=None, placeholder: str = "", **kw):
        super().__init__(parent)
        if placeholder:
            self.setPlaceholderText(placeholder)

    def get(self) -> str:
        return self.text()

    def insert(self, index, text):
        self.setText(text)

    def delete(self, first=0, last=None):
        self.clear()

    def focus_set(self):
        self.setFocus()

    def selection_range(self, start, end):
        if end in ("end", -1):
            self.selectAll()
        else:
            self.setSelection(start, end - start)


class PlaceholderText(QPlainTextEdit):
    """Multi-line hint-text box. Qt's native placeholder text replaces the hand-rolled mechanism the Tkinter build
    needed. It is a QPlainTextEdit rather than a QTextEdit on purpose: only QPlainTextEdit wraps a long hint over
    several lines (QTextEdit draws it on one clipped line), and this box only ever holds plain text anyway."""
    def __init__(self, parent=None, placeholder: str = "", height: Optional[int] = None, **kw):
        super().__init__(parent)
        if placeholder:
            self.setPlaceholderText(placeholder)
        if height:
            fm = self.fontMetrics()
            self.setFixedHeight(fm.lineSpacing() * int(height) + 22)    # + padding, border and document margin

    def setAcceptRichText(self, accept: bool) -> None:
        """No-op: a plain-text box never takes rich text. Kept so code written against QTextEdit still works."""

    def acceptRichText(self) -> bool:
        return False

    def get(self, index1="1.0", index2="end-1c") -> str:
        return self.toPlainText()

    def insert(self, index, text):
        self.setPlainText(text)

    def delete(self, index1="1.0", index2=None):
        self.clear()

    def focus_set(self):
        self.setFocus()


# -----------------------------------------------------------------------------
#  Scrolling, dates and dialogs
# -----------------------------------------------------------------------------
class _WheelGuard(QObject):
    """Stops mouse-wheel events from changing a combo box / spin box / date edit that does
    not have keyboard focus, so scrolling a long form with the wheel never silently changes
    a dropdown's value as the cursor drifts across it. Instead of relying on Qt's implicit
    event propagation, the wheel is handed straight to the enclosing scroll area's vertical
    scroll bar, so the form keeps scrolling smoothly while the cursor is over the field."""

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.Wheel and not obj.hasFocus():
            area = obj.parentWidget()
            while area is not None and not isinstance(area, QAbstractScrollArea):
                area = area.parentWidget()
            if area is not None:
                QApplication.sendEvent(area.verticalScrollBar(), event)
            event.accept()
            return True
        return False


def guard_wheel(widget: QWidget) -> QWidget:
    """Apply the wheel guard to a QComboBox / QAbstractSpinBox / QDateEdit."""
    widget.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
    widget.installEventFilter(_WheelGuard(widget))
    return widget


QT_DISPLAY_DATE_FORMAT: str = "dd MMM yyyy"          # == DISPLAY_DATE_FORMAT "%d %b %Y"


def make_date_edit(value: Optional[datetime.date] = None) -> QDateEdit:
    """A date field that shows e.g. '24 Oct 2026' with an English calendar popup, whatever
    the OS language is (Python's strptime('%b') only understands English month names, and
    the DB layer stores ISO dates, so the displayed text must never be parsed back --
    use date_edit_value()/date_edit_db() instead)."""
    edit = QDateEdit()
    edit.setLocale(QLocale(QLocale.Language.English, QLocale.Country.India))
    edit.setDisplayFormat(QT_DISPLAY_DATE_FORMAT)
    edit.setCalendarPopup(True)
    cal = edit.calendarWidget()
    if cal is not None:
        cal.setLocale(QLocale(QLocale.Language.English, QLocale.Country.India))
        cal.setGridVisible(False)
    value = value or datetime.date.today()
    edit.setDate(QDate(value.year, value.month, value.day))
    guard_wheel(edit)
    return edit


def date_edit_value(edit: QDateEdit) -> datetime.date:
    d = edit.date()
    return datetime.date(d.year(), d.month(), d.day())


def date_edit_db(edit: QDateEdit) -> str:
    """The field's date as the DB layer stores it ('YYYY-MM-DD')."""
    return date_edit_value(edit).strftime(DB_DATE_FORMAT)


def set_date_edit_from_db(edit: QDateEdit, db_date: Optional[str]) -> None:
    """Load a stored 'YYYY-MM-DD' (or empty/invalid -> today) into the field."""
    try:
        d = datetime.datetime.strptime(db_date, DB_DATE_FORMAT).date() if db_date else datetime.date.today()
    except (TypeError, ValueError):
        d = datetime.date.today()
    edit.setDate(QDate(d.year, d.month, d.day))


def documents_dir() -> Path:
    """The user's Documents folder (falls back to the home folder): where file dialogs start when there is no
    better idea of where the user wants to save to."""
    docs = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
    return Path(docs) if docs else Path.home()


def run_dialog(dialog: QDialog) -> int:
    """Run a modal dialog to completion, then release it. Always use this instead of a bare
    `dialog.exec()` in the application.

    * Qt keeps a closed dialog alive as a hidden child of its parent, so without this every
      Add/Edit click would leave one more hidden dialog on the main window for the whole session.
    * PySide6 can crash while tearing down widgets that are kept alive only by lambda slots
      (a reference cycle through a signal connection) when the interpreter exits. Deleting dialogs
      the normal Qt way as soon as they are finished (together with the orderly shutdown in
      main()) keeps that from ever happening.

    Read anything you need from the dialog (e.g. `dialog.job_result`) straight after this returns;
    the C++ object is deleted when control is back in the event loop."""
    try:
        return dialog.exec()
    finally:
        dialog.deleteLater()


def screen_area(widget: Optional[QWidget] = None) -> Optional[QRect]:
    """The usable part (the screen minus the taskbar) of the screen `widget` is on -- the one place every window-size
    decision asks, so they all agree. None when there is no screen."""
    screen = (widget.screen() if widget is not None else None) or QGuiApplication.primaryScreen()
    return screen.availableGeometry() if screen is not None else None


def _fit_to_screen(dialog: QDialog, width: int, height: int, min_width: int, min_height: int) -> None:
    """Sizes a dialog as the Tk build did, but never larger than the screen it opens on (many laptops are 768 px tall)."""
    area = screen_area(dialog)
    if area is not None:
        width, height = min(width, max(320, area.width() - 40)), min(height, max(320, area.height() - 60))
    dialog.setMinimumSize(min(min_width, width), min(min_height, height))
    dialog.resize(width, height)


def _no_default_button(*buttons: QPushButton) -> None:
    for btn in buttons:
        btn.setAutoDefault(False)
        btn.setDefault(False)


class KeyTable(QTableWidget):
    """QTableWidget that turns Delete and Enter/Return into `delete_pressed` / `enter_pressed` signals
    (and swallows those two keys, so a table with no handler connected does nothing, exactly like the
    Tk Treeview). Used by the Dashboard, Clients and Catalog lists."""
    delete_pressed = Signal()
    enter_pressed = Signal()

    def __init__(self, *args):
        super().__init__(*args)
        self._width_floors: Dict[int, int] = {}    # column -> the widest width its content has ever needed (fit_columns)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Delete:
            self.delete_pressed.emit()
            return
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.enter_pressed.emit()
            return
        super().keyPressEvent(event)

    def fit_columns(self, samples: Dict[int, List[str]], padding: int = 24, cap: int = 260) -> None:
        """Widen -- never narrow -- the listed columns so the widest of their sample texts, the header text and
        whatever the column currently shows fit in this table's own font. The designed column widths suit Segoe UI;
        another font, a longer currency symbol or a bigger amount would otherwise end in "…" (a cut-off amount
        reads as a different amount). Only the columns' *needed* width is remembered: a column the user has since
        dragged narrower stays where they put it until something wider than ever before has to be shown. The
        stretch column absorbs the difference."""
        header = self.horizontalHeader()
        self.ensurePolished()                      # the stylesheet's font (not the application default) is what is drawn
        header.ensurePolished()
        body_fm = self.fontMetrics()
        head_font = QFont(header.font())
        head_font.setBold(True)
        head_fm = QFontMetrics(head_font)
        floors = self._width_floors
        for col, texts in samples.items():
            if header.sectionResizeMode(col) == QHeaderView.ResizeMode.Stretch:
                continue
            shown = [self.item(r, col).text() for r in range(self.rowCount()) if self.item(r, col) is not None]
            need = max((body_fm.horizontalAdvance(t) for t in [*texts, *shown]), default=0) + padding
            head = self.horizontalHeaderItem(col)
            if head is not None:                   # room for the sort arrow too (any column can become the sorted one)
                need = max(need, head_fm.horizontalAdvance(head.text().rstrip(" ▲▼") + "  ▼") + padding)
            need = min(need, cap)
            if need > floors.get(col, 0):
                floors[col] = need
                if self.columnWidth(col) < need:
                    self.setColumnWidth(col, need)

    def select_only_row(self, row: int) -> None:
        """Make `row` the one and only selected row, whatever keyboard modifiers are held at that moment.
        (QTableView.selectRow() honours Ctrl = toggle / Shift = extend, which is wrong for "the user
        right-clicked this row".)"""
        index = self.model().index(row, 0)
        sm = self.selectionModel()
        sm.setCurrentIndex(index, QItemSelectionModel.SelectionFlag.NoUpdate)
        sm.select(index, QItemSelectionModel.SelectionFlag.ClearAndSelect | QItemSelectionModel.SelectionFlag.Rows)


# =============================================================================
#  13. PDF VIEWER, PRINTING & SHARING
# =============================================================================
PDF_FIT_FILL: float = 0.97                # Tk: the fitted page fills 97% of the canvas
PDF_ZOOM_STEP: float = 0.15
PDF_PRINT_MAX_DPI: int = 300
PDF_MAX_RENDER_PIXELS: int = 64_000_000   # safety net for the biggest zoom on a 4K / HiDPI screen
_VIEWER_STATE: Dict[str, Any] = {"save_dir": None}      # folder of the last "Save a copy" (this session)


def _is_windows() -> bool:
    """A tiny seam, so the Windows-only printing path can be exercised by tests on any OS."""
    return sys.platform == "win32"


def _load_qtpdf():
    """The QPdfDocument class, or None when PySide6's QtPdf module is not installed (it ships in the
    PySide6-Addons wheel, which a plain `pip install PySide6` includes). Imported on first use so the
    application still starts without it -- the invoice is then simply opened in the default PDF program."""
    try:
        from PySide6.QtPdf import QPdfDocument
    except ImportError:
        return None
    return QPdfDocument


def _load_print_support():
    """(QPrintDialog, QPrinter, QPrinterInfo) from QtPrintSupport, or None when that module is missing."""
    try:
        from PySide6.QtPrintSupport import QPrintDialog, QPrinter, QPrinterInfo
    except ImportError:
        return None
    return QPrintDialog, QPrinter, QPrinterInfo


def open_in_system_viewer(path) -> bool:
    """Open a file in the operating system's default program for it (Tk: os.startfile). True when the OS
    accepted the request."""
    return bool(QDesktopServices.openUrl(QUrl.fromLocalFile(str(path))))


def _pdf_error_text(QPdfDocument, error) -> str:
    """Plain-English reason for a QPdfDocument.load() failure."""
    E = QPdfDocument.Error
    return {
        E.FileNotFound: "The file was not found.",
        E.InvalidFileFormat: "This is not a valid PDF document.",
        E.IncorrectPassword: "This PDF is password-protected.",
        E.UnsupportedSecurityScheme: "This PDF uses an unsupported security scheme.",
    }.get(error, "The PDF could not be loaded.")


def _render_on_white(doc, page: int, size: QSize) -> QImage:
    """One PDF page as an opaque image on white paper. PDFium renders onto a *transparent* bitmap, so everything
    the PDF leaves unpainted (most of an invoice's paper) would show whatever is behind the image -- the page's
    shadow in the viewer, or an undefined background on a print-to-PDF printer. Composite it on white instead."""
    rendered = doc.render(page, size)
    paper = QImage(rendered.size(), QImage.Format.Format_RGB32)
    paper.fill(QColor("#FFFFFF"))
    painter = QPainter(paper)
    painter.drawImage(0, 0, rendered)
    painter.end()
    return paper


def _default_save_dir() -> Path:
    last = _VIEWER_STATE.get("save_dir")
    if last and Path(last).is_dir():
        return Path(last)
    return documents_dir()


class _PageCanvas(QWidget):
    """The viewer's drawing surface: a neutral backdrop, the page's drop shadow and 1 px edge, and the rendered
    page, centred. It lives in a QScrollArea, so a page zoomed beyond the window can be scrolled. The backdrop,
    shadow and edge colours come from the active theme (the Tk canvas used one fixed grey)."""
    zoom_wheel = Signal(int)            # +1 / -1: Ctrl + mouse wheel
    SHADOW = 4

    def __init__(self, parent=None):
        super().__init__(parent)
        self._image: Optional[QImage] = None
        self._page_px = QSize(0, 0)        # the page's size in logical pixels
        self.setMinimumSize(1, 1)

    @property
    def page_size(self) -> QSize:
        return QSize(self._page_px)

    def set_image(self, image: Optional[QImage]) -> None:
        self._image = image
        if image is None:
            self._page_px = QSize(0, 0)
            self.setMinimumSize(1, 1)
        else:
            dpr = image.devicePixelRatio() or 1.0
            self._page_px = QSize(round(image.width() / dpr), round(image.height() / dpr))
            # Never smaller than the page (plus its shadow and edge), so a page zoomed beyond the window gets scroll
            # bars. At "fit" the page is at most 97% of the viewport, so this never forces scroll bars there.
            self.setMinimumSize(self._page_px.width() + self.SHADOW + 1, self._page_px.height() + self.SHADOW + 1)
        self.update()

    def paintEvent(self, event):
        p = QPainter(self)
        p.fillRect(self.rect(), QColor(theme_color("backdrop")))
        if self._image is None:
            return
        w, h = self._page_px.width(), self._page_px.height()
        x = max((self.width() - w) // 2, 0)
        y = max((self.height() - h) // 2, 0)
        shadow = QColor(theme_color("shadow"))
        shadow.setAlpha(110)
        p.fillRect(x, y, w + self.SHADOW, h + self.SHADOW, shadow)       # soft drop shadow, offset to the bottom right
        p.setPen(QPen(QColor(theme_color("border_strong")), 1))
        p.drawRect(x - 1, y - 1, w + 1, h + 1)                            # crisp 1 px edge exactly around the page
        p.drawImage(QPoint(x, y), self._image)

    def wheelEvent(self, event):
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            step = event.angleDelta().y()
            if step:
                self.zoom_wheel.emit(1 if step > 0 else -1)
            event.accept()
            return
        # Plain wheel: hand it straight to the enclosing scroll area's scroll bar (like the form's wheel guard does) --
        # relying on Qt's implicit propagation to the parent is not dependable for every kind of event source.
        area = self.parentWidget()
        while area is not None and not isinstance(area, QAbstractScrollArea):
            area = area.parentWidget()
        if area is None:
            event.ignore()
            return
        delta = event.angleDelta()
        bar = area.verticalScrollBar() if (delta.y() or not delta.x()) else area.horizontalScrollBar()
        QApplication.sendEvent(bar, event)
        event.accept()


class _PrinterDialog(QDialog):
    """`choice` is None (cancelled) or ("print", <printer name>) / ("system", None) / ("manual", None)."""

    def __init__(self, parent, printers: List[str], default_printer: str = ""):
        super().__init__(parent)
        self.setWindowTitle("Select Printer")
        self.setModal(True)
        self.choice: Optional[Tuple[str, Optional[str]]] = None
        self._printers = list(printers)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 15, 20, 15)
        lay.setSpacing(8)
        lay.setSizeConstraint(QLayout.SizeConstraint.SetFixedSize)           # Tk: resizable(False, False)

        lay.addWidget(make_label("🖨️ Choose a printer destination:", "h3"))

        self.combo = QComboBox()
        self.combo.setMinimumWidth(360)                                       # => a 400 px wide dialog, as in Tk
        if self._printers:
            self.combo.addItems(self._printers)
            self.combo.setCurrentText(default_printer if default_printer in self._printers else self._printers[0])
        else:
            self.combo.addItem("No printers detected")
            self.combo.setEnabled(False)
        lay.addWidget(self.combo)

        row = QHBoxLayout()
        row.setContentsMargins(0, 2, 0, 2)
        row.addStretch(1)
        self.print_btn = make_button("Print Document", "success", self._on_print)
        self.cancel_btn = make_button("Cancel", "secondary-outline", self.reject)
        row.addWidget(self.print_btn)
        row.addSpacing(10)
        row.addWidget(self.cancel_btn)
        lay.addLayout(row)
        self.print_btn.setEnabled(bool(self._printers))

        lay.addWidget(hline())
        self.system_btn = make_button("⚙️ Use System Print Dialog", "info-outline", self._on_system_dialog)
        self.manual_btn = make_button("🌐 Open PDF Manually (Browser/Viewer)", "primary-outline", self._on_manual_open)
        lay.addWidget(self.system_btn)
        lay.addWidget(self.manual_btn)

        for btn in (self.print_btn, self.cancel_btn, self.system_btn, self.manual_btn):
            btn.setAutoDefault(False)
        if self._printers:
            self.print_btn.setDefault(True)                                   # Enter prints to the selected printer

    def _on_print(self):
        if not self._printers:
            return
        self.choice = ("print", self.combo.currentText())
        self.accept()

    def _on_system_dialog(self):
        self.choice = ("system", None)
        self.accept()

    def _on_manual_open(self):
        self.choice = ("manual", None)
        self.accept()


class PDFViewer(QDialog):
    MIN_ZOOM = 0.4
    MAX_ZOOM = 3.5

    def __init__(self, parent, pdf_path: Path, invoice_data: Optional[Dict] = None):
        super().__init__(parent)
        self.pdf_path = Path(pdf_path)
        self.invoice_data = invoice_data
        self.doc = None
        self.page_idx = 0
        self.total = 0
        self.zoom = 1.0
        self._failed = False        # the PDF could not be opened: show()/exec() become no-ops
        self._released = False
        self._placed = False

        self.setWindowTitle(f"Preview  —  {self.pdf_path.name}")
        self.setModal(True)                                       # Tk: transient + grab_set() (non-blocking show())
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)    # nobody keeps a reference: free it when it closes
        self.setWindowFlag(Qt.WindowType.WindowMaximizeButtonHint, True)

        # --- Size (Tk: 900x720, also the minimum -- smaller only on a screen that cannot hold it; centred on first show) ---
        _fit_to_screen(self, 900, 720, 900, 720)

        if not self._open_document():
            return

        self._resize_timer = QTimer(self)                         # Tk: after(130, _render) debounce on <Configure>
        self._resize_timer.setSingleShot(True)
        self._resize_timer.timeout.connect(self._render)

        self._build()
        self._bind_keys()

    # ------------------------------------------------------------------ #
    #  Opening the document
    # ------------------------------------------------------------------ #
    def _open_document(self) -> bool:
        QPdfDocument = _load_qtpdf()
        if QPdfDocument is None:
            return self._give_up(
                "The built-in preview needs PySide6's QtPdf module (it is part of `pip install PySide6`).\n"
                "The invoice will open in your default PDF viewer instead.", external=True)
        doc = QPdfDocument(self)
        error = doc.load(str(self.pdf_path))
        if error != QPdfDocument.Error.None_:
            doc.close()
            return self._give_up(f"Cannot open file:\n{self.pdf_path}\n\n{_pdf_error_text(QPdfDocument, error)}")
        if doc.pageCount() < 1:
            doc.close()
            return self._give_up(f"Cannot open file:\n{self.pdf_path}\n\nThe PDF has no pages.")
        self.doc = doc
        self.total = doc.pageCount()
        return True

    def _give_up(self, message: str, external: bool = False) -> bool:
        """Tk: Messagebox.show_error("PDF Error", ...); self.destroy(); return."""
        self._failed = True
        Messagebox.show_error("PDF Error", message, parent=self.parentWidget())
        if external:
            open_in_system_viewer(self.pdf_path)
        self.deleteLater()
        return False

    def setVisible(self, visible: bool):
        if visible and self._failed:
            return                                                 # the Tk window was destroyed in __init__
        super().setVisible(visible)

    def exec(self):
        if self._failed:
            return int(QDialog.DialogCode.Rejected)
        return super().exec()

    # ------------------------------------------------------------------ #
    #  UI
    # ------------------------------------------------------------------ #
    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        bar = QWidget()
        tb = QHBoxLayout(bar)
        tb.setContentsMargins(10, 7, 10, 7)
        tb.setSpacing(4)

        self.prev_btn = make_button("◀  Prev", "primary", self._prev, "Previous page (← or Page Up)")
        self.next_btn = make_button("Next  ▶", "primary", self._next, "Next page (→ or Page Down)")
        for b in (self.prev_btn, self.next_btn):
            b.setMinimumWidth(88)
        tb.addWidget(self.prev_btn)
        tb.addWidget(self.next_btn)
        tb.addSpacing(10)
        self.page_lbl = QLabel("")
        tb.addWidget(self.page_lbl)
        tb.addStretch(1)

        # Right-hand group, left to right: Fit  ⊖  ⊕  100%  |  Print  Save  Share  Close   (Tk packed it right to left)
        self.fit_btn = make_button("Fit", "secondary-outline", self._reset_zoom, "Fit the page to the window (Ctrl+0)")
        self.zoom_out_btn = make_button("⊖", "secondary-outline", self._zoom_out, "Zoom out (Ctrl + −)")
        self.zoom_in_btn = make_button("⊕", "secondary-outline", self._zoom_in, "Zoom in (Ctrl + +)")
        self.fit_btn.setMinimumWidth(48)
        for b in (self.zoom_out_btn, self.zoom_in_btn):
            b.setMinimumWidth(40)
        self.zoom_lbl = QLabel("100%")
        self.zoom_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.zoom_lbl.setMinimumWidth(52)
        self.print_btn = make_button("🖨  Print", "success", self._print, "Print this PDF (Ctrl+P)")
        self.save_btn = make_button("💾  Save", "secondary", self._save_copy, "Save a copy of this PDF (Ctrl+E)")
        self.share_btn = make_button("📤 Share", "info", self._open_share, "Share this invoice by WhatsApp or e-mail")
        self.close_btn = make_button("✕  Close", "danger", self._close, "Close this viewer (Esc)")
        for b in (self.print_btn, self.save_btn, self.share_btn, self.close_btn):
            b.setMinimumWidth(88)

        for w in (self.fit_btn, self.zoom_out_btn, self.zoom_in_btn, self.zoom_lbl):
            tb.addWidget(w)
        tb.addSpacing(6)
        tb.addWidget(vline())
        tb.addSpacing(6)
        for w in (self.print_btn, self.save_btn, self.share_btn, self.close_btn):
            tb.addWidget(w)
        for b in bar.findChildren(QPushButton):
            b.setAutoDefault(False)                                  # Enter must not "click" a random toolbar button

        root.addWidget(bar)
        root.addWidget(hline())

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.canvas = _PageCanvas()
        self.canvas.zoom_wheel.connect(self._wheel_zoom)
        self.scroll.setWidget(self.canvas)
        root.addWidget(self.scroll, 1)

        hint = make_label("← → or Page Up/Down to navigate  •  Ctrl + / − zoom  •  Ctrl+0 fit  •  "
                          "Ctrl+P print  •  Ctrl+E save  •  Esc close", "small", align=Qt.AlignmentFlag.AlignCenter)
        hint.setContentsMargins(0, 3, 0, 3)
        root.addWidget(hint)

        self._update_nav_state()

    def _bind_keys(self):
        """Local shortcuts (Tk: self.bind). WidgetWithChildren context = active whenever this window has the focus."""
        self._shortcuts = []
        for keys, slot in (
            (("Left", "PgUp"), self._prev),
            (("Right", "PgDown"), self._next),
            (("Ctrl+=", "Ctrl++"), self._zoom_in),
            (("Ctrl+-",), self._zoom_out),
            (("Ctrl+0",), self._reset_zoom),
            (("Ctrl+P",), self._print),
            (("Ctrl+E",), self._save_copy),
        ):
            for key in keys:
                sc = QShortcut(QKeySequence(key), self)
                sc.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
                sc.activated.connect(slot)
                self._shortcuts.append(sc)

    # ------------------------------------------------------------------ #
    #  Window events
    # ------------------------------------------------------------------ #
    def showEvent(self, event):
        super().showEvent(event)
        if not self._placed:
            self._placed = True
            area = screen_area(self.parentWidget())
            if area is not None:                                       # Tk: centred on the screen
                self.move(area.center() - QPoint(self.width() // 2, self.height() // 2))
        self.scroll.setFocus()
        self._schedule_render(30)                                      # first picture as soon as the window has its size

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if not self._failed:
            self._schedule_render(130)                                 # re-render once the user stops resizing

    def event(self, event):
        # Dragged to a monitor with another pixel density (100 % -> 150 %): draw the page again at that screen's
        # resolution, otherwise it stays as soft as the old monitor's bitmap.
        if event.type() == QEvent.Type.DevicePixelRatioChange and not getattr(self, "_failed", True) \
                and not getattr(self, "_released", True):
            self._schedule_render(30)
        return super().event(event)

    def _schedule_render(self, delay_ms: int):
        """(Re)start the debounce timer. QTimer.start(msec) also *changes* the interval, so the delay is passed
        every time instead of being set once."""
        self._resize_timer.start(delay_ms)

    def _release_document(self):
        """Let go of the PDF file (on Windows an open handle would block deleting / overwriting the invoice)."""
        if self._released:
            return
        self._released = True
        self._resize_timer.stop()
        if self.doc is not None:
            self.doc.close()

    def done(self, result):
        self._release_document()
        super().done(result)

    def closeEvent(self, event):
        self._release_document()
        super().closeEvent(event)

    def _close(self):
        self.close()

    # ------------------------------------------------------------------ #
    #  Rendering
    # ------------------------------------------------------------------ #
    def _render(self, reset_scroll: bool = False):
        if self._released or self.doc is None:
            return
        viewport = self.scroll.maximumViewportSize()      # the viewport as if there were no scroll bars: stable fit
        cw, ch = viewport.width(), viewport.height()
        if cw <= 1 or ch <= 1:
            return
        page_pt = self.doc.pagePointSize(self.page_idx)
        if page_pt.width() <= 0 or page_pt.height() <= 0:
            return
        fit_z = min(cw / page_pt.width(), ch / page_pt.height()) * PDF_FIT_FILL
        z = fit_z * self.zoom
        w = max(int(page_pt.width() * z), 1)
        h = max(int(page_pt.height() * z), 1)

        dpr = self.devicePixelRatioF() or 1.0
        pw, ph = max(round(w * dpr), 1), max(round(h * dpr), 1)
        if pw * ph > PDF_MAX_RENDER_PIXELS:                 # never ask PDFium for an absurdly large bitmap
            shrink = (PDF_MAX_RENDER_PIXELS / (pw * ph)) ** 0.5
            pw, ph = max(int(pw * shrink), 1), max(int(ph * shrink), 1)
        image = _render_on_white(self.doc, self.page_idx, QSize(pw, ph))
        image.setDevicePixelRatio(pw / w)                    # logical size stays w x h whatever the bitmap size

        # Keep the point at the middle of the window where it is while zooming / resizing.
        hbar, vbar = self.scroll.horizontalScrollBar(), self.scroll.verticalScrollBar()
        view = self.scroll.viewport().size()
        old = self.canvas.size()
        cx = (hbar.value() + view.width() / 2) / max(old.width(), 1)
        cy = (vbar.value() + view.height() / 2) / max(old.height(), 1)

        self.canvas.set_image(image)
        self.canvas.resize(viewport.expandedTo(self.canvas.minimumSize()))   # adopt the new size now (updates the bars)
        QApplication.sendPostedEvents(None, int(QEvent.Type.LayoutRequest))
        new = self.canvas.size()
        view = self.scroll.viewport().size()
        if reset_scroll:
            hbar.setValue(0)
            vbar.setValue(0)
        else:
            hbar.setValue(round(cx * new.width() - view.width() / 2))
            vbar.setValue(round(cy * new.height() - view.height() / 2))

        self.page_lbl.setText(f"Page {self.page_idx + 1} / {self.total}")
        self.zoom_lbl.setText(f"{round(self.zoom * 100)}%")     # Tk: int() -> 1.3 showed as "129%"
        self._update_nav_state()

    def _update_nav_state(self):
        self.prev_btn.setEnabled(self.page_idx > 0)
        self.next_btn.setEnabled(self.page_idx < self.total - 1)

    # ------------------------------------------------------------------ #
    #  Navigation / zoom
    # ------------------------------------------------------------------ #
    def _next(self):
        if self.page_idx < self.total - 1:
            self.page_idx += 1
            self._render(reset_scroll=True)

    def _prev(self):
        if self.page_idx > 0:
            self.page_idx -= 1
            self._render(reset_scroll=True)

    def _zoom(self, delta: float):
        self.zoom = round(max(self.MIN_ZOOM, min(self.MAX_ZOOM, self.zoom + delta)), 4)
        self._render()

    def _zoom_in(self):
        self._zoom(+PDF_ZOOM_STEP)

    def _zoom_out(self):
        self._zoom(-PDF_ZOOM_STEP)

    def _wheel_zoom(self, direction: int):
        self._zoom(PDF_ZOOM_STEP * direction)

    def _reset_zoom(self):
        self.zoom = 1.0
        self._render()

    # ------------------------------------------------------------------ #
    #  Printing
    # ------------------------------------------------------------------ #
    def _print(self):
        if not _is_windows():
            try:
                subprocess.run(["lp", str(self.pdf_path)], check=True)
            except Exception as exc:
                Messagebox.show_error("Print Error", str(exc), parent=self)
            return

        support = _load_print_support()
        if support is None:
            self._print_via_shell()      # Tk: the "PyQt5 is missing" fallback
            return
        QPrinterInfo = support[2]
        printers = list(QPrinterInfo.availablePrinterNames())
        default_printer = QPrinterInfo.defaultPrinterName()

        dialog = _PrinterDialog(self, printers, default_printer)
        run_dialog(dialog)
        choice = dialog.choice
        if choice is None:
            return
        kind, name = choice
        if kind == "print":
            self._print_with_qt(printer_name=name)
        elif kind == "system":
            self._print_with_qt(show_dialog=True)
        else:
            self._open_manually()

    def _print_via_shell(self):
        """Last resort when QtPrintSupport is unavailable: ask Windows to print the file with its default PDF program."""
        try:
            os.startfile(self.pdf_path, "print")
        except OSError:
            Messagebox.show_warning(
                "Manual Print Required",
                "A PDF reader is not set as default.\n"
                "The file will open — press Ctrl+P to print.", parent=self)
            self._open_manually()

    def _open_manually(self):
        if not open_in_system_viewer(self.pdf_path):
            Messagebox.show_error("Open Error", f"Windows could not open the PDF:\n{self.pdf_path}", parent=self)

    def _print_with_qt(self, printer_name: Optional[str] = None, show_dialog: bool = False):
        """Both ways of printing: straight to a named printer, or through the system print dialog."""
        support = _load_print_support()
        if support is None:
            self._print_via_shell()
            return
        QPrintDialog, QPrinter, _ = support

        painter = None
        busy = False
        try:
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            if show_dialog:
                print_dialog = QPrintDialog(printer, self)
                accepted = run_dialog(print_dialog) == int(QDialog.DialogCode.Accepted)
                if not accepted:
                    return
            else:
                printer.setPrinterName(printer_name)

            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            busy = True
            painter = QPainter()
            if not painter.begin(printer):
                raise RuntimeError("The print job could not be started.\n"
                                   "Check that the printer is switched on and not in use by another program.")
            dpi = min(printer.resolution() or PDF_PRINT_MAX_DPI, PDF_PRINT_MAX_DPI)
            scale = dpi / 72.0
            for page_index in range(self.total):
                if page_index > 0:
                    printer.newPage()
                size_pt = self.doc.pagePointSize(page_index)
                image = _render_on_white(self.doc, page_index, QSize(max(round(size_pt.width() * scale), 1),
                                                                     max(round(size_pt.height() * scale), 1)))
                # Fit the page into the printable area, centred. (Tk moved the painter's viewport cumulatively, so
                # every page after the first was drawn at the wrong offset.)
                area = painter.viewport()
                fitted = image.size().scaled(area.size(), Qt.AspectRatioMode.KeepAspectRatio)
                target = QRect(area.x() + (area.width() - fitted.width()) // 2,
                               area.y() + (area.height() - fitted.height()) // 2,
                               fitted.width(), fitted.height())
                painter.drawImage(target, image)
        except Exception as exc:
            if busy:
                QApplication.restoreOverrideCursor()
                busy = False
            Messagebox.show_error("Print Error", str(exc), parent=self)
            return
        finally:
            # Tk BUG FIX kept: cleanup is on every path (a failure part-way used to leave the painter attached to
            # the print spooler).
            if painter is not None and painter.isActive():
                painter.end()
            if busy:
                QApplication.restoreOverrideCursor()
        if not show_dialog:
            Messagebox.show_info("Success", f"Document sent to {printer_name}!", parent=self)

    # ------------------------------------------------------------------ #
    #  Save a copy / share
    # ------------------------------------------------------------------ #
    def _save_copy(self):
        start = str(_default_save_dir() / self.pdf_path.name)
        path, _ = QFileDialog.getSaveFileName(self, "Save As", start, "PDF Documents (*.pdf)")
        if not path:
            return
        if not Path(path).suffix:                    # Tk: defaultextension=".pdf"
            path += ".pdf"
        try:
            shutil.copy(self.pdf_path, path)
            _VIEWER_STATE["save_dir"] = str(Path(path).parent)
            Messagebox.show_info("Saved", f"Invoice saved to:\n{path}", parent=self)
        except Exception as exc:
            Messagebox.show_error("Save Error", str(exc), parent=self)

    def _open_share(self):
        run_dialog(ShareDialog(self, self.pdf_path, self.invoice_data))


class ShareDialog(QDialog):
    """Mini dialog to share the invoice via WhatsApp or the default e-mail program."""

    def __init__(self, parent, pdf_path: Path, invoice_data: Optional[Dict]):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.invoice_data = dict(invoice_data) if invoice_data else {}

        self.setWindowTitle("Share Invoice")
        self.setModal(True)
        self._build()
        self._prefill_data()
        self.phone_entry.setFocus()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(0)
        lay.setSizeConstraint(QLayout.SizeConstraint.SetFixedSize)         # Tk: 420x340, not resizable

        lay.addWidget(make_label("📤 Share Invoice", "h2"))
        lay.addSpacing(15)

        lay.addWidget(make_label("WhatsApp Number (with Country Code):", "caption"))
        lay.addSpacing(2)
        self.phone_entry = QLineEdit()
        self.phone_entry.setMinimumWidth(380)                                # => a 420 px wide dialog, as in Tk
        self.phone_entry.setPlaceholderText("e.g. +91 98765 43210")
        self.phone_entry.returnPressed.connect(self._send_whatsapp)
        lay.addWidget(self.phone_entry)
        lay.addSpacing(10)

        lay.addWidget(make_label("Client Email Address:", "caption"))
        lay.addSpacing(2)
        self.email_entry = QLineEdit()
        self.email_entry.setPlaceholderText("e.g. accounts@client.com")
        self.email_entry.returnPressed.connect(self._send_email)
        lay.addWidget(self.email_entry)
        lay.addSpacing(15)

        row = QHBoxLayout()
        row.setSpacing(10)
        self.whatsapp_btn = make_button("💬 WhatsApp", "success", self._send_whatsapp, min_height=32)
        self.email_btn = make_button("✉️ Email", "primary", self._send_email, min_height=32)
        row.addWidget(self.whatsapp_btn, 1)
        row.addWidget(self.email_btn, 1)
        lay.addLayout(row)
        lay.addSpacing(15)

        note = make_label("Note: The PDF file path will be copied to your clipboard.\n"
                          "Simply paste (Ctrl+V) when attaching the file.", "small",
                          align=Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(note)
        for btn in (self.whatsapp_btn, self.email_btn):
            btn.setAutoDefault(False)

    def _prefill_data(self):
        contact_info = self.invoice_data.get("client_contact", "") or ""
        if not contact_info:
            return
        # Basic regex to pull an e-mail address and a phone number out of the free-text contact box
        email_match = re.search(r'[\w\.-]+@[\w\.-]+', contact_info)
        if email_match:
            self.email_entry.setText(email_match.group(0))
        phone_match = re.search(r'\+?\d[\d\s-]{8,}\d', contact_info)
        if phone_match:
            self.phone_entry.setText(re.sub(r'[^\d+]', '', phone_match.group(0)))   # digits and a leading + only

    def _copy_path_to_clipboard(self):
        QGuiApplication.clipboard().setText(str(self.pdf_path))

    def _send_whatsapp(self):
        phone = self.phone_entry.text().strip()
        clean_phone = re.sub(r'[^\d]', '', phone)
        if not clean_phone:                       # empty, or nothing but punctuation / letters
            Messagebox.show_warning("Missing Info", "Please enter a WhatsApp number.", parent=self)
            return

        self._copy_path_to_clipboard()

        inv_num = self.invoice_data.get('invoice_number', 'Invoice')
        msg = f"Hello! Please find attached {inv_num}. Thank you for your business."
        safe_msg = urllib.parse.quote(msg)
        self._open_link(f"https://wa.me/{clean_phone}?text={safe_msg}")

    def _send_email(self):
        email = self.email_entry.text().strip()
        if not email:
            Messagebox.show_warning("Missing Info", "Please enter an email address.", parent=self)
            return

        self._copy_path_to_clipboard()

        inv_num = self.invoice_data.get('invoice_number', 'Invoice')
        subject = urllib.parse.quote(f"Invoice {inv_num} from {self.invoice_data.get('your_name', 'Us')}")
        body = urllib.parse.quote(f"Hello,\n\nPlease find attached {inv_num}.\n\nThank you!")
        self._open_link(f"mailto:{email}?subject={subject}&body={body}")

    def _open_link(self, url: str):
        if webbrowser.open(url):
            self.accept()                          # Tk: destroy() -- the job is done
        else:
            Messagebox.show_warning(
                "Share", "Could not open your browser / e-mail program.\n"
                "The PDF path is on the clipboard, so you can still paste it into a message.", parent=self)


# =============================================================================
#  14. DIALOGS
# =============================================================================
class _EditDialog(QDialog):
    """Shared chrome for the small modal edit dialogs: a bold header, labelled fields,
    a hairline, and a right-aligned [Save][Cancel] row -- the same arrangement as the
    Tkinter build. Enter saves (the Save button is the dialog's default button) except
    inside a multi-line box, where Enter inserts a new line -- Qt's own behaviour, so the
    Tk build's hand-written <Return> handler is no longer needed. Esc cancels."""

    def __init__(self, parent, window_title: str, header: str, width: int = 440):
        super().__init__(parent)
        self.setWindowTitle(window_title)
        self.setModal(True)
        self.setMinimumWidth(width)
        self._root = QVBoxLayout(self)
        self._root.setContentsMargins(20, 20, 20, 20)
        self._root.setSpacing(0)
        self._root.addWidget(make_label(header, "h2"))
        self._root.addSpacing(14)

    def _add_field(self, label: str, widget: QWidget, space_after: int = 12) -> None:
        self._root.addWidget(make_label(label, "h3"))
        self._root.addSpacing(4)
        self._root.addWidget(widget)
        self._root.addSpacing(space_after)

    @staticmethod
    def _multiline(placeholder: str, rows: int) -> PlaceholderText:
        box = PlaceholderText(placeholder=placeholder, height=rows)
        box.setTabChangesFocus(True)
        box.setAcceptRichText(False)
        return box

    def _add_buttons(self, save_text: str, save_slot: Callable) -> None:
        self._root.addWidget(hline())
        self._root.addSpacing(10)
        row = QHBoxLayout()
        row.addStretch()
        self.save_btn = make_button(save_text, "success", save_slot)
        self.save_btn.setDefault(True)
        self.cancel_btn = make_button("✕ Cancel", "secondary-outline", self.reject)
        self.cancel_btn.setAutoDefault(False)
        row.addWidget(self.save_btn)
        row.addSpacing(8)
        row.addWidget(self.cancel_btn)
        self._root.addLayout(row)


class ClientDialog(_EditDialog):
    """Add or edit a single Address Book entry."""

    def __init__(self, parent, db_manager: "DatabaseManager", client: Optional[Dict], on_saved: Callable):
        super().__init__(parent, "Edit Client" if client else "Add Client",
                         "Edit Client Details" if client else "New Client Details")
        self.db_manager = db_manager
        self.client = client
        self.on_saved = on_saved

        self.name_entry = PlaceholderEntry(placeholder="Enter client name")
        self._add_field("Client / Company Name *", self.name_entry)
        self.contact_text = self._multiline("+91 98765 43210\nname@email.com\nCity, State", 4)
        self._add_field("Contact Info", self.contact_text)
        self.gstin_entry = PlaceholderEntry(placeholder="e.g., 22ABCDE1234F1Z5")
        self._add_field("GSTIN (Optional)", self.gstin_entry, space_after=15)

        if client:
            self.name_entry.setText(client.get("name", ""))
            self.contact_text.setPlainText(client.get("contact", "") or "")
            self.gstin_entry.setText(client.get("gstin", "") or "")

        self._add_buttons("💾 Save", self._save)
        self.name_entry.setFocus()

    def _save(self):
        name = self.name_entry.text().strip()
        if not name:
            return Messagebox.show_error("Validation", "Client name is required.", parent=self)
        contact = self.contact_text.toPlainText().strip()
        gstin = self.gstin_entry.text().strip()
        # IMPROVEMENT over the Tk build: the invoice form refuses to save an invoice whose client
        # GSTIN is malformed, so catch it here instead of at invoicing time.
        if gstin and not is_valid_gstin(gstin):
            return Messagebox.show_error(
                "Invalid GST Number",
                f'"{gstin}" is not a valid GSTIN.\nA GSTIN is 15 characters, for example 22ABCDE1234F1Z5.',
                parent=self)
        cid = self.client.get("id") if self.client else None
        ok, err = self.db_manager.upsert_client_by_id(cid, name, contact, gstin)
        if ok:
            self.on_saved()
            self.accept()
        else:
            Messagebox.show_error("Could Not Save", err, parent=self)


class _UnitRateRow(QWidget):
    """One "unit | rate | x" line of the catalog dialog's "also sold per" list."""
    remove_requested = Signal(object)

    def __init__(self, unit: str = "", rate: float = 0.0, parent: Optional[QWidget] = None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(_ROW_GAP)
        self.unit_combo = make_editable_combo(UNIT_PRESETS, _UNIT_W + 30, "Unit", 14, advance=False)
        self.unit_combo.setCurrentText(unit)
        self.unit_combo.setToolTip("The other unit this item is sold in (pick one or type your own)")
        lay.addWidget(self.unit_combo)
        self.rate_edit = make_money_edit(rate, _AMOUNT_W + 10, "The rate for one of this unit", advance=False)
        lay.addWidget(self.rate_edit)
        lay.addStretch()
        self.remove_btn = make_button("✕", "danger-outline", tooltip="Remove this unit")
        self.remove_btn.setFixedWidth(_REMOVE_W)
        set_flag(self.remove_btn, "glyph", True)
        self.remove_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.remove_btn.clicked.connect(self._on_remove)
        lay.addWidget(self.remove_btn)

    def _on_remove(self, _checked: bool = False):
        self.remove_requested.emit(self)

    def values(self) -> Tuple[str, float]:
        return self.unit_combo.currentText().strip(), safe_float(self.rate_edit.text())


class _UnitRatesEditor(QWidget):
    """The "also sold per ..." list: the other units one catalog item is sold in, each with its own rate
    ("Tent": Nos 500 -- Day 150 -- Week 800)."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.rows: List[_UnitRateRow] = []
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)
        self._rows_layout = QVBoxLayout()
        self._rows_layout.setSpacing(6)
        lay.addLayout(self._rows_layout)
        bar = QHBoxLayout()
        self.add_btn = make_button("＋  Add another unit", "secondary-outline", self._add_clicked,
                                   tooltip="Sell this item in one more unit, with its own rate")
        bar.addWidget(self.add_btn)
        bar.addStretch()
        lay.addLayout(bar)

    def _add_clicked(self):
        self.add_row().unit_combo.setFocus()

    def add_row(self, unit: str = "", rate: float = 0.0) -> _UnitRateRow:
        row = _UnitRateRow(unit, rate)
        row.remove_requested.connect(self.remove_row)
        self._rows_layout.addWidget(row)
        self.rows.append(row)
        return row

    def remove_row(self, row: _UnitRateRow):
        if row in self.rows:
            self.rows.remove(row)
            self._rows_layout.removeWidget(row)
            row.setParent(None)
            row.deleteLater()

    def set_units(self, units: List[Tuple[str, float]]):
        for row in list(self.rows):
            self.remove_row(row)
        for unit, rate in units:
            self.add_row(unit, rate)

    def units(self) -> List[Tuple[str, float]]:
        """The (unit, rate) pairs that have a unit name -- a row left blank is ignored."""
        return [vals for vals in (row.values() for row in self.rows) if vals[0]]


class CatalogItemDialog(_EditDialog):
    """Add or edit a single Product & Service Catalog entry: its name, description, unit and rate, the HSN/SAC code and
    GST rate it is invoiced with, and any other units it is sold in (each with its own rate)."""

    def __init__(self, parent, db_manager: "DatabaseManager", item: Optional[Dict], on_saved: Callable):
        super().__init__(parent, "Edit Catalog Item" if item else "Add Catalog Item",
                         "Edit Catalog Item" if item else "New Catalog Item", width=540)
        self.db_manager = db_manager
        self.item = item
        self.on_saved = on_saved

        self.title_entry = PlaceholderEntry(placeholder="Enter item / service name")
        self._add_field("Item / Service Title *", self.title_entry)
        self.desc_text = self._multiline("Description (optional)", 3)
        self._add_field("Description", self.desc_text)

        self.unit_combo = make_editable_combo(UNIT_PRESETS, 190, "e.g. Nos, Day, Plate", 14, advance=False)
        self.unit_combo.setToolTip("The unit the default rate below is for (pick one or type your own)")
        self.amount_entry = QLineEdit()
        # UX polish: strict numeric controls (same as the Tk build)
        restrict_to_decimal(self.amount_entry)
        format_as_decimal_on_blur(self.amount_entry)
        self._add_pair(("Unit", self.unit_combo), (f"Default Rate ({CURRENCY_SYMBOL}) per unit", self.amount_entry))

        self.hsn_entry = QLineEdit()
        self.hsn_entry.setPlaceholderText("optional, 2 to 8 digits")
        restrict_to_digits(self.hsn_entry, 8)
        self.gst_combo = make_editable_combo(GST_RATE_PRESETS, 190, "not set", 6, advance=False)
        self.gst_combo.lineEdit().setValidator(_DecimalValidator(2, self.gst_combo))
        self.gst_combo.setToolTip("The GST rate this item is usually charged at (left empty: the invoice's own rate is used)")
        self._add_pair(("HSN / SAC Code", self.hsn_entry), ("GST Rate (%)", self.gst_combo))

        self._root.addWidget(make_label("Also Sold Per", "h3"))
        self._root.addSpacing(2)
        self._root.addWidget(make_label("The same item in another unit, with its own rate. On an invoice, choosing that "
                                        "unit fills in its rate.", "muted", wrap=True))
        self._root.addSpacing(6)
        self.units_editor = _UnitRatesEditor()
        self._root.addWidget(self.units_editor)
        self._root.addSpacing(15)

        if item:
            self.title_entry.setText(item.get("title", ""))
            self.desc_text.setPlainText(item.get("description", "") or "")
            self.amount_entry.setText(f"{float(item.get('amount', 0) or 0):.2f}")
            self.unit_combo.setCurrentText(str(item.get("unit") or ""))
            self.hsn_entry.setText(str(item.get("hsn") or ""))
            self.gst_combo.setCurrentText("" if item.get("tax_rate") is None else _num_text(item.get("tax_rate")))
            units = catalog_units(item)
            self.units_editor.set_units(units[1:] if str(item.get("unit") or "").strip() else units)
        else:
            self.amount_entry.setText("0.00")
            self.gst_combo.setCurrentText("")

        self._add_buttons("💾 Save", self._save)
        self.title_entry.setFocus()

    def _add_pair(self, left: Tuple[str, QWidget], right: Tuple[str, QWidget], space_after: int = 12) -> None:
        """Two labelled fields side by side."""
        row = QHBoxLayout()
        row.setSpacing(14)
        for label, widget in (left, right):
            col = QVBoxLayout()
            col.setSpacing(4)
            col.addWidget(make_label(label, "h3"))
            col.addWidget(widget)
            row.addLayout(col, 1)
        self._root.addLayout(row)
        self._root.addSpacing(space_after)

    def _save(self):
        title = self.title_entry.text().strip()
        if not title:
            return Messagebox.show_error("Validation", "Item title is required.", parent=self)
        try:
            amount = float(self.amount_entry.text() or 0)
        except ValueError:
            return Messagebox.show_error("Validation", "Amount must be a valid number.", parent=self)
        desc = self.desc_text.toPlainText().strip()
        unit = self.unit_combo.currentText().strip()
        hsn = self.hsn_entry.text().strip()
        if hsn and not 2 <= len(hsn) <= 8:
            return Messagebox.show_error("Validation", "The HSN / SAC code must be 2 to 8 digits.", parent=self)
        gst_text = self.gst_combo.currentText().strip()
        tax_rate = safe_float(gst_text) if gst_text else None
        if tax_rate is not None and not 0 <= tax_rate <= 100:
            return Messagebox.show_error("Validation", "The GST rate must be between 0 and 100.", parent=self)
        others = self.units_editor.units()
        if others and not unit:
            return Messagebox.show_error("Validation", "Set the item's main unit first: the other units are alternatives to it.",
                                         parent=self)
        seen = {unit.casefold()} if unit else set()
        for name, _rate in others:
            if name.casefold() in seen:
                return Messagebox.show_error("Validation", f'The unit "{name}" is listed more than once.', parent=self)
            seen.add(name.casefold())
        iid = self.item.get("id") if self.item else None
        ok, err = self.db_manager.upsert_catalog_item_by_id(iid, title, desc, amount, unit=unit, hsn=hsn,
                                                            tax_rate=tax_rate, other_units=others)
        if ok:
            self.on_saved()
            self.accept()
        else:
            Messagebox.show_error("Could Not Save", err, parent=self)


class NoteDialog(_EditDialog):
    """Quick popup to add/edit an invoice's internal note straight from the
    invoice list, without opening the full invoice form."""

    def __init__(self, parent, db_manager: "DatabaseManager", invoice_number: str,
                 current_note: str, on_saved: Callable):
        super().__init__(parent, f"Note  —  Invoice #{invoice_number}", f"Invoice #{invoice_number}")
        self.db_manager = db_manager
        self.invoice_number = invoice_number
        self.on_saved = on_saved
        self.setMinimumHeight(320)

        # The header above is followed by a muted explanation, exactly as in the Tk build.
        self._root.insertWidget(1, make_label("Internal note (not printed on the customer's PDF)", "muted"))
        self.text = PlaceholderText(placeholder="Add an internal note...")
        self.text.setTabChangesFocus(True)
        self.text.setAcceptRichText(False)
        self.text.setPlainText(current_note or "")
        self._root.addWidget(self.text, 1)
        self._root.addSpacing(15)

        self._add_buttons("💾 Save Note", self._save)
        self.save_btn.setDefault(False)       # Enter inserts a new line in the note, it does not save
        self.save_btn.setAutoDefault(False)
        self.text.setFocus()

    def _save(self):
        note = self.text.toPlainText().strip()
        self.db_manager.update_notes(self.invoice_number, note)
        self.on_saved()
        self.accept()


class BulkProgressDialog(QDialog):
    """Modal progress window for long bulk jobs (delete / export / print / PDF generation).

    `work_fn(items, report, is_aborted) -> dict` runs on a background thread and must
    never touch a widget; it calls `report(current, total, status_text, log_line=None,
    success_delta=0, failed_delta=0)` to post progress. A QTimer on the GUI thread drains
    the queue every 100 ms -- the only code here allowed to touch widgets -- and calls
    `on_done(result)` exactly once when the job ends. Use as `dlg.exec()`.

    Window X / Esc while running == Cancel (the job stops after the current item); once the
    job has finished they simply close the window, as in the Tkinter build."""

    def __init__(self, parent, title: str, items: list, work_fn: Callable,
                 on_done: Optional[Callable[[dict], None]] = None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        _fit_to_screen(self, 580, 480, 520, 440)

        self._items = items
        self._work_fn = work_fn
        self._on_done = on_done
        self._total = len(items)
        self._abort_event = threading.Event()
        self._q: queue.Queue = queue.Queue()
        self._start_time = time.monotonic()
        self._done = False
        self._success_count = 0
        self._failed_count = 0
        self.job_result: Optional[dict] = None     # NOT `result`: that would shadow QDialog.result()

        self._build(title)
        self._thread = threading.Thread(target=self._run_worker, daemon=True)
        self._thread.start()
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._poll)
        self._timer.start(100)

    def _build(self, title: str):
        main = QVBoxLayout(self)
        main.setContentsMargins(18, 18, 18, 18)

        main.addWidget(make_label(title, "h2"))
        self.status_lbl = make_label("Starting…", "muted")
        main.addSpacing(4)
        main.addWidget(self.status_lbl)
        main.addSpacing(12)

        bar_row = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setProperty("tall", True)
        self.progress.setRange(0, max(1, self._total))
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        self.pct_lbl = make_label("0%", "h3", align=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.pct_lbl.setFixedWidth(52)
        bar_row.addWidget(self.progress, 1)
        bar_row.addSpacing(10)
        bar_row.addWidget(self.pct_lbl)
        main.addLayout(bar_row)

        main.addSpacing(10)
        stats = QHBoxLayout()
        self.processed_lbl = QLabel(f"Processed: 0 / {self._total}")
        self.eta_lbl = make_label("Estimating…", "muted")
        stats.addWidget(self.processed_lbl)
        stats.addStretch()
        stats.addWidget(self.eta_lbl)
        main.addLayout(stats)

        main.addSpacing(4)
        counts = QHBoxLayout()
        self.success_lbl = make_label("✔  Success: 0", "h3", "success")
        self.failed_lbl = make_label("✖  Failed: 0", "h3", "danger")
        counts.addWidget(self.success_lbl)
        counts.addSpacing(18)
        counts.addWidget(self.failed_lbl)
        counts.addStretch()
        main.addLayout(counts)

        main.addSpacing(12)
        main.addWidget(make_label("Details:", "h3"))
        main.addSpacing(2)
        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        self.log_text.setFont(QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont))
        main.addWidget(self.log_text, 1)

        main.addSpacing(14)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.abort_btn = make_button("✕  Cancel", "danger-outline", self._on_abort_click)
        self.close_btn = make_button("Close", "secondary", lambda: QDialog.accept(self))
        self.close_btn.hide()                 # only shown once the operation finishes -- see _on_finished()
        btn_row.addWidget(self.abort_btn)
        btn_row.addWidget(self.close_btn)
        main.addLayout(btn_row)

    def _log(self, line: str):
        self.log_text.appendPlainText(line)
        bar = self.log_text.verticalScrollBar()
        bar.setValue(bar.maximum())

    # ---- window-close semantics -------------------------------------------- #
    def reject(self):                         # Esc
        self._on_abort_click()

    def closeEvent(self, event):              # window X
        if self._done:
            event.accept()
        else:
            event.ignore()
            self._on_abort_click()

    def _on_abort_click(self):
        if self._done:
            QDialog.reject(self)
            return
        self._abort_event.set()
        self.abort_btn.setEnabled(False)
        self.abort_btn.setText("Cancelling…")
        self.status_lbl.setText("Cancelling -- finishing the current item…")

    # ------------------------------------------------------------------ #
    #  Worker thread side -- touches only self._q, never a widget
    # ------------------------------------------------------------------ #
    def _run_worker(self):
        def report(current, total, status_text, log_line=None, success_delta=0, failed_delta=0):
            self._q.put({"type": "progress", "current": current, "total": total,
                         "status_text": status_text, "log_line": log_line,
                         "success_delta": success_delta, "failed_delta": failed_delta})
        try:
            result = self._work_fn(self._items, report, self._abort_event.is_set)
        except Exception as exc:
            logger.error(f"Bulk operation failed: {exc}", exc_info=True)
            result = {"success": self._success_count, "failed": max(0, self._total - self._success_count),
                      "aborted": False, "result": None, "error": str(exc)}
        self._q.put({"type": "done", "result": result})

    # ------------------------------------------------------------------ #
    #  GUI-thread side
    # ------------------------------------------------------------------ #
    def _poll(self):
        try:
            while True:
                msg = self._q.get_nowait()
                if msg["type"] == "progress":
                    self._on_progress(msg)
                elif msg["type"] == "done":
                    self._on_finished(msg["result"])
        except queue.Empty:
            pass

    def _on_progress(self, msg: dict):
        current, total = msg["current"], msg["total"]
        self.progress.setRange(0, max(1, total))
        self.progress.setValue(min(current, max(1, total)))
        pct = int(current / total * 100) if total else 0
        self.pct_lbl.setText(f"{pct}%")
        self.processed_lbl.setText(f"Processed: {current} / {total}")
        self.status_lbl.setText(msg["status_text"])

        if msg.get("success_delta"):
            self._success_count += msg["success_delta"]
            self.success_lbl.setText(f"✔  Success: {self._success_count}")
        if msg.get("failed_delta"):
            self._failed_count += msg["failed_delta"]
            self.failed_lbl.setText(f"✖  Failed: {self._failed_count}")
        if msg.get("log_line"):
            self._log(msg["log_line"])

        elapsed = time.monotonic() - self._start_time
        if current > 0 and elapsed > 0.4:
            rate = current / elapsed
            remaining = max(0, total - current)
            eta_seconds = (remaining / rate) if rate > 0 else 0
            self.eta_lbl.setText(f"~{self._format_duration(eta_seconds)} remaining")
        else:
            self.eta_lbl.setText("Estimating…")

    @staticmethod
    def _format_duration(seconds: float) -> str:
        seconds = max(0, int(seconds))
        if seconds < 60:
            return f"{seconds}s"
        minutes, seconds = divmod(seconds, 60)
        if minutes < 60:
            return f"{minutes}m {seconds}s"
        hours, minutes = divmod(minutes, 60)
        return f"{hours}h {minutes}m"

    def _on_finished(self, result: dict):
        if self._done:
            return
        self._done = True
        self._timer.stop()
        self.job_result = result
        aborted = result.get("aborted", False)
        success = result.get("success", self._success_count)
        failed = result.get("failed", self._failed_count)
        error = result.get("error")

        if error:
            self.status_lbl.setText(f"Stopped due to an error: {error}")
            self._log(f"ERROR: {error}")
        elif aborted:
            self.status_lbl.setText(f"Cancelled -- {success} completed, {failed} failed before stopping.")
        else:
            self.status_lbl.setText(f"Done -- {success} succeeded, {failed} failed.")
            self.progress.setValue(self.progress.maximum())
            self.pct_lbl.setText("100%")
        self.eta_lbl.setText("")

        self.abort_btn.hide()
        self.close_btn.show()
        self.close_btn.setFocus()

        if self._on_done:
            try:
                self._on_done(result)
            except Exception as exc:
                logger.error(f"Bulk operation on_done callback failed: {exc}", exc_info=True)


class ChangelogDialog(QDialog):
    """"What's new": every release since the file this project started from, with all of its changes."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"What's New in {APP_NAME}")
        self.setModal(True)
        _fit_to_screen(self, 560, 640, 420, 420)
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(12)
        root.addWidget(make_label(f"{CHANGE_COUNT} changes since {BASE_VERSION} (build {BASE_BUILD_VERSION})", "h2"))
        self.summary_lbl = make_label(f"You are running {APP_VERSION}, build {APP_BUILD_VERSION}.", "muted")
        root.addWidget(self.summary_lbl)
        area = QScrollArea()
        area.setWidgetResizable(True)
        area.setFrameShape(QFrame.Shape.NoFrame)
        area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        body = QWidget()
        col = QVBoxLayout(body)
        col.setContentsMargins(0, 0, 10, 0)
        col.setSpacing(4)
        self.release_labels: List[QLabel] = []
        self.change_labels: List[QLabel] = []
        for release in CHANGELOG:
            head = make_label(f"{release.version}  —  {release.title}", "h3")
            self.release_labels.append(head)
            col.addSpacing(8)
            col.addWidget(head)
            col.addWidget(make_label(f"{release.date}  ·  {len(release.changes)} changes", "muted"))
            col.addSpacing(4)
            for change in release.changes:
                lbl = make_label(f"•  {change}", wrap=True)
                lbl.setMinimumWidth(120)
                self.change_labels.append(lbl)
                col.addWidget(lbl)
        col.addStretch(1)
        area.setWidget(body)
        root.addWidget(area, 1)
        row = QHBoxLayout()
        row.addStretch(1)
        self.close_btn = make_button("Close", "primary", self.accept)
        self.close_btn.setFixedWidth(130)
        self.close_btn.setDefault(True)
        row.addWidget(self.close_btn)
        root.addLayout(row)


class AboutDialog(QDialog):
    WIDTH = 400

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"About {APP_NAME}")
        self.setModal(True)
        self.setFixedWidth(self.WIDTH)
        self._build()
        self.setFixedHeight(max(420, self.sizeHint().height()))

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(15, 15, 15, 15)
        root.setSpacing(0)

        hero = QFrame()
        hero.setProperty("role", "hero")
        hero_col = QVBoxLayout(hero)
        hero_col.setContentsMargins(15, 15, 15, 15)
        hero_col.setSpacing(4)
        hero_col.addWidget(make_label(APP_NAME, "hero-title", align=Qt.AlignmentFlag.AlignCenter))
        hero_col.addWidget(make_label(APP_TAGLINE, "hero-sub", align=Qt.AlignmentFlag.AlignCenter))
        root.addWidget(hero)

        root.addSpacing(20)
        self.version_label = make_label(APP_VERSION, "display", "info", align=Qt.AlignmentFlag.AlignCenter)
        root.addWidget(self.version_label)
        root.addSpacing(5)
        self.build_label = make_label(f"Build {APP_BUILD_VERSION}", "muted", align=Qt.AlignmentFlag.AlignCenter)
        root.addWidget(self.build_label)
        root.addWidget(make_label(f"Released on {APP_RELEASE_DATE}", "muted", align=Qt.AlignmentFlag.AlignCenter))
        root.addSpacing(6)
        self.headline_label = make_label(CHANGELOG[0].title, "small", align=Qt.AlignmentFlag.AlignCenter, wrap=True)
        root.addWidget(self.headline_label)
        self.changes_label = make_label(f"{CHANGE_COUNT} changes since {BASE_VERSION} (build {BASE_BUILD_VERSION})",
                                        "small", align=Qt.AlignmentFlag.AlignCenter)
        root.addWidget(self.changes_label)
        root.addSpacing(15)
        root.addWidget(hline())
        root.addSpacing(15)

        lines = APP_ATTRIBUTION.split("\n")
        root.addWidget(make_label(lines[0], "h3", align=Qt.AlignmentFlag.AlignCenter))
        for line in lines[1:]:
            root.addSpacing(2)
            root.addWidget(make_label(line, "muted", align=Qt.AlignmentFlag.AlignCenter))

        root.addStretch(1)
        root.addSpacing(20)
        row = QHBoxLayout()
        row.addStretch(1)
        self.whats_new_btn = make_button("What's New", "secondary-outline", self._show_changelog)
        self.whats_new_btn.setFixedWidth(130)
        row.addWidget(self.whats_new_btn)
        row.addSpacing(8)
        self.close_btn = make_button("Close", "primary", self.accept)
        self.close_btn.setFixedWidth(130)
        self.close_btn.setDefault(True)
        row.addWidget(self.close_btn)
        row.addStretch(1)
        root.addLayout(row)

    def _show_changelog(self):
        ChangelogDialog(self).exec()


class ShortcutsDialog(QDialog):
    def __init__(self, parent, groups: Optional[Dict[str, List[Tuple[str, str]]]] = None):
        super().__init__(parent)
        self.setWindowTitle("Keyboard Shortcuts")
        self.setModal(True)
        _fit_to_screen(self, 520, 640, 460, 500)
        if groups is None:
            groups = parent.shortcuts.grouped_reference()
        self._build(groups)

    def _build(self, groups: Dict[str, List[Tuple[str, str]]]):
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(0)

        hero = QFrame()
        hero.setProperty("role", "hero")
        hero_col = QVBoxLayout(hero)
        hero_col.setContentsMargins(20, 20, 20, 20)
        hero_col.setSpacing(5)
        hero_col.addWidget(make_label("⌨  Keyboard Shortcuts", "hero-title", align=Qt.AlignmentFlag.AlignCenter))
        hero_col.addWidget(make_label("Master Vyaprix without touching your mouse.", "hero-sub",
                                      align=Qt.AlignmentFlag.AlignCenter))
        root.addWidget(hero)
        root.addSpacing(10)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        body = QWidget()
        col = QVBoxLayout(body)
        col.setContentsMargins(5, 0, 12, 0)
        col.setSpacing(0)
        self.key_labels: List[QLabel] = []
        for group_title, items in groups.items():
            head = QHBoxLayout()
            head.setSpacing(10)
            head.addWidget(make_label(group_title.upper(), "h3", "info"))
            head.addWidget(hline(), 1, Qt.AlignmentFlag.AlignVCenter)
            col.addSpacing(20)
            col.addLayout(head)
            col.addSpacing(10)
            grid = QGridLayout()
            grid.setContentsMargins(0, 0, 0, 0)
            grid.setHorizontalSpacing(15)
            grid.setVerticalSpacing(10)
            grid.setColumnMinimumWidth(0, 150)
            grid.setColumnStretch(1, 1)
            for row, (keys, desc) in enumerate(items):
                key_lbl = make_label(keys, "key")
                self.key_labels.append(key_lbl)
                grid.addWidget(key_lbl, row, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                desc_lbl = make_label(desc, wrap=True)
                grid.addWidget(desc_lbl, row, 1, Qt.AlignmentFlag.AlignVCenter)
            col.addLayout(grid)
        col.addStretch(1)
        self.scroll.setWidget(body)
        root.addWidget(self.scroll, 1)
        root.addSpacing(15)

        row = QHBoxLayout()
        row.addStretch(1)
        self.close_btn = make_button("Got It", "primary", self.accept)
        self.close_btn.setFixedWidth(130)
        self.close_btn.setDefault(True)
        row.addWidget(self.close_btn)
        row.addStretch(1)
        root.addLayout(row)


def _logo_pixmap_for(widget: QWidget, max_width: int) -> Optional[QPixmap]:
    """The bundled logo, at most `max_width` logical pixels wide, rendered at the screen's native resolution so it stays
    crisp on 125% / 150% / 200% displays. A logo that is already smaller than that is shown 1:1, never enlarged."""
    dpr = widget.devicePixelRatioF() or 1.0
    wanted = int(round(max_width * dpr))
    pixmap = load_app_logo_pixmap(wanted)
    if pixmap is not None and dpr > 1.01 and pixmap.width() == wanted:
        pixmap.setDevicePixelRatio(dpr)
    return pixmap


def pixel_ratio(widget: Optional[QWidget] = None) -> int:
    """The scale to paint a pixmap at for the screen `widget` is on: its device pixel ratio rounded UP, and never below
    2 -- so an image stays sharp at 125 %, 150 %, 200 % or 300 % (Qt scales it down by the exact fraction)."""
    screen = (widget.screen() if widget is not None else None) or QGuiApplication.primaryScreen()
    ratio = screen.devicePixelRatio() if screen is not None else 1.0
    whole = int(ratio)
    return max(2, whole + (1 if ratio - whole > 1e-6 else 0))


class _AssetRow(QWidget):
    """One upload row (Company Logo / Authorized Signature / UPI QR Code): a preview thumbnail, a status line and
    Browse / Clear buttons. The chosen file is copied into Vyaprix's own data folder straight away, so the app keeps
    working even if the original is moved or deleted; the file name is deliberately not shown."""

    THUMB = 64

    def __init__(self, config_key: str, label: str, dest_filename: str):
        super().__init__()
        self.config_key = config_key
        self.dest_filename = dest_filename
        self.rel_path = ""

        row = QHBoxLayout(self)
        row.setContentsMargins(0, 8, 0, 8)
        row.setSpacing(12)

        self.thumb = make_label("—", "thumb", align=Qt.AlignmentFlag.AlignCenter)
        self.thumb.setFixedSize(self.THUMB, self.THUMB)
        row.addWidget(self.thumb)

        info = QVBoxLayout()
        info.setSpacing(2)
        info.addStretch(1)
        info.addWidget(make_label(label, "h3"))
        self.status = make_label("", "muted")
        info.addWidget(self.status)
        info.addStretch(1)
        row.addLayout(info, 1)

        self.browse_btn = make_button("Browse…", "secondary-outline", self.browse)
        self.clear_btn = make_button("✕ Clear", "danger-outline", self.clear)
        _no_default_button(self.browse_btn, self.clear_btn)
        row.addWidget(self.browse_btn)
        row.addWidget(self.clear_btn)

    def set_path(self, rel_path: str) -> None:
        self.rel_path = rel_path or ""
        self.refresh()

    def refresh(self) -> None:
        dpr = self.devicePixelRatioF() or 1.0
        pixmap = None
        if self.rel_path:
            pixmap = load_thumbnail_pixmap(DATA_DIR / self.rel_path, max_size=int(round(44 * dpr)))
        if pixmap is not None:
            pixmap.setDevicePixelRatio(dpr)
            self.thumb.setPixmap(pixmap)
            self.status.setText("Custom file uploaded")
        else:
            self.thumb.clear()
            self.thumb.setText("—")
            self.status.setText("No file selected — using default")

    def browse(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self.window(), "Choose an Image", str(_start_dir("asset_dir")),
                                              "Image Files (*.png *.jpg *.jpeg)")
        if not path:
            return
        _remember_dir("asset_dir", path)
        try:
            dest = ASSETS_DIR / self.dest_filename
            if not (dest.exists() and _same_file(Path(path), dest)):
                shutil.copy(path, dest)
        except Exception as exc:
            Messagebox.show_error("Error", f"Could not save file: {exc}", parent=self.window())
            return
        self.set_path(f"assets/{self.dest_filename}")

    def clear(self) -> None:
        self.set_path("")


def template_thumbnail(tpl: PDFTemplate, width: int = 84, ratio: Optional[int] = None) -> QPixmap:
    """A miniature of an invoice in `tpl`, painted from the same colours the PDF is printed in -- so a brand colour
    shows in it at once. The page is drawn in millimetres (210 x 297), at `ratio` pixels per logical pixel (the
    screen's, by default)."""
    ratio = ratio or pixel_ratio()
    height = round(width * 297 / 210)
    pm = QPixmap(width * ratio, height * ratio)
    pm.setDevicePixelRatio(ratio)
    pm.fill(QColor("#FFFFFF"))
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.scale(width / 210.0, width / 210.0)

    def block(x: float, y: float, w: float, h: float, color: str):
        if color:
            p.fillRect(QRectF(x, y, w, h), QColor(color))

    def rule(x1: float, y1: float, x2: float, y2: float, color: str, points: float = 1.0):
        if color:
            pen = QPen(QColor(color))
            pen.setWidthF(max(points * 0.353, 1.1))          # never thinner than a hair at this size
            p.setPen(pen)
            p.drawLine(QPointF(x1, y1), QPointF(x2, y2))

    ink, soft = "#374151", "#9CA3AF"
    block(0, 0, 210, 6, tpl.bar)
    if tpl.header == "band":
        block(15, 14, 180, 30, tpl.band_bg)
        block(21, 22, 52, 7, tpl.name_color)
        block(21, 32, 34, 3, tpl.tagline_color)
        block(150, 24, 39, 8, tpl.title_color)
        y = 52
    else:
        block(15, 16, 50, 7, tpl.name_color)
        block(15, 26, 32, 3, tpl.tagline_color)
        scale = tpl.title_size / 18.0
        block(195 - 40 * scale, 17, 40 * scale, 8 * scale, tpl.title_color)
        rule(15, 38, 195, 38, tpl.rule, tpl.rule_width)
        y = 46
    # BILL TO and FROM, side by side
    for x in (15, 111):
        block(x, y, 14, 2.6, tpl.label_color)
        block(x, y + 5, 38, 3, ink)
        block(x, y + 10, 46, 2.6, soft)
        block(x, y + 14.5, 34, 2.6, soft)
    # the reference strip: four labelled cells
    y += 25
    block(15, y, 180, 14, tpl.meta_bg)
    if tpl.meta_border:
        p.setPen(QPen(QColor(tpl.meta_border), 1.1))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRect(QRectF(15, y, 180, 14))
    rule(15, y, 195, y, tpl.meta_lines, 0.6)
    rule(15, y + 14, 195, y + 14, tpl.meta_lines, 0.6)
    for k in range(4):
        block(19 + k * 44, y + 3, 14, 1.8, tpl.label_color)
        block(19 + k * 44, y + 7, 26, 2.8, ink)
    # item table: # | item | HSN | qty | rate | amount
    y += 21
    head_h, row_h, rows = 9, 12, 3
    numeric = ((101, 11), (121, 11), (141, 12), (173, 16))
    block(15, y, 180, head_h, tpl.head_bg)
    head_ink = tpl.head_fg
    block(18, y + 3.2, 4, 2.6, head_ink)
    block(28, y + 3.2, 28, 2.6, head_ink)
    for x, w in numeric:
        block(x, y + 3.2, w - 2, 2.6, head_ink)
    rule(15, y + head_h, 195, y + head_h, tpl.head_rule, 1.2)
    for r in range(rows):
        top = y + head_h + r * row_h
        if tpl.zebra and r % 2 == 1:
            block(15, top, 180, row_h, tpl.zebra)
        block(18, top + 4, 4, 2.6, ink)
        block(28, top + 3, 40, 3, ink)
        block(28, top + 7.6, 54, 2.2, soft)
        for x, w in numeric:
            block(x, top + 4, w - 2, 2.6, ink)
        if r < rows - 1:
            rule(15, top + row_h, 195, top + row_h, tpl.row_line, 1.0)
    bottom = y + head_h + rows * row_h
    rule(15, bottom, 195, bottom, tpl.end_line, tpl.end_width)
    if tpl.grid:
        p.setPen(QPen(QColor(tpl.grid_line), 1.0))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRect(QRectF(15, y, 180, bottom - y))
        for x in (24, 98, 118, 138, 168):
            p.drawLine(QPointF(x, y), QPointF(x, bottom))
        for r in range(1, rows):
            p.drawLine(QPointF(15, y + head_h + r * row_h), QPointF(195, y + head_h + r * row_h))
    # summary: the amount in words and the GST summary on the left, the totals on the right
    y = bottom + 7
    block(15, y, 62, 2.4, soft)
    block(15, y + 4.4, 84, 2.4, ink)
    block(15, y + 12, 86, 5, tpl.head_bg or tpl.pay_head_bg)
    for k in range(2):
        block(18, y + 20 + k * 5, 14, 2.2, soft)
        block(48, y + 20 + k * 5, 18, 2.2, ink)
        block(76, y + 20 + k * 5, 18, 2.2, ink)
    left = 108
    for k in range(3):
        block(left + 38, y + k * 6.4 + 1.6, 26, 2.4, soft)
        block(172, y + k * 6.4 + 1.6, 23, 2.4, ink)
    ty = y + 3 * 6.4
    block(left, ty, 195 - left, 7, tpl.total_bg)
    rule(left, ty, 195, ty, tpl.total_rule, 0.8)
    total_ink = tpl.total_fg if tpl.total_bg else ink
    block(left + 38, ty + 2.2, 26, 2.8, total_ink)
    block(172, ty + 2.2, 23, 2.8, total_ink)
    by = ty + 7
    block(left, by, 195 - left, 7, "#FEF2F2")
    block(left + 38, by + 2.2, 26, 2.8, ink)
    block(172, by + 2.2, 23, 2.8, "#DC2626")
    # the page footer: a hairline, the invoice number, the note and "Page 1 of 1"
    rule(15, 281, 195, 281, "#E5E7EB", 0.6)
    block(15, 285, 18, 2.2, soft)
    block(80, 285, 50, 2.2, soft)
    block(177, 285, 18, 2.2, soft)
    p.setPen(QPen(QColor("#CBD5E1"), 0.9))
    p.setBrush(Qt.BrushStyle.NoBrush)
    p.drawRect(QRectF(0.4, 0.4, 209.2, 296.2))
    p.end()
    return pm


# The brand colours offered in Settings (any other colour can be picked with "Custom colour...").
BRAND_COLORS: List[Tuple[str, str]] = [
    ("#2563EB", "Blue"), ("#4F46E5", "Indigo"), ("#7C3AED", "Purple"), ("#DB2777", "Pink"), ("#DC2626", "Red"),
    ("#EA580C", "Orange"), ("#16A34A", "Green"), ("#0D9488", "Teal"), ("#1E3A5F", "Navy"), ("#111827", "Black"),
]


def _swatch_icon(color: str, selected: bool, ratio: Optional[int] = None) -> QIcon:
    """A round colour chip; a selected one has a ring around it."""
    ratio, size = ratio or pixel_ratio(), 24
    pm = QPixmap(size * ratio, size * ratio)
    pm.setDevicePixelRatio(ratio)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    centre = QPointF(size / 2, size / 2)
    if selected:
        p.setPen(QPen(QColor("#374151"), 2))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawEllipse(centre, 10.5, 10.5)
        radius = 7.5
    else:
        radius = 9.0
    p.setPen(QPen(QColor(0, 0, 0, 50), 1))
    p.setBrush(QColor(color))
    p.drawEllipse(centre, radius, radius)
    p.end()
    return QIcon(pm)


class _Swatch(QToolButton):
    """A colour chip in the brand-colour row; clicking it announces its colour."""
    picked = Signal(str)

    def __init__(self, color: str, name: str, parent=None):
        super().__init__(parent)
        self.color = color
        self.setCheckable(True)
        self.setAutoRaise(True)
        self.setToolTip(f"{name}  ({color})")
        self.setIconSize(QSize(24, 24))
        self.setFixedSize(30, 30)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setIcon(_swatch_icon(color, False, pixel_ratio(self)))
        self.clicked.connect(self._on_clicked)

    def _on_clicked(self, _checked: bool = False):
        self.picked.emit(self.color)

    def set_selected(self, selected: bool):
        self.setChecked(selected)
        self.setIcon(_swatch_icon(self.color, selected, pixel_ratio(self)))


class _TemplateCard(QFrame):
    """One template in Settings: a miniature of the invoice, its name (a radio button) and a line about it. Clicking
    anywhere on the card picks it."""
    picked = Signal(str)

    def __init__(self, tpl: PDFTemplate, group: QButtonGroup, parent=None):
        super().__init__(parent)
        self.key = tpl.key
        set_role(self, "tpl")
        set_flag(self, "selected", False)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(12)
        self.thumb = QLabel()
        self.thumb.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        self.thumb.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        lay.addWidget(self.thumb, 0, Qt.AlignmentFlag.AlignTop)
        text = QVBoxLayout()
        text.setSpacing(4)
        self.radio = QRadioButton(tpl.label)
        self.radio.setProperty("role", "h3")
        group.addButton(self.radio)
        self.radio.toggled.connect(self._on_toggled)
        text.addWidget(self.radio)
        self.blurb = make_label(tpl.blurb, "muted", wrap=True)
        self.blurb.setMinimumWidth(90)                  # a wrapped label otherwise asks for far more room than it needs
        text.addWidget(self.blurb)
        text.addStretch(1)
        lay.addLayout(text, 1)
        self.refresh(tpl)

    def refresh(self, tpl: PDFTemplate):
        """Redraw the miniature (the template may have taken another brand colour)."""
        pm = template_thumbnail(tpl, ratio=pixel_ratio(self))
        self.thumb.setPixmap(pm)
        self.thumb.setFixedSize(pm.size() / pm.devicePixelRatio())

    def _on_toggled(self, checked: bool):
        set_flag(self, "selected", checked)
        if checked:
            self.picked.emit(self.key)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.radio.click()
        super().mousePressEvent(event)


class SettingsDialog(QDialog):
    CURRENCY_OPTIONS: List[Tuple[str, str, str]] = [
        ("₹",   "INR", "Indian Rupee"),
        ("$",   "USD", "US Dollar"),
        ("€",   "EUR", "Euro"),
        ("£",   "GBP", "British Pound"),
        ("¥",   "JPY", "Japanese Yen"),
        ("A$",  "AUD", "Australian Dollar"),
        ("C$",  "CAD", "Canadian Dollar"),
        ("S$",  "SGD", "Singapore Dollar"),
        ("CN¥", "CNY", "Chinese Yuan"),
        ("Fr",  "CHF", "Swiss Franc"),
        ("د.إ", "AED", "UAE Dirham"),
        ("R$",  "BRL", "Brazilian Real"),
        ("₽",   "RUB", "Russian Ruble"),
        ("₩",   "KRW", "South Korean Won"),
        ("฿",   "THB", "Thai Baht"),
        ("R",   "ZAR", "South African Rand"),
        ("₺",   "TRY", "Turkish Lira"),
        ("kr",  "SEK", "Swedish Krona"),
        ("zł",  "PLN", "Polish Złoty"),
        ("₫",   "VND", "Vietnamese Dong"),
    ]
    CUSTOM_CURRENCY = "Custom Symbol…"
    TAB_GENERAL, TAB_DETAILS, TAB_PAYMENT, TAB_ASSETS, TAB_DATA = range(5)

    def __init__(self, parent, callback: Callable[[Optional[str]], None]):
        super().__init__(parent)
        self.callback = callback
        self.setWindowTitle(f"{APP_NAME} Settings")
        self.setModal(True)
        _fit_to_screen(self, 680, 780, 620, 700)
        self._cfg = load_config()
        self._asset_rows: Dict[str, _AssetRow] = {}
        self._scroll_areas: Dict[int, QScrollArea] = {}
        self._template_key = PDF_TEMPLATE_CLASSIC.key       # the invoice look and brand colour chosen in this dialog
        self._accent = ""
        self._preview_dirs: List[Path] = []                  # folders holding sample PDFs (removed when the dialog closes)
        self._build()
        self._load()

    def done(self, result: int):
        for folder in self._preview_dirs:
            shutil.rmtree(folder, ignore_errors=True)
        self._preview_dirs.clear()
        super().done(result)

    @classmethod
    def _currency_option_label(cls, opt: Tuple[str, str, str]) -> str:
        symbol, code, name = opt
        return f"{symbol}   {code} — {name}"

    # ------------------------------------------------------------------ #
    #  Small building blocks
    # ------------------------------------------------------------------ #
    @staticmethod
    def _section(title: str) -> QLabel:
        return make_label(title, "h2")

    @staticmethod
    def _note(text: str) -> QLabel:
        return make_label(text, "muted", wrap=True)

    @staticmethod
    def _entry(placeholder: str = "") -> PlaceholderEntry:
        return PlaceholderEntry(placeholder=placeholder)

    @staticmethod
    def _multiline(placeholder: str, rows: int) -> PlaceholderText:
        box = PlaceholderText(placeholder=placeholder, height=rows)
        box.setTabChangesFocus(True)
        box.setAcceptRichText(False)
        return box

    @staticmethod
    def _new_form() -> QFormLayout:
        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(10)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)    # combo boxes fill the row too
        form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.DontWrapRows)
        return form

    def _add_tab(self, title: str, body_layout: QLayout, scroll: bool = False) -> int:
        """Adds a tab whose content is `body_layout` (optionally inside a scroll area) and returns its index."""
        body = QWidget()
        body.setLayout(body_layout)
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.setContentsMargins(16, 16, 16, 16)
        if scroll:
            body_layout.setContentsMargins(0, 0, 10, 0)             # keeps the fields clear of the scroll bar
            area = QScrollArea()
            area.setWidgetResizable(True)
            area.setFrameShape(QFrame.Shape.NoFrame)
            area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            area.setWidget(body)
            outer.addWidget(area)
            index = self.tabs.addTab(page, title)
            self._scroll_areas[index] = area
            return index
        body_layout.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(body)
        return self.tabs.addTab(page, title)

    def _form_tab(self, title: str, form: QFormLayout, scroll: bool = False) -> int:
        column = QVBoxLayout()
        column.addLayout(form)
        column.addStretch(1)
        return self._add_tab(title, column, scroll)

    # ------------------------------------------------------------------ #
    #  Layout
    # ------------------------------------------------------------------ #
    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(16)

        self.tabs = QTabWidget()
        root.addWidget(self.tabs, 1)
        self._build_general_tab()
        self._build_your_details_tab()
        self._build_payment_tab()
        self._build_assets_tab()
        self._build_data_tab()

        row = QHBoxLayout()
        row.addStretch(1)
        self.save_btn = make_button("Save Configuration", "success", self._save)
        self.cancel_btn = make_button("Cancel", "secondary", self.reject)
        _no_default_button(self.save_btn, self.cancel_btn)
        row.addWidget(self.save_btn)
        row.addSpacing(8)
        row.addWidget(self.cancel_btn)
        root.addLayout(row)

    def _build_general_tab(self):
        form = self._new_form()
        form.addRow(self._section("Interface"))
        self.theme_combo = QComboBox()
        for name in THEME_NAMES:
            self.theme_combo.addItem(name.title(), name)
        form.addRow(make_label("App Theme:"), self.theme_combo)
        form.addRow(hline())
        form.addRow(self._section("Company Details"))
        form.addRow(self._note(
            "Your one business identity: shown as the letterhead at the top of every invoice, and as the name and "
            "GSTIN in the FROM block below it. There's only one place to set these, so they can't disagree with "
            "each other."))
        self.co_name = self._entry("Your company or firm name")
        self.co_desc = self._entry("A short tagline, e.g. Business Made Smarter.")
        self.co_gst = self._entry("22ABCDE1234F1Z5")
        form.addRow(make_label("Company Name:"), self.co_name)
        form.addRow(make_label("Company Tagline:"), self.co_desc)
        form.addRow(make_label("Company GSTIN:"), self.co_gst)
        self._form_tab("General", form)

    def _build_your_details_tab(self):
        form = self._new_form()
        form.addRow(self._section("Your Details"))
        form.addRow(self._note(
            "Your personal contact details, shown under the business name in the FROM block on new invoices. Editing "
            "this doesn't change any invoice you've already created. For your business/firm name and GSTIN, see "
            "Company Details on the General tab -- there's just one of each."))
        self.ud_name = self._entry("e.g. Priya Sharma")
        self.ud_phone = self._entry("+91 98765 43210")
        self.ud_email = self._entry("you@example.com")
        self.ud_website = self._entry("www.example.com")
        self.ud_address = self._entry("Street / building / area")
        self.ud_city = self._entry("e.g. Lucknow")
        self.ud_state = self._entry("e.g. Uttar Pradesh")
        # The state decides CGST/SGST or IGST when an invoice's place of supply is chosen, so offer the standard names.
        state_names = QCompleter([name for _code, name in INDIAN_STATES], self.ud_state)
        state_names.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        state_names.setFilterMode(Qt.MatchFlag.MatchContains)
        self.ud_state.setCompleter(state_names)
        self.ud_pin = self._entry("e.g. 226001")
        restrict_to_digits(self.ud_pin, max_length=6)
        for caption, field in (("Your Name:", self.ud_name), ("Phone Number:", self.ud_phone),
                               ("Email Address:", self.ud_email), ("Website (Optional):", self.ud_website),
                               ("Address:", self.ud_address), ("City:", self.ud_city), ("State:", self.ud_state),
                               ("PIN Code:", self.ud_pin)):
            form.addRow(make_label(caption), field)
        self._form_tab("Your Details", form, scroll=True)

    # ---- Invoice template -------------------------------------------------- #
    def _build_template_picker(self) -> QWidget:
        """The templates as cards (two to a row), the brand colour row and the sample-preview button."""
        box = QWidget()
        col = QVBoxLayout(box)
        col.setContentsMargins(0, 0, 0, 0)
        col.setSpacing(10)
        grid = QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(12)
        self.template_group = QButtonGroup(self)
        self.template_cards: Dict[str, _TemplateCard] = {}
        for n, tpl in enumerate(PDF_TEMPLATES.values()):
            card = _TemplateCard(tpl, self.template_group)
            card.picked.connect(self._on_template_picked)
            self.template_cards[tpl.key] = card
            grid.addWidget(card, n // 2, n % 2)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        col.addLayout(grid)

        self.accent_box = QWidget()
        accent_col = QVBoxLayout(self.accent_box)
        accent_col.setContentsMargins(0, 0, 0, 0)
        accent_col.setSpacing(6)
        row = QHBoxLayout()
        row.setSpacing(2)
        row.addWidget(make_label("Brand colour:"))
        row.addSpacing(8)
        self.accent_swatches: Dict[str, _Swatch] = {}
        for color, name in BRAND_COLORS:
            swatch = _Swatch(color, name)
            swatch.picked.connect(self._on_accent_picked)
            self.accent_swatches[color] = swatch
            row.addWidget(swatch)
        row.addStretch(1)
        accent_col.addLayout(row)
        row = QHBoxLayout()
        self.accent_custom_btn = make_button("Custom colour…", "secondary-outline", self._pick_custom_accent,
                                             tooltip="Choose any other colour")
        row.addWidget(self.accent_custom_btn)
        self.accent_reset_btn = make_button("Use the template's colour", "link", self._reset_accent,
                                            tooltip="Go back to the colour the template comes with")
        row.addWidget(self.accent_reset_btn)
        row.addStretch(1)
        accent_col.addLayout(row)
        col.addWidget(self.accent_box)
        self.accent_note = self._note("")
        col.addWidget(self.accent_note)

        line = QHBoxLayout()
        self.preview_template_btn = make_button("Preview sample invoice", "info-outline", self._preview_template,
                                                tooltip="Open a made-up invoice in the template and colour chosen above")
        line.addWidget(self.preview_template_btn)
        line.addStretch(1)
        col.addLayout(line)
        return box

    def _on_template_picked(self, key: str):
        self._template_key = key
        self._sync_template_picker()

    def _on_accent_picked(self, color: str):
        self._accent = color
        self._sync_template_picker()

    def _reset_accent(self):
        self._accent = ""
        self._sync_template_picker()

    def _pick_custom_accent(self):
        start = QColor(self._accent or PDF_TEMPLATES[self._template_key].accent or "#2563EB")
        chosen = QColorDialog.getColor(start, self, "Brand Colour")
        if chosen.isValid():
            self._accent = normalize_hex_color(chosen.name())
            self._sync_template_picker()

    def _sync_template_picker(self):
        """Shows the chosen template, brand colour and miniatures (the colour only recolours the templates that take one)."""
        card = self.template_cards.get(self._template_key)
        if card is not None and not card.radio.isChecked():
            card.radio.setChecked(True)                      # (picks it again: harmless, the state is already set)
        for key, c in self.template_cards.items():
            c.refresh(PDF_TEMPLATES[key].with_accent(self._accent))
        for color, swatch in self.accent_swatches.items():
            swatch.set_selected(color == self._accent)
        brandable = PDF_TEMPLATES[self._template_key].brandable
        for w in (*self.accent_swatches.values(), self.accent_custom_btn, self.accent_reset_btn):
            w.setEnabled(brandable)
        if not brandable:
            self.accent_note.setText("The Classic look is black and grey: it has no brand colour. Pick Modern, Minimal "
                                     "or Corporate to colour your invoices.")
        elif self._accent:
            self.accent_note.setText(f"Your colour {self._accent} is used by Modern, Minimal and Corporate.")
        else:
            self.accent_note.setText("Each template uses its own colour until you pick one. Your colour is shared by "
                                     "Modern, Minimal and Corporate.")

    def _preview_template(self):
        """Opens a made-up invoice in the chosen template and brand colour, under the saved business details."""
        settings = {**(self._cfg.get("pdf_settings") or {}), "template": self._template_key, "accent_color": self._accent}
        cfg = {**self._cfg, "pdf_settings": settings}
        try:
            folder = Path(tempfile.mkdtemp(prefix="vyaprix-template-"))
            self._preview_dirs.append(folder)
            data, payments = sample_invoice_data(cfg)
            path = PDFGenerator(data, cfg, payments=payments, out_dir=folder).generate()
        except Exception as exc:
            logger.exception("Could not make the sample invoice")
            return Messagebox.show_error("Preview Failed", f"The sample invoice could not be made:\n{exc}", parent=self)
        PDFViewer(self, path).show()

    def event(self, event):
        # Dragged to a monitor with another pixel density: paint the miniatures, colour chips and picture previews again
        # at that screen's resolution (a picture keeps the resolution it was made at, so it would stay soft).
        if (event.type() == QEvent.Type.DevicePixelRatioChange and hasattr(self, "_template_key")
                and hasattr(self, "_asset_rows")):
            self._sync_template_picker()
            for row in self._asset_rows.values():
                row.refresh()
        return super().event(event)

    def _build_payment_tab(self):
        form = self._new_form()
        form.addRow(self._section("Invoice Template"))
        form.addRow(self._note("The look of your printed invoices. You can still print a single invoice in another "
                               "template from its own form."))
        form.addRow(self._build_template_picker())
        form.addRow(hline())
        form.addRow(self._section("Payment Settings"))

        # Currency picker: a drop-down of common currencies plus a "Custom" escape hatch (a plain text field) for
        # anything that is not on the list. The extra row only shows once "Custom Symbol…" is picked.
        self.currency_combo = QComboBox()
        guard_wheel(self.currency_combo)
        for opt in self.CURRENCY_OPTIONS:
            self.currency_combo.addItem(self._currency_option_label(opt), opt[0])
        self.currency_combo.addItem(self.CUSTOM_CURRENCY, None)
        self.currency_combo.currentIndexChanged.connect(lambda _i: self._on_currency_selected())
        self.custom_currency_entry = self._entry("e.g. Rs., kr, ₪")
        form.addRow(make_label("Currency:"), self.currency_combo)
        form.addRow(make_label("Custom Symbol:"), self.custom_currency_entry)
        self._payment_form = form

        self.bank_name = self._entry("e.g. State Bank of India")
        self.acc_holder = self._entry("e.g. Priya Sharma")
        self.acc_number = self._entry("e.g. 000123456789")
        self.ifsc_code = self._entry("e.g. SBIN0001234")
        self.upi_id = self._entry("company@upi")
        for caption, field in (("Bank Name:", self.bank_name), ("Account Holder Name:", self.acc_holder),
                               ("Account Number:", self.acc_number), ("IFSC Code:", self.ifsc_code),
                               ("UPI ID:", self.upi_id)):
            form.addRow(make_label(caption), field)
        form.addRow(make_label("Payment Notes:"))
        self.payment_notes_text = self._multiline(
            "Anything else to show on the invoice -- cheque instructions, a late-fee policy, "
            "alternate payment methods, etc.", 3)
        form.addRow(self.payment_notes_text)
        self.pay_terms = self._entry("Due on Receipt")
        form.addRow(make_label("Default Payment Terms:"), self.pay_terms)
        form.addRow(hline())
        form.addRow(self._section("PDF Footer Note"))
        self.footer_text = self._multiline("Thank you for your business!", 3)
        form.addRow(self.footer_text)
        form.addRow(hline())
        form.addRow(self._section("Invoice Printing"))
        self.amount_words_check = QCheckBox("Print the invoice total in words (Rupees ... Only)")
        form.addRow(self.amount_words_check)
        form.addRow(self._note("Only for rupee invoices. It is printed under the totals, and is also shown while you "
                               "fill in an invoice."))
        form.addRow(make_label("Default Terms and Conditions:"))
        self.default_terms_text = self._multiline(
            "Every new invoice starts with this text in its Terms field -- e.g. payment within 15 days, "
            "late-fee policy, cancellation terms. Leave empty to print no terms.", 4)
        form.addRow(self.default_terms_text)
        self._form_tab("Payment && PDF", form, scroll=True)
        self._on_currency_selected()

    def _on_currency_selected(self):
        custom = self.currency_combo.currentData() is None
        self._payment_form.setRowVisible(self.custom_currency_entry, custom)

    def _build_assets_tab(self):
        col = QVBoxLayout()
        col.setSpacing(0)
        col.addWidget(self._section("Upload Assets"))
        col.addSpacing(8)
        col.addWidget(self._note("Files you choose are copied into Vyaprix's own data folder right away, so the app "
                                 "keeps working even if you later move or delete the original."))
        col.addSpacing(6)
        # The save destination is fixed per asset (logo.png, sign.png, upi_qr.png).
        for index, (key, label, dest) in enumerate((("logo", "Company Logo", "logo.png"),
                                                     ("signature", "Authorized Signature", "sign.png"),
                                                     ("upi_qr_code", "UPI QR Code", "upi_qr.png"))):
            if index:
                col.addWidget(hline())
            row = _AssetRow(key, label, dest)
            self._asset_rows[key] = row
            col.addWidget(row)
        col.addSpacing(12)
        col.addWidget(hline())
        col.addSpacing(14)
        col.addWidget(self._section("Payment Details Visibility"))
        col.addSpacing(8)
        self.always_show_qr_check = QCheckBox("Always show payment details, even on fully paid invoices")
        col.addWidget(self.always_show_qr_check)
        col.addSpacing(4)
        col.addWidget(self._note(
            "Off by default: the QR code and bank/UPI details only print on Unpaid or Partially Paid invoices, since "
            "there's nothing left to pay once an invoice is settled. Turn this on to always include them regardless "
            "of payment status."))
        col.addStretch(1)
        self._add_tab("Images && Assets", col)

    def _build_data_tab(self):
        col = QVBoxLayout()
        col.setSpacing(0)
        col.addWidget(self._section("Database Location"))
        col.addSpacing(8)
        self.db_path_label = make_label(str(DB_NAME), "muted")
        self.db_path_label.setWordWrap(True)
        self.db_path_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.db_path_label.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        col.addWidget(self.db_path_label)
        col.addSpacing(12)
        col.addWidget(hline())
        col.addSpacing(14)
        col.addWidget(self._section("Backup & Restore"))
        col.addSpacing(8)
        self.backup_btn = make_button("💾  Save a Backup Copy…  (Ctrl+B)", "secondary-outline", self._backup_db,
                                      tooltip="Back up the database now (Ctrl+B works from anywhere in the app)")
        self.import_btn = make_button("📥  Import Database…", "warning-outline", self._import_db)
        _no_default_button(self.backup_btn, self.import_btn)
        col.addWidget(self.backup_btn, 0, Qt.AlignmentFlag.AlignLeft)
        col.addSpacing(10)
        col.addWidget(self.import_btn, 0, Qt.AlignmentFlag.AlignLeft)
        col.addSpacing(4)
        col.addWidget(self._note(
            "Replaces all invoices, clients, and catalog items with the contents of the file you choose. Your current "
            "database is backed up automatically first, and the app refreshes automatically once the import finishes."))
        col.addStretch(1)
        self._add_tab("Data", col)

    # ------------------------------------------------------------------ #
    #  Backup / import
    # ------------------------------------------------------------------ #
    def _backup_db(self):
        prompt_and_backup_database(self)

    def _import_db(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import Database", str(_start_dir("import_dir")),
                                              "SQLite Database (*.db);;All Files (*)")
        if not path:
            return
        _remember_dir("import_dir", path)
        source = Path(path)

        problem = validate_import_database(source)
        if problem:
            return Messagebox.show_error(
                "Invalid Database", f"This doesn't look like a valid database file:\n{problem}", parent=self)
        if Path(DB_NAME).exists() and _same_file(source, Path(DB_NAME)):
            return Messagebox.show_info(
                "Import Database", "That file is the database Vyaprix is already using, so there is nothing to import.",
                parent=self)

        if not Messagebox.okcancel(
            "Import Database",
            "This replaces all invoices, clients, and catalog items in Vyaprix with the "
            "contents of the selected file.\n\nYour current database will be backed up "
            "automatically first, so you can go back if anything looks wrong.\n\nContinue?",
            parent=self,
        ):
            return

        try:
            backup_path = import_database_from(source)
        except DatabaseImportError as exc:
            if exc.stage == "backup":
                Messagebox.show_error(
                    "Import Failed", f"Could not create a safety backup, so nothing was changed:\n{exc}", parent=self)
            elif exc.restored:
                Messagebox.show_error(
                    "Import Failed",
                    f"The selected file couldn't be used as a Vyaprix database:\n{exc}\n\n"
                    "Your previous database has been restored.", parent=self)
            else:
                Messagebox.show_error(
                    "Import Failed",
                    f"The selected file couldn't be used as a Vyaprix database:\n{exc}\n\n"
                    "Your previous database could not be put back automatically. A copy of it was saved as "
                    f"{exc.backup_path.name if exc.backup_path else 'a Vyaprix-pre-import-backup file'} in the "
                    "Vyaprix data folder.", parent=self)
            return

        Messagebox.show_info(
            "Import Complete",
            f"Database imported successfully.\n\nYour previous database was backed up to:\n{backup_path.name}",
            parent=self)
        self.callback(None)     # reloads CONFIG and refreshes every screen's list

    # ------------------------------------------------------------------ #
    #  Load / save
    # ------------------------------------------------------------------ #
    def _load(self):
        c = self._cfg
        theme = normalize_theme_name(c.get("theme"))
        self.theme_combo.setCurrentIndex(max(0, self.theme_combo.findData(theme)))
        company = c.get("company_details", {})
        self.co_name.setText(company.get("name", ""))
        self.co_desc.setText(company.get("description", ""))
        self.co_gst.setText(company.get("gstin", ""))

        ud = c.get("user_details", {})
        self.ud_name.setText(ud.get("name", ""))
        self.ud_phone.setText(ud.get("phone", ""))
        self.ud_email.setText(ud.get("email", ""))
        self.ud_website.setText(ud.get("website", ""))
        self.ud_address.setText(ud.get("address", ""))
        self.ud_city.setText(ud.get("city", ""))
        self.ud_state.setText(ud.get("state", ""))
        self.ud_pin.setText(ud.get("pin_code", ""))

        current_symbol = c.get("currency_symbol", "₹")
        index = self.currency_combo.findData(current_symbol)
        if index >= 0:
            self.currency_combo.setCurrentIndex(index)
            self.custom_currency_entry.clear()
        else:
            self.currency_combo.setCurrentIndex(self.currency_combo.count() - 1)
            self.custom_currency_entry.setText(current_symbol)
        self._on_currency_selected()

        pd = c.get("payment_details", {})
        self.bank_name.setText(pd.get("bank_name", ""))
        self.acc_holder.setText(pd.get("account_holder_name", ""))
        self.acc_number.setText(pd.get("account_number", ""))
        self.ifsc_code.setText(pd.get("ifsc_code", ""))
        self.upi_id.setText(pd.get("upi_id", ""))
        self.payment_notes_text.setPlainText(pd.get("payment_notes", ""))
        self.pay_terms.setText(pd.get("payment_terms", ""))

        pdf = c.get("pdf_settings", {})
        self.footer_text.setPlainText(pdf.get("footer_message", ""))
        self.always_show_qr_check.setChecked(bool(pdf.get("always_show_qr_code", False)))
        self.amount_words_check.setChecked(bool(pdf.get("amount_in_words", False)))
        self.default_terms_text.setPlainText(pdf.get("default_terms", ""))
        self._template_key = pdf.get("template") if pdf.get("template") in PDF_TEMPLATES else PDF_TEMPLATE_CLASSIC.key
        self._accent = normalize_hex_color(pdf.get("accent_color"))
        self._sync_template_picker()

        paths = c.get("asset_paths", {})
        for key, row in self._asset_rows.items():
            row.set_path(paths.get(key, ""))

    def _reject_input(self, title: str, message: str, tab: int, field: QWidget) -> None:
        """Explains what is wrong, then takes the user to the offending field (it may be on another tab)."""
        Messagebox.show_error(title, message, parent=self)
        self.tabs.setCurrentIndex(tab)
        field.setFocus()
        area = self._scroll_areas.get(tab)
        if area is not None:
            area.ensureWidgetVisible(field)

    def _currency_symbol(self) -> str:
        if self.currency_combo.currentData() is None:
            return self.custom_currency_entry.text().strip() or "₹"
        return self.currency_combo.currentData()

    def _save(self):
        email = self.ud_email.text().strip()
        if email and not is_valid_email(email):
            return self._reject_input("Invalid Email", f'"{email}" doesn\'t look like a valid email address.',
                                      self.TAB_DETAILS, self.ud_email)

        phone = self.ud_phone.text().strip()
        if phone and not is_valid_phone(phone):
            return self._reject_input(
                "Invalid Phone Number",
                f'"{phone}" doesn\'t look like a valid phone number.\n'
                "Use digits only, optionally with +, spaces, or dashes.", self.TAB_DETAILS, self.ud_phone)

        company_gstin = self.co_gst.text().strip()
        if company_gstin and not is_valid_gstin(company_gstin):
            return self._reject_input(
                "Invalid GST Number",
                f'"{company_gstin}" is not a valid GSTIN.\n'
                "A GSTIN is 15 characters, for example 22ABCDE1234F1Z5.", self.TAB_GENERAL, self.co_gst)

        theme = self.theme_combo.currentData() or DEFAULT_THEME
        try:
            save_config({
                "theme": theme,
                "company_details": {
                    "name": self.co_name.text().strip(),
                    "description": self.co_desc.text().strip(),
                    "gstin": company_gstin,
                },
                "user_details": {
                    "name": self.ud_name.text().strip(),
                    "phone": phone,
                    "email": email,
                    "website": self.ud_website.text().strip(),
                    "address": self.ud_address.text().strip(),
                    "city": self.ud_city.text().strip(),
                    "state": self.ud_state.text().strip(),
                    "pin_code": self.ud_pin.text().strip(),
                    # Preserved but no longer editable here directly -- see compose_your_contact()'s
                    # upgrade-path fallback.
                    "contact_info": self._cfg.get("user_details", {}).get("contact_info", ""),
                },
                "payment_details": {
                    "upi_id": self.upi_id.text().strip(),
                    "payment_terms": self.pay_terms.text().strip(),
                    "bank_name": self.bank_name.text().strip(),
                    "account_holder_name": self.acc_holder.text().strip(),
                    "account_number": self.acc_number.text().strip(),
                    "ifsc_code": self.ifsc_code.text().strip(),
                    "payment_notes": self.payment_notes_text.toPlainText().strip(),
                },
                "pdf_settings": {"footer_message": self.footer_text.toPlainText().strip(),
                                 "always_show_qr_code": self.always_show_qr_check.isChecked(),
                                 "amount_in_words": self.amount_words_check.isChecked(),
                                 "default_terms": self.default_terms_text.toPlainText().strip(),
                                 "template": self._template_key,
                                 "accent_color": self._accent},
                "asset_paths": {key: row.rel_path for key, row in self._asset_rows.items()},
                "currency_symbol": self._currency_symbol(),
            })
        except OSError as exc:
            return Messagebox.show_error("Save Failed", f"Your settings could not be saved:\n{exc}", parent=self)

        try:
            self.callback(theme)        # the new theme goes to the app, which applies it and refreshes every screen
        finally:
            self.accept()


# =============================================================================
#  15. SCREENS
# =============================================================================

# -----------------------------------------------------------------------------
#  Clients and Catalog
# -----------------------------------------------------------------------------
class _ManagerFrame(QWidget):
    """Common body of the Clients and Catalog screens: header + add button, search box
    (debounced), multi-select table, Edit/Delete buttons whose enabled state follows the
    selection, right-click menu, double-click / Enter to edit, Delete key to delete."""

    TITLE = ""
    SUBTITLE = ""
    ADD_TEXT = ""
    ADD_TIP = ""
    SEARCH_TIP = ""
    HINT_NONE = ""
    #: (key, header text, width px, alignment)
    COL_DEFS: List[Tuple[str, str, int, Qt.AlignmentFlag]] = []
    STRETCH_COL = 1

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.db_manager = controller.db_manager
        self._build()
        self._debounce_search = Debouncer(self, 200, self._populate)

    # ---- hooks the two screens implement ---------------------------------- #
    def _fetch(self, search: str) -> List[Dict]:
        raise NotImplementedError

    def _cells(self, rec: Dict) -> List[str]:
        raise NotImplementedError

    def _get_by_id(self, rec_id: int) -> Optional[Dict]:
        raise NotImplementedError

    def _open_editor(self, rec: Optional[Dict]):
        raise NotImplementedError

    def _delete_ids(self, ids: List[int]):
        raise NotImplementedError

    def _delete_prompt(self, n: int) -> Tuple[str, str]:
        raise NotImplementedError

    # ---- layout ----------------------------------------------------------- #
    def _build(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 16, 24, 14)

        head = QHBoxLayout()
        left = QVBoxLayout()
        left.setSpacing(2)
        left.addWidget(make_label(self.TITLE, "h1"))
        left.addWidget(make_label(self.SUBTITLE, "muted"))
        head.addLayout(left)
        head.addStretch()
        self.add_btn = make_button(self.ADD_TEXT, "primary", lambda: self._open_editor(None),
                                   tooltip=self.ADD_TIP, min_height=38)
        head.addWidget(self.add_btn, 0, Qt.AlignmentFlag.AlignVCenter)
        outer.addLayout(head)
        outer.addSpacing(18)

        srch = QHBoxLayout()
        srch.addWidget(QLabel("🔍"))
        self.search_entry = QLineEdit()
        self.search_entry.setFixedWidth(300)
        self.search_entry.setClearButtonEnabled(True)
        self.search_entry.setPlaceholderText("Search…")
        self.search_entry.setToolTip(self.SEARCH_TIP)
        self.search_entry.textChanged.connect(lambda _text: self._debounce_search())
        srch.addWidget(self.search_entry)
        srch.addStretch()
        outer.addLayout(srch)
        outer.addSpacing(8)

        self.table = KeyTable(0, len(self.COL_DEFS))
        self.table.setHorizontalHeaderLabels([c[1] for c in self.COL_DEFS])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.setWordWrap(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        header = self.table.horizontalHeader()
        header.setHighlightSections(False)
        header.setSectionsClickable(False)
        for i, (_key, _text, width, align) in enumerate(self.COL_DEFS):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch if i == self.STRETCH_COL
                                        else QHeaderView.ResizeMode.Interactive)
            self.table.setColumnWidth(i, width)
        self.table.itemSelectionChanged.connect(self._on_select)
        self.table.itemDoubleClicked.connect(lambda _item: self._edit_selected())
        self.table.delete_pressed.connect(self._delete_selected)
        self.table.enter_pressed.connect(self._edit_selected)
        self.table.customContextMenuRequested.connect(self._show_ctx_menu)
        outer.addWidget(self.table, 1)

        self._ctx = QMenu(self)
        self._ctx.addAction("✏️  Edit", self._edit_selected)
        self._ctx.addAction("🗑️  Delete", self._delete_selected)

        act = QHBoxLayout()
        act.setContentsMargins(0, 10, 0, 0)
        self.hint_lbl = make_label(self.HINT_NONE, "small")
        act.addWidget(self.hint_lbl)
        act.addStretch()
        self.edit_btn = make_button("✏  Edit", "secondary", self._edit_selected)
        self.del_btn = make_button("🗑  Delete", "danger-outline", self._delete_selected)
        self.edit_btn.setEnabled(False)
        self.del_btn.setEnabled(False)
        act.addWidget(self.edit_btn)
        act.addSpacing(8)
        act.addWidget(self.del_btn)
        outer.addLayout(act)

    # ---- data ------------------------------------------------------------- #
    def on_show(self, *args):
        self._populate()

    def _populate(self):
        records = self._fetch(self.search_entry.text().strip())
        t = self.table
        t.setUpdatesEnabled(False)
        t.blockSignals(True)
        t.clearSelection()
        t.setRowCount(len(records))
        for r, rec in enumerate(records):
            for c, text in enumerate(self._cells(rec)):
                it = QTableWidgetItem(text)
                it.setTextAlignment(self.COL_DEFS[c][3] | Qt.AlignmentFlag.AlignVCenter)
                if c == 0:
                    it.setData(Qt.ItemDataRole.UserRole, int(rec["id"]))
                t.setItem(r, c, it)
        t.blockSignals(False)
        t.setUpdatesEnabled(True)
        self._on_select()

    def _selected_ids(self) -> List[int]:
        rows = sorted({i.row() for i in self.table.selectedIndexes()})
        return [int(self.table.item(r, 0).data(Qt.ItemDataRole.UserRole)) for r in rows]

    def _on_select(self):
        n = len(self._selected_ids())
        self.edit_btn.setEnabled(n == 1)
        self.del_btn.setEnabled(n >= 1)
        self.hint_lbl.setText(f"{n} selected" if n else self.HINT_NONE)

    def _selected_record(self) -> Optional[Dict]:
        ids = self._selected_ids()
        if len(ids) != 1:
            return None
        return self._get_by_id(ids[0])

    def _edit_selected(self):
        if rec := self._selected_record():
            self._open_editor(rec)

    def _delete_selected(self):
        ids = self._selected_ids()
        if not ids:
            return
        title, msg = self._delete_prompt(len(ids))
        if Messagebox.okcancel(title, msg, parent=self.controller):
            self._delete_ids(ids)
            self._populate()

    def focus_search(self):
        """Ctrl+F: jump into the search box, with any existing text selected so typing
        immediately replaces it."""
        self.search_entry.setFocus()
        self.search_entry.selectAll()

    def _show_ctx_menu(self, pos: QPoint):
        item = self.table.itemAt(pos)
        if item is None:
            return
        if not self.table.item(item.row(), 0).isSelected():
            self.table.select_only_row(item.row())
        self._ctx.exec(self.table.viewport().mapToGlobal(pos))


class ClientsFrame(_ManagerFrame):
    TITLE = "Clients"
    SUBTITLE = "Your address book for fast invoice billing."
    ADD_TEXT = "＋  Add Client  (Ctrl+N)"
    ADD_TIP = "Add a new client (Ctrl+N)"
    SEARCH_TIP = "Search clients (Ctrl+F)"
    HINT_NONE = "Select a client to edit or delete"
    COL_DEFS = [
        ("name",    "Client / Company", 260, Qt.AlignmentFlag.AlignLeft),
        ("contact", "Contact Info",     320, Qt.AlignmentFlag.AlignLeft),
        ("gstin",   "GSTIN",            160, Qt.AlignmentFlag.AlignLeft),
    ]

    def _fetch(self, search):
        return self.db_manager.get_all_clients(search)

    def _cells(self, c):
        contact_preview = (c.get("contact") or "").replace("\n", "  ·  ")
        return [c["name"], contact_preview, c.get("gstin") or ""]

    def _get_by_id(self, rec_id):
        return self.db_manager.get_client_by_id(rec_id)

    def _open_editor(self, client: Optional[Dict]):
        run_dialog(ClientDialog(self.controller, self.db_manager, client, self._populate))

    def _delete_ids(self, ids):
        self.db_manager.delete_clients(ids)

    def _delete_prompt(self, n):
        msg = f"Delete {n} client(s) from your address book?" if n > 1 else "Delete this client from your address book?"
        return "Delete Client(s)", msg

    def new_client(self):
        """Ctrl+N on this page: open the Add Client dialog, exactly like the "＋ Add Client" button."""
        self._open_editor(None)


class CatalogFrame(_ManagerFrame):
    TITLE = "Catalog"
    SUBTITLE = "Reusable products & services for faster invoice lines."
    ADD_TEXT = "＋  Add Item  (Ctrl+N)"
    ADD_TIP = "Add a new catalog item (Ctrl+N)"
    SEARCH_TIP = "Search catalog items (Ctrl+F)"
    HINT_NONE = "Select an item to edit or delete"
    COL_DEFS = [
        ("title",  "Item / Service", 230, Qt.AlignmentFlag.AlignLeft),
        ("desc",   "Description",    280, Qt.AlignmentFlag.AlignLeft),
        ("amount", "Default Rate",   130, Qt.AlignmentFlag.AlignRight),
        ("unit",   "Unit",           120, Qt.AlignmentFlag.AlignLeft),
        ("hsn",    "HSN/SAC",         90, Qt.AlignmentFlag.AlignLeft),
        ("gst",    "GST",             64, Qt.AlignmentFlag.AlignRight),
    ]

    def _fetch(self, search):
        return self.db_manager.get_all_catalog_items(search)

    def _cells(self, it):
        names = [u for u, _rate in catalog_units(it)]
        unit_text = ", ".join(names) if len(names) <= 2 else f"{names[0]}, {names[1]} +{len(names) - 2}"
        gst = "" if it.get("tax_rate") is None else f"{safe_float(it['tax_rate']):g}%"
        return [it["title"], (it.get("description") or "").replace("\n", "  ·  "), fmt_money(it.get("amount")),
                unit_text, str(it.get("hsn") or ""), gst]

    def _get_by_id(self, rec_id):
        return self.db_manager.get_catalog_item_by_id(rec_id)

    def _open_editor(self, item: Optional[Dict]):
        run_dialog(CatalogItemDialog(self.controller, self.db_manager, item, self._populate))

    def _delete_ids(self, ids):
        self.db_manager.delete_catalog_items(ids)

    def _delete_prompt(self, n):
        msg = f"Delete {n} catalog item(s)?" if n > 1 else "Delete this catalog item?"
        return "Delete Item(s)", msg

    def new_catalog_item(self):
        """Ctrl+N on this page: open the Add Catalog Item dialog, exactly like the "＋ Add Item" button."""
        self._open_editor(None)


# -----------------------------------------------------------------------------
#  Dashboard (invoice list)
# -----------------------------------------------------------------------------
class DashboardFrame(QWidget):
    STATUS_FILTERS = ["All", "Paid", "Due", "Partial", "Overpaid", "Complimentary"]
    SORT_ASC  = " ▲"
    SORT_DESC = " ▼"
    _L, _C, _R = Qt.AlignmentFlag.AlignLeft, Qt.AlignmentFlag.AlignHCenter, Qt.AlignmentFlag.AlignRight
    #: (column id, header text, width px, alignment) -- the ids are DatabaseManager._SORT_COLUMN_SQL keys
    COL_DEFS  = [
        ("id",      "Invoice #",        112, _C),
        ("client",  "Client / Company", 250, _L),
        ("date",    "Date",             100, _C),
        ("status",  "Status",           140, _C),
        ("total",   "Total",            100, _R),
        ("paid",    "Paid",              94, _R),
        ("balance", "Balance",           94, _R),
        ("note",    "🗒",                 38, _C),
    ]
    STRETCH_COL = 1
    PAGE_SIZE_OPTIONS = [25, 50, 100]
    DEFAULT_PAGE_SIZE = 50

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller     = controller
        self.db_manager     = controller.db_manager
        self._status_filter = "All"
        self._sort_col      = "date"
        self._sort_rev      = True
        self._filter_btns: Dict[str, QPushButton] = {}
        self._row_index: Dict[str, Dict] = {}
        self._page          = 1
        self._page_size     = self.DEFAULT_PAGE_SIZE
        self._total_pages   = 1
        self._bulk_op_running = False  # guards against overlapping bulk operations -- see _start_bulk_op()
        self._scroll_top_pending = False  # set when the result set changes (search/filter/sort/page): next populate starts at the top
        self._build()
        self._debounce_search = Debouncer(self, 200, self._populate)

    # ------------------------------------------------------------------ #
    #  Layout
    # ------------------------------------------------------------------ #
    def _build(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 16, 24, 14)
        outer.setSpacing(0)
        self._header(outer)
        self._stat_cards(outer)
        self._filter_bar(outer)
        self._table(outer)
        self._pagination_bar(outer)
        self._action_bar(outer)

    def _header(self, outer: QVBoxLayout):
        row = QHBoxLayout()
        left = QVBoxLayout()
        left.setSpacing(2)
        left.addWidget(make_label("Dashboard", "h1"))
        left.addWidget(make_label("Every invoice, at a glance.", "muted"))
        row.addLayout(left)
        row.addStretch()

        refresh_btn = make_button("🔄  Refresh  (Ctrl+R)", "secondary-outline", self._act_refresh,
                                  tooltip="Reload the invoice list and stats (Ctrl+R)", min_height=38)
        self.export_csv_btn = make_button("📊  Export CSV", "info-outline", self._export_csv, min_height=38)
        new_inv_btn = make_button("＋  New Invoice  (Ctrl+N)", "primary",
                                  lambda: self.controller.show_frame(InvoiceFormFrame),
                                  tooltip="Start a new invoice (Ctrl+N)", min_height=38)
        row.addWidget(refresh_btn)
        row.addSpacing(8)
        row.addWidget(self.export_csv_btn)
        row.addSpacing(8)
        row.addWidget(new_inv_btn)
        outer.addLayout(row)
        outer.addSpacing(18)

    def _stat_cards(self, outer: QVBoxLayout):
        row = QHBoxLayout()
        row.setSpacing(12)
        self.s_billed   = self._card(row, "💰  Total Billed",   fmt_money(0), "primary")
        self.s_received = self._card(row, "✅  Total Received", fmt_money(0), "success")
        self.s_due      = self._card(row, "⚠️  Balance Due",    fmt_money(0), "danger")
        self.s_count    = self._card(row, "📄  Total Invoices", "0",          "secondary")
        outer.addLayout(row)
        outer.addSpacing(18)

    @staticmethod
    def _card(row: QHBoxLayout, title: str, value: str, tone: str) -> QLabel:
        card = QFrame()
        set_role(card, "stat-solid", tone)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(18, 14, 18, 14)
        lay.setSpacing(4)
        lay.addWidget(make_label(title, "stat-solid-title", tone))
        lbl = make_label(value, "stat-solid-value", tone)
        # A long amount must never force the whole page wider than the window: the label may shrink below
        # its text width, and resizeEvent() switches to a smaller font tier before the text would be clipped.
        lbl.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        lbl.setMinimumWidth(0)
        lay.addWidget(lbl)
        row.addWidget(card, 1)               # equal stretch: the four cards share the row evenly
        return lbl

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_card_density()

    def _update_card_density(self):
        """Step the big stat figures down one font size when four cards no longer have room for a
        ₹ amount of typical length at the full size (narrow windows)."""
        per_card = (self.width() - 48 - 3 * 12) / 4
        compact = per_card < 236
        for lbl in (self.s_billed, self.s_received, self.s_due, self.s_count):
            set_flag(lbl, "compact", compact)

    def _filter_bar(self, outer: QVBoxLayout):
        row = QHBoxLayout()
        row.addWidget(make_label("Filter:", "small"))
        row.addSpacing(6)
        for status in self.STATUS_FILTERS:
            label = "All" if status == "All" else STATUS_DISPLAY_LABELS.get(status, status).title()
            btn = make_button(label, "primary" if status == "All" else "secondary-outline",
                              lambda _=False, s=status: self._set_filter(s))
            self._filter_btns[status] = btn
            row.addWidget(btn)
            row.addSpacing(2)
        row.addStretch(1)

        row.addWidget(QLabel("🔍"))
        self.search_entry = QLineEdit()
        # Flexible: 130-260 px, so the bar still fits when the window is at its minimum width.
        self.search_entry.setMinimumWidth(130)
        self.search_entry.setMaximumWidth(260)
        self.search_entry.setToolTip("Search by invoice number, client name, or note (Ctrl+F)")
        self.search_entry.setPlaceholderText("Search…")
        # Ctrl+A means "Select All invoices" on this screen (see select_all() below), but inside the
        # search box a focused QLineEdit claims Ctrl+A for itself (select all *text*), which is exactly
        # the "don't hijack the keystroke while typing" behaviour the Tk build had to bind by hand.
        self.search_entry.textChanged.connect(lambda _text: self._on_search_changed())
        row.addWidget(self.search_entry, 1)
        row.addSpacing(4)
        clear_btn = make_button("✕", "secondary-outline", lambda: self.search_entry.clear(), tooltip="Clear search")
        clear_btn.setFixedWidth(36)
        set_flag(clear_btn, "glyph", True)
        row.addWidget(clear_btn)
        outer.addLayout(row)
        outer.addSpacing(8)

    def _on_search_changed(self):
        self._page = 1  # a new search query invalidates the old page position
        self._scroll_top_pending = True
        self._debounce_search()

    def _set_filter(self, status: str):
        self._status_filter = status
        self._page = 1  # a new filter invalidates the old page position
        self._scroll_top_pending = True
        for s, btn in self._filter_btns.items():
            set_role(btn, "primary" if s == status else "secondary-outline")
        self._populate()

    def _table(self, outer: QVBoxLayout):
        self.table = KeyTable(0, len(self.COL_DEFS))
        t = self.table
        t.setHorizontalHeaderLabels([c[1] for c in self.COL_DEFS])
        t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        t.setAlternatingRowColors(True)
        t.setShowGrid(False)
        t.setWordWrap(False)
        t.verticalHeader().setVisible(False)
        t.verticalHeader().setDefaultSectionSize(34)
        t.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        header = t.horizontalHeader()
        header.setHighlightSections(False)
        header.setSectionsClickable(True)          # click a header to sort by it
        header.setMinimumSectionSize(40)
        header.setCursor(Qt.CursorShape.PointingHandCursor)
        for i, (_cid, _text, width, align) in enumerate(self.COL_DEFS):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Stretch if i == self.STRETCH_COL
                                        else QHeaderView.ResizeMode.Interactive)
            t.setColumnWidth(i, width)
            t.horizontalHeaderItem(i).setTextAlignment(align | Qt.AlignmentFlag.AlignVCenter)
        header.sectionClicked.connect(lambda idx: self._sort_by(self.COL_DEFS[idx][0]))
        self._refresh_header_arrows()              # shows the initial "Date ▼" (the Tk build sorted this way without saying so)
        self._fit_columns()

        t.itemSelectionChanged.connect(self._on_select)
        t.itemDoubleClicked.connect(lambda _item: self._act_edit())
        t.customContextMenuRequested.connect(self._show_ctx_menu)
        t.delete_pressed.connect(self._act_delete)
        outer.addWidget(t, 1)

        # Context menus
        self._ctx_single = QMenu(self)
        self._ctx_single.addAction("📄  View PDF",   self._act_view_pdf)
        self._ctx_single.addAction("✏️  Edit",       self._act_edit)
        self._ctx_single.addAction("📋  Duplicate",  self._act_duplicate)
        self._ctx_single.addAction("🗒  Edit Note",   self._act_edit_note)
        self._ctx_single.addSeparator()
        self._ctx_single.addAction("🗑️  Delete",     self._act_delete)

        self._ctx_bulk = QMenu(self)
        self._ctx_bulk.addAction("🖨️  Bulk Print",   self._act_bulk_print)
        self._ctx_bulk.addAction("📊  Bulk Export",  self._act_bulk_export)
        self._ctx_bulk.addSeparator()
        self._ctx_bulk.addAction("🗑️  Bulk Delete",  self._act_delete)

    def _fit_columns(self):
        """Widen (never narrow) the fixed-width columns to what they have to show. The designed widths suit Segoe UI;
        in another font, with a longer currency symbol or with a bigger amount, a cell would end in "…" -- and a
        cut-off amount reads as a different amount. The client column stretches, so it gives up the room."""
        dates = [datetime.date(2026, month, 28).strftime(DISPLAY_DATE_FORMAT) for month in range(1, 13)]
        money = fmt_money(999999.99)
        self.table.fit_columns({0: [], 2: dates, 3: list(STATUS_DISPLAY_LABELS.values()),
                                4: [money], 5: [money], 6: [money]})

    def _refresh_header_arrows(self):
        for i, (cid, text, _w, _a) in enumerate(self.COL_DEFS):
            arrow = (self.SORT_DESC if self._sort_rev else self.SORT_ASC) if cid == self._sort_col else ""
            self.table.horizontalHeaderItem(i).setText(f"{text}{arrow}")

    def _sort_by(self, col: str):
        """Sorts the list by a clicked column header and shows the matching arrow on it."""
        if self._sort_col == col:
            self._sort_rev = not self._sort_rev
        else:
            self._sort_col = col
            self._sort_rev = False  # Default to ascending when clicking a new column
        self._page = 1  # a new sort order invalidates the old page position
        self._refresh_header_arrows()
        self._scroll_top_pending = True
        self._populate()

    def _pagination_bar(self, outer: QVBoxLayout):
        outer.addSpacing(8)
        row = QHBoxLayout()
        self.page_info_lbl = make_label("", "small")
        row.addWidget(self.page_info_lbl)
        row.addStretch()

        row.addWidget(make_label("Rows per page:", "small"))
        row.addSpacing(6)
        self.page_size_combo = QComboBox()
        self.page_size_combo.addItems([str(n) for n in self.PAGE_SIZE_OPTIONS])
        self.page_size_combo.setCurrentText(str(self._page_size))
        guard_wheel(self.page_size_combo)
        self.page_size_combo.setFixedWidth(76)
        self.page_size_combo.currentIndexChanged.connect(lambda _i: self._on_page_size_change())
        row.addWidget(self.page_size_combo)
        row.addSpacing(16)

        self.prev_page_btn = make_button("◀  Prev", "secondary-outline", self._prev_page)
        self.next_page_btn = make_button("Next  ▶", "secondary-outline", self._next_page)
        for b in (self.prev_page_btn, self.next_page_btn):
            b.setMinimumWidth(92)
        row.addWidget(self.prev_page_btn)
        row.addSpacing(4)
        row.addWidget(self.next_page_btn)
        outer.addLayout(row)

    def _prev_page(self):
        if self._page > 1:
            self._page -= 1
            self._scroll_top_pending = True
            self._populate()

    def _next_page(self):
        if self._page < self._total_pages:
            self._page += 1
            self._scroll_top_pending = True
            self._populate()

    def _on_page_size_change(self):
        try:
            new_size = int(self.page_size_combo.currentText())
        except (TypeError, ValueError):
            new_size = self.DEFAULT_PAGE_SIZE
        if new_size != self._page_size:
            self._page_size = new_size
            self._page = 1
            self._scroll_top_pending = True
            self._populate()

    def _action_bar(self, outer: QVBoxLayout):
        outer.addSpacing(10)
        row = QHBoxLayout()
        self.hint_lbl = make_label("Select an invoice to take action", "small", wrap=True)
        self.hint_lbl.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        self.hint_lbl.setMinimumWidth(120)            # wraps onto a second line in a narrow window instead of clipping
        row.addWidget(self.hint_lbl, 1)

        # -- Single-selection actions (left to right, as the Tk build packed them) --
        self.del_btn  = make_button("🗑  Delete",    "danger-outline",    self._act_delete,
                                    tooltip="Delete the selected invoice (Delete key)")
        self.note_btn = make_button("🗒  Note",      "secondary-outline", self._act_edit_note)
        self.dup_btn  = make_button("📋  Duplicate",  "secondary-outline", self._act_duplicate,
                                    tooltip="Duplicate this invoice (Ctrl+D)")
        self.edit_btn = make_button("✏  Edit",      "secondary",         self._act_edit)
        self.pdf_btn  = make_button("📄  View PDF",  "primary",           self._act_view_pdf,
                                    tooltip="Open this invoice's PDF, ready to print or export (Ctrl+P / Ctrl+E)")
        self._single_buttons = [self.del_btn, self.note_btn, self.dup_btn, self.edit_btn, self.pdf_btn]

        # -- Bulk-selection actions --
        self.bulk_del_btn    = make_button("🗑  Bulk Delete", "danger-outline",    self._act_bulk_delete,
                                           tooltip="Delete the selected invoices (Delete key)")
        self.bulk_export_btn = make_button("📊  Bulk Export", "secondary-outline", self._act_bulk_export)
        self.bulk_print_btn  = make_button("🖨  Bulk Print",  "primary",           self._act_bulk_print,
                                           tooltip="Merge & open the selected invoices' PDFs, ready to print or export (Ctrl+P / Ctrl+E)")
        self._bulk_group = [self.bulk_del_btn, self.bulk_export_btn, self.bulk_print_btn]

        # Every button that can kick off a bulk-scale operation -- disabled together while one is
        # running (see _start_bulk_op()/_end_bulk_op()), so the user can never have two running at once
        # and always has a clear, consistent signal that one currently is.
        self._bulk_buttons = self._bulk_group + [self.export_csv_btn]

        for b in self._single_buttons + self._bulk_group:
            row.addSpacing(4)
            row.addWidget(b)
        outer.addLayout(row)
        self._show_single_actions()

    def _show_single_actions(self, enabled: bool = False):
        """Show the single-invoice buttons (hide the bulk ones); `enabled` follows whether a row is selected."""
        for b in self._bulk_group:
            b.setVisible(False)
        for b in self._single_buttons:
            b.setVisible(True)
            b.setEnabled(enabled)

    def _show_bulk_actions(self):
        for b in self._single_buttons:
            b.setVisible(False)
        for b in self._bulk_group:
            b.setVisible(True)

    # ------------------------------------------------------------------ #
    #  Data
    # ------------------------------------------------------------------ #
    def on_show(self, *args):
        self._populate()
        self._update_stats()

    def _on_select(self):
        n = len(self._selected_ids())
        if n >= 2:
            self._show_bulk_actions()
            self.hint_lbl.setText(f"{n} invoices selected")
        else:
            self._show_single_actions(enabled=(n == 1))
            self.hint_lbl.setText(
                "Double-click to edit  •  Right-click for options  •  Ctrl/Shift-click to multi-select" if n
                else "Select an invoice to take action"
            )

    def _selected_ids(self) -> List[str]:
        rows = sorted({i.row() for i in self.table.selectedIndexes()})
        return [self.table.item(r, 0).text() for r in rows]

    def _selected_id(self) -> Optional[str]:
        ids = self._selected_ids()
        return ids[0] if ids else None

    def _populate(self):
        t = self.table
        # Remember the selection and scroll position so neither jumps while typing in the search box,
        # editing a note or refreshing.
        keep_selected = set(self._selected_ids())
        scroll_pos = 0 if self._scroll_top_pending else t.verticalScrollBar().value()
        self._scroll_top_pending = False

        # --- Fetch: search + filter + sort + pagination are all done in SQL (see
        # DatabaseManager.get_invoices_page) -- this call never returns more than one page's worth of
        # rows, no matter how many invoices are in the table, which is what actually keeps the Dashboard
        # fast and light on memory at 100,000+ invoices.
        search_query = self.search_entry.text().strip()
        rows, total = self.db_manager.get_invoices_page(
            search_query=search_query, status_filter=self._status_filter,
            sort_col=self._sort_col, sort_desc=self._sort_rev,
            page=self._page, page_size=self._page_size,
        )
        self._total_pages = max(1, -(-total // self._page_size))  # ceiling division
        if self._page > self._total_pages:
            # The page we were on no longer has any rows on it (e.g. the last invoice on the last page
            # just got deleted) -- clamp back into range and re-fetch instead of showing a blank page.
            self._page = self._total_pages
            rows, total = self.db_manager.get_invoices_page(
                search_query=search_query, status_filter=self._status_filter,
                sort_col=self._sort_col, sort_desc=self._sort_rev,
                page=self._page, page_size=self._page_size,
            )

        start = (self._page - 1) * self._page_size
        self._row_index.clear()
        ncols = len(self.COL_DEFS)
        t.setUpdatesEnabled(False)
        t.blockSignals(True)
        try:
            t.clearSelection()
            t.setRowCount(len(rows))
            for r, inv in enumerate(rows):
                row = dict(inv)
                is_comp    = bool(row.get("is_complimentary"))
                status_val = (row["status"] or "Due").upper()
                # Display wording only -- status_val (used below for the row's colour) is untouched.
                status_display = STATUS_DISPLAY_LABELS.get(row["status"] or "Due", status_val)
                date_val   = parse_date(row["invoice_date"], DB_DATE_FORMAT, DISPLAY_DATE_FORMAT)
                total_txt  = "No Charge" if is_comp else fmt_money(row.get("grand_total"))
                paid_txt   = "-"         if is_comp else fmt_money(row.get("total_paid"))
                bal_txt    = "-"         if is_comp else fmt_money(row.get("balance_due"))
                note_text  = (row.get("notes") or "").strip()
                note_icon  = "📝" if note_text else ""

                values = (row["invoice_number"], row["bill_to"], date_val, status_display,
                          total_txt, paid_txt, bal_txt, note_icon)
                badge = STATUS_BADGE.get(status_val)
                brush = QBrush(QColor(badge["foreground"])) if badge and badge.get("foreground") else None
                for c, text in enumerate(values):
                    it = QTableWidgetItem(text)
                    it.setTextAlignment(self.COL_DEFS[c][3] | Qt.AlignmentFlag.AlignVCenter)
                    if brush is not None:
                        it.setForeground(brush)
                    if c == ncols - 1 and note_text:
                        it.setToolTip(note_text if len(note_text) <= 400 else note_text[:400] + "…")
                    elif c <= 1:
                        it.setToolTip(text)        # full invoice number / client name when the column is too narrow for it
                    t.setItem(r, c, it)
                self._row_index[row["invoice_number"]] = row

            # Restore selection (only rows on the current page can still exist)
            if keep_selected:
                for r in range(t.rowCount()):
                    if t.item(r, 0).text() in keep_selected:
                        t.setRangeSelected(QTableWidgetSelectionRange(r, 0, r, ncols - 1), True)
            self._fit_columns()                    # room for the longest date / status / amount now on this page
        finally:
            t.blockSignals(False)
            t.setUpdatesEnabled(True)
        t.verticalScrollBar().setValue(scroll_pos)

        self._update_pagination_controls(total, start, len(rows))
        self._on_select()  # Re-evaluate button states

    def _update_pagination_controls(self, total: int, start: int, shown: int):
        if total == 0:
            self.page_info_lbl.setText("No invoices found")
        else:
            self.page_info_lbl.setText(
                f"Showing {start + 1}–{start + shown} of {total}  •  "
                f"Page {self._page} of {self._total_pages}"
            )
        self.prev_page_btn.setEnabled(self._page > 1)
        self.next_page_btn.setEnabled(self._page < self._total_pages)

    def _update_stats(self):
        s = self.db_manager.get_dashboard_stats()
        self.s_billed.setText(fmt_money(s["total_billed"]))
        self.s_received.setText(fmt_money(s["total_paid"]))
        self.s_due.setText(fmt_money(s["balance_due"]))
        self.s_count.setText(str(s["total"]))

    # ------------------------------------------------------------------ #
    #  Single-invoice actions
    # ------------------------------------------------------------------ #
    def _act_view_pdf(self):
        if inv_id := self._selected_id():
            self.controller.show_progress(True, "Generating PDF…")
            threading.Thread(target=self._pdf_worker, args=(inv_id, self.controller.task_queue), daemon=True).start()

    def _act_edit(self):
        if inv_id := self._selected_id():
            self.controller.show_frame(InvoiceFormFrame, inv_id)

    def _act_edit_note(self):
        if not (inv_id := self._selected_id()):
            return
        current = (self._row_index.get(inv_id, {}) or {}).get("notes", "")
        run_dialog(NoteDialog(self.controller, self.db_manager, inv_id, current, self._populate))

    def _act_delete(self):
        """Single-invoice delete (the small Delete button / Delete key on the action bar). If the Delete
        key fires while multiple rows happen to be selected, this redirects to the bulk-safe path --
        deleting a handful of invoices synchronously is instant and doesn't need a progress dialog, but
        a large multi-select deletion benefits from the same progress/abort treatment as any other
        bulk operation (see _act_bulk_delete)."""
        ids = self._selected_ids()
        if not ids:
            return
        if len(ids) > 1:
            self._act_bulk_delete()
            return
        if Messagebox.okcancel("Delete Invoice", f"Permanently delete Invoice #{ids[0]}?\nThis cannot be undone.",
                               parent=self.controller):
            self.db_manager.delete_invoices(ids)
            self.on_show()

    def _act_duplicate(self):
        if not (inv_id := self._selected_id()):
            return
        row = self.db_manager.get_invoice_by_id(inv_id)
        if not row:
            return

        new_data = {k: row[k] for k in row.keys() if k not in
                    ("grand_total", "total_paid", "balance_due", "created_ts")}
        new_data["invoice_number"] = self.db_manager.get_next_invoice_number()
        new_data["invoice_date"]   = datetime.date.today().strftime(DB_DATE_FORMAT)
        new_data["status"]         = InvoiceStatus.DUE.value
        new_data["amount_paid"]    = 0.0
        new_data["advance_amount"] = 0.0
        new_data["payment_date"]   = None
        new_data["advance_date"]   = None
        if row["due_date"]:                       # the copy is due as long after its date as the original was
            try:
                span = (datetime.datetime.strptime(row["due_date"], DB_DATE_FORMAT)
                        - datetime.datetime.strptime(row["invoice_date"], DB_DATE_FORMAT))
                new_data["due_date"] = (datetime.date.today() + span).strftime(DB_DATE_FORMAT)
            except (TypeError, ValueError):
                new_data["due_date"] = None

        ok, err = self.db_manager.save_invoice(new_data, is_update=False, payments=[])
        if ok:
            self.on_show()
            Messagebox.show_info("Duplicated", f"Created {new_data['invoice_number']} as a copy.", parent=self.controller)
        else:
            Messagebox.show_error("Error", f"Could not duplicate:\n{err}", parent=self.controller)

    # ------------------------------------------------------------------ #
    #  Bulk operation concurrency guard
    # ------------------------------------------------------------------ #
    # Every button that kicks off a bulk-scale operation lives in self._bulk_buttons (built in
    # _action_bar() / _header()). _start_bulk_op()/_end_bulk_op() disable/re-enable all of them together,
    # so only one bulk operation can ever be in flight.
    def _set_bulk_buttons_enabled(self, enabled: bool):
        for btn in getattr(self, "_bulk_buttons", ()):
            btn.setEnabled(enabled)

    def _start_bulk_op(self):
        self._bulk_op_running = True
        self._set_bulk_buttons_enabled(False)

    def _end_bulk_op(self):
        self._bulk_op_running = False
        self._set_bulk_buttons_enabled(True)

    def _run_bulk_dialog(self, title: str, items: list, work_fn: Callable, on_done: Callable[[dict], None]):
        """Open the modal progress dialog for a bulk job. The operation lock is released no matter how the
        dialog ends (the Tk build only released it from on_done, so an error there would have left every
        bulk button disabled for the rest of the session)."""
        try:
            run_dialog(BulkProgressDialog(self.controller, title, items, work_fn=work_fn, on_done=on_done))
        finally:
            if self._bulk_op_running:
                self._end_bulk_op()

    # ------------------------------------------------------------------ #
    #  Bulk actions
    # ------------------------------------------------------------------ #
    def _act_bulk_delete(self):
        if self._bulk_op_running:
            return
        ids = self._selected_ids()
        if not ids:
            return
        if not Messagebox.okcancel("Bulk Delete", f"Permanently delete {len(ids)} invoices?\nThis cannot be undone.",
                                   parent=self.controller):
            return
        self._start_bulk_op()
        self._run_bulk_dialog(f"Deleting {len(ids)} Invoice(s)", ids,
                              self._bulk_delete_work, self._on_bulk_delete_done)

    def _bulk_delete_work(self, ids: List[str], report, is_aborted) -> dict:
        CHUNK = 200
        total = len(ids)
        success = failed = 0
        for i in range(0, total, CHUNK):
            if is_aborted():
                break
            chunk = ids[i:i + CHUNK]
            done = min(i + len(chunk), total)
            try:
                self.db_manager.delete_invoices(chunk, chunk_size=CHUNK)
                success += len(chunk)
                report(done, total, f"Deleting invoices… ({done} of {total})",
                       log_line=f"Deleted {len(chunk)} invoice(s) [batch {i // CHUNK + 1}]",
                       success_delta=len(chunk))
            except Exception as exc:
                failed += len(chunk)
                report(done, total, f"Error deleting batch {i // CHUNK + 1}",
                       log_line=f"FAILED batch {i // CHUNK + 1}: {exc}", failed_delta=len(chunk))
        aborted = is_aborted() and (success + failed) < total
        return {"success": success, "failed": failed, "aborted": aborted, "result": None}

    def _on_bulk_delete_done(self, result: dict):
        try:
            self.on_show()       # refresh the list/stats first...
        finally:
            self._end_bulk_op()  # ...then restore button state, so a stray _on_select() in between can't leave it stale

    def _bulk_export_work_factory(self, file_path: str, ids: Optional[List[str]]) -> Callable:
        """Builds a work_fn that streams matching invoices straight into a CSV file in chunks, writing
        incrementally instead of loading the whole result set into memory first -- the difference between
        "fine at a few hundred invoices" and "fine at 100,000+". Handles both an explicit selection (ids
        given) and "export everything" (ids=None, fed from DatabaseManager.iter_all_invoices_chunked) with
        the same progress-reporting shape either way."""
        def work(items, report, is_aborted):
            total = len(items)
            written = 0
            try:
                with open(file_path, "w", newline="", encoding="utf-8") as fh:
                    writer = None
                    if ids is not None:
                        CHUNK = 500
                        for i in range(0, len(ids), CHUNK):
                            if is_aborted():
                                break
                            chunk_rows = self.db_manager.get_invoices_by_ids(ids[i:i + CHUNK])
                            if writer is None and chunk_rows:
                                writer = csv.writer(fh)
                                writer.writerow(chunk_rows[0].keys())
                            if writer:
                                writer.writerows(chunk_rows)
                            written += len(chunk_rows)
                            report(written, total, f"Exporting… ({written} of {total})",
                                   log_line=f"Wrote {len(chunk_rows)} row(s)", success_delta=len(chunk_rows))
                    else:
                        keys = None
                        for chunk_rows in self.db_manager.iter_all_invoices_chunked():
                            if is_aborted():
                                break
                            if writer is None:
                                keys = [k for k in chunk_rows[0].keys() if k != "_rowid"]
                                writer = csv.writer(fh)
                                writer.writerow(keys)
                            writer.writerows([tuple(r[k] for k in keys) for r in chunk_rows])
                            written += len(chunk_rows)
                            report(written, total, f"Exporting… ({written} of {total})",
                                   log_line=f"Wrote {len(chunk_rows)} row(s)", success_delta=len(chunk_rows))
                aborted = is_aborted() and written < total
                if written == 0 and not aborted:
                    return {"success": 0, "failed": 0, "aborted": False, "result": None,
                            "error": "No invoices to export."}
                return {"success": written, "failed": 0, "aborted": aborted,
                        "result": None if aborted else file_path}
            except Exception as exc:
                return {"success": written, "failed": max(0, total - written), "aborted": False,
                        "result": None, "error": str(exc)}
        return work

    def _on_bulk_export_done(self, result: dict):
        self._end_bulk_op()
        if result.get("error"):
            Messagebox.show_error("Export Failed", result["error"], parent=self.controller)
        elif result.get("result"):
            Messagebox.show_info("Exported", f"Saved {result['success']} invoice(s) to:\n{result['result']}",
                                 parent=self.controller)

    def _ask_csv_path(self, title: str) -> str:
        """Save-as dialog for a CSV export. Adds ".csv" when the user typed a name without an extension
        (what Tk's defaultextension=".csv" did)."""
        path, _filter = QFileDialog.getSaveFileName(self.controller, title, "", "CSV Files (*.csv)")
        if path and not Path(path).suffix:
            path += ".csv"
        return path

    def _act_bulk_export(self):
        if self._bulk_op_running:
            return
        ids = self._selected_ids()
        if not ids:
            return
        path = self._ask_csv_path("Export Selected Invoices")
        if not path:
            return
        self._start_bulk_op()
        self._run_bulk_dialog(f"Exporting {len(ids)} Invoice(s)", ids,
                              self._bulk_export_work_factory(path, ids), self._on_bulk_export_done)

    def _bulk_print_work(self, ids: List[str], report, is_aborted) -> dict:
        """Generates a PDF per invoice and merges them into batches of at most PRINT_BATCH_SIZE each --
        printing 100 invoices produces one manageable merged PDF; printing 1,000 produces ten of them
        instead of a single unwieldy multi-thousand-page file. Each completed batch is flushed to disk as
        soon as it fills, so cancelling partway through still leaves whatever batches finished usable
        rather than discarding all progress."""
        total = len(ids)
        cfg = load_config()
        success = failed = 0
        current_batch: List[Path] = []
        batch_paths: List[Path] = []
        batch_num = 0
        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")

        def flush_batch():
            nonlocal batch_num
            if not current_batch:
                return
            batch_num += 1
            suffix = f"-batch{batch_num}" if total > PRINT_BATCH_SIZE else ""
            merged = merge_pdfs(current_batch, BATCH_DIR / f"Batch-Print-{stamp}{suffix}.pdf")
            batch_paths.append(merged)
            current_batch.clear()

        for idx, inv_id in enumerate(ids, start=1):
            if is_aborted():
                break
            try:
                row = self.db_manager.get_invoice_by_id(inv_id)
                if row is None:
                    raise ValueError("invoice no longer exists")
                payments = self.db_manager.get_payments(inv_id)
                pdf_path = PDFGenerator(dict(row), cfg, payments=payments).generate()
                current_batch.append(pdf_path)
                success += 1
                report(idx, total, f"Generating PDF {idx} of {total}…",
                       log_line=f"Generated {inv_id}", success_delta=1)
            except Exception as exc:
                failed += 1
                report(idx, total, f"Error generating {inv_id}",
                       log_line=f"FAILED {inv_id}: {exc}", failed_delta=1)
            if len(current_batch) >= PRINT_BATCH_SIZE:
                flush_batch()

        flush_batch()  # whatever's left: a partial final batch, or everything if <= PRINT_BATCH_SIZE total
        aborted = is_aborted() and (success + failed) < total
        return {"success": success, "failed": failed, "aborted": aborted, "result": batch_paths}

    def _on_bulk_print_done(self, result: dict, ids: Optional[List[str]] = None):
        self._end_bulk_op()
        paths = result.get("result") or []

        if len(paths) == 1:
            # If exactly one invoice was printed, fetch its data for the Share dialog
            invoice_data = None
            if ids is not None and len(ids) == 1:
                row = self.db_manager.get_invoice_by_id(ids[0])
                if row:
                    invoice_data = dict(row)
            PDFViewer(self.controller, paths[0], invoice_data=invoice_data).show()

        elif len(paths) > 1:
            Messagebox.show_info(
                "Batches Ready",
                f"Generated {len(paths)} batch PDF file(s) (up to {PRINT_BATCH_SIZE} invoices each) in:\n{BATCH_DIR}",
                parent=self.controller,
            )

    def _act_bulk_print(self):
        if self._bulk_op_running:
            return
        ids = self._selected_ids()
        if not ids:
            return
        self._start_bulk_op()
        self._run_bulk_dialog(f"Generating {len(ids)} PDF(s)", ids, self._bulk_print_work,
                              lambda result, _ids=list(ids): self._on_bulk_print_done(result, _ids))

    def _show_ctx_menu(self, pos):
        item = self.table.itemAt(pos)
        if item is not None and not self.table.item(item.row(), 0).isSelected():
            self.table.select_only_row(item.row())
        global_pos = self.table.viewport().mapToGlobal(pos)
        if len(self._selected_ids()) >= 2:
            self._ctx_bulk.exec(global_pos)
        elif item is not None:
            self._ctx_single.exec(global_pos)

    def _export_csv(self):
        """Header 'Export CSV' button -- exports every invoice, not just a selection. Routed through the
        same streaming/progress-dialog path as the bulk actions above (and gated by the same lock) since
        at 100,000+ invoices this is exactly as bulk-scale an operation as an explicit multi-select export."""
        if self._bulk_op_running:
            return
        path = self._ask_csv_path("Export All Invoices")
        if not path:
            return
        total_count = self.db_manager.get_dashboard_stats().get("total", 0)
        if not total_count:
            Messagebox.show_error("Nothing to Export", "There are no invoices to export yet.", parent=self.controller)
            return
        self._start_bulk_op()
        self._run_bulk_dialog(f"Exporting All {total_count} Invoices", list(range(total_count)),
                              self._bulk_export_work_factory(path, None), self._on_bulk_export_done)

    def _act_refresh(self):
        """Reloads the invoice list and summary stats from the database -- used by both the Refresh
        button and Ctrl+R. This is the same query _populate()/_update_stats() already run on every
        navigation; refresh just triggers it on demand (e.g. after editing data outside the app, or just
        to be sure the numbers are current) without changing which page or filter is active."""
        self._populate()
        self._update_stats()

    # ------------------------------------------------------------------ #
    #  Public API used by application-level keyboard shortcuts
    # ------------------------------------------------------------------ #
    def focus_search(self):
        """Ctrl+F: jump into this screen's search box, with any existing text selected so typing
        immediately replaces it."""
        self.search_entry.setFocus()
        self.search_entry.selectAll()

    def refresh(self):
        """Ctrl+R: reload the invoice list and stats, exactly like the Refresh button."""
        self._act_refresh()

    def select_all(self):
        """Ctrl+A: select every invoice on the currently-shown page of the list. Selection lives in the
        table widget itself, so this can only select rows actually on screen -- with pagination, that
        means the current page; switch pages first to select rows elsewhere."""
        self.table.selectAll()

    def prepare_selected_pdf(self):
        """Ctrl+P and Ctrl+E both start here: PDF generation is asynchronous, so both shortcuts generate
        (or re-use) the PDF and open it in the viewer -- the dedicated Print/Save buttons there (and their
        own Ctrl+P/Ctrl+E shortcuts) handle the actual print-or-save step. Mirrors the existing
        single-vs-bulk split already used by the right-click context menu."""
        ids = self._selected_ids()
        if len(ids) >= 2:
            self._act_bulk_print()
        elif ids:
            self._act_view_pdf()

    def duplicate_selected(self):
        """Ctrl+D: duplicate the single selected invoice."""
        self._act_duplicate()

    # ------------------------------------------------------------------ #
    #  Background workers
    # ------------------------------------------------------------------ #
    def _pdf_worker(self, invoice_id: str, q: queue.Queue):
        try:
            row = self.db_manager.get_invoice_by_id(invoice_id)
            if row is None:
                raise ValueError(f"Invoice #{invoice_id} no longer exists.")
            row = dict(row)
            payments = self.db_manager.get_payments(invoice_id)
            path = PDFGenerator(row, load_config(), payments=payments).generate()
            q.put({"status": "success", "kind": "single", "path": path, "invoice_data": row})
        except Exception as exc:
            logger.error(f"PDF generation failed ({invoice_id}): {exc}", exc_info=True)
            q.put({"status": "error", "message": str(exc)})


# -----------------------------------------------------------------------------
#  Invoice form
# -----------------------------------------------------------------------------
# Column widths shared by the item / payment headers and their rows, so the columns always line up.
_ROW_GAP   = 8
_AMOUNT_W  = 112      # an item's rate / a payment's amount
_DISC_W    = 92
_QTY_W     = 72
_UNIT_W    = 100
_LINE_W    = 116      # an item's computed amount
_HSN_W     = 118
_GST_W     = 92
_DATE_W    = 138
_MODE_W    = 140
_REMOVE_W  = 34


def _num_text(value: Any) -> str:
    """A stored number as the plain text a field shows ("18", not "18.0")."""
    text = repr(safe_float(value))
    return text[:-2] if text.endswith(".0") else text


class ToggleSwitch(QCheckBox):
    """A round on/off switch (ttkbootstrap's "round-toggle"). It is still a QCheckBox -- Space toggles it,
    `clicked` / `toggled` work as usual, screen readers see a check box -- but painted with the live theme
    colours, so it follows theme switches with no extra code."""
    TRACK_W, TRACK_H, GAP = 38, 20, 10

    def __init__(self, text: str = "", parent: Optional[QWidget] = None):
        super().__init__(text, parent)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def sizeHint(self) -> QSize:
        self.ensurePolished()
        fm = self.fontMetrics()
        text_w = fm.horizontalAdvance(self.text()) + self.GAP if self.text() else 0
        return QSize(self.TRACK_W + 4 + text_w, max(self.TRACK_H + 8, fm.height() + 8))

    def minimumSizeHint(self) -> QSize:
        return self.sizeHint()

    def hitButton(self, pos) -> bool:                      # the caption toggles it too, not only the switch
        return self.rect().contains(pos)

    def paintEvent(self, _event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        on, enabled = self.isChecked(), self.isEnabled()
        if not enabled:
            track, knob, text = theme_color("border"), theme_color("surface"), theme_color("disabled_fg")
        elif on:
            track, knob, text = theme_color("primary"), "#FFFFFF", theme_color("fg")     # a switch knob is white on every theme
        else:
            track = theme_color("border_strong")
            knob = theme_color("fg") if theme_is_dark() else theme_color("surface")
            text = theme_color("fg")
        w, h = self.TRACK_W, self.TRACK_H
        top = (self.height() - h) / 2
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(track))
        painter.drawRoundedRect(QRectF(2, top, w, h), h / 2, h / 2)
        d = h - 6
        painter.setBrush(QColor(knob))
        painter.drawEllipse(QRectF(2 + (w - d - 3 if on else 3), top + 3, d, d))
        if self.hasFocus() and enabled:
            painter.setPen(QPen(QColor(theme_color("primary")), 1.5))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRoundedRect(QRectF(0.5, top - 1.5, w + 3, h + 3), h / 2 + 1.5, h / 2 + 1.5)
        if self.text():
            painter.setPen(QColor(text))
            painter.drawText(QRectF(self.TRACK_W + 4 + self.GAP, 0, self.width() - self.TRACK_W - 4 - self.GAP, self.height()),
                             int(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft), self.text())


class _ChevronButton(QToolButton):
    """The little ▾ at the right end of a SuggestCombo, painted in the live theme colours."""

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setFixedSize(22, 22)
        self.setCursor(Qt.CursorShape.ArrowCursor)
        self.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.setToolTip("Show suggestions")

    def paintEvent(self, _event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        key = "disabled_fg" if not self.isEnabled() else ("fg" if self.underMouse() else "muted_text")
        painter.setPen(QPen(QColor(theme_color(key)), 1.6, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap,
                            Qt.PenJoinStyle.RoundJoin))
        cx, cy = self.width() / 2, self.height() / 2
        painter.drawPolyline([QPointF(cx - 4, cy - 2), QPointF(cx, cy + 2), QPointF(cx + 4, cy - 2)])


class SuggestCombo(QLineEdit):
    """The free-typed picker used for the client name and the item title: type anything, or choose a known
    value. Typing pops up the known values that *contain* what was typed (case-insensitively); the ▾ button (or
    the Down key) opens the list. `picked(text)` fires when -- and only when -- the user chooses an entry from
    that list; just typing a value never fires it, exactly like the Tk build's <<ComboboxSelected>>.

    It is a QLineEdit + QCompleter rather than an editable QComboBox on purpose: an editable combo box also
    "activates" an item whenever the typed text happens to match one (on Enter and on losing focus), which
    would auto-fill / overwrite the client details behind the user's back."""
    picked = Signal(str)

    def __init__(self, placeholder: str = "", values: Optional[List[str]] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        if placeholder:
            self.setPlaceholderText(placeholder)
        self._model = QStringListModel(list(values or []), self)
        self._completer = QCompleter(self._model, self)
        self._completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self._completer.setFilterMode(Qt.MatchFlag.MatchContains)
        self._completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        self._completer.setMaxVisibleItems(8)
        self.setCompleter(self._completer)
        self._completer.activated[str].connect(self._on_completer_activated)
        self._chevron = _ChevronButton(self)
        self._chevron.clicked.connect(self._show_all)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 3, 0)
        lay.setSpacing(0)
        lay.addStretch()
        lay.addWidget(self._chevron)
        self.setTextMargins(0, 0, 22, 0)

    def _on_completer_activated(self, text: str) -> None:
        # QLineEdit wires "suggestion chosen -> setText" only while it has keyboard focus; doing it here as well
        # makes a pick always land, whatever the focus state is at that instant.
        self.setText(text)
        self.picked.emit(text)

    def set_values(self, values: List[str]) -> None:
        """Replace the list of known values (what has been typed so far is left alone)."""
        self._model.setStringList(list(values))

    def values(self) -> List[str]:
        return self._model.stringList()

    def set_text(self, text: str) -> None:
        self.setText(text)
        self.setCursorPosition(0)

    def _show_all(self) -> None:
        """The ▾ button: open the list -- every known value that contains the text typed so far (all of
        them while the field is empty), like the Tk dropdown filtered by the typed text."""
        self.setFocus()
        self._completer.setCompletionPrefix(self.text())
        self._completer.complete()


class _ReturnToNext(QObject):
    """Event filter: Enter / Return in a single-line field moves focus to the next field in tab order -- the
    same field-to-field flow Tab already gives, from the key people naturally press to "confirm". Applied to
    single-line fields only: the multi-line Contact Info / Notes boxes keep plain Enter for a new line."""

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.KeyPress and event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            blocked = (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.AltModifier
                       | Qt.KeyboardModifier.ShiftModifier | Qt.KeyboardModifier.MetaModifier)
            if event.modifiers() & blocked:
                return False
            completer = obj.completer() if isinstance(obj, (QLineEdit, QComboBox)) else None
            if completer is not None and completer.popup().isVisible() and completer.popup().currentIndex().isValid():
                return False                                   # a highlighted suggestion: the completer takes the key
            # A QDateEdit's own focusNextChild() first steps through its day / month / year sections; asking its
            # parent instead moves straight on to the next field, like Enter in the Tk date entry did.
            target = obj.parentWidget() if isinstance(obj, QAbstractSpinBox) and obj.parentWidget() is not None else obj
            target.focusNextChild()
            return True
        return False


def advance_on_enter(widget: QWidget) -> QWidget:
    widget.installEventFilter(_ReturnToNext(widget))
    return widget


def make_money_edit(value: float, width: int, tooltip: str = "", advance: bool = True) -> QLineEdit:
    """A right-aligned "0.00" amount field: digits and one decimal point only (two decimals at most),
    re-formatted to two decimals when it loses focus."""
    edit = QLineEdit(f"{value:.2f}")
    edit.setAlignment(Qt.AlignmentFlag.AlignRight)
    edit.setFixedWidth(width)
    restrict_to_decimal(edit)
    format_as_decimal_on_blur(edit)
    if advance:
        advance_on_enter(edit)
    if tooltip:
        edit.setToolTip(tooltip)
    return edit


def _caption(text: str) -> QLabel:
    return make_label(text, "caption")


def _retone(widget: QWidget, tone: str) -> None:
    """set_tone() only when the tone actually changes (re-polishing on every keystroke is wasted work)."""
    if widget.property("tone") != tone:
        set_tone(widget, tone)


class _QtyBlurFilter(QObject):
    """A quantity field tidies itself when it loses focus: "2.500" -> "2.5", an empty field -> "1"."""

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.FocusOut:
            text = obj.text().strip()
            tidy = fmt_qty(text) if text not in ("", ".") else "1"
            if tidy != obj.text():
                obj.setText(tidy)
        return False


class _DiscountBlurFilter(QObject):
    """A discount field ("5", "5.5" or "10%") tidies itself when it loses focus: "5" -> "5.00", "10.50%" -> "10.5%"."""

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.FocusOut:
            text = obj.text().strip()
            number = safe_float(text.rstrip("%"))
            tidy = f"{number:g}%" if text.endswith("%") and number > 0 else f"{number:.2f}"
            if tidy != obj.text():
                obj.setText(tidy)
        return False


def make_qty_edit(value: float) -> QLineEdit:
    """A right-aligned quantity field: a plain number with up to three decimals ("2", "0.5", "12.75")."""
    edit = QLineEdit(fmt_qty(value))
    edit.setAlignment(Qt.AlignmentFlag.AlignRight)
    edit.setFixedWidth(_QTY_W)
    edit.setMaxLength(11)
    restrict_to_decimal(edit, 3)
    edit.installEventFilter(_QtyBlurFilter(edit))
    advance_on_enter(edit)
    edit.setToolTip("Quantity (e.g. 2, 0.5, 12.75)")
    return edit


def make_discount_edit(text: str) -> QLineEdit:
    """A discount field: an amount ("150") or a percentage of the line ("10%")."""
    edit = QLineEdit(text)
    edit.setAlignment(Qt.AlignmentFlag.AlignRight)
    edit.setFixedWidth(_DISC_W)
    edit.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d{0,9}(\.\d{0,2})?%?$"), edit))
    edit.installEventFilter(_DiscountBlurFilter(edit))
    advance_on_enter(edit)
    edit.setToolTip("Per-item discount: an amount (150) or a percentage of the line (10%)")
    return edit


def make_editable_combo(presets: List[str], width: int, placeholder: str, max_length: int,
                        advance: bool = True) -> QComboBox:
    """A combo box in which anything can be typed (the list is only a convenience): unit and GST-rate fields."""
    combo = QComboBox()
    combo.setEditable(True)
    combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
    combo.setCompleter(None)                      # what is typed is what is meant -- no auto-completion surprises
    combo.addItems(presets)
    combo.setCurrentIndex(-1)
    combo.lineEdit().setPlaceholderText(placeholder)
    combo.lineEdit().setMaxLength(max_length)
    combo.setFixedWidth(width)
    guard_wheel(combo)
    if advance:
        advance_on_enter(combo)        # the combo itself holds the focus (its line edit's focus proxy)
    return combo


class _EnterAddsRow(QObject):
    """Enter in the last field of the last, filled-in item row starts the next row -- the way a grid is filled in."""

    def __init__(self, row: "_ItemRow", form: "InvoiceFormFrame"):
        super().__init__(row)
        self.row, self.form = row, form

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.KeyPress and event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter) \
                and not (event.modifiers() & (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.AltModifier
                                              | Qt.KeyboardModifier.ShiftModifier | Qt.KeyboardModifier.MetaModifier)):
            rows = self.form.item_rows
            if rows and rows[-1] is self.row and obj is self.row.last_field() and self.row.is_filled():
                self.form.add_item()
                return True
        return False


class _ItemRow(QWidget):
    """One invoice line, on two rows of fields.

        item / service name | qty | unit | rate | discount | line amount | x
        description                              | HSN/SAC | GST %   (GST % only when GST is charged per item)

    The old "amount" field is now the RATE (`amount_edit`): with the quantity left at 1 it is the line's amount, exactly
    as before. The stored line keeps its old keys (title, description, amount = quantity x rate, discount) and only
    gains qty / unit / rate / hsn / tax_rate / disc_type / disc_value when the line uses them, so a plain line is saved
    exactly like one from an older build."""
    changed = Signal()                       # a figure was edited -> the totals must be recalculated
    remove_requested = Signal(object)        # this row asks the form to delete it
    title_picked = Signal(object, str)       # (row, title) -- a catalog title was chosen
    title_edited = Signal(object, str)       # (row, title) -- the title text changed (any way)

    def __init__(self, titles: List[str], title: str = "", description: str = "", amount: float = 0.0,
                 discount: float = 0.0, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._per_item = False               # GST % shown / stored only when the invoice charges GST per item
        self._gst_touched = False            # the user (or a catalog item) chose this line's rate: stop following the default
        self._default_gst = "18"             # the invoice's default rate -- what an emptied GST field means
        self._unit_rates: Dict[str, float] = {}      # the picked catalog item's units -> rate per unit (lower-case keys)
        self._dropdown_units: List[str] = []         # what the unit list currently offers besides the presets
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 7, 0, 0)
        lay.setSpacing(5)

        top = QHBoxLayout()
        top.setSpacing(_ROW_GAP)
        self.title_edit = SuggestCombo("Item / service name", titles)
        self.title_edit.set_text(title)
        self.title_edit.picked.connect(self._on_picked)
        self.title_edit.textChanged.connect(self._on_title_text)
        advance_on_enter(self.title_edit)
        top.addWidget(self.title_edit, 1)

        self.qty_edit = make_qty_edit(1.0)
        self.qty_edit.textChanged.connect(self._on_edited)
        top.addWidget(self.qty_edit)
        self.unit_combo = make_editable_combo(UNIT_PRESETS, _UNIT_W, "Unit", 14)
        self.unit_combo.setToolTip("Unit of measure: pick one or type your own (Nos, Day, Plate, Sq.ft ...)")
        self.unit_combo.activated.connect(self._on_unit_activated)
        top.addWidget(self.unit_combo)
        self.amount_edit = make_money_edit(amount, _AMOUNT_W, "Rate: the price of one unit (before any discount)")
        self.amount_edit.textChanged.connect(self._on_edited)
        top.addWidget(self.amount_edit)
        self.disc_edit = make_discount_edit(f"{discount:.2f}")
        self.disc_edit.textChanged.connect(self._on_edited)
        top.addWidget(self.disc_edit)
        self.line_lbl = make_label("", "muted", align=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.line_lbl.setFixedWidth(_LINE_W)
        self.line_lbl.setToolTip("Quantity x rate, less this line's discount")
        top.addWidget(self.line_lbl)

        self.remove_btn = make_button("✕", "danger-outline", tooltip="Remove this item (Ctrl+Shift+Del or Ctrl+Minus)")
        self.remove_btn.setFixedWidth(_REMOVE_W)
        set_flag(self.remove_btn, "glyph", True)
        self.remove_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)      # Tab skips it, as in the Tk build (takefocus=False)
        self.remove_btn.clicked.connect(self._on_remove)
        top.addWidget(self.remove_btn)
        lay.addLayout(top)

        second = QHBoxLayout()
        second.setSpacing(_ROW_GAP)
        self.desc_edit = QLineEdit(description)
        self.desc_edit.setPlaceholderText("Description (optional)")
        self.desc_edit.setCursorPosition(0)
        advance_on_enter(self.desc_edit)
        second.addWidget(self.desc_edit, 1)
        self.hsn_caption = make_label("HSN/SAC", "muted")
        second.addWidget(self.hsn_caption)
        self.hsn_edit = QLineEdit()
        self.hsn_edit.setFixedWidth(_HSN_W)
        self.hsn_edit.setPlaceholderText("optional")
        self.hsn_edit.setToolTip("HSN code (goods) or SAC code (services): 2 to 8 digits, printed under the item")
        restrict_to_digits(self.hsn_edit, 8)
        advance_on_enter(self.hsn_edit)
        second.addWidget(self.hsn_edit)
        self.gst_caption = make_label("GST %", "muted")
        second.addWidget(self.gst_caption)
        self.gst_combo = make_editable_combo(GST_RATE_PRESETS, _GST_W, "%", 6)
        self.gst_combo.lineEdit().setValidator(_DecimalValidator(2, self.gst_combo))
        self.gst_combo.setToolTip("GST rate of this item (pick one or type your own)")
        self.gst_combo.setCurrentText("18")
        self.gst_combo.lineEdit().textEdited.connect(self._on_gst_touched)
        self.gst_combo.activated.connect(self._on_gst_touched)
        self.gst_combo.currentTextChanged.connect(self._on_edited)
        second.addWidget(self.gst_combo)
        second.addSpacing(_REMOVE_W)         # line the right edge up with the line above
        lay.addLayout(second)
        lay.addSpacing(4)
        lay.addWidget(hline())
        self.set_per_item(False)
        self._refresh_line()

    # ---- signals ------------------------------------------------------------ #
    def _on_picked(self, text: str):
        self.title_picked.emit(self, text)

    def _on_title_text(self, text: str):
        self.title_edited.emit(self, text)

    def _on_remove(self, _checked: bool = False):
        self.remove_requested.emit(self)

    def _on_edited(self, *_):
        self._refresh_line()
        self.changed.emit()

    def _on_gst_touched(self, *_):
        self._gst_touched = True

    def _on_unit_activated(self, index: int):
        """A unit chosen from the list loads that unit's saved rate when the catalog has one for this item."""
        rate = self._unit_rates.get(self.unit_combo.itemText(index).strip().casefold())
        if rate is not None:
            self.amount_edit.setText(f"{rate:.2f}")

    # ---- what the row shows -------------------------------------------------- #
    def _numbers(self) -> Tuple[float, float, float, str, float]:
        """(quantity, rate, line amount, discount kind, discount typed) -- the amount is quantity x rate to the paisa."""
        qty, rate = safe_float(self.qty_edit.text()), safe_float(self.amount_edit.text())
        gross = float(_round2(Decimal(str(qty)) * Decimal(str(rate))))
        text = self.disc_edit.text().strip()
        typed = safe_float(text.rstrip("%"))
        return qty, rate, gross, ("percent" if text.endswith("%") else "amount"), typed

    def _discount(self) -> Tuple[float, str, float]:
        """(discount amount, kind, the number typed) for the current line."""
        _qty, _rate, gross, kind, typed = self._numbers()
        if kind == "percent":
            return float(_round2(Decimal(str(gross)) * Decimal(str(typed)) / Decimal("100"))), kind, typed
        return max(0.0, typed), kind, typed

    def _refresh_line(self):
        _qty, _rate, gross, _kind, _typed = self._numbers()
        self.line_lbl.setText(fmt_money(max(0.0, gross - self._discount()[0])))

    def refresh_line(self) -> None:
        """Redraw the line amount (the currency symbol may have changed in Settings)."""
        self._refresh_line()

    def set_per_item(self, per_item: bool) -> None:
        """Show / hide this line's own GST rate (only an invoice that charges GST per item has one)."""
        self._per_item = per_item
        self.gst_caption.setVisible(per_item)
        self.gst_combo.setVisible(per_item)

    def set_default_gst(self, text: str) -> None:
        """The invoice's default GST rate changed: a line whose rate was never chosen follows it."""
        self._default_gst = text
        if not self._gst_touched and self.gst_combo.currentText() != text:
            self.gst_combo.setCurrentText(text)

    def set_gst(self, text: str) -> None:
        """Give this line its own GST rate (a catalog item's, say) -- it no longer follows the invoice's default."""
        self._gst_touched = True
        self.gst_combo.setCurrentText(text)

    def set_units(self, units: List[Tuple[str, float]]) -> None:
        """Offer the units a catalog item is sold in (first), then the usual ones; remember each unit's rate."""
        self._unit_rates = {u.casefold(): r for u, r in units}
        wanted = [u for u, _ in units]
        if wanted == self._dropdown_units:
            return
        self._dropdown_units = wanted
        current = self.unit_combo.currentText()
        self.unit_combo.blockSignals(True)
        self.unit_combo.clear()
        self.unit_combo.addItems(wanted + [u for u in UNIT_PRESETS if u.casefold() not in {w.casefold() for w in wanted}])
        self.unit_combo.setCurrentIndex(-1)
        self.unit_combo.setCurrentText(current)
        self.unit_combo.blockSignals(False)

    def apply_catalog(self, entry: Dict[str, Any]) -> None:
        """Fill the line from a catalog item: rate, unit, HSN/SAC and GST rate (a quantity already typed stays)."""
        self.set_units(catalog_units(entry))
        self.amount_edit.setText(f"{safe_float(entry.get('amount')):.2f}")
        if unit := str(entry.get("unit") or "").strip():
            self.unit_combo.setCurrentText(unit)
        if hsn := str(entry.get("hsn") or "").strip():
            self.hsn_edit.setText(hsn)
        if entry.get("tax_rate") is not None:
            self.set_gst(_num_text(entry["tax_rate"]))

    def load_item(self, item: Dict[str, Any]) -> None:
        """Fill the line from a stored item (older invoices simply have no quantity / unit / HSN / rate)."""
        self.title_edit.set_text(str(item.get("title") or ""))
        self.desc_edit.setText(str(item.get("description") or ""))
        self.desc_edit.setCursorPosition(0)
        qty = item_qty(item)
        raw_rate = item.get("rate")
        gross = safe_float(item.get("amount"))
        rate = safe_float(raw_rate) if raw_rate not in (None, "") else (gross / qty if qty else gross)
        self.qty_edit.setText(fmt_qty(qty))
        self.unit_combo.setCurrentText(str(item.get("unit") or ""))
        self.amount_edit.setText(f"{rate:.2f}")
        if item.get("disc_type") == "percent":
            self.disc_edit.setText(f"{safe_float(item.get('disc_value')):g}%")
        else:
            self.disc_edit.setText(f"{max(0.0, safe_float(item.get('discount'))):.2f}")
        self.hsn_edit.setText(str(item.get("hsn") or ""))
        if item.get("tax_rate") not in (None, ""):
            self.set_gst(_num_text(item["tax_rate"]))

    def is_filled(self) -> bool:
        """Whether anything has been entered on this line (an untouched row is ignored when the invoice is saved)."""
        return bool(self.title_edit.text().strip() or self.desc_edit.text().strip() or self.hsn_edit.text().strip()
                    or safe_float(self.amount_edit.text()) > 0)

    def last_field(self) -> QWidget:
        return self.gst_combo if self._per_item else self.hsn_edit

    def raw_gst(self) -> float:
        """This line's GST rate as typed (an empty field means the invoice's default), before any clamping."""
        return safe_float(self.gst_combo.currentText().strip() or self._default_gst)

    def focus_widgets(self) -> List[QWidget]:
        chain = [self.title_edit, self.qty_edit, self.unit_combo, self.amount_edit, self.disc_edit, self.desc_edit, self.hsn_edit]
        if self._per_item:
            chain.append(self.gst_combo)
        return chain

    def values(self) -> Dict[str, Any]:
        qty, rate, gross, kind, typed = self._numbers()
        discount = self._discount()[0]
        item: Dict[str, Any] = {
            "title": self.title_edit.text().strip(),
            "description": self.desc_edit.text().strip(),
            "amount": gross,
            "discount": discount,
        }
        unit, hsn = self.unit_combo.currentText().strip(), self.hsn_edit.text().strip()
        tax_rate = (min(100.0, safe_float(self.gst_combo.currentText().strip() or self._default_gst))
                    if self._per_item else None)
        if qty != 1.0 or unit or hsn or tax_rate is not None:           # more than a title and an amount: keep the detail
            item["qty"], item["rate"] = qty, rate
            if unit:
                item["unit"] = unit
            if hsn:
                item["hsn"] = hsn
            if tax_rate is not None:
                item["tax_rate"] = tax_rate
        if kind == "percent":
            item["disc_type"], item["disc_value"] = "percent", typed
        return item


class _PaymentRow(QWidget):
    """One line of the payment ledger: date, amount, mode, note."""
    changed = Signal()
    remove_requested = Signal(object)

    def __init__(self, amount: float = 0.0, payment_date: Optional[str] = None, mode: str = "Cash", note: str = "",
                 payment_id: Optional[int] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.payment_id = payment_id
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 6, 0, 0)
        lay.setSpacing(6)
        row = QHBoxLayout()
        row.setSpacing(_ROW_GAP)

        self.date_edit = make_date_edit()
        if payment_date:
            set_date_edit_from_db(self.date_edit, payment_date)
        self.date_edit.setFixedWidth(_DATE_W)
        advance_on_enter(self.date_edit)
        row.addWidget(self.date_edit)

        self.amount_edit = make_money_edit(amount, _AMOUNT_W)
        self.amount_edit.textChanged.connect(self.changed)
        row.addWidget(self.amount_edit)

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(PAYMENT_MODES)
        self.mode_combo.setCurrentText(mode if mode in PAYMENT_MODES else "Cash")
        self.mode_combo.setFixedWidth(_MODE_W)
        guard_wheel(self.mode_combo)
        row.addWidget(self.mode_combo)

        self.note_edit = QLineEdit(note)
        self.note_edit.setPlaceholderText("Note (optional)")
        self.note_edit.setCursorPosition(0)
        advance_on_enter(self.note_edit)
        row.addWidget(self.note_edit, 1)

        self.remove_btn = make_button("✕", "danger-outline", tooltip="Remove this payment")
        self.remove_btn.setFixedWidth(_REMOVE_W)
        set_flag(self.remove_btn, "glyph", True)
        self.remove_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.remove_btn.clicked.connect(self._on_remove)
        row.addWidget(self.remove_btn)
        lay.addLayout(row)
        lay.addSpacing(4)
        lay.addWidget(hline())

    def _on_remove(self, _checked: bool = False):
        self.remove_requested.emit(self)

    def focus_widgets(self) -> List[QWidget]:
        return [self.date_edit, self.amount_edit, self.mode_combo, self.note_edit]

    def set_editable(self, editable: bool) -> None:
        for w in (self.date_edit, self.amount_edit, self.mode_combo, self.note_edit, self.remove_btn):
            w.setEnabled(editable)

    def values(self) -> Dict[str, Any]:
        return {
            "amount": safe_float(self.amount_edit.text()),
            "payment_date": date_edit_db(self.date_edit),
            "mode": self.mode_combo.currentText(),
            "note": self.note_edit.text().strip(),
            "payment_id": self.payment_id,
        }


class _SummaryRow(QWidget):
    """One "label ............ value" line of the live totals box."""

    def __init__(self, label: str = "", bold: bool = False, parent: Optional[QWidget] = None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 2, 0, 2)
        lay.setSpacing(15)
        role = "h3" if bold else None
        align = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        self.label = make_label(label, role, align=align)
        self.value = make_label("", role, align=align)
        self.label.setMinimumWidth(190)
        self.value.setMinimumWidth(130)
        lay.addWidget(self.label)
        lay.addWidget(self.value)

    def set(self, value: Optional[str] = None, label: Optional[str] = None) -> None:
        if value is not None:
            self.value.setText(value)
        if label is not None:
            self.label.setText(label)


class InvoiceFormFrame(QWidget):
    """Create / edit an invoice: billing info, items, tax & discount, and a
    payment history ledger that supports multiple partial payments."""

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller  = controller
        self.db_manager  = controller.db_manager
        self.editing_id: Optional[str] = None
        self.item_rows:    List[_ItemRow] = []
        self.payment_rows: List[_PaymentRow] = []
        self._catalog_cache: List[Dict] = []
        self._client_cache:  List[Dict] = []
        self._initial_snapshot: Optional[str] = None
        self._catalog_exact: Dict[str, Dict] = {}     # catalog items by title (exact spelling / ignoring case)
        self._catalog_fold:  Dict[str, Dict] = {}
        self._place_auto = False         # the place of supply was filled in from the customer's GSTIN, not typed
        self._due_days: Optional[int] = None       # the due date was set as "n days after the invoice date" (None: by hand)
        self._setting_due = False
        self._loading = False            # True while the fields are being (re)filled: no live recalculation, no "dirty" checks
        self._build()

    # ------------------------------------------------------------------ #
    #  Layout
    # ------------------------------------------------------------------ #
    def _build(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        header = QHBoxLayout()
        header.setContentsMargins(24, 16, 24, 6)
        header.setSpacing(20)
        self.top_back_btn = make_button("←  Back to Dashboard  (Esc)", "secondary-outline", self._go_back,
                                        tooltip="Return to the invoice list (Esc). Unsaved changes will ask first.",
                                        min_height=36)
        header.addWidget(self.top_back_btn)
        self.title_lbl = make_label("New Invoice", "h1")
        header.addWidget(self.title_lbl)
        self.status_badge = make_label("", "badge", "secondary")
        header.addWidget(self.status_badge)
        header.addStretch()
        outer.addLayout(header)

        # The page scrolls as a whole (Tk: ScrolledFrame). Up/Down/PageUp/PageDown scroll it natively whenever focus
        # is in a single-line field; combo boxes, text boxes and the suggestion list keep those keys for themselves.
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.scroll.verticalScrollBar().setSingleStep(48)
        outer.addWidget(self.scroll, 1)

        body = QWidget()
        self.body_layout = QVBoxLayout(body)
        self.body_layout.setContentsMargins(24, 6, 24, 12)
        self.body_layout.setSpacing(16)
        self.scroll.setWidget(body)

        self._section_billing(self.body_layout)
        self._section_details(self.body_layout)
        self._section_items(self.body_layout)
        self._section_payments(self.body_layout)
        self._section_print_notes(self.body_layout)
        self._section_notes(self.body_layout)
        self._section_totals(self.body_layout)
        self._section_save_bar(self.body_layout)
        self.body_layout.addStretch()
        self._restore_tab_order()

    @staticmethod
    def _card(title: str) -> Tuple[QFrame, QVBoxLayout]:
        frame = QFrame()
        frame.setProperty("role", "card")
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(16, 14, 16, 16)
        lay.setSpacing(0)
        lay.addWidget(make_label(title, "h2"))
        lay.addSpacing(10)
        return frame, lay

    @staticmethod
    def _labeled(text: str, widget: QWidget, label_width: int = 140) -> QWidget:
        """A "label | field" row (Tk: a fixed 20-character label column)."""
        row = QWidget()
        lay = QHBoxLayout(row)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(10)
        label = QLabel(text)
        label.setFixedWidth(label_width)
        lay.addWidget(label)
        lay.addWidget(widget, 1)
        row.label = label                      # so a caption can be changed later (e.g. "Default GST Rate")
        return row

    @staticmethod
    def _text_box(placeholder: str) -> PlaceholderText:
        box = PlaceholderText(placeholder=placeholder, height=3)
        box.setAcceptRichText(False)
        box.setTabChangesFocus(True)          # Tab moves on (a literal tab character makes no sense in an address)
        return box

    # ---- Billing (Bill To / From) ------------------------------------ #
    def _section_billing(self, parent: QVBoxLayout):
        row = QHBoxLayout()
        row.setSpacing(16)

        left, lay = self._card("Bill To")
        lay.addWidget(_caption("Client / Company Name *"))
        lay.addSpacing(2)
        self.client_edit = SuggestCombo("Enter client name")
        self.client_edit.picked.connect(self._on_client_picked)
        advance_on_enter(self.client_edit)
        lay.addWidget(self.client_edit)
        lay.addSpacing(8)
        lay.addWidget(_caption("Contact Info"))
        lay.addSpacing(2)
        self.client_contact_text = self._text_box("Phone, email, address...")
        lay.addWidget(self.client_contact_text)
        lay.addSpacing(8)
        lay.addWidget(_caption("GSTIN"))
        lay.addSpacing(2)
        self.client_gstin_entry = QLineEdit()
        self.client_gstin_entry.setPlaceholderText("22ABCDE1234F1Z5 (optional)")
        advance_on_enter(self.client_gstin_entry)
        self.client_gstin_entry.textChanged.connect(self._autofill_place_from_gstin)
        lay.addWidget(self.client_gstin_entry)
        lay.addStretch()
        row.addWidget(left, 1)

        right, lay = self._card("From")
        lay.addWidget(_caption("Your Name / Firm *"))
        lay.addSpacing(2)
        self.your_name_entry = QLineEdit()
        self.your_name_entry.setPlaceholderText("Your business / firm name")
        advance_on_enter(self.your_name_entry)
        lay.addWidget(self.your_name_entry)
        lay.addSpacing(8)
        lay.addWidget(_caption("Contact Info"))
        lay.addSpacing(2)
        self.your_contact_text = self._text_box("Phone, email, address...")
        lay.addWidget(self.your_contact_text)
        lay.addStretch()
        row.addWidget(right, 1)
        parent.addLayout(row)

    # ---- Invoice Details / Tax & Payment Settings ---------------------- #
    def _section_details(self, parent: QVBoxLayout):
        row = QHBoxLayout()
        row.setSpacing(16)

        left, lay = self._card("Invoice Details")
        lay.setSpacing(8)
        self.invoice_number_entry = QLineEdit()
        self.invoice_number_entry.setPlaceholderText("INV-001")
        advance_on_enter(self.invoice_number_entry)
        lay.addWidget(self._labeled("Invoice Number *", self.invoice_number_entry))

        self.invoice_date_edit = make_date_edit()
        self.invoice_date_edit.dateChanged.connect(self._on_invoice_date_changed)
        advance_on_enter(self.invoice_date_edit)
        lay.addWidget(self._labeled("Invoice Date *", self.invoice_date_edit))

        self.due_combo = QComboBox()
        for text, days in DUE_CHOICES:
            self.due_combo.addItem(text, days)
        self.due_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.due_combo.setMinimumContentsLength(10)
        self.due_combo.setToolTip("When payment is due: a number of days after the invoice date, or a date you pick.\n"
                                  "The due date is printed on the invoice.")
        self.due_combo.activated.connect(self._on_due_chosen)
        guard_wheel(self.due_combo)
        lay.addWidget(self._labeled("Due Date", self.due_combo))
        self.due_date_edit = make_date_edit()
        self.due_date_edit.setEnabled(False)
        self.due_date_edit.dateChanged.connect(self._on_due_date_edited)
        advance_on_enter(self.due_date_edit)
        self._due_date_row = self._labeled("Due On", self.due_date_edit)       # shown only while there is a due date
        self._due_date_row.setVisible(False)
        lay.addWidget(self._due_date_row)

        self.payment_terms_entry = QLineEdit()
        self.payment_terms_entry.setPlaceholderText("e.g. Due on Receipt, Net 15")
        advance_on_enter(self.payment_terms_entry)
        lay.addWidget(self._labeled("Payment Terms", self.payment_terms_entry))

        self.po_entry = QLineEdit()
        self.po_entry.setMaxLength(60)
        self.po_entry.setPlaceholderText("Your customer's PO / reference number (optional)")
        advance_on_enter(self.po_entry)
        lay.addWidget(self._labeled("PO / Reference No.", self.po_entry))

        self.template_combo = QComboBox()
        self.template_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.template_combo.setMinimumContentsLength(10)
        self.template_combo.setToolTip("The look this invoice is printed in.\n\"Default\" is the template picked in "
                                       "Settings > Payment & PDF.")
        guard_wheel(self.template_combo)
        self._fill_template_combo()
        lay.addWidget(self._labeled("PDF Template", self.template_combo))

        lay.addSpacing(4)
        self.no_charge_toggle = ToggleSwitch("Mark as Complimentary (No Charge)")
        self.no_charge_toggle.clicked.connect(self._toggle_no_charge)
        lay.addWidget(self.no_charge_toggle)
        self.show_qr_toggle = ToggleSwitch("Show UPI QR Code on PDF")
        self.show_qr_toggle.setChecked(True)
        lay.addWidget(self.show_qr_toggle)
        self.reverse_toggle = ToggleSwitch("Reverse charge applies")
        self.reverse_toggle.setToolTip("The tax is paid by the customer, not by you: prints \"Reverse Charge: Yes\" on a GST invoice")
        lay.addWidget(self.reverse_toggle)
        lay.addStretch()
        row.addWidget(left, 1)

        right, lay = self._card("Tax & Discount")
        lay.setSpacing(8)
        self.gst_toggle = ToggleSwitch("This is a GST Invoice")
        self.gst_toggle.setChecked(True)
        self.gst_toggle.clicked.connect(self._on_gst_toggle)          # toggling GST immediately recalculates the totals
        lay.addWidget(self.gst_toggle)

        self.tax_mode_combo = QComboBox()
        self.tax_mode_combo.addItem("One rate for the whole invoice", TAX_MODE_OVERALL)
        self.tax_mode_combo.addItem("A rate for each item", TAX_MODE_PER_ITEM)
        self.tax_mode_combo.setToolTip("Overall: one GST rate on everything.  Per item: every item line gets its own GST % "
                                       "and the totals show tax per rate.")
        self.tax_mode_combo.currentIndexChanged.connect(self._on_tax_mode_changed)
        self.tax_mode_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.tax_mode_combo.setMinimumContentsLength(12)
        guard_wheel(self.tax_mode_combo)
        lay.addWidget(self._labeled("GST Applied", self.tax_mode_combo))

        self.tax_rate_entry = QLineEdit("18")
        restrict_to_decimal(self.tax_rate_entry, 2)
        format_as_decimal_on_blur(self.tax_rate_entry)
        self.tax_rate_entry.textChanged.connect(self._on_default_rate_changed)
        advance_on_enter(self.tax_rate_entry)
        self._tax_rate_row = self._labeled("GST / Tax Rate (%)", self.tax_rate_entry)
        lay.addWidget(self._tax_rate_row)

        self.gst_type_combo = QComboBox()
        self.gst_type_combo.addItems(["CGST/SGST", "IGST"])
        self.gst_type_combo.currentTextChanged.connect(self._recalc_totals)   # live update when switching CGST <-> IGST
        guard_wheel(self.gst_type_combo)
        lay.addWidget(self._labeled("Tax Split Type", self.gst_type_combo))

        self.place_combo = QComboBox()
        self.place_combo.setEditable(True)
        self.place_combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.place_combo.addItems([state_label(code) for code, _name in INDIAN_STATES])
        self.place_combo.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon)
        self.place_combo.setMinimumContentsLength(12)
        self.place_combo.setCurrentIndex(-1)
        self.place_combo.lineEdit().setPlaceholderText("State the supply is made to (optional)")
        self.place_combo.lineEdit().setMaxLength(60)
        place_completer = QCompleter(self.place_combo.model(), self.place_combo)
        place_completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        place_completer.setFilterMode(Qt.MatchFlag.MatchContains)
        self.place_combo.setCompleter(place_completer)
        self.place_combo.setToolTip("Another state than yours switches the tax split to IGST; your own state to CGST/SGST")
        self.place_combo.currentTextChanged.connect(self._on_place_changed)
        self.place_combo.lineEdit().textEdited.connect(self._on_place_typed)
        self.place_combo.activated.connect(self._on_place_typed)
        guard_wheel(self.place_combo)
        advance_on_enter(self.place_combo)
        lay.addWidget(self._labeled("Place of Supply", self.place_combo))

        disc_wrap = QWidget()
        dl = QHBoxLayout(disc_wrap)
        dl.setContentsMargins(0, 0, 0, 0)
        dl.setSpacing(6)
        self.global_discount_entry = QLineEdit("0")
        restrict_to_decimal(self.global_discount_entry, 2)
        format_as_decimal_on_blur(self.global_discount_entry)
        self.global_discount_entry.textChanged.connect(self._recalc_totals)
        advance_on_enter(self.global_discount_entry)
        dl.addWidget(self.global_discount_entry, 1)
        self.disc_type_combo = QComboBox()
        self.disc_type_combo.addItems(["amount", "%"])
        self.disc_type_combo.setMinimumWidth(92)
        self.disc_type_combo.currentTextChanged.connect(self._recalc_totals)
        guard_wheel(self.disc_type_combo)
        dl.addWidget(self.disc_type_combo)
        lay.addWidget(self._labeled("Overall Discount", disc_wrap))

        self.round_off_toggle = ToggleSwitch("Round off to the nearest rupee")
        self.round_off_toggle.clicked.connect(self._recalc_totals)
        lay.addWidget(self.round_off_toggle)
        lay.addStretch()
        row.addWidget(right, 1)
        parent.addLayout(row)

        self.gst_dependents = [self.tax_mode_combo, self.tax_rate_entry, self.gst_type_combo, self.place_combo,
                               self.reverse_toggle]

    def _fill_template_combo(self):
        """"Default (<the template picked in Settings>)", then every template; the choice is kept across a refill."""
        chosen = self.template_combo.currentData() if self.template_combo.count() else ""
        default = PDF_TEMPLATES.get((CONFIG.get("pdf_settings") or {}).get("template"), PDF_TEMPLATE_CLASSIC)
        self.template_combo.clear()
        self.template_combo.addItem(f"Default ({default.label})", "")
        for tpl in PDF_TEMPLATES.values():
            self.template_combo.addItem(tpl.label, tpl.key)
        self.template_combo.setCurrentIndex(max(0, self.template_combo.findData(chosen or "")))

    def _toggle_no_charge(self, _checked: bool = False):
        self._apply_field_states()
        self._recalc_totals()

    def _on_gst_toggle(self, _checked: bool = False):
        self._apply_field_states()
        self._recalc_totals()

    def _on_tax_mode_changed(self, _index: int = 0):
        self._apply_field_states()
        self._recalc_totals()

    def _on_default_rate_changed(self, *_):
        """The invoice's (default) GST rate was edited: lines whose own rate was never chosen follow it."""
        text = _num_text(safe_float(self.tax_rate_entry.text()))
        for item_row in self.item_rows:
            item_row.set_default_gst(text)
        self._recalc_totals()

    def _on_place_changed(self, text: str):
        """Choosing where the supply goes decides the tax split: another state than the seller's is IGST, the
        seller's own state CGST/SGST. (Loading a saved invoice never changes what was saved.)"""
        if self._loading:
            return
        kind = gst_type_for_place(CONFIG, text)
        if kind and kind != self.gst_type_combo.currentText():
            self.gst_type_combo.setCurrentText(kind)

    def _on_place_typed(self, *_):
        self._place_auto = False             # the user chose the place: a customer's GSTIN no longer overrides it

    def _autofill_place_from_gstin(self, *_):
        """A valid customer GSTIN names the customer's state: it becomes the place of supply, unless the user chose
        one (a place filled in this way follows the GSTIN when another customer is picked)."""
        gstin = self.client_gstin_entry.text().strip().upper()
        if self._loading or not is_valid_gstin(gstin):
            return
        if self.place_combo.currentText().strip() and not self._place_auto:
            return
        if label := state_label(gstin[:2]):
            self._place_auto = True
            self.place_combo.setCurrentText(label)

    def _default_due_days(self) -> int:
        """How many days after the invoice date a due date starts at: read from the payment terms ("Net 15",
        "30 days"), 0 for "on receipt", else 15."""
        terms = self.payment_terms_entry.text()
        found = re.search(r"\bnet\s*(\d{1,3})\b|\b(\d{1,3})\s*days?\b", terms, re.IGNORECASE)
        if found:
            return int(found.group(1) or found.group(2))
        return 0 if re.search(r"receipt|immediate|advance", terms, re.IGNORECASE) else 15

    def _set_due_in(self, days: int):
        due = date_edit_value(self.invoice_date_edit) + datetime.timedelta(days=days)
        self._due_days, self._setting_due = days, True
        try:
            self.due_date_edit.setDate(QDate(due.year, due.month, due.day))
        finally:
            self._setting_due = False

    def _due_on(self) -> bool:
        return self.due_combo.currentData() is not None

    def _set_due_active(self, active: bool):
        """The "Due On" date field exists (and is editable) only while the invoice has a due date."""
        self.due_date_edit.setEnabled(active)
        self._due_date_row.setVisible(active)

    def _on_due_date_edited(self, _date=None):
        """A due date picked in the date field is a "Custom" one: it no longer follows the invoice date."""
        if self._setting_due:
            return
        self._due_days = None
        if self._due_on() and not self._loading:
            self.due_combo.setCurrentIndex(self.due_combo.findData(_DUE_CUSTOM))

    def _on_invoice_date_changed(self, _date=None):
        """A due date set as "n days after the invoice date" keeps that distance when the invoice date moves."""
        if self._due_days is not None and self._due_on() and not self._loading:
            self._set_due_in(self._due_days)

    def _on_due_chosen(self, index: int):
        days = self.due_combo.itemData(index)
        was_on = self.due_date_edit.isEnabled()
        self._set_due_active(days is not None)
        if days is None:
            self._due_days = None
        elif days == _DUE_CUSTOM:
            if not was_on:
                self._set_due_in(self._default_due_days())
            self._due_days = None
        else:
            self._set_due_in(days)

    def _show_stored_due(self, invoice_date: Optional[str], due: Optional[str]):
        """Show a stored due date: as "n days" when it is exactly that long after the invoice date, else as a custom one."""
        self._due_days = None
        if not due:
            self.due_combo.setCurrentIndex(0)
            self._set_due_active(False)
            set_date_edit_from_db(self.due_date_edit, None)
            return
        set_date_edit_from_db(self.due_date_edit, due)
        self._set_due_active(True)
        try:
            days = (datetime.datetime.strptime(due, DB_DATE_FORMAT)
                    - datetime.datetime.strptime(invoice_date or "", DB_DATE_FORMAT)).days
        except (TypeError, ValueError):
            days = -1
        index = self.due_combo.findData(days) if days >= 0 else -1
        if index > 0:
            self.due_combo.setCurrentIndex(index)
            self._due_days = days
        else:
            self.due_combo.setCurrentIndex(self.due_combo.findData(_DUE_CUSTOM))

    def _per_item_active(self) -> bool:
        """True when GST is charged per item (and there is GST to charge at all)."""
        return (self.gst_toggle.isChecked() and not self.no_charge_toggle.isChecked()
                and self.tax_mode_combo.currentData() == TAX_MODE_PER_ITEM)

    def _apply_field_states(self):
        no_charge = self.no_charge_toggle.isChecked()
        is_gst    = self.gst_toggle.isChecked() and not no_charge
        for w in self.gst_dependents:
            w.setEnabled(is_gst)
        self.round_off_toggle.setEnabled(not no_charge)
        per_item = self._per_item_active()
        for item_row in self.item_rows:
            item_row.set_per_item(per_item)
        self._tax_rate_row.label.setText("Default GST Rate (%)" if self.tax_mode_combo.currentData() == TAX_MODE_PER_ITEM
                                         else "GST / Tax Rate (%)")
        self.add_payment_btn.setEnabled(not no_charge)
        for row in self.payment_rows:
            row.set_editable(not no_charge)
        self._restore_tab_order()

    # ---- Items ---------------------------------------------------------- #
    # ---- Items ---------------------------------------------------------- #
    def _section_items(self, parent: QVBoxLayout):
        card, lay = self._card("Items")
        lay.addWidget(make_label("Quantity × rate is the line amount. Pick a unit or type your own; a discount can be an "
                                 "amount (150) or a percentage (10%). Press Enter in an item's last field to start the next one.",
                                 "muted", wrap=True))
        lay.addSpacing(8)
        header = QHBoxLayout()
        header.setSpacing(_ROW_GAP)
        header.addWidget(make_label("ITEM / SERVICE", "section"), 1)
        for text, width in (("QTY", _QTY_W), ("UNIT", _UNIT_W), ("RATE", _AMOUNT_W), ("DISCOUNT", _DISC_W),
                            ("AMOUNT", _LINE_W)):
            caption = make_label(text, "section", align=Qt.AlignmentFlag.AlignCenter)
            caption.setFixedWidth(width)
            header.addWidget(caption)
        header.addSpacing(_REMOVE_W)
        lay.addLayout(header)
        lay.addSpacing(4)
        lay.addWidget(hline())

        self.items_layout = QVBoxLayout()
        self.items_layout.setSpacing(0)
        lay.addLayout(self.items_layout)
        lay.addSpacing(10)
        row = QHBoxLayout()
        self.add_item_btn = make_button("＋  Add Item  (Ctrl+Shift+A)", "secondary-outline", self._add_item_row_clicked,
                                        tooltip="Add a blank item row (Ctrl+Shift+A or Ctrl+Plus)")
        row.addWidget(self.add_item_btn)
        row.addStretch()
        lay.addLayout(row)
        parent.addWidget(card)

    def _add_item_row_clicked(self):
        self._add_item_row()

    def _add_item_row(self, item: Optional[Dict[str, Any]] = None) -> _ItemRow:
        """Adds an item line: blank, or filled in from a stored item (`item`)."""
        row = _ItemRow([c["title"] for c in self._catalog_cache])
        row.set_default_gst(_num_text(safe_float(self.tax_rate_entry.text())))
        row.set_per_item(self._per_item_active())
        if item is not None:
            row.load_item(item)
            self._on_item_title_edited(row, row.title_edit.text())          # offer the units its catalog item is sold in
        row.changed.connect(self._recalc_totals)
        row.remove_requested.connect(self._remove_item_row)
        row.title_picked.connect(self._on_catalog_picked)
        row.title_edited.connect(self._on_item_title_edited)
        for last in (row.hsn_edit, row.gst_combo):                            # Enter in the last field starts the next row
            last.installEventFilter(_EnterAddsRow(row, self))
        self.items_layout.addWidget(row)
        self.item_rows.append(row)
        self._restore_tab_order()
        self._recalc_totals()
        return row

    def _remove_item_row(self, row: _ItemRow):
        if row not in self.item_rows:
            return
        self.item_rows.remove(row)
        self.items_layout.removeWidget(row)
        row.setParent(None)
        row.deleteLater()
        self._restore_tab_order()
        self._recalc_totals()

    def _catalog_entry(self, title: str) -> Optional[Dict[str, Any]]:
        """The catalog item called `title` (the exact spelling first, then ignoring upper / lower case)."""
        entry = self._catalog_exact.get(title)
        return entry if entry is not None else self._catalog_fold.get(title.strip().casefold())

    def _on_catalog_picked(self, row: _ItemRow, chosen: str):
        # Matched against the text the user actually chose, so a still-visible placeholder can never be
        # mistaken for a catalog title.
        entry = self._catalog_exact.get(chosen)
        if entry is not None:
            row.desc_edit.setText(entry.get("description") or "")
            row.desc_edit.setCursorPosition(0)
            row.apply_catalog(entry)                  # rate, unit, HSN/SAC and GST rate of the item
        self._recalc_totals()

    def _on_item_title_edited(self, row: _ItemRow, text: str):
        """Whatever is typed as an item's name: when the catalog knows that item, its units (and their rates) are
        what the row's unit list offers."""
        entry = self._catalog_entry(text)
        row.set_units(catalog_units(entry) if entry is not None else [])

    # ---- Payment History ------------------------------------------------ #
    def _section_payments(self, parent: QVBoxLayout):
        card, lay = self._card("Payment History")
        lay.addWidget(make_label("Log each payment received. Status updates automatically.", "muted"))
        lay.addSpacing(8)
        header = QHBoxLayout()
        header.setSpacing(_ROW_GAP)
        date_hdr = make_label("DATE", "section")
        date_hdr.setFixedWidth(_DATE_W)
        amount_hdr = make_label("AMOUNT", "section", align=Qt.AlignmentFlag.AlignCenter)
        amount_hdr.setFixedWidth(_AMOUNT_W)
        mode_hdr = make_label("MODE", "section")
        mode_hdr.setFixedWidth(_MODE_W)
        header.addWidget(date_hdr)
        header.addWidget(amount_hdr)
        header.addWidget(mode_hdr)
        header.addWidget(make_label("NOTE", "section"), 1)
        header.addSpacing(_REMOVE_W)
        lay.addLayout(header)
        lay.addSpacing(4)
        lay.addWidget(hline())

        self.payments_layout = QVBoxLayout()
        self.payments_layout.setSpacing(0)
        lay.addLayout(self.payments_layout)
        lay.addSpacing(10)
        row = QHBoxLayout()
        self.add_payment_btn = make_button("＋  Add Payment", "secondary-outline", self._add_payment_row_clicked)
        row.addWidget(self.add_payment_btn)
        row.addStretch()
        lay.addLayout(row)
        parent.addWidget(card)

    def _add_payment_row_clicked(self):
        self._add_payment_row()

    def _add_payment_row(self, amount: float = 0.0, payment_date: Optional[str] = None, mode: str = "Cash",
                         note: str = "", payment_id: Optional[int] = None) -> _PaymentRow:
        row = _PaymentRow(amount, payment_date, mode, note, payment_id)
        row.changed.connect(self._recalc_totals)
        row.remove_requested.connect(self._remove_payment_row)
        row.set_editable(not self.no_charge_toggle.isChecked())
        self.payments_layout.addWidget(row)
        self.payment_rows.append(row)
        self._restore_tab_order()
        self._recalc_totals()
        return row

    def _remove_payment_row(self, row: _PaymentRow):
        if row not in self.payment_rows:
            return
        self.payment_rows.remove(row)
        self.payments_layout.removeWidget(row)
        row.setParent(None)
        row.deleteLater()
        self._restore_tab_order()
        self._recalc_totals()

    # ---- Notes & Terms (printed) ------------------------------------------ #
    def _section_print_notes(self, parent: QVBoxLayout):
        card, lay = self._card("Notes & Terms")
        lay.addWidget(make_label("Printed on the invoice, under the totals. Leave a box empty to print nothing.", "muted", wrap=True))
        lay.addSpacing(8)
        row = QHBoxLayout()
        row.setSpacing(16)

        left = QVBoxLayout()
        left.setSpacing(2)
        left.addWidget(_caption("Note to the customer"))
        self.customer_notes_text = self._text_box("e.g. Thank you for your business!")
        left.addWidget(self.customer_notes_text)
        row.addLayout(left, 1)

        right = QVBoxLayout()
        right.setSpacing(2)
        head = QHBoxLayout()
        head.addWidget(_caption("Terms & Conditions"))
        head.addStretch()
        self.default_terms_btn = make_button("Use default terms", "link", self._use_default_terms,
                                             tooltip="Fill this box with the default terms from Settings > Payment & PDF")
        self.default_terms_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        head.addWidget(self.default_terms_btn)
        right.addLayout(head)
        self.terms_text = self._text_box("e.g. Payment within 15 days. Late fee 2% per month.")
        right.addWidget(self.terms_text)
        row.addLayout(right, 1)
        lay.addLayout(row)
        parent.addWidget(card)

    def _use_default_terms(self):
        default = str(CONFIG.get("pdf_settings", {}).get("default_terms", "") or "").strip()
        if not default:
            Messagebox.show_info("No Default Terms",
                                 "No default terms have been set yet.\nAdd them in Settings > Payment & PDF > Invoice Printing.",
                                 parent=self.controller)
            return
        current = self.terms_text.toPlainText().strip()
        if current and current != default and not Messagebox.yesno(
                "Replace Terms?", "Replace the terms written here with the default terms?", parent=self.controller):
            return
        self.terms_text.setPlainText(default)

    # ---- Notes ------------------------------------------------------------ #
    def _section_notes(self, parent: QVBoxLayout):
        card, lay = self._card("Internal Notes")
        lay.addWidget(make_label("Visible only inside Vyaprix — never printed on the invoice PDF.", "muted"))
        lay.addSpacing(6)
        self.notes_text = self._text_box("Internal notes (optional) — not shown on the PDF")
        lay.addWidget(self.notes_text)
        parent.addWidget(card)

    # ---- Totals ----------------------------------------------------------- #
    def _section_totals(self, parent: QVBoxLayout):
        wrap = QHBoxLayout()
        wrap.setContentsMargins(16, 0, 16, 0)
        wrap.addStretch()
        self.summary_box = QWidget()
        self.summary_layout = QVBoxLayout(self.summary_box)
        self.summary_layout.setContentsMargins(0, 0, 0, 0)
        self.summary_layout.setSpacing(0)

        self.row_subtotal    = _SummaryRow("Subtotal:")
        self.row_disc        = _SummaryRow("Discount:")
        self.row_taxable     = _SummaryRow("Taxable Amount:")
        self.row_round       = _SummaryRow("Round Off:")
        self.row_grand_total = _SummaryRow("Grand Total:", bold=True)
        self.row_paid        = _SummaryRow("Total Paid:")
        self.row_balance     = _SummaryRow("Balance Due:", bold=True)     # flips to "Advance / Credit:" when overpaid
        # The tax lines (one CGST and one SGST, or one IGST) are made on demand and slotted in between
        # "Taxable Amount" and "Round Off", so _summary_rows is always in on-screen order.
        self._tax_pool: List[_SummaryRow] = []
        self._summary_rows: List[_SummaryRow] = [self.row_subtotal, self.row_disc, self.row_taxable, self.row_round,
                                                 self.row_grand_total, self.row_paid, self.row_balance]
        for r in self._summary_rows:
            self.summary_layout.addWidget(r)
        self.words_lbl = make_label("", "muted", wrap=True,
                                    align=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        self.words_lbl.setMaximumWidth(440)
        self.words_lbl.setVisible(False)
        self.summary_layout.addSpacing(4)
        self.summary_layout.addWidget(self.words_lbl, 0, Qt.AlignmentFlag.AlignRight)
        # The rows are shown / hidden by _recalc_totals(), always in this fixed order:
        # Subtotal, Discount, Taxable Amount, Tax, Round Off, Grand Total, Total Paid, Balance Due.
        wrap.addWidget(self.summary_box)
        parent.addLayout(wrap)

    def _tax_pool_row(self, index: int) -> _SummaryRow:
        """The `index`-th tax line of the totals box (made the first time an invoice needs it)."""
        while len(self._tax_pool) <= index:
            row = _SummaryRow()
            row.setVisible(False)
            self.summary_layout.insertWidget(self.summary_layout.indexOf(self.row_round), row)
            self._summary_rows.insert(self._summary_rows.index(self.row_round), row)
            self._tax_pool.append(row)
        return self._tax_pool[index]

    def _section_save_bar(self, parent: QVBoxLayout):
        row = QHBoxLayout()
        row.setContentsMargins(0, 4, 0, 12)
        row.addStretch()
        self.save_btn = make_button("💾  Save Invoice  (Ctrl+S)", "success", self._save, tooltip="Save this invoice (Ctrl+S)",
                                    min_height=40)
        row.addWidget(self.save_btn)
        row.addSpacing(8)
        self.back_btn = make_button("←  Back to Dashboard  (Esc)", "secondary", self._go_back,
                                    tooltip="Return to the invoice list (Esc). Unsaved changes will ask first.", min_height=40)
        row.addWidget(self.back_btn)
        parent.addLayout(row)

    def _go_back(self):
        self.controller.show_frame(DashboardFrame)

    def _restore_tab_order(self):
        """Qt appends widgets created later (new item / payment rows) to the END of the focus chain; this rebuilds
        the chain in on-screen order after every structural change, so Tab / Enter always go to the next field."""
        chain: List[QWidget] = [
            self.top_back_btn, self.client_edit, self.client_contact_text, self.client_gstin_entry,
            self.your_name_entry, self.your_contact_text,
            self.invoice_number_entry, self.invoice_date_edit, self.due_combo, self.due_date_edit,
            self.payment_terms_entry, self.po_entry, self.template_combo,
            self.no_charge_toggle, self.show_qr_toggle, self.reverse_toggle,
            self.gst_toggle, self.tax_mode_combo, self.tax_rate_entry, self.gst_type_combo, self.place_combo,
            self.global_discount_entry, self.disc_type_combo, self.round_off_toggle,
        ]
        for r in self.item_rows:
            chain.extend(r.focus_widgets())
        chain.append(self.add_item_btn)
        for r in self.payment_rows:
            chain.extend(r.focus_widgets())
        chain.extend([self.add_payment_btn, self.customer_notes_text, self.terms_text, self.notes_text,
                      self.save_btn, self.back_btn])
        for first, second in zip(chain, chain[1:]):
            QWidget.setTabOrder(first, second)

    # ------------------------------------------------------------------ #
    #  Client autocomplete
    # ------------------------------------------------------------------ #
    def _on_client_picked(self, chosen: str):
        for c in self._client_cache:
            if c["name"] == chosen:
                current_contact = self.client_contact_text.toPlainText().strip()
                current_gstin = self.client_gstin_entry.text().strip()
                new_contact = c.get("contact") or ""
                new_gstin = c.get("gstin") or ""

                # Fields that already hold different, typed-in data must not be wiped silently.
                needs_prompt = bool((current_contact and current_contact != new_contact)
                                    or (current_gstin and current_gstin != new_gstin))
                if needs_prompt and not Messagebox.yesno(
                        "Overwrite Contact Info?",
                        f"Do you want to replace the current contact info and GSTIN with the saved details for {chosen}?",
                        parent=self.controller):
                    break                                      # user clicked No: leave their typed text alone
                self.client_contact_text.setPlainText(new_contact)
                self.client_gstin_entry.setText(new_gstin)
                self.client_gstin_entry.setCursorPosition(0)
                break

    # ------------------------------------------------------------------ #
    #  Collecting the form  /  unsaved-changes guard
    # ------------------------------------------------------------------ #
    def _collect_items(self) -> List[Dict]:
        return [r.values() for r in self.item_rows]

    def _filled_items(self) -> List[Dict]:
        """The lines that make up the invoice: a row nothing has been entered in is ignored (live totals and save)."""
        return [i for i in self._collect_items() if i["title"] or i["description"] or i["amount"]]

    def _collect_payments(self) -> List[Dict]:
        return [r.values() for r in self.payment_rows]

    def _snapshot(self) -> str:
        """The whole editable state as a string, to tell "changed" from "just looked at". Numbers are compared as
        numbers, not as text, so tabbing through a field that re-formats itself ("18" -> "18.00") is not a change."""
        try:
            state = {
                "invoice_number": self.invoice_number_entry.text().strip(),
                "bill_to":        self.client_edit.text().strip(),
                "client_contact": self.client_contact_text.toPlainText().strip(),
                "client_gstin":   self.client_gstin_entry.text().strip(),
                "your_name":      self.your_name_entry.text().strip(),
                "your_contact":   self.your_contact_text.toPlainText().strip(),
                "invoice_date":   date_edit_db(self.invoice_date_edit),
                "due_date":       date_edit_db(self.due_date_edit) if self._due_on() else None,
                "payment_terms":  self.payment_terms_entry.text().strip(),
                "po_number":      self.po_entry.text().strip(),
                "pdf_template":   self.template_combo.currentData() or "",
                "notes":          self.notes_text.toPlainText().strip(),
                "customer_notes": self.customer_notes_text.toPlainText().strip(),
                "terms":          self.terms_text.toPlainText().strip(),
                "no_charge":      self.no_charge_toggle.isChecked(),
                "show_qr":        self.show_qr_toggle.isChecked(),
                "reverse":        self.reverse_toggle.isChecked(),
                "gst":            self.gst_toggle.isChecked(),
                "tax_mode":       self.tax_mode_combo.currentData(),
                "tax_rate":       safe_float(self.tax_rate_entry.text()),
                "gst_type":       self.gst_type_combo.currentText(),
                "place":          self.place_combo.currentText().strip(),
                "g_disc":         safe_float(self.global_discount_entry.text()),
                "g_disc_type":    self.disc_type_combo.currentText(),
                "round_off":      self.round_off_toggle.isChecked(),
                "items":          self._collect_items(),
                "payments":       self._collect_payments(),
            }
            return json.dumps(state, sort_keys=True, default=str)
        except Exception:
            return self._initial_snapshot or ""

    def has_unsaved_changes(self) -> bool:
        if self._initial_snapshot is None or self._loading:
            return False
        return self._snapshot() != self._initial_snapshot

    def confirm_navigate_away(self) -> bool:
        if not self.has_unsaved_changes():
            return True
        return Messagebox.yesno(
            "Discard Changes?",
            "This invoice has unsaved changes. Leave without saving?",
            parent=self.controller,
        )

    # ------------------------------------------------------------------ #
    #  Totals calculation (live, mirrors DatabaseManager.calculate_totals)
    # ------------------------------------------------------------------ #
    def _set_badge(self, text: str, tone: str):
        if self.status_badge.property("tone") != tone:
            set_role(self.status_badge, "badge", tone)
        self.status_badge.setText(text)
        self.status_badge.setVisible(bool(text))

    def _recalc_totals(self, *_):
        """Calculates live totals and updates the summary box.

        Rows are only shown / hidden -- never re-created (the tax lines excepted: they are made when an invoice first
        needs them) -- so the summary always lays out in the fixed, correct order (Subtotal, Discount, Taxable Amount,
        Tax, Round Off, Grand Total, Total Paid, Balance Due) whichever rows were visible on the previous call.

        The figures come from DatabaseManager.calculate_totals() -- the engine the database layer and the PDF use --
        and the balance / status from DatabaseManager.compute_payment_status(), so this live preview can never show a
        different number or status than what actually gets saved and printed.
        """
        if self._loading:
            return
        for row in self._summary_rows:
            row.setVisible(False)
        self.words_lbl.setVisible(False)

        if self.no_charge_toggle.isChecked():
            self.row_subtotal.set(fmt_money(0))
            self.row_grand_total.set(fmt_money(0))
            self.row_balance.set(fmt_money(0), "Balance Due:")
            _retone(self.row_balance.value, "secondary")
            for r in (self.row_subtotal, self.row_grand_total, self.row_balance):
                r.setVisible(True)
            self._set_badge(STATUS_DISPLAY_LABELS[InvoiceStatus.COMPLIMENTARY.value], STATUS_TONES["Complimentary"])
            return

        items    = self._filled_items()
        is_gst   = self.gst_toggle.isChecked()
        tax_rate = safe_float(self.tax_rate_entry.text()) if is_gst else 0.0
        g_disc   = safe_float(self.global_discount_entry.text())
        gst_type = self.gst_type_combo.currentText()

        totals = DatabaseManager.calculate_totals(
            items, tax_rate, g_disc, self.disc_type_combo.currentText(),
            tax_mode=TAX_MODE_PER_ITEM if self._per_item_active() else TAX_MODE_OVERALL,
            gst_type=gst_type, round_off=self.round_off_toggle.isChecked())
        paid   = sum(r.values()["amount"] for r in self.payment_rows)
        status, balance, credit = DatabaseManager.compute_payment_status(totals["grand_total"], paid)

        # Grand Total = Subtotal - Discount + Tax (+ Round Off): already true here, since calculate_totals() derives
        # grand_total as taxable + cgst + sgst, where taxable is subtotal minus total_discount.
        self.row_subtotal.set(fmt_money(totals["subtotal"]))
        self.row_taxable.set(fmt_money(totals["taxable"]))
        self.row_grand_total.set(fmt_money(totals["grand_total"]))

        # Balance Due is never negative: compute_payment_status() floors it at zero and returns any overpayment
        # separately as `credit`, shown as "Advance / Credit" in a positive colour instead of a confusing
        # negative balance.
        if credit > 0:
            self.row_balance.set(fmt_money(credit), "Advance / Credit:")
            _retone(self.row_balance.value, "info")
        else:
            self.row_balance.set(fmt_money(balance), "Balance Due:")
            _retone(self.row_balance.value, "danger" if balance > 0 else "success")

        # 1. Subtotal always shows first.
        self.row_subtotal.setVisible(True)

        # 2. Discount -- shown as a negative value, only when there is one.
        if totals["total_discount"] > 0:
            self.row_disc.set(f"− {fmt_money(totals['total_discount'])}")
            self.row_disc.setVisible(True)

        # 3. Taxable Amount -- Subtotal minus Discount, the base tax is actually calculated on. Only worth showing
        # once there's a discount or tax to make it different from the Subtotal above.
        taxed = [g for g in totals["tax_groups"] if g["rate"] > 0] if is_gst else []
        if totals["total_discount"] > 0 or taxed:
            self.row_taxable.setVisible(True)

        # 4. Tax -- always ONE CGST line and ONE SGST line (or one IGST line), however many GST rates the invoice has.
        # The rate is part of the label when there is a single rate (18% -> CGST 9% + SGST 9%); with several rates the
        # lines are the totals, and hovering one lists the split by rate (the printed invoice has a GST summary).
        tax_lines: List[Tuple[str, float, str]] = []
        if taxed:
            several = len(taxed) > 1
            if gst_type == "IGST":
                tax_lines.append(("IGST:" if several else f"IGST ({taxed[0]['rate']:g}%):", totals["tax"],
                                  gst_breakdown_text(taxed, "tax", 1.0) if several else ""))
            else:
                half = taxed[0]["rate"] / 2.0
                tax_lines.append(("CGST:" if several else f"CGST ({half:g}%):", totals["cgst"],
                                  gst_breakdown_text(taxed, "cgst", 0.5) if several else ""))
                tax_lines.append(("SGST:" if several else f"SGST ({half:g}%):", totals["sgst"],
                                  gst_breakdown_text(taxed, "sgst", 0.5) if several else ""))
        for line, (label, amount, tip) in enumerate(tax_lines):
            tax_row = self._tax_pool_row(line)
            tax_row.set(fmt_money(amount), label)
            tax_row.setToolTip(tip)
            tax_row.setVisible(True)

        # 5. Round Off -- only when the total was actually rounded.
        if totals["round_off"]:
            self.row_round.set(f"{'+' if totals['round_off'] > 0 else '−'} {fmt_money(abs(totals['round_off']))}")
            self.row_round.setVisible(True)

        # 6. Grand Total always shows.
        self.row_grand_total.setVisible(True)

        # 7. Total Paid -- only when something has actually been paid.
        if paid > 0:
            self.row_paid.set(f"− {fmt_money(paid)}")
            self.row_paid.setVisible(True)

        # 8. Balance Due (or Advance / Credit) always shows last.
        self.row_balance.setVisible(True)

        # The total in words, when Settings asks for it (rupee invoices only).
        if CONFIG.get("pdf_settings", {}).get("amount_in_words", False):
            words = amount_in_words(totals["grand_total"], CURRENCY_SYMBOL)
            self.words_lbl.setText(words)
            self.words_lbl.setVisible(bool(words))

        # Status badge -- UNPAID / PARTIALLY PAID / PAID / OVERPAID, from the exact same status
        # compute_payment_status() just returned, so the badge can never disagree with the numbers above it.
        self._set_badge(STATUS_DISPLAY_LABELS.get(status, status.upper()), STATUS_TONES.get(status, "secondary"))

    # ------------------------------------------------------------------ #
    #  Page life-cycle
    # ------------------------------------------------------------------ #
    def _refresh_suggestions(self):
        self._catalog_cache = self.db_manager.get_all_catalog_items()
        self._catalog_exact = {c["title"]: c for c in self._catalog_cache}
        self._catalog_fold  = {c["title"].casefold(): c for c in self._catalog_cache}
        self._client_cache  = self.db_manager.get_all_clients()
        self.client_edit.set_values([c["name"] for c in self._client_cache])
        titles = [c["title"] for c in self._catalog_cache]
        for row in self.item_rows:
            row.title_edit.set_values(titles)

    def refresh_after_settings_change(self):
        self._fill_template_combo()            # "Default (...)" names the template Settings now has
        self._refresh_suggestions()
        for row in self.item_rows:
            row.refresh_line()                 # the currency symbol may have changed
        self._recalc_totals()

    def on_show(self, *args):
        invoice_id: Optional[str] = args[0] if args else None
        self._initial_snapshot = None          # nothing to compare against (and no "unsaved changes") while loading
        self._loading = True
        try:
            self._refresh_suggestions()
            self._reset()
            if invoice_id:
                self._load_for_edit(invoice_id)
            else:
                self.editing_id = None
                self.title_lbl.setText("New Invoice")
                self.save_btn.setText("💾  Save Invoice  (Ctrl+S)")
                self.invoice_number_entry.setText(self.db_manager.get_next_invoice_number())
                set_date_edit_from_db(self.invoice_date_edit, None)
                self.payment_terms_entry.setText(CONFIG.get("payment_details", {}).get("payment_terms", ""))
                self.your_name_entry.setText(compose_your_name(CONFIG))
                self.your_contact_text.setPlainText(compose_your_contact(CONFIG))
                self.terms_text.setPlainText(str(CONFIG.get("pdf_settings", {}).get("default_terms", "") or ""))
                self._add_item_row()
        finally:
            self._loading = False
        self._apply_field_states()
        self._recalc_totals()
        self._initial_snapshot = self._snapshot()
        self.scroll.verticalScrollBar().setValue(0)
        QTimer.singleShot(0, self._focus_first_field)

    def _focus_first_field(self):
        if self.isVisible():
            self.scroll.verticalScrollBar().setValue(0)
            self.client_edit.setFocus()

    def _clear_rows(self):
        for row in list(self.item_rows):
            self.items_layout.removeWidget(row)
            row.setParent(None)
            row.deleteLater()
        self.item_rows.clear()
        for row in list(self.payment_rows):
            self.payments_layout.removeWidget(row)
            row.setParent(None)
            row.deleteLater()
        self.payment_rows.clear()

    def _reset(self):
        self._clear_rows()
        self.client_edit.clear()
        self.client_contact_text.clear()
        self.client_gstin_entry.clear()
        self.your_name_entry.clear()
        self.your_contact_text.clear()
        self.invoice_number_entry.clear()
        set_date_edit_from_db(self.invoice_date_edit, None)
        self.payment_terms_entry.clear()
        self.notes_text.clear()
        self.no_charge_toggle.setChecked(False)
        self.show_qr_toggle.setChecked(True)
        self.gst_toggle.setChecked(True)
        self.tax_rate_entry.setText("18")
        self.gst_type_combo.setCurrentText("CGST/SGST")
        self.global_discount_entry.setText("0")
        self.disc_type_combo.setCurrentText("amount")
        self.tax_mode_combo.setCurrentIndex(0)
        self.round_off_toggle.setChecked(False)
        self.place_combo.setCurrentIndex(-1)
        self.place_combo.setEditText("")
        self._place_auto = False
        self._show_stored_due(None, None)
        self.po_entry.clear()
        self.template_combo.setCurrentIndex(0)
        self.reverse_toggle.setChecked(False)
        self.customer_notes_text.clear()
        self.terms_text.clear()
        self._set_badge("", "secondary")
        self._restore_tab_order()

    def _load_for_edit(self, invoice_id: str):
        row = self.db_manager.get_invoice_by_id(invoice_id)
        if not row:
            Messagebox.show_error("Not Found", f"Invoice {invoice_id} no longer exists.", parent=self.controller)
            self.controller.show_frame(DashboardFrame)
            return
        d = dict(row)
        self.editing_id = invoice_id
        self.title_lbl.setText(f"Editing Invoice #{invoice_id}")
        self.save_btn.setText("💾  Update Invoice  (Ctrl+S)")

        self.client_edit.set_text(d.get("bill_to") or "")
        self.client_contact_text.setPlainText(d.get("client_contact") or "")
        self.client_gstin_entry.setText(d.get("client_gstin") or "")
        self.your_name_entry.setText(d.get("your_name") or "")
        self.your_contact_text.setPlainText(d.get("your_contact") or "")
        self.invoice_number_entry.setText(d.get("invoice_number") or "")
        set_date_edit_from_db(self.invoice_date_edit, d.get("invoice_date"))
        self.payment_terms_entry.setText(d.get("payment_terms") or "")
        self.notes_text.setPlainText(d.get("notes") or "")
        self.no_charge_toggle.setChecked(bool(d.get("is_complimentary")))
        self.show_qr_toggle.setChecked(bool(d.get("show_qr_code", 1)))
        self.gst_toggle.setChecked(bool(d.get("is_gst_invoice", 1)))
        self.tax_rate_entry.setText(_num_text(d.get("tax_rate")))
        self.gst_type_combo.setCurrentText(d.get("gst_type") or "CGST/SGST")
        self.global_discount_entry.setText(_num_text(d.get("global_discount_value")))
        self.disc_type_combo.setCurrentText(d.get("global_discount_type") or "amount")
        self.tax_mode_combo.setCurrentIndex(max(0, self.tax_mode_combo.findData(d.get("tax_mode") or TAX_MODE_OVERALL)))
        self.round_off_toggle.setChecked(bool(d.get("round_off")))
        self.place_combo.setEditText(d.get("place_of_supply") or "")
        self._place_auto = False
        self.po_entry.setText(d.get("po_number") or "")
        self.template_combo.setCurrentIndex(max(0, self.template_combo.findData(d.get("pdf_template") or "")))
        self.reverse_toggle.setChecked(bool(d.get("reverse_charge")))
        self._show_stored_due(d.get("invoice_date"), d.get("due_date"))
        self.customer_notes_text.setPlainText(d.get("customer_notes") or "")
        self.terms_text.setPlainText(d.get("terms") or "")

        try:
            stored_items = json.loads(d.get("items") or "[]")
        except (TypeError, ValueError):
            stored_items = []
        for item in stored_items if isinstance(stored_items, list) else []:
            if isinstance(item, dict):
                self._add_item_row(item)
        if not self.item_rows:
            self._add_item_row()

        for p in self.db_manager.get_payments(invoice_id):
            self._add_payment_row(safe_float(p.get("amount")), p.get("payment_date"),
                                  p.get("mode") or "Cash", p.get("note") or "", payment_id=p.get("id"))

    # ------------------------------------------------------------------ #
    #  Validation & Save
    # ------------------------------------------------------------------ #
    def _error(self, title: str, message: str) -> None:
        Messagebox.show_error(title, message, parent=self.controller)

    def _save(self):
        inv_num = self.invoice_number_entry.text().strip()
        bill_to = self.client_edit.text().strip()
        your_name = self.your_name_entry.text().strip()
        if not inv_num or not bill_to or not your_name:
            return self._error("Validation", "Invoice Number, Client Name, and Your Name are all required.")
        if INVALID_FILENAME_CHARS.search(inv_num):
            return self._error(
                "Validation",
                'Invoice Number cannot contain any of:  \\  /  :  * ?  "  <  >  |\n'
                "(these characters break the PDF file that gets generated from it).",
            )
        if inv_num != self.editing_id and self.db_manager.get_invoice_by_id(inv_num) is not None:
            return self._error(
                "Duplicate Invoice Number",
                f'Invoice number "{inv_num}" is already used by another invoice.\n'
                "Please choose a different invoice number.",
            )

        client_gstin = self.client_gstin_entry.text().strip()
        if client_gstin and not is_valid_gstin(client_gstin):
            return self._error(
                "Invalid GST Number",
                f'"{client_gstin}" is not a valid GSTIN.\n'
                "A GSTIN is 15 characters, for example 22ABCDE1234F1Z5.",
            )

        # The date field is a calendar-backed QDateEdit: it can only ever hold a real date, so the Tk build's
        # "Invoice Date is required." / "is not a valid date." checks cannot be reached any more.
        inv_date = date_edit_db(self.invoice_date_edit)

        no_charge = self.no_charge_toggle.isChecked()
        tax_rate  = 0.0
        if not no_charge and self.gst_toggle.isChecked():
            tax_rate = safe_float(self.tax_rate_entry.text())
            if not (0 <= tax_rate <= 100):
                return self._error("Validation", "GST / Tax Rate must be between 0 and 100.")

        g_disc = safe_float(self.global_discount_entry.text())
        if g_disc < 0:
            return self._error("Validation", "Overall discount cannot be negative.")
        if self.disc_type_combo.currentText() == "%" and g_disc > 100:
            return self._error("Validation", "A percentage discount cannot exceed 100%.")

        due_date: Optional[str] = None
        if self._due_on() and not no_charge:
            due_date = date_edit_db(self.due_date_edit)
            if due_date < inv_date:
                return self._error("Validation", "The due date cannot be earlier than the invoice date.")

        items = self._filled_items()
        if not no_charge and not items:
            return self._error("Validation", "Add at least one item to the invoice.")
        per_item = self._per_item_active()
        for i in items:
            name = i["title"] or "an item"
            if i["amount"] < 0 or i["discount"] < 0:
                return self._error("Validation", "Item amounts and discounts cannot be negative.")
            if i.get("qty", 1.0) <= 0:
                return self._error("Validation", f'The quantity of "{name}" must be greater than zero.')
            if i["discount"] > i["amount"]:
                return self._error(
                    "Validation",
                    f'The discount on "{name}" cannot be larger than its amount.',
                )
        for i in items:
            if i.get("hsn") and not 2 <= len(i["hsn"]) <= 8:
                return self._error("Validation", f'The HSN / SAC code of "{i["title"] or "an item"}" must be 2 to 8 digits.')
        if per_item:
            for row in self.item_rows:
                if row.is_filled() and not 0 <= row.raw_gst() <= 100:
                    return self._error("Validation", f'The GST rate of "{row.title_edit.text().strip() or "an item"}" '
                                                     "must be between 0 and 100.")

        payments = []
        if not no_charge:
            payments = [p for p in self._collect_payments() if p["amount"] > 0]

        data = {
            "invoice_number": inv_num, "bill_to": bill_to,
            "client_contact": self.client_contact_text.toPlainText().strip(),
            "client_gstin": client_gstin,
            "your_name": your_name,
            "your_contact": self.your_contact_text.toPlainText().strip(),
            "items": json.dumps(items), "amount_paid": 0.0, "tax_rate": tax_rate,
            "invoice_date": inv_date, "payment_date": None,
            "status": InvoiceStatus.DUE.value, "payment_terms": self.payment_terms_entry.text().strip(),
            "payment_mode": None, "show_qr_code": int(self.show_qr_toggle.isChecked()),
            "is_complimentary": int(no_charge), "is_gst_invoice": int(self.gst_toggle.isChecked()),
            "global_discount_value": g_disc, "global_discount_type": self.disc_type_combo.currentText(),
            "advance_amount": 0.0, "advance_date": None, "gst_type": self.gst_type_combo.currentText(),
            "notes": self.notes_text.toPlainText().strip(),
            "tax_mode": TAX_MODE_PER_ITEM if per_item else TAX_MODE_OVERALL,
            "round_off": int(self.round_off_toggle.isChecked() and not no_charge),
            "place_of_supply": self.place_combo.currentText().strip(),
            "po_number": self.po_entry.text().strip(),
            "pdf_template": self.template_combo.currentData() or "",
            "due_date": due_date,
            "reverse_charge": int(self.reverse_toggle.isChecked()),
            "customer_notes": self.customer_notes_text.toPlainText().strip(),
            "terms": self.terms_text.toPlainText().strip(),
        }

        is_update = self.editing_id is not None
        ok, err = self.db_manager.save_invoice(data, is_update, original_id=self.editing_id,
                                               payments=None if no_charge else payments)
        if not ok:
            return self._error("Save Failed", err)

        if bill_to:
            self.db_manager.save_client(bill_to, data["client_contact"], data["client_gstin"])
        for item in items:
            if item["title"]:                      # what this line teaches the catalog: its rate in this unit, HSN/SAC, GST rate
                self.db_manager.remember_catalog_item(item["title"], item["description"],
                                                      safe_float(item.get("rate", item["amount"])),
                                                      item.get("unit", ""), item.get("hsn", ""), item.get("tax_rate"))

        Messagebox.show_info("Saved", f"Invoice #{inv_num} saved successfully.", parent=self.controller)
        self._initial_snapshot = self._snapshot()
        self.controller.show_frame(DashboardFrame)

    # ------------------------------------------------------------------ #
    #  Public API used by application-level keyboard shortcuts
    # ------------------------------------------------------------------ #
    def save(self):
        """Ctrl+S: validates and saves the invoice, exactly like the Save button."""
        self._save()

    def _focused_item_row(self) -> Optional[_ItemRow]:
        """Walks up from the current keyboard focus to find which item row (if any) owns it, so shortcuts can act
        on "whichever row the user is in" the same way F2/Ctrl+Shift+Delete do in a spreadsheet."""
        widget = QApplication.focusWidget()
        while widget is not None:
            if isinstance(widget, _ItemRow) and widget in self.item_rows:
                return widget
            widget = widget.parentWidget()
        return None

    def _show_row(self, widget: QWidget):
        """Focus a field and make sure it is on screen. A freshly added row only gets its geometry (and the page its
        new scroll range) when the pending layout request is processed, so that is done before scrolling."""
        widget.setFocus()
        QTimer.singleShot(0, lambda: self._reveal(widget))

    def _reveal(self, widget: QWidget):
        if widget.isVisible():
            QCoreApplication.sendPostedEvents(None, int(QEvent.Type.LayoutRequest))
            self.scroll.ensureWidgetVisible(widget, 0, 60)

    def add_item(self):
        """Ctrl+Shift+A / Ctrl+Plus: adds a new blank item row and moves focus straight into it, ready to type."""
        row = self._add_item_row()
        self._show_row(row.title_edit)

    def remove_selected_item(self):
        """Ctrl+Shift+Delete / Ctrl+Minus: removes whichever item row currently has focus, or the last row if none
        does -- the same row a row's own "✕" button would remove."""
        row = self._focused_item_row() or (self.item_rows[-1] if self.item_rows else None)
        if row is not None:
            self._remove_item_row(row)

    def edit_selected_item(self):
        """F2: jumps focus into the title field of the currently focused/last item row, ready to edit."""
        row = self._focused_item_row() or (self.item_rows[-1] if self.item_rows else None)
        if row is not None:
            self._show_row(row.title_edit)


# =============================================================================
#  16. APPLICATION SHELL
# =============================================================================
class SplashScreen(QWidget):
    SIZE = 420
    FADE_STEP = 0.05
    FADE_INTERVAL_MS = 15
    FADE_TARGET = 0.97          # a very subtle, modern transparency
    MESSAGE_INTERVAL_MS = 400   # six messages: the last one appears right as the 2.5 s splash ends
    LOADING_MESSAGES: List[str] = [
        "Waking up the database...",
        "Calculating dashboard stats...",
        "Preparing invoice registry...",
        "Loading your address book...",
        "Polishing the interface...",
        "Almost ready...",
    ]

    def __init__(self):
        super().__init__(None, Qt.WindowType.SplashScreen | Qt.WindowType.FramelessWindowHint
                         | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.setWindowTitle(APP_NAME)
        self.setFixedSize(self.SIZE, self.SIZE)
        self.setWindowOpacity(0.0)                  # starts invisible for the fade-in
        self.loading_messages = list(self.LOADING_MESSAGES)
        self.msg_index = 0
        self._build()
        self._center_on_screen()

        self._fade_timer = QTimer(self)
        self._fade_timer.setInterval(self.FADE_INTERVAL_MS)
        self._fade_timer.timeout.connect(self._fade_in)
        self._fade_timer.start()
        self._message_timer = QTimer(self)
        self._message_timer.setInterval(self.MESSAGE_INTERVAL_MS)
        self._message_timer.timeout.connect(self._cycle_messages)
        self._cycle_messages()

    def _center_on_screen(self) -> None:
        screen = QApplication.primaryScreen()
        if screen is not None:
            geo = screen.geometry()
            self.move(geo.center() - QPoint(self.SIZE // 2, self.SIZE // 2))

    def _build(self):
        # A 2 px border in the accent colour: a frameless window has no OS outline or shadow, and this keeps the splash
        # from bleeding into a white desktop.
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        border = QFrame()
        border.setProperty("role", "splash-border")
        outer.addWidget(border)
        border_col = QVBoxLayout(border)
        border_col.setContentsMargins(2, 2, 2, 2)
        panel = QFrame()
        panel.setProperty("role", "splash")
        border_col.addWidget(panel)

        col = QVBoxLayout(panel)
        col.setContentsMargins(0, 0, 0, 0)
        col.setSpacing(0)

        # Branding, a little above the exact centre for better visual balance.
        col.addStretch(9)
        logo = _logo_pixmap_for(self, 160)
        self.logo_label: Optional[QLabel] = None
        if logo is not None:
            self.logo_label = QLabel()
            self.logo_label.setPixmap(logo)
            self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            col.addWidget(self.logo_label)
            col.addSpacing(20)
        col.addWidget(make_label(APP_NAME, "display", align=Qt.AlignmentFlag.AlignCenter))
        col.addSpacing(5)
        col.addWidget(make_label(APP_TAGLINE, "muted", align=Qt.AlignmentFlag.AlignCenter))
        col.addStretch(11)

        # Loading status.
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)                # indeterminate
        self.progress.setTextVisible(False)
        bar_row = QHBoxLayout()
        bar_row.setContentsMargins(20, 0, 20, 0)
        bar_row.addWidget(self.progress)
        col.addLayout(bar_row)

        bottom = QHBoxLayout()
        bottom.setContentsMargins(20, 12, 20, 15)
        self.status_lbl = make_label("Starting up...", "section")
        bottom.addWidget(self.status_lbl)
        bottom.addStretch(1)
        bottom.addWidget(make_label(APP_VERSION, "small"))
        col.addLayout(bottom)

    def _fade_in(self) -> None:
        opacity = min(self.FADE_TARGET, self.windowOpacity() + self.FADE_STEP)
        self.setWindowOpacity(opacity)
        if opacity >= self.FADE_TARGET:
            self._fade_timer.stop()

    def _cycle_messages(self) -> None:
        """Walks through the friendly loading messages so the app feels alive while it starts."""
        if self.msg_index < len(self.loading_messages):
            self.status_lbl.setText(self.loading_messages[self.msg_index])
            self.msg_index += 1
            self._message_timer.start()
        else:
            self._message_timer.stop()

    def closeEvent(self, event):
        self._fade_timer.stop()
        self._message_timer.stop()
        super().closeEvent(event)


class Sidebar(QFrame):
    """The navigation column. It has three looks, chosen by the window it sits in (see InvoiceApp._apply_density): the
    full one (240 px wide), a compact one that shows only the icons (72 px wide, for narrow windows) and a short one that
    drops the logo (for windows that are not tall enough for it). Every button keeps its full name as its tooltip and
    its accessible name, so nothing is lost when the labels are hidden."""
    NAV_ITEMS: List[Tuple[str, str]] = [
        ("⊞   Dashboard", "DashboardFrame"),
        ("👥   Clients",   "ClientsFrame"),
        ("📑   Catalog",   "CatalogFrame"),
    ]
    WIDTH = 240
    COMPACT_WIDTH = 72
    LOGO_WIDTH = 160
    COMPACT_LOGO_WIDTH = 40

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.setProperty("role", "sidebar")
        self.setFixedWidth(self.WIDTH)
        self._nav_btns: Dict[str, QPushButton] = {}
        self._compact = False
        self._short = False
        self._wide_only: List[QWidget] = []                       # labels that only the full sidebar shows
        self._icon_buttons: List[Tuple[QPushButton, str, str, str]] = []   # (button, full text, icon, tooltip)
        self._margin_layouts: List[Tuple[QVBoxLayout, Tuple[int, int], Tuple[int, int]]] = []   # (layout, full v, short v)
        self._build()

    @property
    def compact(self) -> bool:
        return self._compact

    @property
    def short(self) -> bool:
        return self._short

    @staticmethod
    def _left_aligned(btn: QPushButton) -> QPushButton:
        btn.setProperty("align", "left")
        return btn

    def _register_button(self, btn: QPushButton) -> QPushButton:
        """Remembers how a button looks in each density: its text splits into the icon (the first word) and the name."""
        full = btn.text()
        icon, _sep, name = full.partition(" ")
        name = name.strip() or full
        btn.setAccessibleName(name)
        self._icon_buttons.append((btn, full, icon, btn.toolTip()))
        return btn

    def _layout(self, full_v: Tuple[int, int], short_v: Optional[Tuple[int, int]] = None) -> QVBoxLayout:
        lay = QVBoxLayout()
        self._margin_layouts.append((lay, full_v, short_v or full_v))
        lay.setContentsMargins(20, full_v[0], 20, full_v[1])
        return lay

    def _build(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # --- branding header ---------------------------------------------
        header = self._layout((25, 10), (10, 6))
        header.setSpacing(0)
        self.logo_label: Optional[QLabel] = None
        if _logo_pixmap_for(self, self.LOGO_WIDTH) is not None:
            self.logo_label = QLabel()
            self.logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            header.addWidget(self.logo_label)
            header.addSpacing(10)
        self.name_label = make_label("Vyaprix", "h1", align=Qt.AlignmentFlag.AlignCenter)
        self.byline_label = make_label("by EllowLabs", "section", align=Qt.AlignmentFlag.AlignCenter)
        self.mark_label = make_label("V", "h1", align=Qt.AlignmentFlag.AlignCenter)       # stands in for the name when compact
        self.mark_label.setVisible(False)
        header.addWidget(self.name_label)
        header.addWidget(self.byline_label)
        header.addWidget(self.mark_label)
        self._wide_only += [self.name_label, self.byline_label]
        root.addLayout(header)

        divider = self._layout((15, 15), (8, 8))
        divider.addWidget(hline())
        root.addLayout(divider)

        # --- quick action --------------------------------------------------
        quick = self._layout((0, 15), (0, 8))
        self.new_invoice_btn = make_button(
            "＋  New Invoice", "success", lambda: self.controller.show_frame_by_name("InvoiceFormFrame"),
            tooltip="Instantly start a new invoice (Ctrl+N)", min_height=44)
        self._register_button(self.new_invoice_btn)
        quick.addWidget(self.new_invoice_btn)
        root.addLayout(quick)

        # --- main navigation -----------------------------------------------
        nav = self._layout((0, 0))
        nav.setSpacing(6)
        menu_label = make_label("MAIN MENU", "section")
        menu_label.setContentsMargins(4, 0, 0, 2)
        nav.addWidget(menu_label)
        self._wide_only.append(menu_label)
        for label, frame_name in self.NAV_ITEMS:
            btn = make_button(label, "secondary-outline", lambda n=frame_name: self.controller.show_frame_by_name(n))
            self._left_aligned(btn)
            self._register_button(btn)
            nav.addWidget(btn)
            self._nav_btns[frame_name] = btn
        root.addLayout(nav)

        root.addStretch(1)

        # --- system / bottom -----------------------------------------------
        bottom = self._layout((15, 15), (8, 8))
        bottom.setSpacing(4)
        system_label = make_label("SYSTEM", "section")
        system_label.setContentsMargins(4, 0, 0, 4)
        bottom.addWidget(system_label)
        self._wide_only.append(system_label)
        self.settings_btn = self._left_aligned(make_button("⚙️  Settings", "info-outline",
                                                           lambda: self.controller.open_settings()))
        self.shortcuts_btn = self._left_aligned(make_button("⌨️  Shortcuts", "info-outline",
                                                            lambda: self.controller.open_shortcuts()))
        self.about_btn = self._left_aligned(make_button("ℹ️  About", "info-outline",
                                                        lambda: self.controller.open_about()))
        for btn in (self.settings_btn, self.shortcuts_btn, self.about_btn):
            self._register_button(btn)
            bottom.addWidget(btn)
        bottom.addSpacing(11)
        self.version_label = make_label(f"Version {APP_VERSION}", "small", align=Qt.AlignmentFlag.AlignCenter)
        self.version_short_label = make_label(APP_VERSION, "small", align=Qt.AlignmentFlag.AlignCenter)
        self.version_short_label.setVisible(False)
        bottom.addWidget(self.version_label)
        bottom.addWidget(self.version_short_label)
        self._wide_only.append(self.version_label)
        root.addLayout(bottom)
        self._refresh_logo()

    def set_density(self, compact: bool, short: bool) -> None:
        """Switches between the full, the compact (icons only) and the short (no logo) look; does nothing when the look
        is already the requested one."""
        compact, short = bool(compact), bool(short)
        if (compact, short) == (self._compact, self._short):
            return
        self._compact, self._short = compact, short
        self.setFixedWidth(self.COMPACT_WIDTH if compact else self.WIDTH)
        side = 8 if compact else 20
        for lay, full_v, short_v in self._margin_layouts:
            top, bottom = short_v if short else full_v
            lay.setContentsMargins(side, top, side, bottom)
        for widget in self._wide_only:
            widget.setVisible(not compact)
        self.version_short_label.setVisible(compact)
        for btn, full, icon, tip in self._icon_buttons:
            btn.setText(icon if compact else full)
            btn.setToolTip(f"{btn.accessibleName()}  —  {tip}" if (compact and tip) else (btn.accessibleName() if compact else tip))
            if btn.property("align") != ("icon" if compact else "left") and btn is not self.new_invoice_btn:
                btn.setProperty("align", "icon" if compact else "left")
                _repolish(btn)
        self._refresh_logo()
        self.updateGeometry()

    def _refresh_logo(self) -> None:
        """The logo at the size the current look wants, rendered for the screen the window is on now (so it stays sharp
        after the window is dragged to a monitor with another scale). With no logo file, compact shows a "V" instead."""
        show_logo = self.logo_label is not None and not self._short
        if self.logo_label is not None:
            self.logo_label.setVisible(show_logo)
            if show_logo:
                pixmap = _logo_pixmap_for(self, self.COMPACT_LOGO_WIDTH if self._compact else self.LOGO_WIDTH)
                if pixmap is not None:
                    self.logo_label.setPixmap(pixmap)
        self.mark_label.setVisible(self._compact and not show_logo)

    def refresh_for_screen(self) -> None:
        """Called when the window's device pixel ratio changes (it was moved to another monitor)."""
        self._refresh_logo()

    def set_active(self, frame_name: str) -> None:
        """Highlights the navigation button of the page being shown: solid for the active page, outlined for the rest."""
        for name, btn in self._nav_btns.items():
            set_role(btn, "primary" if name == frame_name else "secondary-outline")


class ShortcutManager:
    """One place that defines what every app-wide key combination does. A shortcut either always means the same thing
    (register_global) or means something different on each page (register_contextual: Ctrl+N is "new invoice" on the
    Dashboard, "new client" on Clients, "new catalog item" on Catalog). Registering the same sequence twice for the same
    scope -- or mixing a global and a contextual registration of one sequence -- raises ValueError at once, so a
    copy-paste slip is caught the moment the app starts instead of showing up later as "that key stopped working".

    Qt notes (what replaced the Tk build's machinery):
      * exactly one QShortcut per key sequence, owned by the main window with the WindowShortcut context. Routing to the
        right handler for whichever page is showing happens in _dispatch(), exactly as in the Tk build.
      * WindowShortcut also gives "no firing under an open dialog" for free: a QShortcut is only live while the main
        window is the active window, and Settings / About / Shortcuts / the PDF viewer are windows of their own. The
        Tk build needed an explicit toplevel check for this.
      * the Tk build had to re-bind Tk's own Ctrl+D / Ctrl+O class bindings (install_class_overrides); Qt text fields
        have no such bindings, and the editing keys they do own (Ctrl+A = select all, Ctrl+C/V/X/Z/Y ...) win while a
        text field has the focus, via Qt's ShortcutOverride -- which is what a user typing in a box expects.
      * Qt key-sequence strings ("Ctrl+N", "F2") already are the readable form shown in the reference card, so no
        _humanize_sequence(); the few keys with several physical forms (Ctrl+Plus) pass their own display= text.
      * auto-repeat is off by default (holding Ctrl+D must not stamp out dozens of duplicate invoices); the two
        row-adding / removing shortcuts on the invoice form switch it on, since repeating them is useful.
    """

    def __init__(self, app: "InvoiceApp"):
        self.app = app
        self._owner: Dict[str, str] = {}                                # sequence -> "global" | "contextual"
        self._global_handlers: Dict[str, Callable[[], None]] = {}
        self._contextual_handlers: Dict[str, Dict[Type, Callable[[], None]]] = {}
        self._qshortcuts: Dict[str, QShortcut] = {}
        self._reference: List[Dict[str, str]] = []                      # feeds ShortcutsDialog

    @staticmethod
    def canonical(sequence: str) -> str:
        """The portable text of a key sequence ("ctrl+n" -> "Ctrl+N"); ValueError when Qt does not understand it."""
        text = QKeySequence(sequence).toString(QKeySequence.SequenceFormat.PortableText)
        if not text:
            raise ValueError(f"Not a valid key sequence: {sequence!r}")
        return text

    @classmethod
    def _sequences(cls, sequences) -> Tuple[str, ...]:
        seqs = (sequences,) if isinstance(sequences, str) else tuple(sequences)
        return tuple(cls.canonical(s) for s in seqs)

    # ---- registration --------------------------------------------------- #
    def register_global(self, sequences, handler: Callable[[], None], description: str,
                        group: str, display: Optional[str] = None, auto_repeat: bool = False):
        """Registers a shortcut whose action never changes, whichever page is showing. `sequences` is one key sequence
        or a tuple of equivalent ones (they are one shortcut, so they share one reference-card entry)."""
        seqs = self._sequences(sequences)
        for s in seqs:
            if s in self._owner:
                raise ValueError(f"Shortcut conflict: {s} is already registered ({self._owner[s]}).")
        for s in seqs:
            self._owner[s] = "global"
            self._global_handlers[s] = handler
            self._bind(s, auto_repeat)
        self._reference.append({"keys": display or seqs[0], "description": description, "group": group})

    def register_contextual(self, sequences, frame_class: Type, handler: Callable[[], None],
                            description: str, group: str, display: Optional[str] = None,
                            show_in_reference: bool = True, auto_repeat: bool = False):
        """Registers a shortcut whose action depends on which page is showing. The same sequence can be registered for
        several page classes (that is the point), but not twice for the same one, and never mixed with a global
        registration of the same sequence."""
        seqs = self._sequences(sequences)
        for s in seqs:
            if self._owner.get(s, "contextual") != "contextual":
                raise ValueError(f"Shortcut conflict: {s} is already registered ({self._owner[s]}).")
            if frame_class in self._contextual_handlers.get(s, {}):
                raise ValueError(f"Shortcut conflict: {s} is already registered for {frame_class.__name__}.")
        for s in seqs:
            self._owner[s] = "contextual"
            self._contextual_handlers.setdefault(s, {})[frame_class] = handler
            self._bind(s, auto_repeat)
        if show_in_reference:
            page = frame_class.__name__.replace("Frame", "") or frame_class.__name__
            self._reference.append({"keys": display or seqs[0],
                                    "description": f"{description} ({page})", "group": group})

    def register_info(self, display: str, description: str, group: str):
        """Documents a key that a widget handles natively (Tab order, per-field Enter-to-advance ...) rather than through
        this registry -- purely so the reference card stays the one complete list of what the keyboard can do."""
        self._reference.append({"keys": display, "description": description, "group": group})

    def _bind(self, sequence: str, auto_repeat: bool = False):
        existing = self._qshortcuts.get(sequence)
        if existing is not None:
            existing.setAutoRepeat(existing.autoRepeat() or auto_repeat)
            return
        shortcut = QShortcut(QKeySequence(sequence), self.app)
        shortcut.setContext(Qt.ShortcutContext.WindowShortcut)
        shortcut.setAutoRepeat(auto_repeat)
        shortcut.activated.connect(lambda seq=sequence: self._dispatch(seq))
        self._qshortcuts[sequence] = shortcut

    # ---- dispatch -------------------------------------------------------- #
    def _dispatch(self, sequence: str) -> None:
        handler = self._global_handlers.get(sequence)
        if handler is None:
            per_frame = self._contextual_handlers.get(sequence)
            if per_frame:
                handler = per_frame.get(self.app.active_frame_class())
        if handler:
            handler()

    # ---- introspection, used by ShortcutsDialog --------------------------- #
    def grouped_reference(self) -> Dict[str, List[Tuple[str, str]]]:
        """{group name: [(key label, description), ...]} built straight from the live registry, in registration order,
        so the Shortcuts reference can never drift out of sync with what the keys actually do."""
        groups: Dict[str, List[Tuple[str, str]]] = {}
        for entry in self._reference:
            groups.setdefault(entry["group"], []).append((entry["keys"], entry["description"]))
        return groups


class InvoiceApp(QMainWindow):
    SPLASH_DURATION_MS = 2500
    POLL_INTERVAL_MS = 200
    DEFAULT_SIZE = (1280, 800)
    MIN_SIZE = (1080, 680)          # the smallest window the layout was designed for
    FLOOR_SIZE = (880, 540)         # the smallest it shrinks to on a screen that has no room for MIN_SIZE
    COMPACT_BELOW = 1080            # a window narrower than this shows the icons-only sidebar
    SHORT_BELOW = 760               # a window lower than this drops the logo from the sidebar

    def __init__(self):
        super().__init__()
        cfg = load_config()
        apply_app_style(QApplication.instance(), cfg.get("theme", DEFAULT_THEME))
        configure_windows_app_identity()
        self._closing_for_good = False

        # 1. SHOW THE SPLASH SCREEN first (the main window stays hidden until start-up is done) and let it paint now.
        self.splash: Optional[SplashScreen] = SplashScreen()
        self.splash.show()
        QApplication.processEvents()

        self.setWindowTitle(f"{APP_NAME} • {APP_TAGLINE}")
        self._place_on_screen()
        self._set_icon()

        self.db_manager = DatabaseManager(DB_NAME)
        self.task_queue: queue.Queue = queue.Queue()

        self._current_frame_class: Optional[Type] = None
        self._build_shell()
        self._setup_keyboard_shortcuts()
        self.show_frame(DashboardFrame)

        # Worker threads hand results back through task_queue; this timer picks them up on the GUI thread. It is
        # re-armed after each pass (not a repeating timer) so a blocking message box inside a pass can never re-enter it.
        self._poll_timer = QTimer(self)
        self._poll_timer.setSingleShot(True)
        self._poll_timer.timeout.connect(self._poll_queue)
        self._poll_timer.start(self.POLL_INTERVAL_MS)

        # 2. SCHEDULE THE HAND-OVER from splash to main window.
        self._splash_timer = QTimer(self)
        self._splash_timer.setSingleShot(True)
        self._splash_timer.timeout.connect(self._close_splash_and_show_main)
        self._splash_timer.start(self.SPLASH_DURATION_MS)

    # ------------------------------------------------------------------ #
    #  Window set-up / start-up hand-over
    # ------------------------------------------------------------------ #
    def _usable_area(self) -> Optional[QRect]:
        """The part of the screen this window is on that is not taken by the taskbar (None when there is no screen)."""
        return screen_area(self)

    def _minimum_for(self, area: Optional[QRect]) -> Tuple[int, int]:
        """The smallest window allowed on a screen with this usable area: the designed 1080 x 680 where the screen has
        room for it, smaller (down to FLOOR_SIZE, with the compact sidebar) on a small one -- so the window never hangs
        off the edge of a laptop display where Save and Back would be out of reach."""
        min_w, min_h = self.MIN_SIZE
        if area is not None:
            min_w = max(self.FLOOR_SIZE[0], min(min_w, area.width()))
            min_h = max(self.FLOOR_SIZE[1], min(min_h, area.height()))
        return min_w, min_h

    def _frame_extra(self) -> Tuple[int, int]:
        """How much the window frame (borders and title bar) adds to the client size. The system only says how thick the
        frame is once the native window exists, so it is created here (not shown) when it is not there yet."""
        if self.windowHandle() is None:
            self.create()
        frame, client = self.frameGeometry(), self.geometry()
        return max(0, frame.width() - client.width()), max(0, frame.height() - client.height())

    @staticmethod
    def _client_room(area: Optional[QRect], extra: Tuple[int, int]) -> Optional[QRect]:
        """The part of `area` the window's client size may use: what is left once the frame has taken its share."""
        if area is None:
            return None
        return QRect(area.x(), area.y(), max(1, area.width() - extra[0]), max(1, area.height() - extra[1]))

    def _place_on_screen(self, area: Optional[QRect] = None) -> None:
        """1280 x 800, centred on the screen it opens on -- never larger than that screen's usable area (the title bar and
        borders included, so a 1366 x 768 laptop does not push the window's bottom behind the taskbar), and never below
        the minimum that screen allows (1080 x 680 where it fits)."""
        area = area if area is not None else self._usable_area()
        extra = self._frame_extra()
        room = self._client_room(area, extra)
        min_w, min_h = self._minimum_for(room)
        width, height = self.DEFAULT_SIZE
        if room is not None:
            width, height = max(min_w, min(width, room.width())), max(min_h, min(height, room.height()))
        self.setMinimumSize(min_w, min_h)
        self.resize(width, height)
        if area is not None:
            self.move(area.x() + (area.width() - (width + extra[0])) // 2, area.y() + (area.height() - (height + extra[1])) // 2)
        self._apply_density()

    def _fit_into(self, area: Optional[QRect]) -> None:
        """Brings the window -- its frame included -- inside `area` (the usable part of a screen): the limits follow the
        screen, a window that is too big shrinks and one that hangs over an edge moves back. A maximized or full-screen
        window is left alone."""
        if area is None:
            return
        extra = self._frame_extra()
        min_w, min_h = self._minimum_for(self._client_room(area, extra))
        self.setMinimumSize(min_w, min_h)
        if self.isMaximized() or self.isFullScreen():
            return
        width = max(min_w, min(self.width(), area.width() - extra[0]))
        height = max(min_h, min(self.height(), area.height() - extra[1]))
        if (width, height) != (self.width(), self.height()):
            self.resize(width, height)
        frame = self.frameGeometry()
        x = max(area.x(), min(frame.x(), area.x() + area.width() - frame.width()))
        y = max(area.y(), min(frame.y(), area.y() + area.height() - frame.height()))
        if (x, y) != (frame.x(), frame.y()):
            self.move(x, y)

    def _apply_density(self) -> None:
        """The sidebar follows the window: icons only when the window is narrower than COMPACT_BELOW, no logo when it is
        lower than SHORT_BELOW."""
        sidebar = getattr(self, "sidebar", None)
        if sidebar is not None:
            sidebar.set_density(self.width() < self.COMPACT_BELOW, self.height() < self.SHORT_BELOW)

    def _on_screen_changed(self, screen) -> None:
        """The window was dragged to another monitor: fit the limits to its usable area and shrink the window into it if
        it no longer fits (a maximized or full-screen window is left alone)."""
        self._fit_into(screen.availableGeometry() if screen is not None else None)
        self._apply_density()

    def showEvent(self, event):
        super().showEvent(event)
        handle = self.windowHandle()
        if handle is not None and not getattr(self, "_screen_watch", False):
            self._screen_watch = True
            handle.screenChanged.connect(self._on_screen_changed)
            QTimer.singleShot(0, self._fit_shown_window)         # now that the window is up, its real frame is known

    def _fit_shown_window(self) -> None:
        """Once, right after the first show: the frame measured before the window existed may have been an estimate."""
        if not self._closing_for_good:
            self._fit_into(self._usable_area())

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._apply_density()

    def event(self, event):
        if event.type() == QEvent.Type.DevicePixelRatioChange and hasattr(self, "sidebar"):
            self.sidebar.refresh_for_screen()
        return super().event(event)

    def _set_icon(self) -> None:
        candidates = ([APP_LOGO_ICO] if sys.platform == "win32" else []) + [APP_LOGO_PNG]
        for path in candidates:
            try:
                if path.exists():
                    icon = QIcon(str(path))
                    if not icon.isNull():
                        self.setWindowIcon(icon)
                        QApplication.setWindowIcon(icon)
                        return
            except Exception:
                logger.warning("Could not load the application icon %s", path, exc_info=True)

    def _close_splash_and_show_main(self) -> None:
        # Show the main window BEFORE closing the splash: closing the last visible window would quit the application.
        self.show()
        self.raise_()
        self.activateWindow()
        self._close_splash()
        self._warn_if_fallback_font()

    def _close_splash(self) -> None:
        splash, self.splash = self.splash, None
        if splash is not None:
            try:
                splash.close()
            except RuntimeError:
                pass            # Qt already closed and deleted it (QApplication.quit() closes every visible window first)

    def _warn_if_fallback_font(self) -> None:
        if PDF_FONT == "Helvetica":
            Messagebox.show_warning(
                "Font Files Missing",
                "Vyaprix could not load its bundled font (NotoSans).\n\n"
                "PDFs will still generate, but the currency symbol "
                f"({CURRENCY_SYMBOL}) and any non-English characters may "
                "not display correctly until this is fixed.\n\n"
                "Try reinstalling Vyaprix. Details were written to the log:\n"
                f"{LOG_FILE}",
                parent=self,
            )

    # ------------------------------------------------------------------ #
    #  Shell: sidebar | divider | (page stack, progress strip)
    # ------------------------------------------------------------------ #
    def _build_shell(self) -> None:
        central = QWidget()
        outer = QHBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        self.sidebar = Sidebar(central, self)
        outer.addWidget(self.sidebar)
        outer.addWidget(vline())
        self._apply_density()

        right = QWidget()
        column = QVBoxLayout(right)
        column.setContentsMargins(0, 0, 0, 0)
        column.setSpacing(0)
        outer.addWidget(right, 1)

        self.container = QStackedWidget()
        column.addWidget(self.container, 1)

        # Progress strip under the pages -- only visible while a background job runs.
        self.progress_panel = QWidget()
        strip = QVBoxLayout(self.progress_panel)
        strip.setContentsMargins(16, 8, 16, 8)
        strip.setSpacing(2)
        self.progress_lbl = make_label("", "small")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)            # indeterminate: animates by itself while visible
        self.progress_bar.setTextVisible(False)
        strip.addWidget(self.progress_lbl)
        strip.addWidget(self.progress_bar)
        self.progress_panel.setVisible(False)
        column.addWidget(self.progress_panel)

        self.frames: Dict[Type, QWidget] = {}
        for FrameClass in (DashboardFrame, ClientsFrame, CatalogFrame, InvoiceFormFrame):
            frame = FrameClass(self.container, self)
            self.frames[FrameClass] = frame
            self.container.addWidget(frame)
        self.setCentralWidget(central)

    # ------------------------------------------------------------------ #
    #  Keyboard shortcuts
    # ------------------------------------------------------------------ #
    def _setup_keyboard_shortcuts(self) -> None:
        """Builds the ShortcutManager and registers every keyboard shortcut through it -- the single place that defines
        what every key combo does, in every context. See ShortcutManager for the safety behaviours every registration
        gets automatically."""
        sc = self.shortcuts = ShortcutManager(self)

        # ---- Ctrl+N: context-aware "create a new thing for this page" ---
        # Registered once per page, each with its own handler and reference-card line -- this is what makes one
        # physical key combo mean three different (correct) things depending on where you are, instead of an if/elif
        # chain buried in a handler.
        sc.register_contextual("Ctrl+N", DashboardFrame,   self._sc_new_invoice,      "New Invoice",      "Navigation", display="Ctrl+N")
        sc.register_contextual("Ctrl+N", ClientsFrame,     self._sc_new_client,       "New Client",       "Navigation", display="Ctrl+N")
        sc.register_contextual("Ctrl+N", CatalogFrame,     self._sc_new_catalog_item, "New Catalog Item", "Navigation", display="Ctrl+N")
        # Also meaningful while already on the invoice form (start another blank one); kept out of the reference card
        # to avoid a confusing near-duplicate line right under "New Invoice (Dashboard)".
        sc.register_contextual("Ctrl+N", InvoiceFormFrame, self._sc_new_invoice,      "New Invoice",      "Navigation", show_in_reference=False)

        # ---- Global: identical action wherever you are ----------------
        sc.register_global("Ctrl+O", self._sc_open_invoice,    "Open Invoice",    "Navigation", display="Ctrl+O")
        sc.register_global("Ctrl+Q", self._sc_quit,            "Quit",            "Navigation", display="Ctrl+Q")
        sc.register_global("Ctrl+B", self._sc_backup_database, "Backup Database", "Navigation", display="Ctrl+B")

        # ---- Contextual: Dashboard --------------------------------------
        sc.register_contextual("Ctrl+F", DashboardFrame, lambda: self.frames[DashboardFrame].focus_search(), "Search",            "Navigation", display="Ctrl+F")
        sc.register_contextual("Ctrl+D", DashboardFrame, self._sc_duplicate_invoice,                         "Duplicate Invoice", "Navigation", display="Ctrl+D")
        sc.register_contextual("Ctrl+P", DashboardFrame, self._sc_print_pdf,                                 "Print PDF",         "Navigation", display="Ctrl+P")
        sc.register_contextual("Ctrl+E", DashboardFrame, self._sc_export_pdf,                                "Export PDF",        "Navigation", display="Ctrl+E")
        sc.register_contextual("Ctrl+R", DashboardFrame, self._sc_refresh_dashboard,                         "Refresh List",      "Navigation", display="Ctrl+R")
        sc.register_contextual("Ctrl+A", DashboardFrame, self._sc_select_all_invoices,                       "Select All",        "Navigation", display="Ctrl+A")

        # ---- Contextual: Clients / Catalog search boxes ------------------
        sc.register_contextual("Ctrl+F", ClientsFrame, lambda: self.frames[ClientsFrame].focus_search(), "Search", "Navigation", display="Ctrl+F")
        sc.register_contextual("Ctrl+F", CatalogFrame, lambda: self.frames[CatalogFrame].focus_search(), "Search", "Navigation", display="Ctrl+F")

        # ---- Contextual: Invoice Form -------------------------------------
        sc.register_contextual("Ctrl+S", InvoiceFormFrame, self._sc_save_invoice, "Save Invoice", "Invoice Form", display="Ctrl+S")
        sc.register_contextual("Ctrl+Shift+A", InvoiceFormFrame, self._sc_add_item, "Add Item", "Invoice Form", display="Ctrl+Shift+A")
        sc.register_contextual(("Ctrl+=", "Ctrl++"),                    # Ctrl+Plus: the physical =/+ key, with or without Shift, or the keypad's +
                               InvoiceFormFrame, self._sc_add_item, "Add Item (Row)", "Invoice Form", display="Ctrl+Plus", auto_repeat=True)
        sc.register_contextual("Ctrl+Shift+Del", InvoiceFormFrame, self._sc_remove_item, "Delete Selected Item", "Invoice Form",
                               display="Ctrl+Shift+Del", auto_repeat=True)
        sc.register_contextual("Ctrl+-",                                # main-row minus and the keypad's -
                               InvoiceFormFrame, self._sc_remove_item, "Remove Item (Row)", "Invoice Form", display="Ctrl+Minus", auto_repeat=True)
        sc.register_contextual("F2", InvoiceFormFrame, self._sc_edit_item, "Edit Selected Item", "Invoice Form", display="F2")
        # Esc on the invoice form == clicking "Back to Dashboard": same method, same unsaved-changes confirmation.
        sc.register_contextual("Esc", InvoiceFormFrame, self._sc_back_to_dashboard, "Back to Dashboard", "Invoice Form", display="Esc")
        # Native Qt field traversal / per-widget Enter handling, not a central registration -- documented here anyway
        # so the reference card stays the one complete list of shortcuts.
        sc.register_info("Tab / Shift+Tab", "Move Between Fields", "Invoice Form")
        sc.register_info("↑ / ↓  •  Page Up / Down", "Scroll the Form", "Invoice Form")
        sc.register_info("Enter", "Next Field", "Invoice Form")

        # ---- Anywhere -------------------------------------------------------
        sc.register_global("F1", self._sc_show_shortcuts, "Show This Help", "Anywhere", display="F1")
        # Every dialog (Settings, About, PDF Viewer, Add Client, ...) handles its own Esc -> close/cancel.
        sc.register_info("Esc", "Close Dialog / PDF Viewer", "Anywhere")

    def active_frame_class(self) -> Optional[Type]:
        return self._current_frame_class

    def _active_frame(self) -> Optional[QWidget]:
        return self.frames.get(self._current_frame_class)

    # Handlers below are only ever invoked once ShortcutManager has confirmed the right context (or none, for
    # globals), so they need no isinstance / hasattr guards of their own: that check lives in the registrations above.
    def _sc_new_invoice(self):
        """Ctrl+N on the Dashboard, or on the invoice form itself: start a fresh invoice. When the invoice form is
        already open with unsaved edits, show_frame() asks before discarding them."""
        self.show_frame(InvoiceFormFrame)

    def _sc_new_client(self):
        self.frames[ClientsFrame].new_client()

    def _sc_new_catalog_item(self):
        self.frames[CatalogFrame].new_catalog_item()

    def _sc_save_invoice(self):
        self.frames[InvoiceFormFrame].save()

    def _sc_print_pdf(self):
        self.frames[DashboardFrame].prepare_selected_pdf()

    def _sc_export_pdf(self):
        self.frames[DashboardFrame].prepare_selected_pdf()

    def _sc_open_invoice(self):
        """Ctrl+O: always jumps to the invoice list and into its search box, regardless of which screen you are on --
        unlike Ctrl+F, which searches whatever screen is already active."""
        self.show_frame(DashboardFrame)
        if self._current_frame_class is DashboardFrame:
            self.frames[DashboardFrame].focus_search()

    def _sc_quit(self):
        self.close()

    def _sc_duplicate_invoice(self):
        self.frames[DashboardFrame].duplicate_selected()

    def _sc_refresh_dashboard(self):
        self.frames[DashboardFrame].refresh()

    def _sc_select_all_invoices(self):
        self.frames[DashboardFrame].select_all()

    def _sc_backup_database(self):
        prompt_and_backup_database(self)

    def _sc_add_item(self):
        self.frames[InvoiceFormFrame].add_item()

    def _sc_remove_item(self):
        self.frames[InvoiceFormFrame].remove_selected_item()

    def _sc_edit_item(self):
        self.frames[InvoiceFormFrame].edit_selected_item()

    def _sc_back_to_dashboard(self):
        """Esc on the invoice form: identical to the "Back to Dashboard" button -- show_frame() already asks about
        unsaved changes on its own when leaving the invoice form for a different page."""
        self.show_frame(DashboardFrame)

    def _sc_show_shortcuts(self):
        self.open_shortcuts()

    # ------------------------------------------------------------------ #
    #  Navigation
    # ------------------------------------------------------------------ #
    def show_frame(self, frame_class: Type, *args) -> None:
        # Give the currently visible page a chance to block the switch (an invoice form with unsaved edits). That includes
        # asking for the invoice form while it is already showing -- the sidebar's "New Invoice" button, Ctrl+N: opening
        # it again re-initialises the form, which would silently throw the edits away. (The Tk build only guarded Ctrl+N.)
        current_frame = self.frames.get(self._current_frame_class)
        if current_frame is not None and hasattr(current_frame, "confirm_navigate_away"):
            if not current_frame.confirm_navigate_away():
                return

        frame = self.frames[frame_class]
        self.container.setCurrentWidget(frame)
        # The invoice form belongs to the Dashboard in the menu: that button stays highlighted while it is open.
        self.sidebar.set_active(frame_class.__name__ if frame_class in
                                (DashboardFrame, ClientsFrame, CatalogFrame) else "DashboardFrame")
        self._current_frame_class = frame_class
        if hasattr(frame, "on_show"):
            frame.on_show(*args)

    def show_frame_by_name(self, name: str) -> None:
        mapping = {
            "DashboardFrame": DashboardFrame,
            "ClientsFrame": ClientsFrame,
            "CatalogFrame": CatalogFrame,
            "InvoiceFormFrame": InvoiceFormFrame,
        }
        self.show_frame(mapping[name])

    def open_settings(self) -> None:
        run_dialog(SettingsDialog(self, self._on_settings_saved))

    def open_about(self) -> None:
        run_dialog(AboutDialog(self))

    def open_shortcuts(self) -> None:
        run_dialog(ShortcutsDialog(self, self.shortcuts.grouped_reference()))

    def _on_settings_saved(self, new_theme: Optional[str] = None) -> None:
        # 1. LIVE THEME SWITCHING: one application-wide stylesheet + palette swap re-colours every widget that exists,
        #    including the multi-line boxes the Tk build had to walk and re-theme by hand (_retheme_text_widgets).
        if new_theme and normalize_theme_name(new_theme) != active_theme_name():
            apply_app_style(QApplication.instance(), new_theme)

        # 2. Refresh the globals so the currency symbol updates live.
        global CONFIG, CURRENCY_SYMBOL
        CONFIG = load_config()
        CURRENCY_SYMBOL = CONFIG.get("currency_symbol", "₹")

        # 3. Make every screen redraw with the new colours / symbol. The invoice form is special-cased: on_show()
        #    resets and reloads the whole form, which would silently wipe an in-progress invoice if the user opened
        #    Settings mid-edit. It gets a lightweight refresh that only updates currency/theme-dependent display.
        for frame in self.frames.values():
            if isinstance(frame, InvoiceFormFrame):
                frame.refresh_after_settings_change()
            elif hasattr(frame, "on_show"):
                try:
                    frame.on_show()
                except TypeError:
                    pass

    # ------------------------------------------------------------------ #
    #  Progress / background task plumbing
    # ------------------------------------------------------------------ #
    def show_progress(self, active: bool, message: str = "") -> None:
        if active:
            self.progress_lbl.setText(message)
        self.progress_panel.setVisible(bool(active))

    def _poll_queue(self) -> None:
        try:
            while True:
                msg = self.task_queue.get_nowait()
                self.show_progress(False)
                if msg["status"] == "success":
                    PDFViewer(self, msg["path"], invoice_data=msg.get("invoice_data")).show()
                    self.frames[DashboardFrame].on_show()
                else:
                    Messagebox.show_error("Error", msg.get("message", "An unknown error occurred."), parent=self)
        except queue.Empty:
            pass
        finally:
            if not self._closing_for_good:
                self._poll_timer.start(self.POLL_INTERVAL_MS)

    # ------------------------------------------------------------------ #
    #  Closing
    # ------------------------------------------------------------------ #
    def closeEvent(self, event) -> None:
        if not self._closing_for_good:
            frame = self._active_frame()
            if isinstance(frame, InvoiceFormFrame) and not frame.confirm_navigate_away():
                event.ignore()
                return
        self._poll_timer.stop()
        self._splash_timer.stop()
        self._close_splash()
        super().closeEvent(event)


# =============================================================================
#  17. START-UP, SHUTDOWN & ENTRY POINT
# =============================================================================
def startup_background_task(app_instance):
    """Put whatever heavy lifting you need at start-up here. It runs in a thread, so it never freezes the splash screen
    or the GUI. If it needs to update the GUI it MUST go through the app's task_queue, never touch a widget itself."""
    logger.info("Background thread started successfully.")


def _orderly_shutdown(app: QApplication) -> None:
    """Closes and deletes every window in order, then flushes the pending deletions, BEFORE the interpreter exits.
    PySide6 can crash while it destroys widgets at interpreter shutdown that are only kept alive by a Python slot
    (lambda) connected to one of their own signals; tearing everything down here, while the event loop machinery is
    still intact, avoids that. Qt-owned popups (completer lists, tooltips) are left alone: Qt frees those itself."""
    for window in app.topLevelWidgets():
        if isinstance(window, InvoiceApp):
            window._closing_for_good = True         # no "unsaved changes" prompt while the process is going away
        if window.windowType() in (Qt.WindowType.Popup, Qt.WindowType.ToolTip):
            continue
        window.close()
        window.deleteLater()
    app.sendPostedEvents(None, QEvent.Type.DeferredDelete)
    app.processEvents()


def _log_unhandled_exception(exc_type, exc_value, exc_tb) -> None:
    """sys.excepthook. PySide6 reports an exception raised inside a Qt slot or event handler through sys.excepthook and
    carries on running -- and a windowed (pythonw) build has no console to show it on. Writing it to the log file as well
    means "I clicked it and nothing happened" always leaves a trace to look into."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_tb)
        return
    try:
        logger.critical("Unhandled exception", exc_info=(exc_type, exc_value, exc_tb))   # file + console (stream handler)
    except Exception:
        sys.__excepthook__(exc_type, exc_value, exc_tb)


def _log_unhandled_thread_exception(args) -> None:
    """threading.excepthook: the same for an exception that escapes a worker thread."""
    if args.exc_type is SystemExit:
        return
    try:
        logger.critical("Unhandled exception in thread %s", getattr(args.thread, "name", "?"),
                        exc_info=(args.exc_type, args.exc_value, args.exc_traceback))
    except Exception:
        pass


def _close_splash_windows() -> None:
    """The splash stays on top of every other window, so it must go before an error box can be seen."""
    for widget in QApplication.topLevelWidgets():
        if isinstance(widget, SplashScreen):
            widget.close()


def configure_windows_app_identity() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
            "EllowLabs.Vyaprix"
        )
    except Exception as exc:
        logger.warning("Could not set Windows application identity: %s", exc)


def configure_high_dpi() -> None:
    """Per-monitor, exact-fraction scaling. Qt 6 already scales every window by the DPI of the screen it is on; this pins
    the rounding policy to "pass through", so a 125 % or 150 % Windows scale stays 125 % / 150 % instead of being rounded
    to 100 % / 200 % (which would make the whole window too small or too large), and asks Windows for per-monitor-v2
    awareness, so the window re-scales as it is dragged between monitors. Must run before the QApplication exists."""
    try:
        QGuiApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    except Exception:
        logger.warning("Could not set the high-DPI scale policy", exc_info=True)
    if sys.platform == "win32":
        os.environ.setdefault("QT_QPA_PLATFORM", "windows:dpiawareness=2")


def main() -> int:
    app = None
    try:
        if QApplication.instance() is None:
            configure_high_dpi()
        app = QApplication.instance() or QApplication(sys.argv)
        app.setApplicationName(APP_NAME)
        sys.excepthook = _log_unhandled_exception
        threading.excepthook = _log_unhandled_thread_exception
        window = InvoiceApp()

        # Start the background task as a daemon thread before the event loop runs. The event loop stays on the main thread.
        bg_thread = threading.Thread(target=startup_background_task, args=(window,), daemon=True)  # daemon: dies with the app
        bg_thread.start()

        exit_code = app.exec()
        del window
        return exit_code
    except Exception:
        logger.critical("Fatal error during application startup", exc_info=True)
        try:
            _close_splash_windows()
            QMessageBox.critical(None, "Vyaprix — Fatal Error",
                                 "Vyaprix could not start. See the log file for details:\n"
                                 f"{LOG_FILE}")
        except Exception:
            pass
        raise
    finally:
        if app is not None:
            _orderly_shutdown(app)


if __name__ == "__main__":
    sys.exit(main())
