# =============================================================================
#  Vyaprix By EllowLabs  –  Professional Edition v3.0
#  Complete rewrite: faster data layer, redesigned navigation, payment
#  history, dedicated client/catalog managers, bulk actions, invoice notes.
# =============================================================================

# --- 1. Standard Library Imports ---
import csv
import datetime
import io
import json
import logging
import os
import queue
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import tkinter as tk
from contextlib import closing
from decimal import Decimal, ROUND_HALF_UP
from enum import Enum
from logging.handlers import RotatingFileHandler
from pathlib import Path
from tkinter import filedialog
from typing import Any, Callable, Dict, List, Optional, Tuple, Type
from xml.sax.saxutils import escape as xml_escape

# --- 2. Third-Party GUI Imports (ttkbootstrap) ---
import ttkbootstrap as ttk
from ttkbootstrap.dialogs import Messagebox as _TTKMessagebox
from ttkbootstrap.widgets import DateEntry
from ttkbootstrap.widgets.scrolled import ScrolledFrame

# --- 3. Third-Party Document Generation (ReportLab) ---
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.fonts import addMapping
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# --- 4. Optional & Critical Dependencies ---

# PyMuPDF (Required for PDF viewing/manipulation)
try:
    import fitz
except ImportError:
    print("CRITICAL: PyMuPDF not installed. Run: pip install PyMuPDF")
    sys.exit(1)

# Python-Barcode (Optional)
try:
    import barcode
    from barcode.writer import ImageWriter
    BARCODE_AVAILABLE = True
except ImportError:
    BARCODE_AVAILABLE = False
    print("WARNING: python-barcode not installed. Barcodes will be skipped.")

# =============================================================================
#  1. APPLICATION METADATA
# =============================================================================
APP_NAME: str          = "Vyaprix"
APP_VERSION: str       = "v1.2"
APP_BUILD_VERSION: str = "1.2.2"
APP_RELEASE_DATE: str  = "15 July 2026"
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

# Mutable application state mapped to Windows ProgramData
PROGRAM_DATA_DIR: Path = Path(os.environ.get("PUBLIC", r"C:\Users\Public"))
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
def _copy_bundled_files(src_dir: Path, dest_dir: Path) -> None:
    """Helper method to securely copy bundled files to the writable directory."""
    if not src_dir.exists():
        return
        
    for item in src_dir.iterdir():
        if item.is_file():
            dest = dest_dir / item.name
            if not dest.exists():
                try:
                    shutil.copy2(item, dest)
                except Exception as e:
                    # Fallback to print since the logger is not initialized yet
                    print(f"Failed to copy {item.name}: {e}")

def _initialize_environment() -> None:
    """Create mutable directories and copy default assets/fonts on first launch."""
    # 1. Ensure all required writable directories exist
    directories_to_create = (
        DATA_DIR, ASSETS_DIR, DATA_FONTS_DIR, 
        INVOICES_DIR, BATCH_DIR, LOGS_DIR
    )
    for directory in directories_to_create:
        directory.mkdir(parents=True, exist_ok=True)

    # 2. Extract defaults from EXE to Writable paths
    _copy_bundled_files(RESOURCE_ASSETS_DIR, ASSETS_DIR)
    _copy_bundled_files(FONTS_DIR, DATA_FONTS_DIR)

# Trigger initialization immediately on module load
_initialize_environment()

# =============================================================================
#  4. UI CONSTANTS & FORMATTING
# =============================================================================
FONT_FAMILY: str         = "Segoe UI"
DB_DATE_FORMAT: str      = "%Y-%m-%d"
DISPLAY_DATE_FORMAT: str = "%d %b %Y"
PAYMENT_MODES: list[str] = ['Cash', 'UPI', 'Bank Transfer', 'Card', 'Cheque', 'Other']

# Regex to safely sanitize invoice numbers before converting them to PDF filenames.
# Prevents path traversal and Windows filesystem crashes.
INVALID_FILENAME_CHARS: re.Pattern = re.compile(r'[\\/:*?"<>|]')

STATUS_BADGE: dict[str, dict[str, str]] = {
    "PAID":          {"foreground": "#10B981", "background": ""}, # Soft Green
    "DUE":           {"foreground": "#EF4444", "background": ""}, # Soft Red
    "PARTIAL":       {"foreground": "#F59E0B", "background": ""}, # Orange/Amber
    "OVERPAID":      {"foreground": "#8B5CF6", "background": ""}, # Violet -- distinct from "Paid" green so an overpayment stands out
    "COMPLIMENTARY": {"foreground": "#3B82F6", "background": ""}, # Blue
}

# User-facing wording for the "Status" column/badge. Kept separate from
# InvoiceStatus's own values (which are what's actually stored in the
# database and must not change, or every invoice saved before this
# update would suddenly show a status the code no longer recognises).
# This dict is the ONLY place the display wording lives, so the
# Dashboard list, the live invoice-form badge, and the PDF status line
# can never say three different things for the same status.
STATUS_DISPLAY_LABELS: Dict[str, str] = {
    "Due": "UNPAID", "Partial": "PARTIALLY PAID", "Paid": "PAID",
    "Overpaid": "OVERPAID", "Complimentary": "COMPLIMENTARY",
}

# ttkbootstrap style for the live invoice-form status badge (a different
# widget/mechanism than the Treeview's STATUS_BADGE hex colors above).
STATUS_BOOTSTYLE: Dict[str, str] = {
    "Paid": "success-inverse", "Due": "danger-inverse", "Partial": "warning-inverse",
    "Overpaid": "info-inverse", "Complimentary": "info-inverse",
}

class Messagebox:
    """Drop-in replacement matching how this app actually calls Messagebox:
    title first, then message -- and real booleans from confirmation dialogs."""

    @staticmethod
    def show_info(title, message, *args, **kwargs):
        return _TTKMessagebox.show_info(message, title, *args, **kwargs)

    @staticmethod
    def show_warning(title, message, *args, **kwargs):
        return _TTKMessagebox.show_warning(message, title, *args, **kwargs)

    @staticmethod
    def show_error(title, message, *args, **kwargs):
        return _TTKMessagebox.show_error(message, title, *args, **kwargs)

    @staticmethod
    def show_question(title, message, *args, **kwargs):
        return _TTKMessagebox.show_question(message, title, *args, **kwargs)

    @staticmethod
    def ok(title, message, *args, **kwargs):
        return _TTKMessagebox.ok(message, title, *args, **kwargs)

    @staticmethod
    def okcancel(title, message, *args, **kwargs) -> bool:
        """True only if the user actually clicked OK."""
        kwargs.setdefault("localize", False)  # keep button text predictable for comparison
        result = _TTKMessagebox.okcancel(message, title, *args, **kwargs)
        return result == "OK"

    @staticmethod
    def yesno(title, message, *args, **kwargs) -> bool:
        """True only if the user actually clicked Yes."""
        kwargs.setdefault("localize", False)
        result = _TTKMessagebox.yesno(message, title, *args, **kwargs)
        return result == "Yes"


class InvoiceStatus(Enum):
    PAID          = "Paid"
    DUE           = "Due"
    PARTIAL       = "Partial"
    OVERPAID      = "Overpaid"
    COMPLIMENTARY = "Complimentary"

# =============================================================================
#  LOGGING
# =============================================================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        # Rotate at 5MB, keep 5 backups -- prevents the log from growing
        # forever on a machine that runs Vyaprix for years.
        RotatingFileHandler(LOG_FILE, mode='a', maxBytes=5 * 1024 * 1024,
                            backupCount=5, encoding='utf-8'),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

# =============================================================================
#  FONTS
# =============================================================================
def register_fonts() -> str:
    """Registers the bundled NotoSans font (needed for the currency
    symbol and non-Latin characters). Tries the writable ProgramData
    copy first, then the read-only copy bundled in the EXE, so a failed
    first-run copy doesn't take the real font down with it.

    Falling back to "Helvetica" is a real correctness problem: it has no
    glyph for "₹", and ReportLab silently substitutes an unrelated
    character with no error -- see InvoiceApp._warn_if_fallback_font(),
    which surfaces this to the user once at startup.
    """
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
#  CONFIGURATION
# =============================================================================
def get_default_config() -> Dict[str, Any]:
    return {
        "theme": "litera",
        "user_details": {
            # Structured "Your Details" fields (Settings > Your Details).
            # Deliberately does NOT include a business/firm name or GST
            # number: there is only one business identity, and
            # company_details below (Settings > General) is its single
            # source of truth -- it's what the invoice letterhead uses,
            # so it's also what the FROM block and GSTIN line use (see
            # compose_your_name()/compose_your_contact()). Keeping a
            # second copy of the same two facts here is exactly what
            # caused the same GSTIN to print twice on one invoice.
            "name": "", "phone": "", "email": "",
            "website": "", "address": "", "city": "",
            "state": "", "pin_code": "",
            # Legacy free-text field from older builds. No longer edited
            # by SettingsDialog, but kept (and still readable) so nobody's
            # existing data silently disappears on upgrade -- see
            # compose_your_contact()'s fallback below.
            "contact_info": "",
        },
        "company_details": {"name": APP_NAME, "description": APP_TAGLINE, "gstin": "22AAAAA0000A1Z5"},
        "payment_details": {
            "upi_id": "yourupi@bank", "payment_terms": "Due upon receipt",
            "bank_name": "", "account_holder_name": "", "account_number": "",
            "ifsc_code": "", "payment_notes": "",
        },
        "asset_paths":     {"logo": "assets/app_logo.png",    "signature": "assets/sign.png", "upi_qr_code": "assets/upi_qr.png"},
        "pdf_settings":    {"footer_message": "Generated with Vyaprix • Business Made Smarter.", "always_show_qr_code": False},
        "currency_symbol": "₹",
    }

def _migrate_legacy_business_identity(config: Dict[str, Any]) -> bool:
    """One-release upgrade path. A prior build stored a business/firm
    name and GST number under user_details ('firm_name'/'gst_number'),
    separate from company_details.name/gstin -- the fields the invoice
    letterhead and the dedicated GSTIN line actually use. That let
    someone fill in only one of the two and still expect it everywhere,
    and it printed the GSTIN twice when both happened to be filled in.

    There is only one business identity. This folds any already-entered
    legacy firm_name/gst_number into company_details, but only where
    company_details hasn't already been customized away from the
    built-in default -- a real, deliberately-set company name/GSTIN is
    never overwritten. The old keys are simply left in place afterwards;
    nothing reads them anymore, and they drop out of the file on the
    next Settings save (SettingsDialog._save() no longer writes them).

    Returns True if it changed anything, so load_config() knows to
    persist the fix back to disk immediately rather than waiting for the
    user to open Settings and hit Save.
    """
    ud   = config.get("user_details", {})
    comp = config.setdefault("company_details", {})
    changed = False

    legacy_firm = (ud.get("firm_name") or "").strip()
    if legacy_firm and (not comp.get("name") or comp.get("name") == APP_NAME):
        comp["name"] = legacy_firm
        changed = True

    legacy_gst = (ud.get("gst_number") or "").strip()
    default_gstin = get_default_config()["company_details"]["gstin"]
    if legacy_gst and (not comp.get("gstin") or comp.get("gstin") == default_gstin):
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
                _save_config_raw(data)
            return data
    except Exception:
        return get_default_config()

def _save_config_raw(data: Dict[str, Any]):
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
#  HELPERS
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

_CENTS = Decimal("0.01")

def _round2(value: Any) -> Decimal:
    """Rounds a number to 2 decimal places using ROUND_HALF_UP -- the
    convention accountants and tax software expect (0.005 always rounds
    up to 0.01, not down to 0.00 the way banker's rounding would, and
    not left to whatever a raw float happens to do). Converts through
    str() first: Decimal(0.1) reproduces float's imprecise binary
    representation, while Decimal(str(0.1)) does not."""
    if isinstance(value, Decimal):
        return value.quantize(_CENTS, rounding=ROUND_HALF_UP)
    return Decimal(str(value or 0)).quantize(_CENTS, rounding=ROUND_HALF_UP)

def safe_float(value: Any, fallback: float = 0.0) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return fallback

DECIMAL_INPUT_PATTERN = re.compile(r"^\d*\.?\d*$")
INTEGER_INPUT_PATTERN = re.compile(r"^\d*$")
EMAIL_FORMAT_PATTERN  = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]{2,}$")
GSTIN_FORMAT_PATTERN  = re.compile(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$")
PHONE_FORMAT_PATTERN  = re.compile(r"^[+]?[\d\s().-]+$")

def restrict_entry_to_decimal(entry: tk.Widget, max_decimal_places: int = 2) -> tk.Widget:
    def validate(proposed: str) -> bool:
        if proposed == "":
            return True
        if not DECIMAL_INPUT_PATTERN.match(proposed):
            return False
        return len(proposed.partition(".")[2]) <= max_decimal_places
    command = entry.register(validate)
    entry.configure(validate="key", validatecommand=(command, "%P"))
    return entry

def restrict_entry_to_digits(entry: tk.Widget, max_length: Optional[int] = None) -> tk.Widget:
    def validate(proposed: str) -> bool:
        if proposed == "":
            return True
        if not INTEGER_INPUT_PATTERN.match(proposed):
            return False
        return max_length is None or len(proposed) <= max_length
    command = entry.register(validate)
    entry.configure(validate="key", validatecommand=(command, "%P"))
    return entry

def format_entry_as_decimal_on_blur(entry: tk.Widget, places: int = 2) -> tk.Widget:
    def on_focus_out(_event=None):
        try:
            current = entry.get()
        except tk.TclError:
            return
        formatted = f"{safe_float(current):.{places}f}"
        if formatted != current:
            entry.delete(0, "end")
            entry.insert(0, formatted)
    entry.bind("<FocusOut>", on_focus_out, add="+")
    return entry

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
    strings, which parse a small XML/HTML-like markup subset. Without this,
    any free-text field containing '&', '<', or '>' (e.g. a client name like
    "Smith & Sons") throws a parse error and PDF generation fails outright."""
    return xml_escape(str(value or ""))

def load_app_logo_image(master: tk.Misc) -> Optional[tk.PhotoImage]:
    if APP_LOGO_PNG.exists():
        try:
            # First, try to use Pillow (PIL) for high-quality, smooth resizing
            try:
                from PIL import Image, ImageTk
                img = Image.open(str(APP_LOGO_PNG))
                
                # We want the logo to be exactly 160 pixels wide to fit the sidebar
                max_width = 160
                w, h = img.size
                if w > max_width:
                    new_height = int(h * (max_width / w))
                    # Use LANCZOS for high-quality downsampling
                     # Try modern Pillow, fallback to legacy Pillow
                    resample_filter = getattr(Image, 'Resampling', Image).LANCZOS
                    img = img.resize((max_width, new_height), resample_filter)
                
                return ImageTk.PhotoImage(image=img, master=master)
            
            except ImportError:
                # Fallback to Tkinter's basic built-in resizing if Pillow is not installed
                img = tk.PhotoImage(master=master, file=str(APP_LOGO_PNG))
                w = img.width()
                if w > 160:
                    factor = w // 160
                    if factor > 1:
                        img = img.subsample(factor, factor)
                return img
                
        except Exception as exc:
            logger.warning(f"Could not load app logo image: {exc}")
    return None

def apply_theme_to_text_widget(widget: tk.Text, colors=None) -> None:
    """Applies the active ttkbootstrap theme's input colors to a raw
    tk.Text widget.

    ttkbootstrap only themes ttk.* widgets automatically -- a plain
    tk.Text (which PlaceholderText is built on, and which nothing else
    in ttk provides a multi-line-editable equivalent of) is invisible to
    that machinery and just keeps Tk's system-default white background
    with black text forever, regardless of which theme is active. On a
    dark theme that's a jarring white box in an otherwise dark window;
    combined with the placeholder mixin's muted grey hint text, it's
    also the most likely source of the "some text becomes hard to read"
    symptom, since a light grey hint on white is much lower-contrast
    than the same grey was ever designed to be shown against.

    Safe to call on any tk.Text, and safe to call repeatedly -- used
    both at widget-creation time and again on every live theme switch
    (see InvoiceApp._on_settings_saved), since switching themes doesn't
    retroactively fix widgets that already exist.
    """
    try:
        if colors is None:
            colors = ttk.Style().colors
        fg = colors.inputfg
        widget.configure(
            background=colors.inputbg, foreground=fg,
            insertbackground=fg,  # the text cursor/caret -- defaults to black, invisible on a dark background
            selectbackground=colors.selectbg, selectforeground=colors.selectfg,
            highlightbackground=colors.border, highlightcolor=colors.primary,
        )
        # A placeholder-enabled Text currently showing its hint must keep
        # the muted placeholder colour, not the "real text" colour just
        # applied above.
        if getattr(widget, "_showing_placeholder", False):
            widget.configure(foreground=themed_placeholder_color(widget))
        if hasattr(widget, "_real_fg"):
            widget._real_fg = fg  # keep in sync so a later focus-out uses the CURRENT theme's colour, not a stale one from creation time
    except Exception as exc:
        logger.warning(f"Could not theme a text widget: {exc}")

def themed_placeholder_color(widget: tk.Misc) -> str:
    """Picks a muted 'hint text' grey with guaranteed-readable contrast
    against the CURRENT theme's actual input background, rather than one
    fixed grey for every theme: a fixed medium grey looks fine against
    most bundled themes' near-white or near-black input fields, but is
    hard to read against at least one (a pale, not-quite-white input
    background blends a medium grey in too much). Looked up fresh on
    every call, not cached, so it stays correct across live theme
    switches too -- the same reasoning as _PlaceholderMixin._lookup_
    normal_fg() just above, applied to the hint color instead of the
    real one."""
    try:
        colors = ttk.Style().colors
        r, g, b = widget.winfo_rgb(colors.inputbg)
        def lin(c):
            c = c / 65535.0
            return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
        rel_luminance = 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)
        return "#6b7280" if rel_luminance >= 0.5 else "#9ca3af"
    except Exception:
        return "#9CA3AF"

def load_thumbnail_image(path: Optional[Path], master: tk.Misc, max_size: int = 44) -> Optional[tk.PhotoImage]:
    """Loads an arbitrary image file as a small square-ish thumbnail, for
    the asset-upload preview rows in Settings. Same Pillow-with-Tk-fallback
    approach as load_app_logo_image() above, just generalised to any path
    and any target size instead of being hardcoded to the app logo."""
    if not path or not path.exists():
        return None
    try:
        try:
            from PIL import Image, ImageTk
            img = Image.open(str(path))
            img.thumbnail((max_size, max_size), getattr(Image, "Resampling", Image).LANCZOS)
            return ImageTk.PhotoImage(image=img, master=master)
        except ImportError:
            img = tk.PhotoImage(master=master, file=str(path))
            longest = max(img.width(), img.height())
            if longest > max_size:
                factor = max(1, longest // max_size)
                img = img.subsample(factor, factor)
            return img
    except Exception as exc:
        logger.warning(f"Could not load thumbnail for {path}: {exc}")
        return None

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
    dedicated line elsewhere on the same invoice (see _build_address() /
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

# =============================================================================
#  KEYBOARD-SHORTCUT HELPERS
# =============================================================================
def _bind_action(func: Callable[[], Any]) -> Callable[..., str]:
    """Wraps a zero-argument callable into a Tk key-event handler that runs
    it and then returns "break".

    Every application-wide shortcut needs that "break": without it, a key
    combo bound with bind_all still falls through to whatever a widget's
    own class bindings do with the same keys (for example Tk's built-in
    Ctrl+D, which deletes the character after the cursor in Entry/Text
    widgets) -- so the shortcut and the widget's default behaviour would
    both fire from a single keystroke. Centralising the "break" here means
    every shortcut below gets that guarantee the same way, instead of each
    binding site having to remember it individually.
    """
    def _handler(event=None):
        func()
        return "break"
    return _handler


def prompt_and_backup_database(parent: tk.Misc) -> None:
    """Asks where to save a copy of the live database, then copies it
    there. Shared by the Settings dialog's "Save a Backup Copy" button and
    the Ctrl+B keyboard shortcut, so there is one place that defines what
    "back up the database" means rather than two copies that could drift
    apart.
    """
    path = filedialog.asksaveasfilename(
        parent=parent,
        initialfile=f"Vyaprix-backup-{datetime.date.today().isoformat()}.db",
        defaultextension=".db", filetypes=[("SQLite Database", "*.db")],
    )
    if not path:
        return
    try:
        shutil.copy(DB_NAME, path)
        Messagebox.show_info("Backup Saved", f"Database backed up to:\n{path}")
    except Exception as exc:
        Messagebox.show_error("Backup Failed", str(exc))

class ToolTip:
    def __init__(self, widget: tk.Widget, text: str, delay: int = 500):
        self.widget = widget
        self.text   = text
        self.delay  = delay
        self._id    = None
        self._win   = None
        widget.bind("<Enter>",       self._schedule, add="+")
        widget.bind("<Leave>",       self._cancel,   add="+")
        widget.bind("<ButtonPress>", self._cancel,   add="+")

    def _schedule(self, _=None):
        self._cancel()
        self._id = self.widget.after(self.delay, self._show)

    def _cancel(self, _=None):
        if self._id:
            self.widget.after_cancel(self._id)
            self._id = None
        if self._win:
            self._win.destroy()
            self._win = None

    def _show(self):
        x = self.widget.winfo_rootx() + 22
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self._win = win = tk.Toplevel(self.widget)
        win.wm_overrideredirect(True)
        win.wm_geometry(f"+{x}+{y}")
        tk.Label(win, text=self.text, background="#FFFDE7", foreground="#212121", relief="flat",
                 borderwidth=1, padx=7, pady=4, font=(FONT_FAMILY, 9),
                 justify="left", wraplength=280).pack()


class Debouncer:
    """Coalesces rapid repeated calls (e.g. keystrokes) into one after `delay` ms."""
    def __init__(self, widget: tk.Misc, delay: int, func: Callable):
        self.widget = widget
        self.delay  = delay
        self.func   = func
        self._id    = None

    def __call__(self, *args):
        if self._id:
            self.widget.after_cancel(self._id)
        self._id = self.widget.after(self.delay, self.func)

    def cancel(self):
        if self._id:
            self.widget.after_cancel(self._id)
            self._id = None


# =============================================================================
#  PLACEHOLDER-AWARE INPUT WIDGETS
# =============================================================================
class _PlaceholderMixin:
    """Shared placeholder behaviour for the Entry/Combobox subclasses below.

    The hint text is drawn directly into the widget -- so it automatically
    inherits the widget's real font, padding, and per-theme styling instead
    of needing a pixel-matched overlay -- but every read path is guaranteed
    to see an empty string while the hint is showing. get() is overridden to
    return "" in that state, so every existing call site in this file that
    does `widget.get()` keeps working completely unchanged: it can never
    mistake the hint for real user input and accidentally save it.

    Widgets bound to a `textvariable` still sync that variable's text to the
    Tk layer as usual (that's necessary for autocomplete filtering elsewhere
    in the app to keep working), so code that reads the *variable* directly
    rather than the widget must not do that for a placeholder-enabled field --
    this file always reads such fields via the widget itself.
    """
    _PLACEHOLDER_COLOR = "#9CA3AF"  # fallback only -- see themed_placeholder_color()

    def _init_placeholder(self, placeholder: str):
        self._placeholder = placeholder
        self._showing_placeholder = False
        self._real_fg = self._lookup_normal_fg()
        self.bind("<FocusIn>", self._ph_on_focus_in, add="+")
        self.bind("<FocusOut>", self._ph_on_focus_out, add="+")
        if placeholder:
            self._ph_draw()

    def _lookup_normal_fg(self) -> str:
        """Looks the *current* theme's real foreground color up fresh (rather
        than caching it once) so placeholders keep rendering correctly even
        after the user switches themes live from Settings."""
        try:
            color = ttk.Style().lookup(self.winfo_class(), "foreground")
            return str(color) if color else "black"
        except Exception:
            return "black"

    def _ph_draw(self):
        if not self._placeholder or self._raw_get():
            return
        self._showing_placeholder = True
        try:
            self.configure(foreground=themed_placeholder_color(self))
        except tk.TclError:
            pass
        self._raw_insert(self._placeholder)

    def _ph_clear(self):
        if self._showing_placeholder:
            self._raw_clear()
            try:
                self.configure(foreground=self._lookup_normal_fg())
            except tk.TclError:
                pass
            self._showing_placeholder = False

    def _ph_on_focus_in(self, _event=None):
        self._ph_clear()

    def _ph_on_focus_out(self, _event=None):
        if not self._raw_get():
            self._ph_draw()

    def set_placeholder(self, text: str):
        """Changes the hint text after construction (e.g. an item row whose
        purpose changes) and redraws it immediately if applicable."""
        showing = self._showing_placeholder
        if showing:
            self._ph_clear()
        self._placeholder = text
        if showing or not self._raw_get():
            self._ph_draw()


class PlaceholderEntry(_PlaceholderMixin, ttk.Entry):
    """Drop-in replacement for ttk.Entry that shows greyed-out hint text
    while empty and unfocused. get() transparently returns "" while the
    hint is showing."""

    def __init__(self, master=None, placeholder: str = "", **kw):
        ttk.Entry.__init__(self, master, **kw)
        self._init_placeholder(placeholder)

    def _raw_get(self) -> str:
        return ttk.Entry.get(self)

    def _raw_insert(self, text: str):
        ttk.Entry.insert(self, 0, text)

    def _raw_clear(self):
        ttk.Entry.delete(self, 0, "end")

    def get(self) -> str:
        return "" if self._showing_placeholder else ttk.Entry.get(self)

    def insert(self, index, text):
        self._ph_clear()
        if text:
            ttk.Entry.insert(self, index, text)
        if not ttk.Entry.get(self):
            self._ph_draw()

    def delete(self, first, last=None):
        ttk.Entry.delete(self, first, last)
        if not ttk.Entry.get(self):
            self._ph_draw()


class PlaceholderCombobox(_PlaceholderMixin, ttk.Combobox):
    """Same placeholder behaviour as PlaceholderEntry, for the app's two
    *editable* (non-readonly) comboboxes -- client name and catalog item
    title -- which can be freely typed into as well as picked from a list."""

    def __init__(self, master=None, placeholder: str = "", **kw):
        ttk.Combobox.__init__(self, master, **kw)
        self._init_placeholder(placeholder)

    def _raw_get(self) -> str:
        return ttk.Combobox.get(self)

    def _raw_insert(self, text: str):
        ttk.Combobox.insert(self, 0, text)

    def _raw_clear(self):
        ttk.Combobox.delete(self, 0, "end")

    def get(self) -> str:
        # --- THE FIX ---
        # Smart detection: If the native text in the box no longer matches 
        # the placeholder, the user MUST have selected an item from the list natively.
        current_text = ttk.Combobox.get(self)
        if self._showing_placeholder and current_text != self._placeholder:
            self._showing_placeholder = False
            try:
                self.configure(foreground=self._lookup_normal_fg())
            except tk.TclError:
                pass
                
        return "" if self._showing_placeholder else current_text

    def insert(self, index, text):
        self._ph_clear()
        if text:
            ttk.Combobox.insert(self, index, text)
        if not ttk.Combobox.get(self):
            self._ph_draw()

    def delete(self, first, last=None):
        ttk.Combobox.delete(self, first, last)
        if not ttk.Combobox.get(self):
            self._ph_draw()

class PlaceholderText(tk.Text):
    """A tk.Text box with the same placeholder behaviour, adapted to Text's
    "1.0" / "end-1c" indexing instead of Entry's integer indexing."""

    _PLACEHOLDER_COLOR = "#9CA3AF"  # fallback only -- see themed_placeholder_color()

    def __init__(self, master=None, placeholder: str = "", **kw):
        tk.Text.__init__(self, master, **kw)
        self._placeholder = placeholder
        self._showing_placeholder = False
        apply_theme_to_text_widget(self)
        try:
            fg = self.cget("foreground")
            self._real_fg = str(fg) if fg else "black"
        except tk.TclError:
            self._real_fg = "black"
        self.bind("<FocusIn>", self._ph_on_focus_in, add="+")
        self.bind("<FocusOut>", self._ph_on_focus_out, add="+")
        if placeholder:
            self._ph_draw()

    def _raw_get(self) -> str:
        return tk.Text.get(self, "1.0", "end-1c")

    def _ph_draw(self):
        if not self._placeholder or self._raw_get():
            return
        self._showing_placeholder = True
        self.configure(foreground=themed_placeholder_color(self))
        tk.Text.insert(self, "1.0", self._placeholder)

    def _ph_clear(self):
        if self._showing_placeholder:
            tk.Text.delete(self, "1.0", "end")
            self.configure(foreground=self._real_fg)
            self._showing_placeholder = False

    def _ph_on_focus_in(self, _event=None):
        self._ph_clear()

    def _ph_on_focus_out(self, _event=None):
        if not self._raw_get():
            self._ph_draw()

    def get(self, index1="1.0", index2="end-1c"):
        norm = lambda ix: "end-1c" if str(ix).startswith("end") else str(ix)
        if self._showing_placeholder and norm(index1) == "1.0" and norm(index2) == "end-1c":
            return ""
        return tk.Text.get(self, index1, index2)

    def insert(self, index, text, *args):
        self._ph_clear()
        if text:
            tk.Text.insert(self, index, text, *args)
        if not self._raw_get():
            self._ph_draw()

    def delete(self, index1, index2=None):
        tk.Text.delete(self, index1, index2)
        if not self._raw_get():
            self._ph_draw()

# =============================================================================
#  DATABASE
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
                         global_disc_type: str = "amount") -> Dict[str, float]:
        """Computes every figure on the invoice summary, in this fixed
        order: Subtotal -> Discount -> Taxable Amount -> CGST/SGST ->
        Grand Total. All arithmetic happens in Decimal with
        ROUND_HALF_UP, rounding to the nearest paisa/cent at each
        intermediate step rather than only at the very end -- that's
        what guarantees CGST + SGST always sum to exactly the combined
        tax shown elsewhere, instead of occasionally being a paisa off
        from ordinary floating-point drift.

        Returns plain floats (already rounded to 2dp), not Decimal
        objects -- every existing caller (SQLite writes, fmt_money(),
        the PDF generator) keeps working exactly as before; only the
        arithmetic *inside* this function changed.
        """
        def d(x: Any) -> Decimal:
            return Decimal(str(x or 0))

        subtotal   = _round2(sum((d(i.get("amount", 0)) for i in items), Decimal("0")))
        item_disc  = _round2(sum((d(i.get("discount", 0)) for i in items), Decimal("0")))
        after_item = max(Decimal("0"), subtotal - item_disc)

        if global_disc_type == "%":
            global_disc = _round2(after_item * d(global_disc_val) / Decimal("100"))
        else:
            global_disc = _round2(d(global_disc_val))
        global_disc = max(Decimal("0"), min(global_disc, after_item))

        total_disc = item_disc + global_disc
        taxable    = max(Decimal("0"), subtotal - total_disc)

        half_rate  = d(tax_rate) / Decimal("2")
        cgst = sgst = _round2(taxable * half_rate / Decimal("100"))
        grand_total = taxable + cgst + sgst

        return {
            "subtotal": float(subtotal), "item_discount": float(item_disc),
            "global_discount": float(global_disc), "total_discount": float(total_disc),
            "taxable": float(taxable), "cgst": float(cgst), "sgst": float(sgst),
            "grand_total": float(grand_total),
        }

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
        try:
            items = json.loads(d.get("items") or "[]")
        except json.JSONDecodeError:
            items = []
        tax_rate = float(d.get("tax_rate") or 0) if d.get("is_gst_invoice", 1) else 0.0
        totals = self.calculate_totals(
            items, tax_rate,
            float(d.get("global_discount_value") or 0),
            d.get("global_discount_type") or "amount",
        )
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
                q = f"%{search}%"
                rows = c.execute(
                    "SELECT * FROM clients WHERE name LIKE ? ORDER BY name", (q,)
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
                q = f"%{search}%"
                rows = c.execute(
                    "SELECT * FROM catalog WHERE title LIKE ? ORDER BY title", (q,)
                ).fetchall()
            else:
                rows = c.execute("SELECT * FROM catalog ORDER BY title").fetchall()
            return [dict(r) for r in rows]

    def get_catalog_item_by_id(self, item_id: int) -> Optional[Dict]:
        with closing(self._conn()) as c:
            row = c.execute("SELECT * FROM catalog WHERE id=?", (item_id,)).fetchone()
            return dict(row) if row else None

    def save_catalog_item(self, title: str, description: str, amount: float):
        if not title:
            return
        with closing(self._conn()) as c:
            with c:
                c.execute("""
                    INSERT INTO catalog (title, description, amount)
                    VALUES (?, ?, ?)
                    ON CONFLICT(title) DO UPDATE SET
                    description=excluded.description, amount=excluded.amount
                """, (title, description, amount))

    def upsert_catalog_item_by_id(self, item_id: Optional[int], title: str,
                                   description: str, amount: float) -> Tuple[bool, Optional[str]]:
        if not title.strip():
            return False, "Item title is required."
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
                        c.execute("INSERT INTO catalog (title, description, amount) VALUES (?,?,?)",
                                  (title, description, amount))
                    else:
                        c.execute("UPDATE catalog SET title=?, description=?, amount=? WHERE id=?",
                                  (title, description, amount, item_id))
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
            clauses.append("(invoice_number LIKE ? OR bill_to LIKE ? OR notes LIKE ?)")
            q = f"%{search_query}%"
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
            clauses.append("(invoice_number LIKE ? OR bill_to LIKE ? OR notes LIKE ?)")
            q = f"%{search_query}%"
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

# =============================================================================
#  PDF GENERATOR
# =============================================================================
class PDFGenerator:
    def __init__(self, data: Dict[str, Any], config: Dict[str, Any],
                 payments: Optional[List[Dict]] = None, out_dir: Optional[Path] = None):
        self.data     = data
        self.config   = config
        self.payments = payments or []

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
        self.styles   = self._setup_styles()

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
        ss   = getSampleStyleSheet()
        base = PDF_FONT
        bold = f"{base}-Bold" if base != "Helvetica" else "Helvetica-Bold"
        ss.add(ParagraphStyle("NC",          fontName=base, fontSize=9,  parent=ss["Normal"],   leading=11))
        ss.add(ParagraphStyle("Right",       parent=ss["NC"], alignment=TA_RIGHT))
        ss.add(ParagraphStyle("Center",      parent=ss["NC"], alignment=TA_CENTER))
        ss.add(ParagraphStyle("Small",       parent=ss["NC"], fontSize=7.5, textColor=colors.HexColor("#6B7280")))
        ss.add(ParagraphStyle("SmallRight",  parent=ss["Small"], alignment=TA_RIGHT))
        ss.add(ParagraphStyle("YourContact", parent=ss["NC"], alignment=TA_RIGHT, leading=11))
        ss.add(ParagraphStyle("BillTo",      parent=ss["NC"], leading=11))
        ss.add(ParagraphStyle("ItemDesc",    parent=ss["NC"], textColor=colors.HexColor("#6B7280"), fontSize=8, leading=10))
        ss.add(ParagraphStyle("CompanyName", fontName=bold,   fontSize=16, textColor=colors.HexColor("#1F2937"), leading=18))
        ss.add(ParagraphStyle("Tagline",     parent=ss["NC"], fontSize=8.5, textColor=colors.HexColor("#6B7280")))
        ss.add(ParagraphStyle("InvoiceTitle",fontName=bold,   fontSize=18, alignment=TA_RIGHT, textColor=colors.HexColor("#111827")))
        ss.add(ParagraphStyle("NoteBox",     parent=ss["NC"], fontSize=8, textColor=colors.HexColor("#374151")))
        return ss

    def _build_header(self) -> Table:
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

        if logo := self._asset("logo"):
            try:
                img = Image(str(logo), 2.0*cm, 2.0*cm, kind="proportional")
                left_side = Table([[img, left]], colWidths=[2.5*cm, 6.5*cm],
                                  style=[("VALIGN",(0,0),(-1,-1),"MIDDLE"), ("LEFTPADDING",(0,0),(-1,-1),0)])
            except Exception:
                left_side = left
        else:
            left_side = left

        return Table([[left_side, title_block]], colWidths=[9.0*cm, 9.0*cm], style=[
            ("VALIGN",(0,0),(-1,-1),"TOP"), ("ALIGN",(1,0),(1,-1),"RIGHT"),
            ("LEFTPADDING",(0,0),(-1,-1),0),   ("RIGHTPADDING",(0,0),(-1,-1),0),
            ("LINEBELOW",(0,0),(-1,-1),1,colors.HexColor("#E5E7EB")),
            ("BOTTOMPADDING",(0,0),(-1,-1),10),
        ])
    
    def _draw_footer(self, canvas, doc):
        canvas.saveState()
        canvas.setFont(PDF_FONT, 8)
        canvas.setFillColor(colors.HexColor("#6B7280"))
        canvas.drawCentredString(A4[0]/2, 1.0*cm,
                                 self.config.get("pdf_settings", {}).get("footer_message", "Thank you!"))
        canvas.drawRightString(A4[0]-1.5*cm, 1.0*cm, f"Page {doc.page}")
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

    def on_each_page(self, canvas, doc):
        self._apply_watermark(canvas, doc)
        self._draw_footer(canvas, doc)

    def _build_address(self) -> Table:
        cg = yg = ""
        if self.is_gst:
            if cgstin := self.data.get("client_gstin"):
                cg = f"<br/>GSTIN: {safe_markup(cgstin)}"
            if ygstin := self.config.get("company_details", {}).get("gstin"):
                yg = f"<br/>GSTIN: {safe_markup(ygstin)}"
        # Escape first, THEN turn newlines into <br/> -- safe_markup() only
        # touches &/</>, so it won't disturb the newlines we're about to replace.
        your_c   = safe_markup(self.data.get("your_contact")).replace("\n", "<br/>")
        client_c = safe_markup(self.data.get("client_contact")).replace("\n", "<br/>")
        your_name = safe_markup(self.data.get("your_name"))
        bill_to   = safe_markup(self.data.get("bill_to"))
        your_p   = Paragraph(f"<font color='#6B7280'>FROM:</font><br/><b>{your_name}</b><br/>{your_c}{yg}",     self.styles["YourContact"])
        bill_p   = Paragraph(f"<font color='#6B7280'>BILL TO:</font><br/><b>{bill_to}</b><br/>{client_c}{cg}", self.styles["BillTo"])
        return Table([[bill_p, your_p]], colWidths=["50%","50%"], style=[
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(-1,-1),0),
        ])

    def _build_meta(self) -> Table:
        date_str = parse_date(self.data.get("invoice_date"), DB_DATE_FORMAT, DISPLAY_DATE_FORMAT)
        meta     = f"<b>Invoice #:</b> {safe_markup(self.data.get('invoice_number'))}<br/><b>Invoice Date:</b> {date_str}"
        if not self.data.get("is_complimentary"):
            meta += f"<br/><b>Payment Terms:</b> {safe_markup(self.data.get('payment_terms') or 'N/A')}"
        return Table([[Paragraph(meta, self.styles["NC"])]], style=[
            ("BOX",(0,0),(-1,-1),1,colors.HexColor("#E5E7EB")),
            ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#F9FAFB")),
            ("LEFTPADDING",(0,0),(-1,-1),8), ("TOPPADDING",(0,0),(-1,-1),6),
            ("BOTTOMPADDING",(0,0),(-1,-1),6),
        ])

    def _build_items(self) -> Table:
        hs  = ParagraphStyle("HS",  parent=self.styles["NC"], textColor=colors.white)
        hsr = ParagraphStyle("HSR", parent=hs, alignment=TA_RIGHT)
        rows = [[Paragraph("<b>ITEM / DESCRIPTION</b>", hs), Paragraph("<b>AMOUNT</b>", hsr)]]
        for item in json.loads(self.data.get("items") or "[]"):
            title  = item.get("title")       or ""
            desc   = item.get("description") or ""
            amount = float(item.get("amount", 0))
            if not title.strip() and not desc.strip() and amount == 0:
                rows.append([[Paragraph("<br/>", self.styles["NC"])], Paragraph("", self.styles["Right"])])
            else:
                cell = [Paragraph(f"<b>{safe_markup(title)}</b>", self.styles["NC"])]
                if desc.strip():
                    cell.append(Paragraph(safe_markup(desc), self.styles["ItemDesc"]))
                rows.append([cell, Paragraph(f"{CURRENCY_SYMBOL}{amount:,.2f}", self.styles["Right"])])
        return Table(rows, colWidths=["80%","20%"], repeatRows=1, style=TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F2937")),
            ("VALIGN",(0,0),(-1,-1),"TOP"),   ("ALIGN",(1,0),(1,-1),"RIGHT"),
            ("LINEBELOW",(0,1),(-1,-2),1,colors.HexColor("#F3F4F6")),
            ("LINEBELOW",(0,-1),(-1,-1),2,colors.HexColor("#D1D5DB")),
            ("TOPPADDING",(0,0),(-1,-1),5),   ("BOTTOMPADDING",(0,0),(-1,-1),5),
            ("LEFTPADDING",(0,0),(-1,-1),6),  ("RIGHTPADDING",(0,0),(-1,-1),6),
        ]))

    def _build_totals(self) -> Table:
        if self.data.get("is_complimentary"):
            return Table(
                [[Paragraph("<b>Total</b>", self.styles["Right"]),
                  Paragraph("<b>No Charge</b>", self.styles["Right"])]],
                colWidths=["75%","25%"],
                style=TableStyle([("ALIGN",(0,0),(-1,-1),"RIGHT"), ("LEFTPADDING",(0,0),(-1,-1),10)]),
            )
        items_list = json.loads(self.data.get("items") or "[]")
        tax_rate   = float(self.data.get("tax_rate") or 0) if self.is_gst else 0.0
        g_disc     = float(self.data.get("global_discount_value") or 0)
        # Same calculation engine the invoice form uses live, so the
        # printed PDF can never show different numbers than what the
        # person saw on screen when they saved.
        totals     = DatabaseManager.calculate_totals(items_list, tax_rate, g_disc, self.data.get("global_discount_type") or "amount")
        total_paid = sum(safe_float(p.get("amount")) for p in self.payments)
        _status, balance, credit = DatabaseManager.compute_payment_status(totals["grand_total"], total_paid)

        rows = [[Paragraph("Subtotal", self.styles["Right"]),
                 Paragraph(fmt_money(totals['subtotal']), self.styles["Right"])]]
        if totals["total_discount"] > 0:
            rows.append([Paragraph("Discount", self.styles["Right"]),
                         Paragraph(f"- {fmt_money(totals['total_discount'])}", self.styles["Right"])])

        # --- IGST vs CGST/SGST PDF Rendering Logic ---
        if self.is_gst and tax_rate > 0:
            rows.append([Paragraph("Taxable Amount", self.styles["Right"]),
                         Paragraph(fmt_money(totals['taxable']), self.styles["Right"])])

            half_rate = tax_rate / 2.0
            if self.data.get("gst_type") == "IGST":
                rows.append([Paragraph(f"IGST ({tax_rate:g}%)", self.styles["Right"]),
                             Paragraph(fmt_money(totals['cgst'] + totals['sgst']), self.styles["Right"])])
            else:
                rows += [
                    [Paragraph(f"CGST ({half_rate:g}%)", self.styles["Right"]),
                     Paragraph(fmt_money(totals['cgst']), self.styles["Right"])],
                    [Paragraph(f"SGST ({half_rate:g}%)", self.styles["Right"]),
                     Paragraph(fmt_money(totals['sgst']), self.styles["Right"])],
                ]
        # ---------------------------------------------

        rows.append([Paragraph("<b>Invoice Total</b>", self.styles["Right"]),
                     Paragraph(f"<b>{fmt_money(totals['grand_total'])}</b>", self.styles["Right"])])
        if total_paid > 0:
            rows.append([Paragraph("Total Paid", self.styles["Right"]),
                         Paragraph(f"- {fmt_money(total_paid)}", self.styles["Right"])])

        # Balance Due is never negative -- an overpayment prints as
        # "Advance / Credit" instead, in green, the same way the
        # invoice-form badge shows it on screen.
        if credit > 0:
            bc = "#059669"
            rows.append([Paragraph("<b>Advance / Credit</b>", self.styles["Right"]),
                         Paragraph(f"<font color='{bc}'><b>{fmt_money(credit)}</b></font>", self.styles["Right"])])
        else:
            bc = "#059669" if balance <= 0 else "#DC2626"
            rows.append([Paragraph("<b>Balance Due</b>", self.styles["Right"]),
                         Paragraph(f"<font color='{bc}'><b>{fmt_money(balance)}</b></font>", self.styles["Right"])])
        return Table(rows, colWidths=["75%","25%"], style=TableStyle([
            ("ALIGN",(0,0),(-1,-1),"RIGHT"), ("LEFTPADDING",(0,0),(-1,-1),10),
            ("TOPPADDING",(0,0),(-1,-1),3),  ("BOTTOMPADDING",(0,0),(-1,-1),3),
            ("BACKGROUND",(0,-1),(1,-1),colors.HexColor("#FEF2F2" if (balance > 0 and credit <= 0) else "#ECFDF5")),
            ("GRID",(0,-1),(1,-1),1,colors.HexColor("#E5E7EB")),
        ]))

    def _build_payment_history(self) -> Optional[Table]:
        """Small table listing each individual payment -- omitted entirely if
        there is no payment history to show (keeps simple invoices simple)."""
        if self.data.get("is_complimentary") or not self.payments:
            return None
        hs = ParagraphStyle("PHS", parent=self.styles["Small"], textColor=colors.white)
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
        return Table(rows, colWidths=["22%","20%","33%","25%"], style=TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#374151")),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LINEBELOW",(0,0),(-1,-2),0.5,colors.HexColor("#F3F4F6")),
            ("TOPPADDING",(0,0),(-1,-1),3),   ("BOTTOMPADDING",(0,0),(-1,-1),3),
            ("LEFTPADDING",(0,0),(-1,-1),6),  ("RIGHTPADDING",(0,0),(-1,-1),6),
        ]))

    def _should_show_qr(self) -> bool:
        """Whether the QR code prints on this invoice.

        Always requires the per-invoice "Show QR Code" toggle to be on
        and the invoice not to be complimentary -- both pre-existing
        controls, unchanged. On top of those, by default, the QR only
        appears on invoices that still owe money (Unpaid/Partially
        Paid): there's nothing left to pay on a fully settled invoice,
        so a "scan to pay" code there is just noise. Settings > Images
        & Assets has an "Always show the QR code" override for anyone
        who wants it printed regardless of payment status.
        """
        if not self.data.get("show_qr_code") or self.data.get("is_complimentary"):
            return False
        if self.config.get("pdf_settings", {}).get("always_show_qr_code", False):
            return True
        return self.data.get("status") in (InvoiceStatus.DUE.value, InvoiceStatus.PARTIAL.value)

    def _build_payment_lines(self) -> Optional[str]:
        """Composes the configured bank/UPI/payment-notes details (Settings
        > Payment & PDF) into PDF markup, or None if nothing is configured.
        Deliberately independent of _should_show_qr(): a business that
        only takes bank transfers (no QR/UPI) should still see its bank
        details printed. Skipped entirely for complimentary invoices,
        since there's nothing to pay on those."""
        if self.data.get("is_complimentary"):
            return None
        pd = self.config.get("payment_details", {})
        lines = []
        for label, key in (("Bank", "bank_name"), ("A/C Holder", "account_holder_name"),
                            ("A/C No", "account_number"), ("IFSC", "ifsc_code"),
                            ("UPI ID", "upi_id")):
            if val := pd.get(key):
                lines.append(f"{label}: {safe_markup(val)}")
        if notes := pd.get("payment_notes"):
            lines.append(safe_markup(notes).replace("\n", "<br/>"))
        return "<br/>".join(lines) if lines else None

    def _build_footer_section(self) -> Table:
        qr_cell = sig_cell = Spacer(0, 0)
        payment_lines = self._build_payment_lines()

        if self._should_show_qr() and (qr := self._asset("upi_qr_code")):
            body = payment_lines or "Scan to pay."
            detail = Paragraph(f"<b>PAYMENT DETAILS</b><br/>{body}", self.styles["NC"])
            qr_cell = Table([[Image(str(qr), 2.2*cm, 2.2*cm, kind="proportional"), detail]],
                            colWidths=[2.5*cm,"100%"], style=[("VALIGN",(0,0),(-1,-1),"TOP")])
        elif payment_lines:
            qr_cell = Paragraph(f"<b>PAYMENT DETAILS</b><br/>{payment_lines}", self.styles["NC"])

        if sig := self._asset("signature"):
            name = safe_markup(self.data.get("your_name"))
            sig_cell = Table(
                [[Image(str(sig), 4.0*cm, 2.0*cm, kind="proportional")],
                 [Paragraph(f"<b>{name}</b>", self.styles["Center"])]],
                style=[("LINEABOVE",(0,1),(0,1),1,colors.black),
                       ("TOPPADDING",(0,1),(0,1),3),
                       ("ALIGN",(0,0),(-1,-1),"CENTER")],
            )
        return Table([[qr_cell, sig_cell]], colWidths=["60%","40%"],
                     style=[("VALIGN",(0,0),(-1,-1),"BOTTOM")])

    def generate(self) -> Path:
        doc   = SimpleDocTemplate(str(self.filename), pagesize=A4,
                                  rightMargin=1.5*cm, leftMargin=1.5*cm,
                                  topMargin=1.5*cm,  bottomMargin=1.5*cm)
        story = [
            self._build_header(),         Spacer(1, 0.4*cm),
            self._build_address(),        Spacer(1, 0.4*cm),
            self._build_meta(),           Spacer(1, 0.4*cm),
            self._build_items(),
            self._build_totals(),         Spacer(1, 0.3*cm),
        ]
        if history := self._build_payment_history():
            story += [Paragraph("<b>Payment History</b>", self.styles["NoteBox"]), Spacer(1, 0.15*cm),
                      history, Spacer(1, 0.3*cm)]
        # Internal notes (self.data["notes"]) are intentionally never printed
        # on the customer-facing PDF -- they're for Vyaprix's eyes only.
        story += [Spacer(1, 0.3*cm), self._build_footer_section()]
        doc.build(story, onFirstPage=self.on_each_page, onLaterPages=self.on_each_page)
        return self.filename


def merge_pdfs(paths: List[Path], out_path: Path) -> Path:
    """Combine multiple single-invoice PDFs into one file for batch printing."""
    merged = fitz.open()
    try:
        for p in paths:
            with fitz.open(p) as src:
                merged.insert_pdf(src)
        merged.save(str(out_path))
    finally:
        merged.close()
    return out_path

# =============================================================================
#  SETTINGS DIALOG
# =============================================================================
class SettingsDialog(tk.Toplevel):
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

    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback  = callback
        self.title(f"{APP_NAME} Settings")
        self.geometry("680x780")
        self.minsize(620, 700)
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self._cfg = load_config()
        self._asset_rows: Dict[str, Tuple[tk.StringVar, ttk.Label, ttk.Label]] = {}
        self._currency_label_by_symbol = {o[0]: self._currency_option_label(o) for o in self.CURRENCY_OPTIONS}
        self._currency_symbol_by_label = {v: k for k, v in self._currency_label_by_symbol.items()}
        self._build()
        self._load()
        self.focus_force()
        self.bind("<Escape>", lambda e: self.destroy())

    @classmethod
    def _currency_option_label(cls, opt: Tuple[str, str, str]) -> str:
        symbol, code, name = opt
        return f"{symbol}   {code} — {name}"

    def _section(self, parent, title: str):
        ttk.Label(parent, text=title, font=(FONT_FAMILY, 11, "bold")).pack(anchor="w", pady=(0, 8))

    def _labeled_entry(self, parent, label: str, placeholder: str = "") -> "PlaceholderEntry":
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)
        ttk.Label(row, text=label, width=24, anchor="w").pack(side="left")
        entry = PlaceholderEntry(row, placeholder=placeholder)
        entry.pack(side="left", fill="x", expand=True)
        return entry

    def _labeled_combo(self, parent, label: str, var: tk.StringVar, values: list):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)
        ttk.Label(row, text=label, width=24, anchor="w").pack(side="left")
        ttk.Combobox(row, textvariable=var, values=values, state="readonly").pack(side="left", fill="x", expand=True)

    # ------------------------------------------------------------------ #
    #  Currency picker: a dropdown of common currencies, plus a "Custom"
    #  escape hatch (an ordinary text entry) for anything not on the list.
    # ------------------------------------------------------------------ #
    def _build_currency_picker(self, parent):
        all_labels = [self._currency_option_label(o) for o in self.CURRENCY_OPTIONS] + [self.CUSTOM_CURRENCY]

        self._currency_row = ttk.Frame(parent)
        self._currency_row.pack(fill="x", pady=4)
        ttk.Label(self._currency_row, text="Currency:", width=24, anchor="w").pack(side="left")
        self.currency_var = tk.StringVar()
        self.currency_combo = ttk.Combobox(self._currency_row, textvariable=self.currency_var,
                                            values=all_labels, state="readonly")
        self.currency_combo.pack(side="left", fill="x", expand=True)
        self.currency_combo.bind("<<ComboboxSelected>>", lambda e: self._on_currency_selected())

        self.custom_currency_row = ttk.Frame(parent)
        ttk.Label(self.custom_currency_row, text="Custom Symbol:", width=24, anchor="w").pack(side="left")
        self.custom_currency_entry = PlaceholderEntry(self.custom_currency_row, placeholder="e.g. Rs., kr, ₪")
        self.custom_currency_entry.pack(side="left", fill="x", expand=True)
        # Packed on demand by _on_currency_selected() / _load() -- only
        # shown once "Custom Symbol…" is actually chosen, right after the
        # picker row above.

    def _on_currency_selected(self):
        if self.currency_var.get() == self.CUSTOM_CURRENCY:
            self.custom_currency_row.pack(fill="x", pady=4, after=self._currency_row)
        else:
            self.custom_currency_row.pack_forget()

    # ------------------------------------------------------------------ #
    #  Asset upload rows: a thumbnail preview, a Browse button, and a
    #  Clear button -- deliberately no visible file name or path. The
    #  chosen file is still copied into Vyaprix's own data folder
    #  immediately on Browse, exactly as before, so the app keeps working
    #  even if the original file is later moved or deleted.
    # ------------------------------------------------------------------ #
    def _asset_row(self, parent, config_key: str, label: str, dest_filename: str):
        var = tk.StringVar()
        row = ttk.Frame(parent, padding=(0, 8))
        row.pack(fill="x")

        thumb = ttk.Label(row, relief="solid", borderwidth=1, anchor="center", width=8)
        thumb.pack(side="left", padx=(0, 12), ipady=14)

        info = ttk.Frame(row)
        info.pack(side="left", fill="x", expand=True)
        ttk.Label(info, text=label, font=(FONT_FAMILY, 10, "bold")).pack(anchor="w")
        status = ttk.Label(info, text="", font=(FONT_FAMILY, 9), bootstyle="secondary")
        status.pack(anchor="w")

        btns = ttk.Frame(row)
        btns.pack(side="right")
        ttk.Button(btns, text="Browse…", width=10, bootstyle="secondary-outline",
                   command=lambda k=config_key, d=dest_filename: self._browse_asset(k, d)).pack(side="left", padx=(0, 4))
        ttk.Button(btns, text="✕ Clear", width=8, bootstyle="danger-outline",
                   command=lambda k=config_key: self._clear_asset(k)).pack(side="left")

        self._asset_rows[config_key] = (var, thumb, status)

    def _refresh_asset_row(self, config_key: str):
        var, thumb, status = self._asset_rows[config_key]
        rel_path = var.get()
        full_path = (DATA_DIR / rel_path) if rel_path else None
        thumb_img = load_thumbnail_image(full_path, thumb, max_size=44) if full_path else None
        if thumb_img:
            thumb.configure(image=thumb_img, text="")
            thumb.image = thumb_img  # keep a reference alive, or Tk garbage-collects it
            status.configure(text="Custom file uploaded")
        else:
            thumb.configure(image="", text="—")
            thumb.image = None
            status.configure(text="No file selected — using default")

    def _browse_asset(self, config_key: str, dest_filename: str):
        path = filedialog.askopenfilename(parent=self, title="Choose an Image",
                                           filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
        if not path:
            return
        try:
            shutil.copy(path, ASSETS_DIR / dest_filename)
        except Exception as exc:
            Messagebox.show_error("Error", f"Could not save file: {exc}", parent=self)
            return
        var, _, _ = self._asset_rows[config_key]
        var.set(f"assets/{dest_filename}")
        self._refresh_asset_row(config_key)

    def _clear_asset(self, config_key: str):
        var, _, _ = self._asset_rows[config_key]
        var.set("")
        self._refresh_asset_row(config_key)

    # ------------------------------------------------------------------ #
    #  Tabs
    # ------------------------------------------------------------------ #
    def _build(self):
        main = ttk.Frame(self, padding=20)
        main.pack(fill="both", expand=True)

        nb = ttk.Notebook(main)
        nb.pack(fill="both", expand=True, pady=(0, 16))

        self._build_general_tab(nb)
        self._build_your_details_tab(nb)
        self._build_payment_tab(nb)
        self._build_assets_tab(nb)
        self._build_data_tab(nb)

        btn_row = ttk.Frame(main)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Cancel",             bootstyle="secondary", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(btn_row, text="Save Configuration", bootstyle="success",   command=self._save).pack(side="right")

    def _build_general_tab(self, nb):
        t1 = ttk.Frame(nb, padding=16)
        nb.add(t1, text="  General  ")
        self._section(t1, "Interface")
        self.theme_var = tk.StringVar()
        self._labeled_combo(t1, "App Theme:", self.theme_var,
                            ["litera","cosmo","flatly","darkly","cyborg","superhero","morph","pulse"])
        ttk.Separator(t1, bootstyle="secondary").pack(fill="x", pady=12)
        self._section(t1, "Company Details")
        ttk.Label(t1, text="Your one business identity: shown as the letterhead at the top of every "
                            "invoice, and as the name and GSTIN in the FROM block below it. There's "
                            "only one place to set these, so they can't disagree with each other.",
                  font=(FONT_FAMILY, 9), bootstyle="secondary", wraplength=520, justify="left").pack(anchor="w", pady=(0, 6))
        self.co_name = self._labeled_entry(t1, "Company Name:",    "Your company or firm name")
        self.co_desc = self._labeled_entry(t1, "Company Tagline:", "A short tagline, e.g. Business Made Smarter.")
        self.co_gst  = self._labeled_entry(t1, "Company GSTIN:",   "22ABCDE1234F1Z5")

    def _build_your_details_tab(self, nb):
        t2 = ttk.Frame(nb, padding=(16, 16, 4, 16))
        nb.add(t2, text="  Your Details  ")
        scroller = ScrolledFrame(t2, autohide=True)
        scroller.pack(fill="both", expand=True)

        self._section(scroller, "Your Details")
        ttk.Label(scroller, text="Your personal contact details, shown under the business name in the "
                                  "FROM block on new invoices. Editing this doesn't change any invoice "
                                  "you've already created. For your business/firm name and GSTIN, see "
                                  "Company Details on the General tab -- there's just one of each.",
                  font=(FONT_FAMILY, 9), bootstyle="secondary", wraplength=500, justify="left").pack(anchor="w", pady=(0, 10))
        self.ud_name    = self._labeled_entry(scroller, "Your Name:",           "e.g. Priya Sharma")
        self.ud_phone   = self._labeled_entry(scroller, "Phone Number:",        "+91 98765 43210")
        self.ud_email   = self._labeled_entry(scroller, "Email Address:",       "you@example.com")
        self.ud_website = self._labeled_entry(scroller, "Website (Optional):",  "www.example.com")
        self.ud_address = self._labeled_entry(scroller, "Address:",             "Street / building / area")
        self.ud_city    = self._labeled_entry(scroller, "City:",                "e.g. Lucknow")
        self.ud_state   = self._labeled_entry(scroller, "State:",               "e.g. Uttar Pradesh")
        self.ud_pin     = self._labeled_entry(scroller, "PIN Code:",            "e.g. 226001")

    def _build_payment_tab(self, nb):
        t3 = ttk.Frame(nb, padding=(16, 16, 4, 16))
        nb.add(t3, text="  Payment & PDF  ")
        scroller = ScrolledFrame(t3, autohide=True)
        scroller.pack(fill="both", expand=True)

        self._section(scroller, "Payment Settings")
        self._build_currency_picker(scroller)
        self.bank_name  = self._labeled_entry(scroller, "Bank Name:",           "e.g. State Bank of India")
        self.acc_holder = self._labeled_entry(scroller, "Account Holder Name:", "e.g. Priya Sharma")
        self.acc_number = self._labeled_entry(scroller, "Account Number:",      "e.g. 000123456789")
        self.ifsc_code  = self._labeled_entry(scroller, "IFSC Code:",           "e.g. SBIN0001234")
        self.upi_id     = self._labeled_entry(scroller, "UPI ID:",              "company@upi")
        ttk.Label(scroller, text="Payment Notes:", font=(FONT_FAMILY, 9)).pack(anchor="w", pady=(9, 2))
        self.payment_notes_text = PlaceholderText(
            scroller, placeholder="Anything else to show on the invoice -- cheque instructions, "
                                   "a late-fee policy, alternate payment methods, etc.",
            height=3, relief="flat", highlightthickness=1, wrap="word")
        self.payment_notes_text.pack(fill="x")
        self.pay_terms = self._labeled_entry(scroller, "Default Payment Terms:", "Due on Receipt")

        ttk.Separator(scroller, bootstyle="secondary").pack(fill="x", pady=12)
        self._section(scroller, "PDF Footer Note")
        self.footer_text = PlaceholderText(scroller, placeholder="Thank you for your business!",
                                            height=3, relief="flat", highlightthickness=1, wrap="word")
        self.footer_text.pack(fill="x")

    def _build_assets_tab(self, nb):
        t4 = ttk.Frame(nb, padding=16)
        nb.add(t4, text="  Images & Assets  ")
        self._section(t4, "Upload Assets")
        ttk.Label(t4, text="Files you choose are copied into Vyaprix's own data folder right away, "
                            "so the app keeps working even if you later move or delete the original.",
                  font=(FONT_FAMILY, 9), bootstyle="secondary", wraplength=540, justify="left").pack(anchor="w", pady=(0, 2))

        self._asset_row(t4, "logo",        "Company Logo",         "app_logo.png")
        ttk.Separator(t4, bootstyle="secondary").pack(fill="x", pady=2)
        self._asset_row(t4, "signature",   "Authorized Signature", "sign.png")
        ttk.Separator(t4, bootstyle="secondary").pack(fill="x", pady=2)
        self._asset_row(t4, "upi_qr_code", "UPI QR Code",          "upi_qr.png")

        ttk.Separator(t4, bootstyle="secondary").pack(fill="x", pady=12)
        self._section(t4, "QR Code Visibility")
        self.always_show_qr_var = tk.BooleanVar()
        qr_check = ttk.Checkbutton(t4, text="Always show the QR code, even on fully paid invoices",
                                    variable=self.always_show_qr_var, bootstyle="round-toggle")
        qr_check.pack(anchor="w")
        ttk.Label(t4, text="Off by default: the QR code only prints on Unpaid or Partially Paid invoices, "
                            "since there's nothing left to pay once an invoice is settled. Turn this on to "
                            "always include it regardless of payment status.",
                  font=(FONT_FAMILY, 9), bootstyle="secondary", wraplength=540, justify="left").pack(anchor="w", pady=(4, 0))

    def _build_data_tab(self, nb):
        t5 = ttk.Frame(nb, padding=16)
        nb.add(t5, text="  Data  ")
        self._section(t5, "Database Location")
        ttk.Label(t5, text=str(DB_NAME), font=(FONT_FAMILY, 9), bootstyle="secondary", wraplength=540).pack(anchor="w")
        ttk.Separator(t5, bootstyle="secondary").pack(fill="x", pady=12)

        self._section(t5, "Backup & Restore")
        backup_btn = ttk.Button(t5, text="💾  Save a Backup Copy…  (Ctrl+B)", bootstyle="secondary-outline",
                                 command=self._backup_db)
        backup_btn.pack(anchor="w")
        ToolTip(backup_btn, "Back up the database now (Ctrl+B works from anywhere in the app)")

        import_btn = ttk.Button(t5, text="📥  Import Database…", bootstyle="warning-outline",
                                 command=self._import_db)
        import_btn.pack(anchor="w", pady=(10, 0))
        ttk.Label(t5, text="Replaces all invoices, clients, and catalog items with the contents of "
                            "the file you choose. Your current database is backed up automatically "
                            "first, and the app refreshes automatically once the import finishes.",
                  font=(FONT_FAMILY, 9), bootstyle="secondary", wraplength=540, justify="left").pack(anchor="w", pady=(4, 0))

    def _backup_db(self):
        prompt_and_backup_database(self)

    def _import_db(self):
        """Lets the user replace the live database with another .db file:
        validate it opens as SQLite, confirm with the user (this is a
        destructive action), back up the CURRENT database first, copy the
        new file into place, re-run schema migrations against it (so an
        older or partial .db still works with the current app), and
        refresh every screen. Any failure along the way restores the
        pre-import backup automatically, so a bad file can't leave the
        app worse off than before the import was attempted.
        """
        path = filedialog.askopenfilename(
            parent=self, title="Import Database",
            filetypes=[("SQLite Database", "*.db"), ("All Files", "*.*")],
        )
        if not path:
            return
        path = Path(path)

        try:
            with closing(sqlite3.connect(str(path))) as test_conn:
                test_conn.execute("PRAGMA schema_version")
        except sqlite3.Error as exc:
            Messagebox.show_error("Invalid Database",
                                   f"This doesn't look like a valid database file:\n{exc}", parent=self)
            return

        if not Messagebox.okcancel(
            "Import Database",
            "This replaces all invoices, clients, and catalog items in Vyaprix with the "
            "contents of the selected file.\n\nYour current database will be backed up "
            "automatically first, so you can go back if anything looks wrong.\n\nContinue?",
            parent=self,
        ):
            return

        backup_path = DATA_DIR / f"Vyaprix-pre-import-backup-{datetime.datetime.now().strftime('%Y%m%d-%H%M%S')}.db"
        try:
            if DB_NAME.exists():
                shutil.copy2(DB_NAME, backup_path)
        except Exception as exc:
            Messagebox.show_error("Import Failed",
                                   f"Could not create a safety backup, so nothing was changed:\n{exc}", parent=self)
            return

        try:
            shutil.copy2(path, DB_NAME)
            # Re-run schema setup/migrations against the newly-imported
            # file (adds any missing tables/columns/indexes). Schema
            # changes apply to the database FILE itself, and every
            # DatabaseManager call reopens its own connection to that
            # same path, so the app's existing db_manager picks up the
            # migrated data automatically -- no object references to
            # swap here.
            DatabaseManager(DB_NAME)
        except Exception as exc:
            try:
                shutil.copy2(backup_path, DB_NAME)
            except Exception:
                pass
            Messagebox.show_error(
                "Import Failed",
                f"The selected file couldn't be used as a Vyaprix database:\n{exc}\n\n"
                "Your previous database has been restored.",
                parent=self,
            )
            return

        Messagebox.show_info(
            "Import Complete",
            f"Database imported successfully.\n\nYour previous database was backed up to:\n{backup_path.name}",
            parent=self,
        )
        self.callback(None)  # reloads CONFIG and refreshes every screen's list

    def _load(self):
        c = self._cfg
        self.theme_var.set(c.get("theme", "litera"))
        self.co_name.insert(0, c.get("company_details", {}).get("name", ""))
        self.co_desc.insert(0, c.get("company_details", {}).get("description", ""))
        self.co_gst.insert(0, c.get("company_details", {}).get("gstin", ""))

        ud = c.get("user_details", {})
        self.ud_name.insert(0, ud.get("name", ""))
        self.ud_phone.insert(0, ud.get("phone", ""))
        self.ud_email.insert(0, ud.get("email", ""))
        self.ud_website.insert(0, ud.get("website", ""))
        self.ud_address.insert(0, ud.get("address", ""))
        self.ud_city.insert(0, ud.get("city", ""))
        self.ud_state.insert(0, ud.get("state", ""))
        self.ud_pin.insert(0, ud.get("pin_code", ""))

        current_symbol = c.get("currency_symbol", "₹")
        if current_symbol in self._currency_label_by_symbol:
            self.currency_var.set(self._currency_label_by_symbol[current_symbol])
            self.custom_currency_row.pack_forget()
        else:
            self.currency_var.set(self.CUSTOM_CURRENCY)
            self.custom_currency_entry.insert(0, current_symbol)
            self.custom_currency_row.pack(fill="x", pady=4, after=self._currency_row)

        pd = c.get("payment_details", {})
        self.bank_name.insert(0, pd.get("bank_name", ""))
        self.acc_holder.insert(0, pd.get("account_holder_name", ""))
        self.acc_number.insert(0, pd.get("account_number", ""))
        self.ifsc_code.insert(0, pd.get("ifsc_code", ""))
        self.upi_id.insert(0, pd.get("upi_id", ""))
        self.payment_notes_text.insert("1.0", pd.get("payment_notes", ""))
        self.pay_terms.insert(0, pd.get("payment_terms", ""))

        self.footer_text.insert("1.0", c.get("pdf_settings", {}).get("footer_message", ""))
        self.always_show_qr_var.set(bool(c.get("pdf_settings", {}).get("always_show_qr_code", False)))

        ap = c.get("asset_paths", {})
        for key in ("logo", "signature", "upi_qr_code"):
            var, _, _ = self._asset_rows[key]
            var.set(ap.get(key, ""))
            self._refresh_asset_row(key)

    def _save(self):
        if self.currency_var.get() == self.CUSTOM_CURRENCY:
            currency_symbol = self.custom_currency_entry.get().strip() or "₹"
        else:
            currency_symbol = self._currency_symbol_by_label.get(self.currency_var.get(), "₹")

        save_config({
            "theme": self.theme_var.get(),
            "company_details": {
                "name": self.co_name.get().strip(),
                "description": self.co_desc.get().strip(),
                "gstin": self.co_gst.get().strip(),
            },
            "user_details": {
                "name": self.ud_name.get().strip(),
                "phone": self.ud_phone.get().strip(),
                "email": self.ud_email.get().strip(),
                "website": self.ud_website.get().strip(),
                "address": self.ud_address.get().strip(),
                "city": self.ud_city.get().strip(),
                "state": self.ud_state.get().strip(),
                "pin_code": self.ud_pin.get().strip(),
                # Preserved but no longer editable here directly -- see
                # compose_your_contact()'s upgrade-path fallback.
                "contact_info": self._cfg.get("user_details", {}).get("contact_info", ""),
            },
            "payment_details": {
                "upi_id": self.upi_id.get().strip(),
                "payment_terms": self.pay_terms.get().strip(),
                "bank_name": self.bank_name.get().strip(),
                "account_holder_name": self.acc_holder.get().strip(),
                "account_number": self.acc_number.get().strip(),
                "ifsc_code": self.ifsc_code.get().strip(),
                "payment_notes": self.payment_notes_text.get("1.0", "end-1c").strip(),
            },
            "pdf_settings": {"footer_message": self.footer_text.get("1.0", "end-1c").strip(),
                              "always_show_qr_code": self.always_show_qr_var.get()},
            "asset_paths": {key: self._asset_rows[key][0].get() for key in ("logo", "signature", "upi_qr_code")},
            "currency_symbol": currency_symbol,
        })

        # PASS THE NEW THEME TO THE CALLBACK
        self.callback(self.theme_var.get())
        self.destroy()

# =============================================================================
#  ABOUT DIALOG
# =============================================================================
class AboutDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title(f"About {APP_NAME}")
        self.geometry("400x420")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self._build()
        self.focus_force()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build(self):
        style = ttk.Style()
        colors = style.colors
        self.configure(bg=colors.bg)
        container = ttk.Frame(self, padding=15)
        container.pack(fill="both", expand=True)

        hero = ttk.Frame(container, padding=15, bootstyle="primary")
        hero.pack(fill="x")
        ttk.Label(hero, text=APP_NAME, font=(FONT_FAMILY, 18, "bold"), bootstyle="inverse-primary", anchor="center", justify="center").pack(fill="x")
        ttk.Label(hero, text=APP_TAGLINE, font=(FONT_FAMILY, 9), bootstyle="inverse-primary", anchor="center", justify="center").pack(fill="x", pady=(4, 0))

        ttk.Label(container, text=APP_VERSION, font=(FONT_FAMILY, 28, "bold"), bootstyle="info").pack(pady=(20, 5))
        ttk.Label(container, text=f"Released on {APP_RELEASE_DATE}", font=(FONT_FAMILY, 9), bootstyle="secondary").pack()
        ttk.Label(container, text="Now with client & catalog managers, payment\nhistory tracking, and bulk invoice actions.",
                  font=(FONT_FAMILY, 8), bootstyle="secondary", justify="center", anchor="center").pack(pady=(6, 0))
        ttk.Separator(container).pack(fill="x", pady=15)

        lines = APP_ATTRIBUTION.split("\n")
        ttk.Label(container, text=lines[0], font=(FONT_FAMILY, 11, "bold"), anchor="center", justify="center").pack(fill="x")
        for line in lines[1:]:
            ttk.Label(container, text=line, font=(FONT_FAMILY, 9), bootstyle="secondary", anchor="center", justify="center").pack(fill="x", pady=1)

        ttk.Button(container, text="Close", width=14, bootstyle="primary", command=self.destroy).pack(side="bottom", pady=(20, 0))

# =============================================================================
#  KEYBOARD SHORTCUTS REFERENCE
# =============================================================================
class ShortcutsDialog(tk.Toplevel):
    """A reference card listing every keyboard shortcut the app responds
    to. Reads directly from the app's ShortcutManager (self.shortcuts)
    rather than a hardcoded list, so this dialog can never drift out of
    sync with what a key combo actually does -- add or change a
    shortcut in InvoiceApp._setup_keyboard_shortcuts and this dialog
    updates itself automatically. Opened from the sidebar's
    "⌨  Shortcuts" button or by pressing F1 anywhere."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Keyboard Shortcuts")
        self.geometry("440x600")
        self.minsize(400, 420)
        self.transient(parent)
        self.grab_set()
        self._build(parent.shortcuts.grouped_reference())
        self.focus_force()
        self.bind("<Escape>", lambda e: self.destroy())

    def _build(self, groups: Dict[str, List[Tuple[str, str]]]):
        style = ttk.Style()
        self.configure(bg=style.colors.bg)
        container = ttk.Frame(self, padding=15)
        container.pack(fill="both", expand=True)

        hero = ttk.Frame(container, padding=15, bootstyle="primary")
        hero.pack(fill="x")
        ttk.Label(hero, text="⌨  Keyboard Shortcuts", font=(FONT_FAMILY, 15, "bold"),
                  bootstyle="inverse-primary", anchor="center", justify="center").pack(fill="x")
        ttk.Label(hero, text="Everything below works without touching the mouse",
                  font=(FONT_FAMILY, 9), bootstyle="inverse-primary", anchor="center", justify="center").pack(fill="x", pady=(4, 0))

        scroller = ScrolledFrame(container, autohide=True)
        scroller.pack(fill="both", expand=True, pady=(15, 0))

        for group_title, items in groups.items():
            ttk.Label(scroller, text=group_title, font=(FONT_FAMILY, 11, "bold"),
                      bootstyle="secondary").pack(anchor="w", pady=(12, 6))
            for keys, desc in items:
                row = ttk.Frame(scroller)
                row.pack(fill="x", pady=3)
                ttk.Label(row, text=keys, font=(FONT_FAMILY, 9, "bold"), width=16,
                          bootstyle="inverse-secondary", anchor="center", padding=(4, 3)).pack(side="left")
                ttk.Label(row, text=desc, font=(FONT_FAMILY, 10)).pack(side="left", padx=(10, 0))

        ttk.Button(container, text="Close", width=14, bootstyle="primary", command=self.destroy).pack(side="bottom", pady=(15, 0))

# =============================================================================
#  PDF VIEWER
# =============================================================================
class PDFViewer(tk.Toplevel):
    MIN_ZOOM = 0.4
    MAX_ZOOM = 3.5

    def __init__(self, parent, pdf_path: Path):
        super().__init__(parent)
        self.pdf_path   = pdf_path
        self.title(f"Preview  —  {pdf_path.name}")
        self.geometry("740x720")
        self.minsize(740, 720)
        self.maxsize(1920, 1080)

        try:
            self.doc = fitz.open(pdf_path)
        except Exception as exc:
            Messagebox.show_error("PDF Error", f"Cannot open file:\n{exc}")
            self.destroy(); return

        self.page_idx   = 0
        self.total      = len(self.doc)
        self.zoom       = 1.0
        self._img       = None
        self._resize_id = None

        self._build()
        self.bind("<Configure>",     self._on_resize)
        self.bind("<Left>",          lambda _: self._prev())
        self.bind("<Right>",         lambda _: self._next())
        self.bind("<Prior>",         lambda _: self._prev())
        self.bind("<Next>",          lambda _: self._next())
        self.bind("<Control-equal>", lambda _: self._zoom(+0.15))
        self.bind("<Control-minus>", lambda _: self._zoom(-0.15))
        self.bind("<Control-0>",     lambda _: self._reset_zoom())
        # Bound locally (rather than left to the app-wide shortcuts) so
        # printing/saving act on *this* PDF the moment the viewer has
        # focus; _bind_action's "break" keeps the keystroke from also
        # reaching the Dashboard's own Ctrl+P/Ctrl+E handling underneath.
        self.bind("<Control-p>", _bind_action(self._print))
        self.bind("<Control-e>", _bind_action(self._save_copy))
        self.bind("<Escape>",    _bind_action(self._close))
        self.protocol("WM_DELETE_WINDOW", self._close)
        self.transient(parent)
        self.grab_set()
        self.after(120, self._render)

    def _build(self):
        tb = ttk.Frame(self, padding=(10, 7))
        tb.pack(side="top", fill="x")

        prev_btn = ttk.Button(tb, text="◀  Prev",  command=self._prev,         width=9)
        next_btn = ttk.Button(tb, text="Next  ▶",  command=self._next,         width=9)
        prev_btn.pack(side="left", padx=(0, 4))
        next_btn.pack(side="left", padx=(0, 14))
        ToolTip(prev_btn, "Previous page (← or Page Up)")
        ToolTip(next_btn, "Next page (→ or Page Down)")
        self.page_lbl = ttk.Label(tb, text="", font=(FONT_FAMILY, 10))
        self.page_lbl.pack(side="left")

        close_btn = ttk.Button(tb, text="✕  Close",  bootstyle="danger",    command=self._close,      width=9)
        save_btn  = ttk.Button(tb, text="💾  Save",   bootstyle="secondary", command=self._save_copy,  width=9)
        print_btn = ttk.Button(tb, text="🖨  Print",  bootstyle="success",   command=self._print,      width=9)
        close_btn.pack(side="right", padx=(4, 0))
        save_btn.pack(side="right", padx=4)
        print_btn.pack(side="right", padx=4)
        ToolTip(close_btn, "Close this viewer (Esc)")
        ToolTip(save_btn,  "Save a copy of this PDF (Ctrl+E)")
        ToolTip(print_btn, "Print this PDF (Ctrl+P)")

        ttk.Separator(tb, orient="vertical").pack(side="right", fill="y", padx=10)
        self.zoom_lbl = ttk.Label(tb, text="100%", font=(FONT_FAMILY, 9), width=6, anchor="center")
        self.zoom_lbl.pack(side="right")
        ttk.Button(tb, text="⊕", command=lambda: self._zoom(+0.15), width=3, bootstyle="secondary-outline").pack(side="right", padx=2)
        ttk.Button(tb, text="⊖", command=lambda: self._zoom(-0.15), width=3, bootstyle="secondary-outline").pack(side="right", padx=2)
        ttk.Button(tb, text="Fit",command=self._reset_zoom,        width=4, bootstyle="secondary-outline").pack(side="right", padx=2)

        ttk.Separator(self, bootstyle="secondary").pack(fill="x")
        self.canvas = tk.Canvas(self, bg="#555555", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        ttk.Label(self, text="← → or Page Up/Down to navigate  •  Ctrl + / − zoom  •  Ctrl+0 fit  •  "
                             "Ctrl+P print  •  Ctrl+E save  •  Esc close",
                  font=(FONT_FAMILY, 8), bootstyle="secondary").pack(side="bottom", pady=3)

    def _on_resize(self, _=None):
        if self._resize_id:
            self.after_cancel(self._resize_id)
        self._resize_id = self.after(130, self._render)

    def _render(self):
        cw = self.canvas.winfo_width()
        ch = self.canvas.winfo_height()
        if cw <= 1 or ch <= 1:
            return
        page    = self.doc.load_page(self.page_idx)
        fit_z   = min(cw / page.rect.width, ch / page.rect.height) * 0.97
        z       = fit_z * self.zoom
        pix     = page.get_pixmap(matrix=fitz.Matrix(z, z))
        self._img = tk.PhotoImage(data=pix.tobytes("ppm"))
        self.canvas.delete("all")
        x = max((cw - pix.width) / 2, 0)
        y = max((ch - pix.height) / 2, 0)
        self.canvas.create_image(x, y, anchor="nw", image=self._img)
        self.page_lbl.config(text=f"Page {self.page_idx + 1} / {self.total}")
        self.zoom_lbl.config(text=f"{int(self.zoom * 100)}%")

    def _next(self):
        if self.page_idx < self.total - 1:
            self.page_idx += 1
            self._render()

    def _prev(self):
        if self.page_idx > 0:
            self.page_idx -= 1
            self._render()

    def _zoom(self, delta: float):
        self.zoom = max(self.MIN_ZOOM, min(self.MAX_ZOOM, self.zoom + delta))
        self._render()

    def _reset_zoom(self):
        self.zoom = 1.0
        self._render()

    def _print(self):
        if sys.platform != "win32":
            try:
                subprocess.run(["lp", str(self.pdf_path)], check=True)
            except Exception as exc:
                Messagebox.show_error("Print Error", str(exc), parent=self)
            return

        try:
            from PyQt5.QtWidgets import QApplication
            from PyQt5.QtPrintSupport import QPrinterInfo

            # QApplication must exist before we can query the hardware for printers
            qapp = QApplication.instance()
            if qapp is None:
                qapp = QApplication(sys.argv)

            printers = QPrinterInfo.availablePrinterNames()
            default_printer = QPrinterInfo.defaultPrinterName()

        except ImportError:
            # Fallback to standard Windows print if PyQt5 is missing
            try:
                os.startfile(self.pdf_path, "print")
            except OSError:
                Messagebox.show_warning("Manual Print Required",
                    "A PDF reader is not set as default.\n"
                    "The file will open — press Ctrl+P to print.", parent=self)
                os.startfile(self.pdf_path)
            return

        # --- Create Custom Tkinter/ttkbootstrap Printer Dialog ---
        dialog = tk.Toplevel(self)
        dialog.title("Select Printer")
        dialog.geometry("400x260")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        ttk.Label(dialog, text="🖨️ Choose a printer destination:", font=("Segoe UI", 10, "bold")).pack(pady=(15, 5), padx=20, anchor="w")

        combo = ttk.Combobox(dialog, values=printers, state="readonly", font=("Segoe UI", 9))
        combo.pack(fill="x", padx=20, pady=5)

        if printers:
            if default_printer in printers:
                combo.set(default_printer)
            else:
                combo.set(printers[0])
        else:
            combo.set("No printers detected")
            combo.config(state="disabled")

        # --- Button Commands ---
        def on_print():
            if not printers:
                return
            selected_printer = combo.get()
            dialog.destroy()
            self._print_with_pyqt5(printer_name=selected_printer)

        def on_system_dialog():
            dialog.destroy()
            self._print_with_pyqt5(show_dialog=True)

        def on_manual_open():
            dialog.destroy()
            os.startfile(self.pdf_path)

        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill="x", padx=20, pady=10)

        ttk.Button(btn_frame, text="Cancel", bootstyle="secondary-outline", command=dialog.destroy).pack(side="right", padx=(10, 0))

        print_btn = ttk.Button(btn_frame, text="Print Document", bootstyle="success", command=on_print)
        print_btn.pack(side="right")
        if not printers:
            print_btn.config(state="disabled")

        ttk.Separator(dialog, bootstyle="secondary").pack(fill="x", padx=20, pady=5)

        alt_frame = ttk.Frame(dialog)
        alt_frame.pack(fill="x", padx=20, pady=5)

        ttk.Button(alt_frame, text="⚙️ Use System Print Dialog", bootstyle="info-outline", command=on_system_dialog).pack(fill="x", pady=2)
        ttk.Button(alt_frame, text="🌐 Open PDF Manually (Browser/Viewer)", bootstyle="primary-outline", command=on_manual_open).pack(fill="x", pady=2)

    def _print_with_pyqt5(self, printer_name=None, show_dialog=False):
        """Handles both headless printing and system dialog printing using the PyQt5 engine"""
        from PyQt5.QtWidgets import QApplication
        from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
        from PyQt5.QtGui import QImage, QPixmap, QPainter
        from PyQt5.QtCore import Qt

        painter = None
        doc = None
        try:
            qapp = QApplication.instance()
            if qapp is None:
                qapp = QApplication(sys.argv)

            printer = QPrinter(QPrinter.HighResolution)

            if show_dialog:
                print_dlg = QPrintDialog(printer)
                if print_dlg.exec_() != QPrintDialog.Accepted:
                    return
            else:
                printer.setPrinterName(printer_name)

            painter = QPainter(printer)
            doc = fitz.open(self.pdf_path)

            for page_index in range(len(doc)):
                if page_index > 0:
                    printer.newPage()

                page = doc.load_page(page_index)
                pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
                fmt = QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888
                qimage = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt)
                pixmap = QPixmap.fromImage(qimage)

                rect = painter.viewport()
                size = pixmap.size()
                size.scale(rect.size(), Qt.KeepAspectRatio)

                x_offset = int((rect.width() - size.width()) / 2)
                y_offset = int((rect.height() - size.height()) / 2)

                painter.setViewport(x_offset, y_offset, size.width(), size.height())
                painter.setWindow(pixmap.rect())
                painter.drawPixmap(0, 0, pixmap)

            if not show_dialog:
                Messagebox.show_info("Success", f"Document sent to {printer_name}!", parent=self)

        except Exception as exc:
            Messagebox.show_error("Print Error", str(exc), parent=self)
        finally:
            # BUG FIX: cleanup used to live only on the success path, so a
            # failure partway through the page loop (e.g. one corrupt page)
            # skipped painter.end()/doc.close() -- leaking the PDF file
            # handle and leaving the painter attached to the print spooler.
            if painter is not None:
                try:
                    painter.end()
                except Exception:
                    pass
            if doc is not None:
                doc.close()

    def _save_copy(self):
        if path := filedialog.asksaveasfilename(
            initialfile=self.pdf_path.name, defaultextension=".pdf",
            filetypes=[("PDF Documents", "*.pdf")]
        ):
            try:
                shutil.copy(self.pdf_path, path)
                Messagebox.show_info("Saved", f"Invoice saved to:\n{path}")
            except Exception as exc:
                Messagebox.show_error("Save Error", str(exc))

    def _close(self):
        self.doc.close()
        self.destroy()

# =============================================================================
#  SMALL RECORD DIALOGS  (Client / Catalog Item / Invoice Note)
# =============================================================================
class ClientDialog(tk.Toplevel):
    """Add or edit a single Address Book entry."""
    def __init__(self, parent, db_manager: "DatabaseManager", client: Optional[Dict], on_saved: Callable):
        super().__init__(parent)
        self.db_manager = db_manager
        self.client     = client
        self.on_saved   = on_saved
        self.title("Edit Client" if client else "Add Client")
        self.geometry("420x380")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        pad = ttk.Frame(self, padding=18)
        pad.pack(fill="both", expand=True)

        ttk.Label(pad, text="Client / Company Name *", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.name_entry = PlaceholderEntry(pad, placeholder="Enter client name")
        self.name_entry.pack(fill="x", pady=(2, 10))

        ttk.Label(pad, text="Contact Info", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.contact_text = PlaceholderText(pad, placeholder="+91 98765 43210, name@email.com, City",
                                             height=5, relief="flat", highlightthickness=1, wrap="word")
        self.contact_text.pack(fill="x", pady=(2, 10))

        ttk.Label(pad, text="GSTIN (Optional)", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.gstin_entry = PlaceholderEntry(pad, placeholder="22ABCDE1234F1Z5")
        self.gstin_entry.pack(fill="x", pady=(2, 10))

        if client:
            self.name_entry.insert(0, client.get("name", ""))
            self.contact_text.insert("1.0", client.get("contact", "") or "")
            self.gstin_entry.insert(0, client.get("gstin", "") or "")

        btns = ttk.Frame(pad)
        btns.pack(fill="x", pady=(10, 0), side="bottom")
        ttk.Button(btns, text="Cancel", bootstyle="secondary", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(btns, text="Save", bootstyle="success", command=self._save).pack(side="right")
        self.name_entry.focus_set()
        
        # --- THE FIX IS HERE ---
        # Instead of calling _save() directly, we call a new handler
        self.bind("<Return>", self._on_enter_pressed)
        self.bind("<Escape>", lambda _: self.destroy())

    # --- ADD THIS NEW METHOD ---
    def _on_enter_pressed(self, event):
        """Handle the Enter key smartly based on what the user is doing."""
        # If the user is inside the multi-line text box, let them make a new line!
        if self.focus_get() == self.contact_text:
            return 
        
        # Otherwise, save the form.
        self._save()

    def _save(self):
        name = self.name_entry.get().strip()
        if not name:
            return Messagebox.show_error("Validation", "Client name is required.", parent=self)
        contact = self.contact_text.get("1.0", "end-1c").strip()
        gstin   = self.gstin_entry.get().strip()
        cid = self.client.get("id") if self.client else None
        ok, err = self.db_manager.upsert_client_by_id(cid, name, contact, gstin)
        if ok:
            self.on_saved()
            self.destroy()
        else:
            Messagebox.show_error("Could Not Save", err, parent=self)

class CatalogItemDialog(tk.Toplevel):
    """Add or edit a single Product & Service Catalog entry."""
    def __init__(self, parent, db_manager: "DatabaseManager", item: Optional[Dict], on_saved: Callable):
        super().__init__(parent)
        self.db_manager = db_manager
        self.item       = item
        self.on_saved   = on_saved
        self.title("Edit Catalog Item" if item else "Add Catalog Item")
        self.geometry("420x360")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        pad = ttk.Frame(self, padding=18)
        pad.pack(fill="both", expand=True)

        ttk.Label(pad, text="Item / Service Title *", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.title_entry = PlaceholderEntry(pad, placeholder="Enter item / service name")
        self.title_entry.pack(fill="x", pady=(2, 10))

        ttk.Label(pad, text="Description", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.desc_text = PlaceholderText(pad, placeholder="Description (optional)",
                                          height=5, relief="flat", highlightthickness=1, wrap="word")
        self.desc_text.pack(fill="x", pady=(2, 10))

        ttk.Label(pad, text=f"Default Amount ({CURRENCY_SYMBOL})", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.amount_entry = ttk.Entry(pad)
        self.amount_entry.pack(fill="x", pady=(2, 10))

        if item:
            self.title_entry.insert(0, item.get("title", ""))
            self.desc_text.insert("1.0", item.get("description", "") or "")
            self.amount_entry.insert(0, str(item.get("amount", 0) or 0))
        else:
            self.amount_entry.insert(0, "0")

        btns = ttk.Frame(pad)
        btns.pack(fill="x", pady=(10, 0), side="bottom")
        ttk.Button(btns, text="Cancel", bootstyle="secondary", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(btns, text="Save", bootstyle="success", command=self._save).pack(side="right")
        self.title_entry.focus_set()
        self.bind("<Return>", lambda _: self._save())
        self.bind("<Escape>", lambda _: self.destroy())

    def _save(self):
        title = self.title_entry.get().strip()
        if not title:
            return Messagebox.show_error("Validation", "Item title is required.", parent=self)
        try:
            amount = float(self.amount_entry.get() or 0)
        except ValueError:
            return Messagebox.show_error("Validation", "Amount must be a number.", parent=self)
        desc = self.desc_text.get("1.0", "end-1c").strip()
        iid = self.item.get("id") if self.item else None
        ok, err = self.db_manager.upsert_catalog_item_by_id(iid, title, desc, amount)
        if ok:
            self.on_saved()
            self.destroy()
        else:
            Messagebox.show_error("Could Not Save", err, parent=self)

class NoteDialog(tk.Toplevel):
    """Quick popup to add/edit an invoice's internal note straight from the
    invoice list, without opening the full invoice form."""
    def __init__(self, parent, db_manager: "DatabaseManager", invoice_number: str,
                 current_note: str, on_saved: Callable):
        super().__init__(parent)
        self.db_manager     = db_manager
        self.invoice_number = invoice_number
        self.on_saved       = on_saved
        self.title(f"Note  —  Invoice #{invoice_number}")
        self.geometry("420x260")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        pad = ttk.Frame(self, padding=18)
        pad.pack(fill="both", expand=True)
        ttk.Label(pad, text="Internal note (not printed on the invoice PDF)",
                  font=(FONT_FAMILY, 9), bootstyle="secondary").pack(anchor="w", pady=(0, 6))
        self.text = PlaceholderText(pad, placeholder="Add an internal note...",
                                     height=7, relief="flat", highlightthickness=1, wrap="word")
        self.text.pack(fill="both", expand=True)
        self.text.insert("1.0", current_note or "")

        btns = ttk.Frame(pad)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Cancel", bootstyle="secondary", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(btns, text="Save Note", bootstyle="success", command=self._save).pack(side="right")
        self.text.focus_set()
        self.bind("<Escape>", lambda _: self.destroy())

    def _save(self):
        note = self.text.get("1.0", "end-1c").strip()
        self.db_manager.update_notes(self.invoice_number, note)
        self.on_saved()
        self.destroy()

# =============================================================================
#  BULK PROGRESS DIALOG
# =============================================================================
class BulkProgressDialog(tk.Toplevel):
    """Generic progress popup for any bulk operation (Delete, Export,
    Print/Generate). Runs work_fn in a background thread and polls its
    own queue via .after(), so the UI never blocks.

    work_fn(items, report, is_aborted) -> dict, called on the worker
    thread:
      report(current, total, status_text, log_line=None,
             success_delta=0, failed_delta=0) -- one progress update;
             deltas let a chunked operation report >1 item per call.
      is_aborted() -- True once Cancel was clicked; check periodically
             and stop promptly.
      Returns {"success": int, "failed": int, "aborted": bool,
               "result": Any, "error": Optional[str]}.
    """

    def __init__(self, parent, title: str, items: list, work_fn: Callable,
                 on_done: Optional[Callable[[dict], None]] = None):
        super().__init__(parent)
        self.title(title)
        self.geometry("580x480")
        self.minsize(520, 440)
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_abort_click)

        self._items        = items
        self._work_fn       = work_fn
        self._on_done       = on_done
        self._total         = len(items)
        self._abort_event   = threading.Event()
        self._q: queue.Queue = queue.Queue()
        self._start_time    = time.monotonic()
        self._done           = False
        self._success_count  = 0
        self._failed_count   = 0

        self._build(title)
        threading.Thread(target=self._run_worker, daemon=True).start()
        self.after(100, self._poll)
        self.bind("<Escape>", lambda e: self._on_abort_click())

    def _build(self, title: str):
        main = ttk.Frame(self, padding=18)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text=title, font=(FONT_FAMILY, 13, "bold")).pack(anchor="w")
        self.status_lbl = ttk.Label(main, text="Starting…", font=(FONT_FAMILY, 10), bootstyle="secondary")
        self.status_lbl.pack(anchor="w", pady=(4, 12))

        bar_row = ttk.Frame(main)
        bar_row.pack(fill="x")
        self.progress = ttk.Progressbar(bar_row, mode="determinate", maximum=max(1, self._total),
                                         bootstyle="success-striped")
        self.progress.pack(side="left", fill="x", expand=True)
        self.pct_lbl = ttk.Label(bar_row, text="0%", width=6, anchor="e", font=(FONT_FAMILY, 10, "bold"))
        self.pct_lbl.pack(side="left", padx=(10, 0))

        stats = ttk.Frame(main)
        stats.pack(fill="x", pady=(10, 4))
        self.processed_lbl = ttk.Label(stats, text=f"Processed: 0 / {self._total}", font=(FONT_FAMILY, 9))
        self.processed_lbl.pack(side="left")
        self.eta_lbl = ttk.Label(stats, text="Estimating…", font=(FONT_FAMILY, 9), bootstyle="secondary")
        self.eta_lbl.pack(side="right")

        counts = ttk.Frame(main)
        counts.pack(fill="x", pady=(0, 12))
        self.success_lbl = ttk.Label(counts, text="✔  Success: 0", font=(FONT_FAMILY, 9, "bold"), bootstyle="success")
        self.success_lbl.pack(side="left")
        self.failed_lbl = ttk.Label(counts, text="✖  Failed: 0", font=(FONT_FAMILY, 9, "bold"), bootstyle="danger")
        self.failed_lbl.pack(side="left", padx=(18, 0))

        ttk.Label(main, text="Details:", font=(FONT_FAMILY, 9, "bold")).pack(anchor="w")
        log_frame = ttk.Frame(main, relief="solid", borderwidth=1)
        log_frame.pack(fill="both", expand=True, pady=(2, 14))
        self.log_text = tk.Text(log_frame, height=8, wrap="word", state="disabled",
                                 relief="flat", font=("Consolas", 9), padx=6, pady=4)
        apply_theme_to_text_widget(self.log_text)
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

        btn_row = ttk.Frame(main)
        btn_row.pack(fill="x")
        self.abort_btn = ttk.Button(btn_row, text="✕  Cancel", bootstyle="danger-outline",
                                     command=self._on_abort_click)
        self.abort_btn.pack(side="right")
        self.close_btn = ttk.Button(btn_row, text="Close", bootstyle="secondary", command=self.destroy)
        # close_btn is only shown once the operation finishes -- see _on_finished()

    def _log(self, line: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _on_abort_click(self):
        if self._done:
            self.destroy()
            return
        self._abort_event.set()
        self.abort_btn.config(state="disabled", text="Cancelling…")
        self.status_lbl.config(text="Cancelling -- finishing the current item…")

    # ------------------------------------------------------------------ #
    #  Worker thread side -- touches only self._q, never a Tk widget
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
    #  Main-thread side -- the only code here allowed to touch widgets
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
        if not self._done:
            self.after(100, self._poll)

    def _on_progress(self, msg: dict):
        current, total = msg["current"], msg["total"]
        self.progress.configure(maximum=max(1, total))
        self.progress["value"] = current
        pct = int(current / total * 100) if total else 0
        self.pct_lbl.config(text=f"{pct}%")
        self.processed_lbl.config(text=f"Processed: {current} / {total}")
        self.status_lbl.config(text=msg["status_text"])

        if msg.get("success_delta"):
            self._success_count += msg["success_delta"]
            self.success_lbl.config(text=f"✔  Success: {self._success_count}")
        if msg.get("failed_delta"):
            self._failed_count += msg["failed_delta"]
            self.failed_lbl.config(text=f"✖  Failed: {self._failed_count}")
        if msg.get("log_line"):
            self._log(msg["log_line"])

        elapsed = time.monotonic() - self._start_time
        if current > 0 and elapsed > 0.4:
            rate = current / elapsed
            remaining = max(0, total - current)
            eta_seconds = (remaining / rate) if rate > 0 else 0
            self.eta_lbl.config(text=f"~{self._format_duration(eta_seconds)} remaining")
        else:
            self.eta_lbl.config(text="Estimating…")

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
        self._done = True
        aborted = result.get("aborted", False)
        success = result.get("success", self._success_count)
        failed  = result.get("failed", self._failed_count)
        error   = result.get("error")

        if error:
            self.status_lbl.config(text=f"Stopped due to an error: {error}")
            self._log(f"ERROR: {error}")
        elif aborted:
            self.status_lbl.config(text=f"Cancelled -- {success} completed, {failed} failed before stopping.")
        else:
            self.status_lbl.config(text=f"Done -- {success} succeeded, {failed} failed.")
            self.progress["value"] = self.progress["maximum"]
            self.pct_lbl.config(text="100%")
        self.eta_lbl.config(text="")

        self.abort_btn.pack_forget()
        self.close_btn.pack(side="right")

        if self._on_done:
            try:
                self._on_done(result)
            except Exception as exc:
                logger.error(f"Bulk operation on_done callback failed: {exc}", exc_info=True)

# =============================================================================
#  DASHBOARD  /  INVOICE LIST (IMPROVED)
# =============================================================================
class DashboardFrame(ttk.Frame):
    STATUS_FILTERS = ["All", "Paid", "Due", "Partial", "Overpaid", "Complimentary"]
    SORT_ASC  = " ▲"
    SORT_DESC = " ▼"
    COL_DEFS  = [
        ("id",      "Invoice #",        110, "center"),
        ("client",  "Client / Company", 250, "w"),
        ("date",    "Date",             100, "center"),
        ("status",  "Status",            92, "center"),
        ("total",   "Total",            118, "e"),
        ("paid",    "Paid",             108, "e"),
        ("balance", "Balance",          108, "e"),
        ("note",    "🗒",                40, "center"),
    ]
    PAGE_SIZE_OPTIONS = [25, 50, 100]
    DEFAULT_PAGE_SIZE = 50

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller     = controller
        self.db_manager     = controller.db_manager
        self._status_filter = "All"
        self._sort_col      = "date"
        self._sort_rev      = True
        self._filter_btns: Dict[str, ttk.Button] = {}
        self._row_index: Dict[str, Dict] = {}
        self._page          = 1
        self._page_size     = self.DEFAULT_PAGE_SIZE
        self._total_pages   = 1
        self._bulk_op_running = False  # guards against overlapping bulk operations -- see _start_bulk_op()
        self._build()
        self._debounce_search = Debouncer(self, 200, self._populate)

    def _build(self):
        outer = ttk.Frame(self, padding=(24, 16, 24, 14))
        outer.pack(fill="both", expand=True)
        self._header(outer)
        self._stat_cards(outer)
        self._filter_bar(outer)
        self._table(outer)
        self._pagination_bar(outer)
        self._action_bar(outer)

    def _header(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(0, 18))
        
        left = ttk.Frame(row)
        left.pack(side="left")
        ttk.Label(left, text="Dashboard", font=(FONT_FAMILY, 24, "bold")).pack(anchor="w")
        ttk.Label(left, text="Every invoice, at a glance.", font=(FONT_FAMILY, 10), bootstyle="secondary").pack(anchor="w")

        right = ttk.Frame(row)
        right.pack(side="right")
        refresh_btn = ttk.Button(right, text="🔄  Refresh  (Ctrl+R)", bootstyle="secondary-outline", command=self._act_refresh)
        refresh_btn.pack(side="left", padx=4)
        ToolTip(refresh_btn, "Reload the invoice list and stats (Ctrl+R)")
        self.export_csv_btn = ttk.Button(right, text="📊  Export CSV", bootstyle="info-outline", command=self._export_csv)
        self.export_csv_btn.pack(side="left", padx=4)
        new_inv_btn = ttk.Button(right, text="＋  New Invoice  (Ctrl+N)", bootstyle="primary",
                                  command=lambda: self.controller.show_frame(InvoiceFormFrame))
        new_inv_btn.pack(side="left", padx=(4, 0))
        ToolTip(new_inv_btn, "Start a new invoice (Ctrl+N)")

    def _stat_cards(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(0, 18))
        for i in range(4):
            row.columnconfigure(i, weight=1)
            
        self.s_billed   = self._card(row, "💰  Total Billed",   fmt_money(0), "primary",   0)
        self.s_received = self._card(row, "✅  Total Received", fmt_money(0), "success",   1)
        self.s_due      = self._card(row, "⚠️  Balance Due",    fmt_money(0), "danger",    2)
        self.s_count    = self._card(row, "📄  Total Invoices", "0",          "secondary", 3)

    def _card(self, parent, title: str, value: str, style: str, col: int) -> ttk.Label:
        card = ttk.Frame(parent, padding=(18, 14), bootstyle=style)
        card.grid(row=0, column=col, sticky="ew", padx=6)
        ttk.Label(card, text=title, font=(FONT_FAMILY, 10), bootstyle=f"inverse-{style}").pack(anchor="w")
        lbl = ttk.Label(card, text=value, font=(FONT_FAMILY, 20, "bold"), bootstyle=f"inverse-{style}")
        lbl.pack(anchor="w", pady=(4, 0))
        return lbl

    def _filter_bar(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(0, 8))

        filter_group = ttk.Frame(row)
        filter_group.pack(side="left")
        ttk.Label(filter_group, text="Filter:", font=(FONT_FAMILY, 9), bootstyle="secondary").pack(side="left", padx=(0, 8))
        
        for status in self.STATUS_FILTERS:
            label = "All" if status == "All" else STATUS_DISPLAY_LABELS.get(status, status).title()
            btn = ttk.Button(
                filter_group, text=label, width=13,
                bootstyle="primary" if status == "All" else "secondary-outline",
                command=lambda s=status: self._set_filter(s),
            )
            btn.pack(side="left", padx=2)
            self._filter_btns[status] = btn

        srch = ttk.Frame(row)
        srch.pack(side="right")
        ttk.Label(srch, text="🔍", font=(FONT_FAMILY, 12)).pack(side="left", padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._on_search_changed())
        
        self.search_entry = ttk.Entry(srch, textvariable=self.search_var, width=28)
        self.search_entry.pack(side="left")
        ToolTip(self.search_entry, "Search by invoice number, client name, or note (Ctrl+F)")
        # Ctrl+A means "Select All invoices" on this screen (see select_all()
        # below), but inside the search box specifically that would be a
        # surprising side effect of just moving your cursor while typing --
        # this local, higher-priority binding keeps the native "jump to
        # start of text" behaviour here instead, without needing to touch
        # every Entry/Combobox app-wide the way the app's Ctrl+D handling
        # does for its own (more disruptive) native-binding collision.
        self.search_entry.bind("<Control-a>", self._search_entry_ctrl_a)
        ttk.Button(srch, text="✕", width=3, bootstyle="secondary-outline", command=lambda: self.search_var.set("")).pack(side="left", padx=(4, 0))

    @staticmethod
    def _search_entry_ctrl_a(event):
        event.widget.icursor(0)
        return "break"

    def _on_search_changed(self):
        self._page = 1  # a new search query invalidates the old page position
        self._debounce_search()

    def _set_filter(self, status: str):
        self._status_filter = status
        self._page = 1  # a new filter invalidates the old page position
        for s, btn in self._filter_btns.items():
            btn.configure(bootstyle="primary" if s == status else "secondary-outline")
        self._populate()

    def _table(self, parent):
        wrap = ttk.Frame(parent)
        wrap.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(wrap, columns=[c[0] for c in self.COL_DEFS],
                                 show="headings", bootstyle="primary",
                                 selectmode="extended", height=16)
                                 
        for cid, text, width, anchor in self.COL_DEFS:
            # Set up column headers with sorting commands
            self.tree.heading(cid, text=text, anchor=anchor, command=lambda c=cid: self._sort_by(c))
            self.tree.column(cid, width=width, anchor=anchor, minwidth=40)

        for tag, colors_map in STATUS_BADGE.items():
            self.tree.tag_configure(tag, foreground=colors_map["foreground"], background=colors_map["background"])

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Double-1>",          lambda _: self._act_edit())
        self.tree.bind("<<TreeviewSelect>>",  self._on_select)
        self.tree.bind("<Button-3>",          self._show_ctx_menu)
        self.tree.bind("<Delete>",            lambda _: self._act_delete())

        # Context Menus
        self._ctx_single = tk.Menu(self, tearoff=0)
        self._ctx_single.add_command(label="📄  View PDF",   command=self._act_view_pdf)
        self._ctx_single.add_command(label="✏️  Edit",       command=self._act_edit)
        self._ctx_single.add_command(label="📋  Duplicate",  command=self._act_duplicate)
        self._ctx_single.add_command(label="🗒  Edit Note",   command=self._act_edit_note)
        self._ctx_single.add_separator()
        self._ctx_single.add_command(label="🗑️  Delete",     command=self._act_delete)

        self._ctx_bulk = tk.Menu(self, tearoff=0)
        self._ctx_bulk.add_command(label="🖨️  Bulk Print",   command=self._act_bulk_print)
        self._ctx_bulk.add_command(label="📊  Bulk Export",  command=self._act_bulk_export)
        self._ctx_bulk.add_separator()
        self._ctx_bulk.add_command(label="🗑️  Bulk Delete",  command=self._act_delete)

    def _sort_by(self, col: str):
        """Sorts the treeview by clicking a column header and handles UI arrows."""
        if self._sort_col == col:
            self._sort_rev = not self._sort_rev
        else:
            self._sort_col = col
            self._sort_rev = False # Default to ascending when clicking a new column
        self._page = 1  # a new sort order invalidates the old page position

        # Update column headers to show the correct arrow and clear others
        for cid, text, _, _ in self.COL_DEFS:
            if cid == col:
                arrow = self.SORT_DESC if self._sort_rev else self.SORT_ASC
                self.tree.heading(cid, text=f"{text}{arrow}")
            else:
                self.tree.heading(cid, text=text) 
                
        self._populate()

    def _pagination_bar(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(8, 0))

        self.page_info_lbl = ttk.Label(row, text="", font=(FONT_FAMILY, 9), bootstyle="secondary")
        self.page_info_lbl.pack(side="left")

        nav = ttk.Frame(row)
        nav.pack(side="right")

        ttk.Label(nav, text="Rows per page:", font=(FONT_FAMILY, 9), bootstyle="secondary").pack(side="left", padx=(0, 6))
        self.page_size_var = tk.StringVar(value=str(self._page_size))
        page_size_combo = ttk.Combobox(nav, textvariable=self.page_size_var, state="readonly", width=5,
                                        values=[str(n) for n in self.PAGE_SIZE_OPTIONS])
        page_size_combo.pack(side="left", padx=(0, 16))
        page_size_combo.bind("<<ComboboxSelected>>", self._on_page_size_change)

        self.prev_page_btn = ttk.Button(nav, text="◀  Prev", width=8, bootstyle="secondary-outline", command=self._prev_page)
        self.prev_page_btn.pack(side="left", padx=2)
        self.next_page_btn = ttk.Button(nav, text="Next  ▶", width=8, bootstyle="secondary-outline", command=self._next_page)
        self.next_page_btn.pack(side="left", padx=2)

    def _prev_page(self):
        if self._page > 1:
            self._page -= 1
            self._populate()

    def _next_page(self):
        if self._page < self._total_pages:
            self._page += 1
            self._populate()

    def _on_page_size_change(self, _=None):
        try:
            new_size = int(self.page_size_var.get())
        except (TypeError, ValueError):
            new_size = self.DEFAULT_PAGE_SIZE
        if new_size != self._page_size:
            self._page_size = new_size
            self._page = 1
            self._populate()

    def _action_bar(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(10, 0))

        self.hint_lbl = ttk.Label(row, text="Select an invoice to take action", bootstyle="secondary", font=(FONT_FAMILY, 9))
        self.hint_lbl.pack(side="left")

        # -- Single-selection actions --
        self.del_btn  = ttk.Button(row, text="🗑  Delete",    bootstyle="danger-outline",   state="disabled", command=self._act_delete)
        self.note_btn = ttk.Button(row, text="🗒  Note",      bootstyle="secondary-outline", state="disabled", command=self._act_edit_note)
        self.dup_btn  = ttk.Button(row, text="📋  Duplicate",  bootstyle="secondary-outline", state="disabled", command=self._act_duplicate)
        self.edit_btn = ttk.Button(row, text="✏  Edit",      bootstyle="secondary",         state="disabled", command=self._act_edit)
        self.pdf_btn  = ttk.Button(row, text="📄  View PDF",  bootstyle="primary",           state="disabled", command=self._act_view_pdf)
        ToolTip(self.del_btn,  "Delete the selected invoice (Delete key)")
        ToolTip(self.dup_btn,  "Duplicate this invoice (Ctrl+D)")
        ToolTip(self.pdf_btn,  "Open this invoice's PDF, ready to print or export (Ctrl+P / Ctrl+E)")

        # -- Bulk-selection actions --
        self.bulk_del_btn    = ttk.Button(row, text="🗑  Bulk Delete", bootstyle="danger-outline",    command=self._act_bulk_delete)
        self.bulk_export_btn = ttk.Button(row, text="📊  Bulk Export", bootstyle="secondary-outline",  command=self._act_bulk_export)
        self.bulk_print_btn  = ttk.Button(row, text="🖨  Bulk Print",  bootstyle="primary",            command=self._act_bulk_print)
        ToolTip(self.bulk_del_btn,   "Delete the selected invoices (Delete key)")
        ToolTip(self.bulk_print_btn, "Merge & open the selected invoices' PDFs, ready to print or export (Ctrl+P / Ctrl+E)")

        # Every button that can kick off a bulk-scale operation -- disabled
        # together while one is running (see _start_bulk_op()/_end_bulk_op()),
        # so the user can never have two running at once and always has a
        # clear, consistent signal that one currently is.
        self._bulk_buttons = [self.bulk_del_btn, self.bulk_export_btn, self.bulk_print_btn, self.export_csv_btn]

        self._show_single_actions()

    def _show_single_actions(self, state="disabled"):
        """Helper to cleanly display single-selection buttons."""
        for b in [self.bulk_print_btn, self.bulk_export_btn, self.bulk_del_btn]:
            b.pack_forget()
        for b in [self.pdf_btn, self.edit_btn, self.dup_btn, self.note_btn, self.del_btn]:
            b.pack(side="right", padx=4)
            b.config(state=state)

    def _show_bulk_actions(self):
        """Helper to cleanly display bulk-selection buttons."""
        for b in [self.pdf_btn, self.edit_btn, self.dup_btn, self.note_btn, self.del_btn]:
            b.pack_forget()
        for b in [self.bulk_print_btn, self.bulk_export_btn, self.bulk_del_btn]:
            b.pack(side="right", padx=4)

    def on_show(self, _=None):
        self._populate()
        self._update_stats()

    def _on_select(self, _=None):
        n = len(self.tree.selection())
        
        if n >= 2:
            self._show_bulk_actions()
            self.hint_lbl.config(text=f"{n} invoices selected")
        else:
            state = "normal" if n == 1 else "disabled"
            self._show_single_actions(state=state)
            self.hint_lbl.config(
                text="Double-click to edit  •  Right-click for options  •  Ctrl/Shift-click to multi-select" if n
                     else "Select an invoice to take action"
            )

    def _selected_ids(self) -> List[str]:
        return [self.tree.item(iid, "values")[0] for iid in self.tree.selection()]

    def _selected_id(self) -> Optional[str]:
        ids = self._selected_ids()
        return ids[0] if ids else None

    def _populate(self):
        # Remember selection so it doesn't drop when typing in search or clicking sort
        current_selection = self.tree.selection()

        self.tree.delete(*self.tree.get_children())
        self._row_index.clear()

        # --- Fetch: search + filter + sort + pagination are all done in
        # SQL (see DatabaseManager.get_invoices_page) -- this call never
        # returns more than one page's worth of rows, no matter how many
        # invoices are in the table, which is what actually keeps the
        # Dashboard fast and light on memory at 100,000+ invoices. The
        # old approach (fetch every matching row, sort/paginate in
        # Python) is still fine for the small "export everything" case
        # elsewhere, just not for populating this list.
        search_query  = self.search_var.get().strip()
        rows, total = self.db_manager.get_invoices_page(
            search_query=search_query, status_filter=self._status_filter,
            sort_col=self._sort_col, sort_desc=self._sort_rev,
            page=self._page, page_size=self._page_size,
        )
        self._total_pages = max(1, -(-total // self._page_size))  # ceiling division
        if self._page > self._total_pages:
            # The page we were on no longer has any rows on it (e.g. the
            # last invoice on the last page just got deleted) -- clamp
            # back into range and re-fetch instead of showing a blank
            # page or a stale one.
            self._page = self._total_pages
            rows, total = self.db_manager.get_invoices_page(
                search_query=search_query, status_filter=self._status_filter,
                sort_col=self._sort_col, sort_desc=self._sort_rev,
                page=self._page, page_size=self._page_size,
            )

        start = (self._page - 1) * self._page_size
        for inv in rows:
            row = dict(inv)
            is_comp    = bool(row.get("is_complimentary"))
            status_val = (row["status"] or "Due").upper()
            # Display wording only -- status_val (used below for the
            # row's colour tag) is untouched, so this is purely cosmetic.
            status_display = STATUS_DISPLAY_LABELS.get(row["status"] or "Due", status_val)
            date_val   = parse_date(row["invoice_date"], DB_DATE_FORMAT, DISPLAY_DATE_FORMAT)
            total_txt  = "No Charge" if is_comp else fmt_money(row.get("grand_total"))
            paid_txt   = "-"         if is_comp else fmt_money(row.get("total_paid"))
            bal_txt    = "-"         if is_comp else fmt_money(row.get("balance_due"))
            note_icon  = "📝" if (row.get("notes") or "").strip() else ""

            values = (row["invoice_number"], row["bill_to"], date_val, status_display,
                      total_txt, paid_txt, bal_txt, note_icon)
            badge = status_val if status_val in STATUS_BADGE else ""
            iid = row["invoice_number"]
            self.tree.insert("", "end", iid=iid, values=values, tags=(badge,) if badge else ())
            self._row_index[iid] = row

        # Restore selection (only rows on the current page can still exist)
        for item in current_selection:
            if self.tree.exists(item):
                self.tree.selection_add(item)

        self._update_pagination_controls(total, start, len(rows))
        self._on_select() # Re-evaluate button states

    def _update_pagination_controls(self, total: int, start: int, shown: int):
        if total == 0:
            self.page_info_lbl.config(text="No invoices found")
        else:
            self.page_info_lbl.config(
                text=f"Showing {start + 1}–{start + shown} of {total}  •  "
                     f"Page {self._page} of {self._total_pages}"
            )
        self.prev_page_btn.config(state="normal" if self._page > 1 else "disabled")
        self.next_page_btn.config(state="normal" if self._page < self._total_pages else "disabled")

    def _update_stats(self):
        s = self.db_manager.get_dashboard_stats()
        self.s_billed.config(text=fmt_money(s['total_billed']))
        self.s_received.config(text=fmt_money(s['total_paid']))
        self.s_due.config(text=fmt_money(s['balance_due']))
        self.s_count.config(text=str(s["total"]))

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
        NoteDialog(self.controller, self.db_manager, inv_id, current, self._populate)

    def _act_delete(self):
        """Single-invoice delete (the small Delete button / Delete key on
        the action bar). If the Delete key fires while multiple rows
        happen to be selected, this redirects to the bulk-safe path --
        deleting a handful of invoices synchronously is instant and
        doesn't need a progress dialog, but a large multi-select
        deletion benefits from the same progress/abort treatment as any
        other bulk operation (see _act_bulk_delete)."""
        ids = self._selected_ids()
        if not ids:
            return
        if len(ids) > 1:
            self._act_bulk_delete()
            return
        if Messagebox.okcancel("Delete Invoice", f"Permanently delete Invoice #{ids[0]}?\nThis cannot be undone."):
            self.db_manager.delete_invoices(ids)
            self.on_show()

    def _act_duplicate(self):
        if not (inv_id := self._selected_id()): return
        row = self.db_manager.get_invoice_by_id(inv_id)
        if not row: return
        
        new_data = {k: row[k] for k in row.keys() if k not in
                    ("grand_total", "total_paid", "balance_due", "created_ts")}
        new_data["invoice_number"] = self.db_manager.get_next_invoice_number()
        new_data["invoice_date"]   = datetime.date.today().strftime(DB_DATE_FORMAT)
        new_data["status"]         = InvoiceStatus.DUE.value
        new_data["amount_paid"]    = 0.0
        new_data["advance_amount"] = 0.0
        new_data["payment_date"]   = None
        new_data["advance_date"]   = None
        
        ok, err = self.db_manager.save_invoice(new_data, is_update=False, payments=[])
        if ok:
            self.on_show()
            Messagebox.show_info("Duplicated", f"Created {new_data['invoice_number']} as a copy.")
        else:
            Messagebox.show_error("Error", f"Could not duplicate:\n{err}")

    # ------------------------------------------------------------------ #
    #  Bulk operation concurrency guard
    # ------------------------------------------------------------------ #
    # Every button that kicks off a bulk-scale operation lives in this
    # list, populated as each one is built (see _action_bar()/_header()).
    # _start_bulk_op()/_end_bulk_op() disable/re-enable all of them
    # together, so only one bulk operation can ever be in flight and the
    # user always has a clear, consistent signal that one is running.
    def _set_bulk_buttons_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for btn in getattr(self, "_bulk_buttons", ()):
            btn.config(state=state)

    def _start_bulk_op(self):
        self._bulk_op_running = True
        self._set_bulk_buttons_enabled(False)

    def _end_bulk_op(self):
        self._bulk_op_running = False
        self._set_bulk_buttons_enabled(True)

    # ------------------------------------------------------------------ #
    #  Bulk actions
    # ------------------------------------------------------------------ #
    def _act_bulk_delete(self):
        if self._bulk_op_running:
            return
        ids = self._selected_ids()
        if not ids:
            return
        if not Messagebox.okcancel("Bulk Delete", f"Permanently delete {len(ids)} invoices?\nThis cannot be undone."):
            return
        self._start_bulk_op()
        BulkProgressDialog(self.controller, f"Deleting {len(ids)} Invoice(s)", ids,
                            work_fn=self._bulk_delete_work, on_done=self._on_bulk_delete_done)

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
        self.on_show()      # refresh the list/stats first...
        self._end_bulk_op() # ...then restore button state, so a stray _on_select() in between can't leave it stale

    def _bulk_export_work_factory(self, file_path: str, ids: Optional[List[str]]) -> Callable:
        """Builds a work_fn that streams matching invoices straight into
        a CSV file in chunks, writing incrementally instead of loading
        the whole result set into memory first -- the difference between
        "fine at a few hundred invoices" and "fine at 100,000+". Handles
        both an explicit selection (ids given) and "export everything"
        (ids=None, fed from DatabaseManager.iter_all_invoices_chunked)
        with the same progress-reporting shape either way."""
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
            Messagebox.show_error("Export Failed", result["error"])
        elif result.get("result"):
            Messagebox.show_info("Exported", f"Saved {result['success']} invoice(s) to:\n{result['result']}")

    def _act_bulk_export(self):
        if self._bulk_op_running:
            return
        ids = self._selected_ids()
        if not ids:
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")],
                                             title="Export Selected Invoices")
        if not path:
            return
        self._start_bulk_op()
        BulkProgressDialog(self.controller, f"Exporting {len(ids)} Invoice(s)", ids,
                            work_fn=self._bulk_export_work_factory(path, ids),
                            on_done=self._on_bulk_export_done)

    def _bulk_print_work(self, ids: List[str], report, is_aborted) -> dict:
        """Generates a PDF per invoice and merges them into batches of at
        most PRINT_BATCH_SIZE each -- printing 100 invoices produces one
        manageable merged PDF; printing 1,000 produces ten of them
        instead of a single unwieldy multi-thousand-page file. Each
        completed batch is flushed to disk as soon as it fills, so
        cancelling partway through still leaves whatever batches finished
        usable rather than discarding all progress."""
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

    def _on_bulk_print_done(self, result: dict):
        self._end_bulk_op()
        paths = result.get("result") or []
        if len(paths) == 1:
            PDFViewer(self.controller, paths[0])
        elif len(paths) > 1:
            Messagebox.show_info(
                "Batches Ready",
                f"Generated {len(paths)} batch PDF file(s) (up to {PRINT_BATCH_SIZE} invoices each) in:\n{BATCH_DIR}",
            )

    def _act_bulk_print(self):
        if self._bulk_op_running:
            return
        ids = self._selected_ids()
        if not ids:
            return
        self._start_bulk_op()
        BulkProgressDialog(self.controller, f"Generating {len(ids)} PDF(s)", ids,
                            work_fn=self._bulk_print_work, on_done=self._on_bulk_print_done)

    def _show_ctx_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row and row not in self.tree.selection():
            self.tree.selection_set(row)
        if len(self.tree.selection()) >= 2:
            self._ctx_bulk.post(event.x_root, event.y_root)
        elif row:
            self._ctx_single.post(event.x_root, event.y_root)

    def _export_csv(self):
        """Header 'Export CSV' button -- exports every invoice, not just
        a selection. Routed through the same streaming/progress-dialog
        path as the bulk actions below (and gated by the same mutex)
        since at 100,000+ invoices this is exactly as bulk-scale an
        operation as an explicit multi-select export."""
        if self._bulk_op_running:
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")],
                                             title="Export All Invoices")
        if not path:
            return
        total_count = self.db_manager.get_dashboard_stats().get("total", 0)
        if not total_count:
            Messagebox.show_error("Nothing to Export", "There are no invoices to export yet.")
            return
        self._start_bulk_op()
        BulkProgressDialog(self.controller, f"Exporting All {total_count} Invoices", list(range(total_count)),
                            work_fn=self._bulk_export_work_factory(path, None),
                            on_done=self._on_bulk_export_done)

    def _act_refresh(self):
        """Reloads the invoice list and summary stats from the database --
        used by both the Refresh button and Ctrl+R. This is the same
        query _populate()/_update_stats() already run on every
        navigation; refresh just triggers it on demand (e.g. after
        editing data outside the app, or just to be sure the numbers
        are current) without changing which page or filter is active."""
        self._populate()
        self._update_stats()

    # ------------------------------------------------------------------ #
    #  Public API used by application-level keyboard shortcuts
    # ------------------------------------------------------------------ #
    def focus_search(self):
        """Ctrl+F: jump into this screen's search box, with any existing
        text selected so typing immediately replaces it."""
        self.search_entry.focus_set()
        self.search_entry.selection_range(0, "end")

    def refresh(self):
        """Ctrl+R: reload the invoice list and stats, exactly like the
        Refresh button."""
        self._act_refresh()

    def select_all(self):
        """Ctrl+A: select every invoice on the currently-shown page of
        the list. Selection lives in the Treeview widget itself, so
        this can only select rows actually on screen -- with
        pagination, that means the current page; switch pages first to
        select rows elsewhere."""
        self.tree.selection_set(self.tree.get_children())

    def prepare_selected_pdf(self):
        """Ctrl+P and Ctrl+E both start here: PDF generation is
        asynchronous, so both shortcuts generate (or re-use) the PDF and
        open it in the viewer -- the dedicated Print/Save buttons there
        (and their own Ctrl+P/Ctrl+E shortcuts) handle the actual
        print-or-save step. Mirrors the existing single-vs-bulk split
        already used by the right-click context menu."""
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
            row = dict(self.db_manager.get_invoice_by_id(invoice_id))
            payments = self.db_manager.get_payments(invoice_id)
            path = PDFGenerator(row, load_config(), payments=payments).generate()
            q.put({"status": "success", "kind": "single", "path": path})
        except Exception as exc:
            logger.error(f"PDF generation failed ({invoice_id}): {exc}", exc_info=True)
            q.put({"status": "error", "message": str(exc)})

# =============================================================================
#  CLIENTS  (Address Book manager)
# =============================================================================
class ClientsFrame(ttk.Frame):
    COL_DEFS = [
        ("name",    "Client / Company", 260, "w"),
        ("contact", "Contact Info",     320, "w"),
        ("gstin",   "GSTIN",            160, "w"),
    ]

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.db_manager = controller.db_manager
        self._build()
        self._debounce_search = Debouncer(self, 200, self._populate)

    def _build(self):
        outer = ttk.Frame(self, padding=(24, 16, 24, 14))
        outer.pack(fill="both", expand=True)

        row = ttk.Frame(outer); row.pack(fill="x", pady=(0, 18))
        left = ttk.Frame(row); left.pack(side="left")
        ttk.Label(left, text="Clients", font=(FONT_FAMILY, 24, "bold")).pack(anchor="w")
        ttk.Label(left, text="Your address book for fast invoice billing.", font=(FONT_FAMILY, 10), bootstyle="secondary").pack(anchor="w")
        add_client_btn = ttk.Button(row, text="＋  Add Client  (Ctrl+N)", bootstyle="primary",
                                     command=lambda: self._open_editor(None))
        add_client_btn.pack(side="right")
        ToolTip(add_client_btn, "Add a new client (Ctrl+N)")

        srch = ttk.Frame(outer); srch.pack(fill="x", pady=(0, 8))
        ttk.Label(srch, text="🔍", font=(FONT_FAMILY, 12)).pack(side="left", padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._debounce_search())
        self.search_entry = ttk.Entry(srch, textvariable=self.search_var, width=32)
        self.search_entry.pack(side="left")
        ToolTip(self.search_entry, "Search clients (Ctrl+F)")

        wrap = ttk.Frame(outer); wrap.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(wrap, columns=[c[0] for c in self.COL_DEFS], show="headings",
                                 bootstyle="primary", selectmode="extended", height=18)
        for cid, text, width, anchor in self.COL_DEFS:
            self.tree.heading(cid, text=text, anchor=anchor)
            self.tree.column(cid, width=width, anchor=anchor, minwidth=60)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Double-1>",         lambda _: self._edit_selected())
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Delete>",           lambda _: self._delete_selected())

        self._ctx = tk.Menu(self, tearoff=0)
        self._ctx.add_command(label="✏️  Edit",   command=self._edit_selected)
        self._ctx.add_command(label="🗑️  Delete", command=self._delete_selected)
        self.tree.bind("<Button-3>", self._show_ctx_menu)

        act = ttk.Frame(outer); act.pack(fill="x", pady=(10, 0))
        self.hint_lbl = ttk.Label(act, text="Select a client to edit or delete", bootstyle="secondary", font=(FONT_FAMILY, 9))
        self.hint_lbl.pack(side="left")
        self.del_btn  = ttk.Button(act, text="🗑  Delete", bootstyle="danger-outline", state="disabled", command=self._delete_selected)
        self.del_btn.pack(side="right", padx=4)
        self.edit_btn = ttk.Button(act, text="✏  Edit", bootstyle="secondary", state="disabled", command=self._edit_selected)
        self.edit_btn.pack(side="right", padx=4)

    def on_show(self, _=None):
        self._populate()

    def _populate(self):
        self.tree.delete(*self.tree.get_children())
        for c in self.db_manager.get_all_clients(self.search_var.get().strip()):
            contact_preview = (c.get("contact") or "").replace("\n", "  ·  ")
            self.tree.insert("", "end", iid=str(c["id"]),
                              values=(c["name"], contact_preview, c.get("gstin") or ""))

    def _on_select(self, _=None):
        n = len(self.tree.selection())
        self.edit_btn.config(state="normal" if n == 1 else "disabled")
        self.del_btn.config(state="normal" if n >= 1 else "disabled")
        self.hint_lbl.config(text=f"{n} selected" if n else "Select a client to edit or delete")

    def _selected_client(self) -> Optional[Dict]:
        sel = self.tree.selection()
        if len(sel) != 1:
            return None
        return self.db_manager.get_client_by_id(int(sel[0]))

    def _open_editor(self, client: Optional[Dict]):
        ClientDialog(self.controller, self.db_manager, client, self._populate)

    def new_client(self):
        """Ctrl+N on this page: open the Add Client dialog, exactly like
        the "＋ Add Client" button."""
        self._open_editor(None)

    def _edit_selected(self):
        if client := self._selected_client():
            self._open_editor(client)

    def _delete_selected(self):
        ids = [int(i) for i in self.tree.selection()]
        if not ids:
            return
        msg = f"Delete {len(ids)} client(s) from your address book?" if len(ids) > 1 else \
              "Delete this client from your address book?"
        if Messagebox.okcancel("Delete Client(s)", msg):
            self.db_manager.delete_clients(ids)
            self._populate()

    def focus_search(self):
        """Ctrl+F: jump into the search box, with any existing text
        selected so typing immediately replaces it."""
        self.search_entry.focus_set()
        self.search_entry.selection_range(0, "end")

    def _show_ctx_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            if row not in self.tree.selection():
                self.tree.selection_set(row)
            self._ctx.post(event.x_root, event.y_root)

# =============================================================================
#  CATALOG  (Product & Service manager)
# =============================================================================
class CatalogFrame(ttk.Frame):
    COL_DEFS = [
        ("title",  "Item / Service", 260, "w"),
        ("desc",   "Description",    340, "w"),
        ("amount", "Default Amount", 140, "e"),
    ]

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.db_manager = controller.db_manager
        self._build()
        self._debounce_search = Debouncer(self, 200, self._populate)

    def _build(self):
        outer = ttk.Frame(self, padding=(24, 16, 24, 14))
        outer.pack(fill="both", expand=True)

        row = ttk.Frame(outer); row.pack(fill="x", pady=(0, 18))
        left = ttk.Frame(row); left.pack(side="left")
        ttk.Label(left, text="Catalog", font=(FONT_FAMILY, 24, "bold")).pack(anchor="w")
        ttk.Label(left, text="Reusable products & services for faster invoice lines.", font=(FONT_FAMILY, 10), bootstyle="secondary").pack(anchor="w")
        add_catalog_btn = ttk.Button(row, text="＋  Add Item  (Ctrl+N)", bootstyle="primary",
                                      command=lambda: self._open_editor(None))
        add_catalog_btn.pack(side="right")
        ToolTip(add_catalog_btn, "Add a new catalog item (Ctrl+N)")

        srch = ttk.Frame(outer); srch.pack(fill="x", pady=(0, 8))
        ttk.Label(srch, text="🔍", font=(FONT_FAMILY, 12)).pack(side="left", padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._debounce_search())
        self.search_entry = ttk.Entry(srch, textvariable=self.search_var, width=32)
        self.search_entry.pack(side="left")
        ToolTip(self.search_entry, "Search catalog items (Ctrl+F)")

        wrap = ttk.Frame(outer); wrap.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(wrap, columns=[c[0] for c in self.COL_DEFS], show="headings",
                                 bootstyle="primary", selectmode="extended", height=18)
        for cid, text, width, anchor in self.COL_DEFS:
            self.tree.heading(cid, text=text, anchor=anchor)
            self.tree.column(cid, width=width, anchor=anchor, minwidth=60)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Double-1>",         lambda _: self._edit_selected())
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Delete>",           lambda _: self._delete_selected())

        self._ctx = tk.Menu(self, tearoff=0)
        self._ctx.add_command(label="✏️  Edit",   command=self._edit_selected)
        self._ctx.add_command(label="🗑️  Delete", command=self._delete_selected)
        self.tree.bind("<Button-3>", self._show_ctx_menu)

        act = ttk.Frame(outer); act.pack(fill="x", pady=(10, 0))
        self.hint_lbl = ttk.Label(act, text="Select an item to edit or delete", bootstyle="secondary", font=(FONT_FAMILY, 9))
        self.hint_lbl.pack(side="left")
        self.del_btn  = ttk.Button(act, text="🗑  Delete", bootstyle="danger-outline", state="disabled", command=self._delete_selected)
        self.del_btn.pack(side="right", padx=4)
        self.edit_btn = ttk.Button(act, text="✏  Edit", bootstyle="secondary", state="disabled", command=self._edit_selected)
        self.edit_btn.pack(side="right", padx=4)

    def on_show(self, _=None):
        self._populate()

    def _populate(self):
        self.tree.delete(*self.tree.get_children())
        for it in self.db_manager.get_all_catalog_items(self.search_var.get().strip()):
            self.tree.insert("", "end", iid=str(it["id"]),
                              values=(it["title"], it.get("description") or "", fmt_money(it.get("amount"))))

    def _on_select(self, _=None):
        n = len(self.tree.selection())
        self.edit_btn.config(state="normal" if n == 1 else "disabled")
        self.del_btn.config(state="normal" if n >= 1 else "disabled")
        self.hint_lbl.config(text=f"{n} selected" if n else "Select an item to edit or delete")

    def _selected_item(self) -> Optional[Dict]:
        sel = self.tree.selection()
        if len(sel) != 1:
            return None
        return self.db_manager.get_catalog_item_by_id(int(sel[0]))

    def _open_editor(self, item: Optional[Dict]):
        CatalogItemDialog(self.controller, self.db_manager, item, self._populate)

    def new_catalog_item(self):
        """Ctrl+N on this page: open the Add Catalog Item dialog, exactly
        like the "＋ Add Item" button."""
        self._open_editor(None)

    def _edit_selected(self):
        if item := self._selected_item():
            self._open_editor(item)

    def _delete_selected(self):
        ids = [int(i) for i in self.tree.selection()]
        if not ids:
            return
        msg = f"Delete {len(ids)} catalog item(s)?" if len(ids) > 1 else "Delete this catalog item?"
        if Messagebox.okcancel("Delete Item(s)", msg):
            self.db_manager.delete_catalog_items(ids)
            self._populate()

    def focus_search(self):
        """Ctrl+F: jump into the search box, with any existing text
        selected so typing immediately replaces it."""
        self.search_entry.focus_set()
        self.search_entry.selection_range(0, "end")

    def _show_ctx_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            if row not in self.tree.selection():
                self.tree.selection_set(row)
            self._ctx.post(event.x_root, event.y_root)

# =============================================================================
#  INVOICE FORM (MODERNIZED & REACTIVE)
# =============================================================================
class InvoiceFormFrame(ttk.Frame):
    """Create / edit an invoice: billing info, items, tax & discount, and a
    payment history ledger that supports multiple partial payments."""

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller  = controller
        self.db_manager  = controller.db_manager
        self.editing_id: Optional[str] = None
        self.item_rows:    List[Dict[str, Any]] = []
        self.payment_rows: List[Dict[str, Any]] = []
        self._catalog_cache: List[Dict] = []
        self._client_cache:  List[Dict] = []
        self._initial_snapshot: Optional[str] = None   
        self._build()

    # ------------------------------------------------------------------ #
    #  Layout
    # ------------------------------------------------------------------ #
    def _build(self):
        header = ttk.Frame(self, padding=(24, 16, 24, 6))
        header.pack(fill="x")
        top_back_btn = ttk.Button(header, text="←  Back to Dashboard  (Esc)", bootstyle="secondary-outline",
                                   command=lambda: self.controller.show_frame(DashboardFrame))
        top_back_btn.pack(side="left")
        ToolTip(top_back_btn, "Return to the invoice list (Esc). Unsaved changes will ask first.")
        self.title_lbl = ttk.Label(header, text="New Invoice", font=(FONT_FAMILY, 20, "bold"))
        self.title_lbl.pack(side="left", padx=20)
        self.status_badge = ttk.Label(header, text="", font=(FONT_FAMILY, 10, "bold"), padding=(10, 3))
        self.status_badge.pack(side="left")

        self.scroller = ScrolledFrame(self, autohide=True)
        self.scroller.pack(fill="both", expand=True, padx=24, pady=(0, 10))
        body = self.scroller
        self._bind_arrow_scrolling()

        self._section_billing(body)
        self._section_details(body)
        self._section_items(body)
        self._section_payments(body)
        self._section_notes(body)
        self._section_totals(body)
        self._section_save_bar(body)

    def _card(self, parent, title: str) -> ttk.Frame:
        wrap = ttk.Frame(parent, padding=16) 
        wrap.pack(fill="x", pady=(0, 16))
        ttk.Label(wrap, text=title, font=(FONT_FAMILY, 12, "bold")).pack(anchor="w", pady=(0, 10))
        return wrap

    def _prevent_scroll_hijack(self, widget):
        """Prevents Comboboxes from changing values, and forwards the scroll to the main page."""
        
        def on_mousewheel(event):
            # Windows and macOS scrolling
            if event.delta:
                # Calculate direction and scroll the parent ScrolledFrame directly
                self.scroller.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break" # Stop the combobox data from changing
            
        def on_linux_scroll_up(event):
            # Linux scroll up
            self.scroller.yview_scroll(-1, "units")
            return "break"
            
        def on_linux_scroll_down(event):
            # Linux scroll down
            self.scroller.yview_scroll(1, "units")
            return "break"

        # Apply the bindings to the widget
        widget.bind("<MouseWheel>", on_mousewheel)
        widget.bind("<Button-4>", on_linux_scroll_up)
        widget.bind("<Button-5>", on_linux_scroll_down)

    def _bind_arrow_scrolling(self):
        """↑/↓ scroll by one "unit" (matching the mouse wheel), Page Up/
        Down by one "page". Bound at the app level since Tk has no
        "only while visible" binding scope -- the handler itself checks
        the invoice form is the active page before acting. Leaves combo
        boxes, Text fields, and Treeviews alone since they already use
        ↑/↓ for their own purpose (cycling values / cursor movement /
        row selection).
        """
        def on_scroll_key(event, direction: int, what: str):
            if self.controller._active_frame() is not self:
                return  # a different page is showing; don't touch this keystroke
            try:
                if isinstance(event.widget, (ttk.Combobox, tk.Text, ttk.Treeview)):
                    return  # these widgets already use Up/Down for their own purpose
            except Exception:
                return  # can't safely resolve this event's widget; do nothing rather than risk a crash
            self.scroller.yview_scroll(direction, what)
            return "break"

        self.controller.bind_all("<Up>",    lambda e: on_scroll_key(e, -1, "units"))
        self.controller.bind_all("<Down>",  lambda e: on_scroll_key(e,  1, "units"))
        self.controller.bind_all("<Prior>", lambda e: on_scroll_key(e, -1, "pages"))  # Page Up
        self.controller.bind_all("<Next>",  lambda e: on_scroll_key(e,  1, "pages"))  # Page Down

    def _advance_focus_on_enter(self, widget):
        """Binds <Return> so pressing Enter moves focus to the next widget
        in tab order -- the same field-to-field flow Tab already gives,
        just from the keyboard's most natural "confirm" key. Applied to
        single-line Entry/editable-Combobox fields only: multi-line Text
        widgets (Contact Info, Notes) deliberately keep plain Enter so it
        can insert a newline instead."""
        widget.bind("<Return>", self._focus_next_widget)

    @staticmethod
    def _focus_next_widget(event):
        event.widget.tk_focusNext().focus_set()
        return "break"

    # ---- Billing (Bill To / From) ------------------------------------ #
    def _section_billing(self, parent):
        wrap = ttk.Frame(parent)
        wrap.pack(fill="x", pady=(0, 0))
        wrap.columnconfigure(0, weight=1)
        wrap.columnconfigure(1, weight=1)

        left = ttk.Frame(wrap, padding=16)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ttk.Label(left, text="Bill To", font=(FONT_FAMILY, 12, "bold")).pack(anchor="w", pady=(0, 10))
        ttk.Label(left, text="Client / Company Name *", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.client_name_var = tk.StringVar()
        self.client_combo = PlaceholderCombobox(left, placeholder="Enter client name",
                                                 textvariable=self.client_name_var, height=6)
        self.client_combo.pack(fill="x", pady=(2, 8))
        self._prevent_scroll_hijack(self.client_combo)
        self.client_combo.bind("<KeyRelease>", self._filter_client_suggestions)
        self.client_combo.bind("<<ComboboxSelected>>", self._on_client_picked)
        self._advance_focus_on_enter(self.client_combo)
        
        ttk.Label(left, text="Contact Info", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.client_contact_text = PlaceholderText(left, placeholder="Phone, email, address...",
                                                     height=3, relief="flat", highlightthickness=1, wrap="word")
        self.client_contact_text.pack(fill="x", pady=(2, 8))
        ttk.Label(left, text="GSTIN", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.client_gstin_entry = PlaceholderEntry(left, placeholder="22ABCDE1234F1Z5 (optional)")
        self.client_gstin_entry.pack(fill="x", pady=(2, 0))
        self._advance_focus_on_enter(self.client_gstin_entry)

        right = ttk.Frame(wrap, padding=16)
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ttk.Label(right, text="From", font=(FONT_FAMILY, 12, "bold")).pack(anchor="w", pady=(0, 10))
        ttk.Label(right, text="Your Name / Firm *", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.your_name_entry = PlaceholderEntry(right, placeholder="Your business / firm name")
        self.your_name_entry.pack(fill="x", pady=(2, 8))
        self._advance_focus_on_enter(self.your_name_entry)
        ttk.Label(right, text="Contact Info", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.your_contact_text = PlaceholderText(right, placeholder="Phone, email, address...",
                                                   height=3, relief="flat", highlightthickness=1, wrap="word")
        self.your_contact_text.pack(fill="x", pady=(2, 0))

    # ---- Invoice Details / Tax & Payment Settings ---------------------- #
    def _section_details(self, parent):
        wrap = ttk.Frame(parent)
        wrap.pack(fill="x", pady=(16, 0))
        wrap.columnconfigure(0, weight=1)
        wrap.columnconfigure(1, weight=1)

        left = ttk.Frame(wrap, padding=16)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ttk.Label(left, text="Invoice Details", font=(FONT_FAMILY, 12, "bold")).pack(anchor="w", pady=(0, 10))

        r1 = ttk.Frame(left); r1.pack(fill="x", pady=4)
        ttk.Label(r1, text="Invoice Number *", width=20, anchor="w").pack(side="left")
        self.invoice_number_entry = PlaceholderEntry(r1, placeholder="INV-001")
        self.invoice_number_entry.pack(side="left", fill="x", expand=True)
        self._advance_focus_on_enter(self.invoice_number_entry)

        r2 = ttk.Frame(left); r2.pack(fill="x", pady=4)
        ttk.Label(r2, text="Invoice Date *", width=20, anchor="w").pack(side="left")
        self.invoice_date_entry = DateEntry(r2, dateformat=DISPLAY_DATE_FORMAT)
        self.invoice_date_entry.pack(side="left", fill="x", expand=True)
        self._advance_focus_on_enter(self.invoice_date_entry.entry)

        r3 = ttk.Frame(left); r3.pack(fill="x", pady=4)
        ttk.Label(r3, text="Payment Terms", width=20, anchor="w").pack(side="left")
        self.payment_terms_entry = PlaceholderEntry(r3, placeholder="e.g. Due on Receipt, Net 15")
        self.payment_terms_entry.pack(side="left", fill="x", expand=True)
        self._advance_focus_on_enter(self.payment_terms_entry)

        r4 = ttk.Frame(left); r4.pack(fill="x", pady=(10, 0))
        self.no_charge_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(r4, text="Mark as Complimentary (No Charge)", variable=self.no_charge_var,
                        bootstyle="round-toggle", command=self._toggle_no_charge).pack(side="left")

        r5 = ttk.Frame(left); r5.pack(fill="x", pady=(6, 0))
        self.show_qr_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(r5, text="Show UPI QR Code on PDF", variable=self.show_qr_var,
                        bootstyle="round-toggle").pack(side="left")

        right = ttk.Frame(wrap, padding=16)
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ttk.Label(right, text="Tax & Discount", font=(FONT_FAMILY, 12, "bold")).pack(anchor="w", pady=(0, 10))

        rg = ttk.Frame(right); rg.pack(fill="x", pady=4)
        self.gst_var = tk.BooleanVar(value=True)
        # BUG FIX: Ensure toggling GST immediately recalculates totals
        ttk.Checkbutton(rg, text="This is a GST Invoice", variable=self.gst_var,
                        bootstyle="round-toggle", command=self._on_gst_toggle).pack(side="left")

        rt = ttk.Frame(right); rt.pack(fill="x", pady=4)
        ttk.Label(rt, text="GST / Tax Rate (%)", width=20, anchor="w").pack(side="left")
        self.tax_rate_entry = ttk.Entry(rt)
        self.tax_rate_entry.insert(0, "18")
        self.tax_rate_entry.pack(side="left", fill="x", expand=True)
        self.tax_rate_entry.bind("<KeyRelease>", lambda _: self._recalc_totals())
        self._advance_focus_on_enter(self.tax_rate_entry)

        rgt = ttk.Frame(right); rgt.pack(fill="x", pady=4)
        ttk.Label(rgt, text="Tax Split Type", width=20, anchor="w").pack(side="left")
        self.gst_type_var = tk.StringVar(value="CGST/SGST")
        # BUG FIX: Live update when switching between CGST and IGST
        self.gst_type_var.trace_add("write", lambda *_: self._recalc_totals())
        self.gst_type_combo = ttk.Combobox(rgt, textvariable=self.gst_type_var, state="readonly",
                                           values=["CGST/SGST", "IGST"])
        self.gst_type_combo.pack(side="left", fill="x", expand=True)
        self._prevent_scroll_hijack(self.gst_type_combo)

        rd = ttk.Frame(right); rd.pack(fill="x", pady=4)
        ttk.Label(rd, text="Overall Discount", width=20, anchor="w").pack(side="left")
        disc_wrap = ttk.Frame(rd); disc_wrap.pack(side="left", fill="x", expand=True)
        self.global_discount_entry = ttk.Entry(disc_wrap)
        self.global_discount_entry.insert(0, "0")
        self.global_discount_entry.pack(side="left", fill="x", expand=True)
        self.global_discount_entry.bind("<KeyRelease>", lambda _: self._recalc_totals())
        self._advance_focus_on_enter(self.global_discount_entry)
        
        self.global_discount_type_var = tk.StringVar(value="amount")
        
        # 1. CREATE and SAVE the combobox to a variable first
        self.disc_type_combo = ttk.Combobox(disc_wrap, textvariable=self.global_discount_type_var, state="readonly", values=["amount", "%"], width=7)
        
        # 2. Pack it on the next line
        self.disc_type_combo.pack(side="left", padx=(6, 0))
        
        # 3. Now the scroll fix knows exactly what self.disc_type_combo is!
        self._prevent_scroll_hijack(self.disc_type_combo) 
        
        self.global_discount_type_var.trace_add("write", lambda *_: self._recalc_totals())

        self.gst_dependents = [self.tax_rate_entry, self.gst_type_combo]

    def _toggle_no_charge(self):
        self._apply_field_states()
        self._recalc_totals()
        
    def _on_gst_toggle(self):
        self._apply_field_states()
        self._recalc_totals()

    def _apply_field_states(self):
        no_charge = self.no_charge_var.get()
        is_gst    = self.gst_var.get() and not no_charge

        for w in self.gst_dependents:
            w.configure(state="normal" if is_gst else "disabled")
        self.gst_type_combo.configure(state="readonly" if is_gst else "disabled")

        pay_state = "disabled" if no_charge else "normal"
        self.add_payment_btn.configure(state=pay_state)
        for row in self.payment_rows:
            row["amount_entry"].configure(state=pay_state)
            row["mode_combo"].configure(state=("readonly" if not no_charge else "disabled"))
            row["note_entry"].configure(state=pay_state)
            row["remove_btn"].configure(state=pay_state)

    # ---- Items ---------------------------------------------------------- #
    def _section_items(self, parent):
        wrap = self._card(parent, "Items")
        header = ttk.Frame(wrap)
        header.pack(fill="x", pady=(0, 6))
        ttk.Label(header, text="ITEM / SERVICE", font=(FONT_FAMILY, 8, "bold"),
                  bootstyle="secondary").pack(side="left", fill="x", expand=True)
        ttk.Label(header, text="AMOUNT", font=(FONT_FAMILY, 8, "bold"), bootstyle="secondary",
                  width=12, anchor="center").pack(side="left", padx=6)
        ttk.Label(header, text="DISCOUNT", font=(FONT_FAMILY, 8, "bold"), bootstyle="secondary",
                  width=10, anchor="center").pack(side="left", padx=(0, 6))
        ttk.Label(header, text="", width=3).pack(side="left")
        ttk.Separator(wrap, bootstyle="secondary").pack(fill="x", pady=(0, 4))

        self.items_container = ttk.Frame(wrap)
        self.items_container.pack(fill="x")
        add_item_btn = ttk.Button(wrap, text="＋  Add Item  (Ctrl+Shift+A)", bootstyle="secondary-outline",
                                   command=lambda: self._add_item_row())
        add_item_btn.pack(anchor="w", pady=(10, 0))
        ToolTip(add_item_btn, "Add a blank item row (Ctrl+Shift+A or Ctrl+Plus)")

    def _add_item_row(self, title: str = "", description: str = "", amount: float = 0.0, discount: float = 0.0):
        row_frame = ttk.Frame(self.items_container, padding=(0, 7))
        row_frame.pack(fill="x")

        top = ttk.Frame(row_frame); top.pack(fill="x")
        title_var = tk.StringVar(value=title)
        title_combo = PlaceholderCombobox(top, placeholder="Item / service name", textvariable=title_var,
                                           values=[c["title"] for c in self._catalog_cache], height=6)
        title_combo.pack(side="left", fill="x", expand=True, ipady=2)
        self._prevent_scroll_hijack(title_combo)
        title_combo.bind("<<ComboboxSelected>>", lambda e, tc=title_combo: self._on_catalog_picked(tc))
        title_combo.bind("<KeyRelease>", lambda e, tc=title_combo: self._filter_catalog_suggestions(tc))
        self._advance_focus_on_enter(title_combo)

        amount_entry = ttk.Entry(top, width=12, justify="right")
        amount_entry.insert(0, f"{amount:.2f}")
        amount_entry.pack(side="left", padx=6, ipady=2)
        restrict_entry_to_decimal(amount_entry)
        format_entry_as_decimal_on_blur(amount_entry)
        amount_entry.bind("<KeyRelease>", lambda _: self._recalc_totals())
        self._advance_focus_on_enter(amount_entry)
        ToolTip(amount_entry, "Item amount, before any discount")

        disc_entry = ttk.Entry(top, width=10, justify="right")
        disc_entry.insert(0, f"{discount:.2f}")
        disc_entry.pack(side="left", padx=(0, 6), ipady=2)
        restrict_entry_to_decimal(disc_entry)
        format_entry_as_decimal_on_blur(disc_entry)
        disc_entry.bind("<KeyRelease>", lambda _: self._recalc_totals())
        self._advance_focus_on_enter(disc_entry)
        ToolTip(disc_entry, "Per-item discount amount")

        remove_btn = ttk.Button(top, text="✕", width=3, bootstyle="danger-outline", takefocus=False)
        remove_btn.pack(side="left")
        ToolTip(remove_btn, "Remove this item (Ctrl+Shift+Del or Ctrl+Minus)")

        desc_entry = PlaceholderEntry(row_frame, placeholder="Description (optional)")
        desc_entry.insert(0, description)
        desc_entry.pack(fill="x", pady=(5, 0), ipady=1)
        self._advance_focus_on_enter(desc_entry)

        ttk.Separator(row_frame, bootstyle="secondary").pack(fill="x", pady=(9, 0))

        entry = {
            "frame": row_frame, "title_var": title_var, "title_combo": title_combo,
            "amount_entry": amount_entry, "disc_entry": disc_entry, "desc_entry": desc_entry,
        }
        remove_btn.configure(command=lambda e=entry: self._remove_item_row(e))
        self.item_rows.append(entry)
        self._recalc_totals()
        return entry

    def _remove_item_row(self, entry: Dict):
        entry["frame"].destroy()
        self.item_rows.remove(entry)
        self._recalc_totals()

    def _on_catalog_picked(self, title_combo: "PlaceholderCombobox"):
        # Matched and read via the combo widget itself (not the linked
        # StringVar) so a still-visible placeholder can never be mistaken
        # for a chosen catalog title.
        chosen = title_combo.get()
        for c in self._catalog_cache:
            if c["title"] == chosen:
                for row in self.item_rows:
                    if row["title_combo"] is title_combo:
                        row["desc_entry"].delete(0, "end")
                        row["desc_entry"].insert(0, c.get("description") or "")
                        row["amount_entry"].delete(0, "end")
                        row["amount_entry"].insert(0, f"{float(c.get('amount') or 0):.2f}")
                break
        self._recalc_totals()

    def _filter_catalog_suggestions(self, combo: ttk.Combobox):
        text = combo.get().lower()
        combo["values"] = [c["title"] for c in self._catalog_cache if text in c["title"].lower()] if text else \
                           [c["title"] for c in self._catalog_cache]

    # ---- Payment History ------------------------------------------------ #
    def _section_payments(self, parent):
        wrap = self._card(parent, "Payment History")
        ttk.Label(wrap, text="Log each payment received. Status updates automatically.",
                  font=(FONT_FAMILY, 9), bootstyle="secondary").pack(anchor="w", pady=(0, 8))
        header = ttk.Frame(wrap)
        header.pack(fill="x", pady=(0, 4))
        ttk.Label(header, text="DATE", font=(FONT_FAMILY, 8, "bold"), bootstyle="secondary",
                  width=12, anchor="w").pack(side="left", padx=(0, 6))
        ttk.Label(header, text="AMOUNT", font=(FONT_FAMILY, 8, "bold"), bootstyle="secondary",
                  width=12, anchor="center").pack(side="left", padx=(0, 6))
        ttk.Label(header, text="MODE", font=(FONT_FAMILY, 8, "bold"), bootstyle="secondary",
                  width=13, anchor="w").pack(side="left", padx=(0, 6))
        ttk.Label(header, text="NOTE", font=(FONT_FAMILY, 8, "bold"),
                  bootstyle="secondary").pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Label(header, text="", width=3).pack(side="left")
        ttk.Separator(wrap, bootstyle="secondary").pack(fill="x", pady=(0, 2))
        self.payments_container = ttk.Frame(wrap)
        self.payments_container.pack(fill="x")
        self.add_payment_btn = ttk.Button(wrap, text="＋  Add Payment", bootstyle="secondary-outline",
                                          command=lambda: self._add_payment_row())
        self.add_payment_btn.pack(anchor="w", pady=(10, 0))

    def _add_payment_row(self, amount: float = 0.0, payment_date: Optional[str] = None,
                          mode: str = "Cash", note: str = "", payment_id: Optional[int] = None):
        row_frame = ttk.Frame(self.payments_container, padding=(0, 6))
        row_frame.pack(fill="x")

        date_entry = DateEntry(row_frame, dateformat=DISPLAY_DATE_FORMAT, width=12)
        if payment_date:
            try:
                date_entry.entry.delete(0, "end")
                date_entry.entry.insert(0, parse_date(payment_date, DB_DATE_FORMAT, DISPLAY_DATE_FORMAT))
            except Exception:
                pass
        date_entry.pack(side="left", padx=(0, 6), ipady=1)
        self._advance_focus_on_enter(date_entry.entry)

        amount_entry = ttk.Entry(row_frame, width=12, justify="right")
        amount_entry.insert(0, f"{amount:.2f}")
        amount_entry.pack(side="left", padx=(0, 6), ipady=2)
        restrict_entry_to_decimal(amount_entry)
        format_entry_as_decimal_on_blur(amount_entry)
        amount_entry.bind("<KeyRelease>", lambda _: self._recalc_totals())
        self._advance_focus_on_enter(amount_entry)

        mode_combo = ttk.Combobox(row_frame, values=PAYMENT_MODES, state="readonly", width=13)
        mode_combo.set(mode if mode in PAYMENT_MODES else "Cash")
        mode_combo.pack(side="left", padx=(0, 6), ipady=1)
        self._prevent_scroll_hijack(mode_combo)

        note_entry = PlaceholderEntry(row_frame, placeholder="Note (optional)")
        note_entry.insert(0, note)
        note_entry.pack(side="left", fill="x", expand=True, padx=(0, 6), ipady=1)
        self._advance_focus_on_enter(note_entry)

        remove_btn = ttk.Button(row_frame, text="✕", width=3, bootstyle="danger-outline", takefocus=False)
        remove_btn.pack(side="left")

        entry = {
            "frame": row_frame, "date_entry": date_entry, "amount_entry": amount_entry,
            "mode_combo": mode_combo, "note_entry": note_entry, "remove_btn": remove_btn,
            "payment_id": payment_id,
        }
        remove_btn.configure(command=lambda e=entry: self._remove_payment_row(e))
        self.payment_rows.append(entry)
        self._recalc_totals()
        return entry

    def _remove_payment_row(self, entry: Dict):
        entry["frame"].destroy()
        self.payment_rows.remove(entry)
        self._recalc_totals()

    # ---- Notes ------------------------------------------------------------ #
    def _section_notes(self, parent):
        wrap = self._card(parent, "Internal Notes")
        ttk.Label(wrap, text="Visible only inside Vyaprix — never printed on the invoice PDF.",
                  font=(FONT_FAMILY, 9), bootstyle="secondary").pack(anchor="w", pady=(0, 6))
        self.notes_text = PlaceholderText(wrap, placeholder="Internal notes (optional) — not shown on the PDF",
                                           height=3, relief="flat", highlightthickness=1, wrap="word")
        self.notes_text.pack(fill="x")

    def _section_totals(self, parent):
        wrap = ttk.Frame(parent, padding=16)
        wrap.pack(fill="x", pady=(0, 16))
        
        self.summary_box = ttk.Frame(wrap)
        self.summary_box.pack(side="right", anchor="e")
        
        # Initialize variables
        self.var_subtotal    = tk.StringVar(value="0.00")
        self.var_discount    = tk.StringVar(value="")
        self.var_taxable     = tk.StringVar(value="0.00")
        self.var_cgst_lbl    = tk.StringVar(value="CGST:")
        self.var_cgst_val    = tk.StringVar(value="")
        self.var_sgst_lbl    = tk.StringVar(value="SGST:")
        self.var_sgst_val    = tk.StringVar(value="")
        self.var_igst_lbl    = tk.StringVar(value="IGST:")
        self.var_igst_val    = tk.StringVar(value="")
        self.var_grand_total = tk.StringVar(value="0.00")
        self.var_paid        = tk.StringVar(value="")
        self.var_balance_lbl = tk.StringVar(value="Balance Due:")   # flips to "Advance / Credit:" when overpaid
        self.var_balance     = tk.StringVar(value="0.00")

        # Helper to create consistent rows with a fixed label. Returns
        # (row, value_label) -- callers that never need to recolour the
        # value at runtime can just ignore the second item.
        def add_row(lbl_text, var_val, is_bold=False, is_color=None):
            row = ttk.Frame(self.summary_box)
            font_style = (FONT_FAMILY, 11, "bold") if is_bold else (FONT_FAMILY, 10)
            
            lbl = ttk.Label(row, text=lbl_text, width=25, anchor="e", font=font_style)
            lbl.pack(side="left", padx=(0, 15))
            
            val_lbl = ttk.Label(row, textvariable=var_val, width=15, anchor="e", font=font_style)
            val_lbl.pack(side="right")
            if is_color: val_lbl.configure(bootstyle=is_color)
            return row, val_lbl

        # Same as add_row, but for rows whose *label* also changes at
        # runtime -- the GST rows (percentage depends on the tax rate)
        # and Balance Due (becomes "Advance / Credit" when overpaid).
        def add_dynamic_row(lbl_var, var_val, is_bold=False, is_color=None):
            row = ttk.Frame(self.summary_box)
            font_style = (FONT_FAMILY, 11, "bold") if is_bold else (FONT_FAMILY, 10)
            ttk.Label(row, textvariable=lbl_var, width=25, anchor="e", font=font_style).pack(side="left", padx=(0, 15))
            val_lbl = ttk.Label(row, textvariable=var_val, width=15, anchor="e", font=font_style)
            val_lbl.pack(side="right")
            if is_color: val_lbl.configure(bootstyle=is_color)
            return row, val_lbl

        # Rows are created here but deliberately left unpacked.
        # _recalc_totals() rebuilds the stacking order from scratch on
        # every call -- unpacking every row, then re-packing only the
        # ones currently relevant, in this fixed order:
        #   Subtotal -> Discount -> Taxable Amount -> Tax (CGST+SGST or
        #   IGST) -> Grand Total -> Total Paid -> Balance Due
        # A row that toggles from hidden to visible after later rows are
        # already packed would otherwise land at the *end* of the stack
        # instead of its correct place -- that's what used to put
        # Discount below Balance Due.
        self.row_subtotal,    _                  = add_row("Subtotal:", self.var_subtotal)
        self.row_disc,        _                  = add_row("Discount:", self.var_discount)
        self.row_taxable,     _                  = add_row("Taxable Amount:", self.var_taxable)
        self.row_cgst,        _                  = add_dynamic_row(self.var_cgst_lbl, self.var_cgst_val)
        self.row_sgst,        _                  = add_dynamic_row(self.var_sgst_lbl, self.var_sgst_val)
        self.row_igst,        _                  = add_dynamic_row(self.var_igst_lbl, self.var_igst_val)
        self.row_grand_total, _                  = add_row("Grand Total:", self.var_grand_total, is_bold=True)
        self.row_paid,        _                  = add_row("Total Paid:", self.var_paid)
        self.row_balance,     self._balance_lbl_widget = add_dynamic_row(
            self.var_balance_lbl, self.var_balance, is_bold=True, is_color="danger")

        self._summary_rows = (self.row_subtotal, self.row_disc, self.row_taxable, self.row_cgst, self.row_sgst,
                               self.row_igst, self.row_grand_total, self.row_paid, self.row_balance)

    def _section_save_bar(self, parent):
        row = ttk.Frame(parent, padding=(0, 4, 0, 24))
        row.pack(fill="x")
        back_btn = ttk.Button(row, text="←  Back to Dashboard  (Esc)", bootstyle="secondary",
                               command=lambda: self.controller.show_frame(DashboardFrame))
        back_btn.pack(side="right", padx=(8, 0))
        ToolTip(back_btn, "Return to the invoice list (Esc). Unsaved changes will ask first.")

        self.save_btn = ttk.Button(row, text="💾  Save Invoice  (Ctrl+S)", bootstyle="success", command=self._save)
        self.save_btn.pack(side="right")
        ToolTip(self.save_btn, "Save this invoice (Ctrl+S)")

    # ------------------------------------------------------------------ #
    #  Client autocomplete
    # ------------------------------------------------------------------ #
    def _filter_client_suggestions(self, _=None):
        text = self.client_combo.get().lower()
        self.client_combo["values"] = [c["name"] for c in self._client_cache if text in c["name"].lower()] if text else \
                                       [c["name"] for c in self._client_cache]

    def _on_client_picked(self, _=None):
        chosen = self.client_combo.get()
        for c in self._client_cache:
            if c["name"] == chosen:
                self.client_contact_text.delete("1.0", "end")
                self.client_contact_text.insert("1.0", c.get("contact") or "")
                self.client_gstin_entry.delete(0, "end")
                self.client_gstin_entry.insert(0, c.get("gstin") or "")
                break

    # ------------------------------------------------------------------ #
    #  Totals calculation (live, mirrors DatabaseManager.calculate_totals)
    # ------------------------------------------------------------------ #
    def _collect_items(self) -> List[Dict]:
        items = []
        for row in self.item_rows:
            items.append({
                "title": row["title_combo"].get().strip(),
                "description": row["desc_entry"].get().strip(),
                "amount": safe_float(row["amount_entry"].get()),
                "discount": max(0.0, safe_float(row["disc_entry"].get())),
            })
        return items

    def _collect_payments(self) -> List[Dict]:
        payments = []
        for row in self.payment_rows:
            amt = safe_float(row["amount_entry"].get())
            date_txt = row["date_entry"].entry.get().strip()
            db_date = None
            if date_txt:
                try:
                    db_date = datetime.datetime.strptime(date_txt, DISPLAY_DATE_FORMAT).strftime(DB_DATE_FORMAT)
                except ValueError:
                    db_date = None
            payments.append({
                "amount": amt, "payment_date": db_date,
                "mode": row["mode_combo"].get(), "note": row["note_entry"].get().strip(),
                "payment_id": row["payment_id"],
            })
        return payments

    def _snapshot(self) -> str:
        try:
            state = {
                "invoice_number": self.invoice_number_entry.get().strip(),
                "bill_to":        self.client_combo.get().strip(),
                "client_contact": self.client_contact_text.get("1.0", "end-1c").strip(),
                "client_gstin":   self.client_gstin_entry.get().strip(),
                "your_name":      self.your_name_entry.get().strip(),
                "your_contact":   self.your_contact_text.get("1.0", "end-1c").strip(),
                "invoice_date":   self.invoice_date_entry.entry.get().strip(),
                "payment_terms":  self.payment_terms_entry.get().strip(),
                "notes":          self.notes_text.get("1.0", "end-1c").strip(),
                "no_charge":      self.no_charge_var.get(),
                "show_qr":        self.show_qr_var.get(),
                "gst":            self.gst_var.get(),
                "tax_rate":       self.tax_rate_entry.get().strip(),
                "gst_type":       self.gst_type_var.get(),
                "g_disc":         self.global_discount_entry.get().strip(),
                "g_disc_type":    self.global_discount_type_var.get(),
                "items":          self._collect_items(),
                "payments":       self._collect_payments(),
            }
            return json.dumps(state, sort_keys=True, default=str)
        except Exception:
            return self._initial_snapshot or ""

    def has_unsaved_changes(self) -> bool:
        if self._initial_snapshot is None:
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

    def _recalc_totals(self, *_):
        """Calculates live totals and instantly updates the UI grid.

        Every row is unpacked and then re-packed from scratch on each
        call, rather than toggled in place, so the summary always lays
        out in the fixed, correct order -- Subtotal, Discount, Taxable
        Amount, Tax, Grand Total, Total Paid, Balance Due -- no matter
        which rows were visible on the previous call. This is a handful
        of Frame widgets, so repacking them on every keystroke has no
        noticeable cost.

        Balance/status math goes through DatabaseManager.compute_payment_
        status() -- the same function _recompute_one() uses when saving
        and the PDF uses when printing -- so this live preview can never
        show a different number or status than what actually gets saved.
        """
        for row in self._summary_rows:
            row.pack_forget()

        if self.no_charge_var.get():
            self.var_subtotal.set("0.00")
            self.var_grand_total.set("0.00")
            self.var_balance_lbl.set("Balance Due:")
            self.var_balance.set("0.00")
            self._balance_lbl_widget.configure(bootstyle="secondary")

            self.row_subtotal.pack(fill="x", pady=2)
            self.row_grand_total.pack(fill="x", pady=2)
            self.row_balance.pack(fill="x", pady=2)

            self.status_badge.config(text=STATUS_DISPLAY_LABELS[InvoiceStatus.COMPLIMENTARY.value], bootstyle="info-inverse")
            return

        items    = self._collect_items()
        tax_rate = safe_float(self.tax_rate_entry.get()) if self.gst_var.get() else 0.0
        g_disc   = safe_float(self.global_discount_entry.get())

        totals = DatabaseManager.calculate_totals(items, tax_rate, g_disc, self.global_discount_type_var.get())
        paid   = sum(safe_float(r["amount_entry"].get()) for r in self.payment_rows)
        status, balance, credit = DatabaseManager.compute_payment_status(totals["grand_total"], paid)

        # Grand Total = Subtotal - Discount + Tax: already true here,
        # since calculate_totals() derives grand_total as taxable +
        # cgst + sgst, where taxable is subtotal minus total_discount.
        self.var_subtotal.set(fmt_money(totals['subtotal']))
        self.var_taxable.set(fmt_money(totals['taxable']))
        self.var_grand_total.set(fmt_money(totals['grand_total']))

        # Balance Due is never negative: compute_payment_status() floors
        # it at zero and returns any overpayment separately as `credit`,
        # which is shown as "Advance / Credit" in a positive colour
        # instead of a confusing negative balance.
        if credit > 0:
            self.var_balance_lbl.set("Advance / Credit:")
            self.var_balance.set(fmt_money(credit))
            self._balance_lbl_widget.configure(bootstyle="info")
        else:
            self.var_balance_lbl.set("Balance Due:")
            self.var_balance.set(fmt_money(balance))
            self._balance_lbl_widget.configure(bootstyle="danger" if balance > 0 else "success")

        # 1. Subtotal always shows first.
        self.row_subtotal.pack(fill="x", pady=2)

        # 2. Discount -- shown as a negative value, only when there is one.
        if totals["total_discount"] > 0:
            self.var_discount.set(f"− {fmt_money(totals['total_discount'])}")
            self.row_disc.pack(fill="x", pady=2)

        # 3. Taxable Amount -- Subtotal minus Discount, the base tax is
        # actually calculated on. Only worth showing once there's a
        # discount or tax to make it different from the Subtotal above.
        if (totals["total_discount"] > 0 or (self.gst_var.get() and tax_rate > 0)):
            self.row_taxable.pack(fill="x", pady=2)

        # 4. Tax -- CGST/SGST is split equally and shown as two separate
        # lines (e.g. 18% -> CGST 9% + SGST 9%); IGST is shown as a
        # single line at the full rate. calculate_totals() always splits
        # the tax into equal cgst/sgst halves regardless of type, so
        # IGST's single amount is just those two halves added back
        # together at the full rate.
        if self.gst_var.get() and tax_rate > 0:
            half_rate = tax_rate / 2.0
            if self.gst_type_var.get() == "IGST":
                self.var_igst_lbl.set(f"IGST ({tax_rate:g}%):")
                self.var_igst_val.set(fmt_money(totals['cgst'] + totals['sgst']))
                self.row_igst.pack(fill="x", pady=2)
            else:
                self.var_cgst_lbl.set(f"CGST ({half_rate:g}%):")
                self.var_cgst_val.set(fmt_money(totals['cgst']))
                self.var_sgst_lbl.set(f"SGST ({half_rate:g}%):")
                self.var_sgst_val.set(fmt_money(totals['sgst']))
                self.row_cgst.pack(fill="x", pady=2)
                self.row_sgst.pack(fill="x", pady=2)

        # 5. Grand Total always shows.
        self.row_grand_total.pack(fill="x", pady=2)

        # 6. Total Paid -- only when something has actually been paid.
        if paid > 0:
            self.var_paid.set(f"− {fmt_money(paid)}")
            self.row_paid.pack(fill="x", pady=2)

        # 7. Balance Due (or Advance / Credit) always shows last.
        self.row_balance.pack(fill="x", pady=2)

        # Status badge -- UNPAID / PARTIALLY PAID / PAID / OVERPAID,
        # from the exact same status compute_payment_status() just
        # returned, so the badge can never disagree with the numbers
        # right above it.
        self.status_badge.config(text=STATUS_DISPLAY_LABELS.get(status, status.upper()),
                                  bootstyle=STATUS_BOOTSTYLE.get(status, "secondary-inverse"))

    def refresh_after_settings_change(self):
        self._catalog_cache = self.db_manager.get_all_catalog_items()
        self._client_cache  = self.db_manager.get_all_clients()
        self.client_combo["values"] = [c["name"] for c in self._client_cache]
        self._recalc_totals() 

    def on_show(self, invoice_id: Optional[str] = None):
        self._catalog_cache = self.db_manager.get_all_catalog_items()
        self._client_cache  = self.db_manager.get_all_clients()
        self.client_combo["values"] = [c["name"] for c in self._client_cache]
        self._reset()
        if invoice_id:
            self._load_for_edit(invoice_id)
        else:
            self.editing_id = None
            self.title_lbl.config(text="New Invoice")
            self.save_btn.config(text="💾  Save Invoice  (Ctrl+S)")
            self.invoice_number_entry.insert(0, self.db_manager.get_next_invoice_number())
            self.invoice_date_entry.entry.delete(0, "end")
            self.invoice_date_entry.entry.insert(0, datetime.date.today().strftime(DISPLAY_DATE_FORMAT))
            self.payment_terms_entry.insert(0, CONFIG.get("payment_details", {}).get("payment_terms", ""))
            self.your_name_entry.insert(0, compose_your_name(CONFIG))
            self.your_contact_text.insert("1.0", compose_your_contact(CONFIG))
            self._add_item_row()
            self._apply_field_states()
            self._recalc_totals()
        self._initial_snapshot = self._snapshot()

    def _reset(self):
        for entry in list(self.item_rows):
            entry["frame"].destroy()
        self.item_rows.clear()
        for entry in list(self.payment_rows):
            entry["frame"].destroy()
        self.payment_rows.clear()

        self.client_combo.delete(0, "end")
        self.client_contact_text.delete("1.0", "end")
        self.client_gstin_entry.delete(0, "end")
        self.your_name_entry.delete(0, "end")
        self.your_contact_text.delete("1.0", "end")
        self.invoice_number_entry.delete(0, "end")
        self.invoice_date_entry.entry.delete(0, "end")
        self.payment_terms_entry.delete(0, "end")
        self.notes_text.delete("1.0", "end")
        self.no_charge_var.set(False)
        self.show_qr_var.set(True)
        self.gst_var.set(True)
        self.tax_rate_entry.delete(0, "end"); self.tax_rate_entry.insert(0, "18")
        self.gst_type_var.set("CGST/SGST")
        self.global_discount_entry.delete(0, "end"); self.global_discount_entry.insert(0, "0")
        self.global_discount_type_var.set("amount")
        self.status_badge.config(text="", bootstyle="default")

    def _load_for_edit(self, invoice_id: str):
        row = self.db_manager.get_invoice_by_id(invoice_id)
        if not row:
            Messagebox.show_error("Not Found", f"Invoice {invoice_id} no longer exists.")
            self.controller.show_frame(DashboardFrame)
            return
        d = dict(row)
        self.editing_id = invoice_id
        self.title_lbl.config(text=f"Editing Invoice #{invoice_id}")
        self.save_btn.config(text="💾  Update Invoice  (Ctrl+S)")

        self.client_combo.insert(0, d.get("bill_to") or "")
        self.client_contact_text.insert("1.0", d.get("client_contact") or "")
        self.client_gstin_entry.insert(0, d.get("client_gstin") or "")
        self.your_name_entry.insert(0, d.get("your_name") or "")
        self.your_contact_text.insert("1.0", d.get("your_contact") or "")
        self.invoice_number_entry.insert(0, d.get("invoice_number") or "")
        self.invoice_date_entry.entry.insert(0, parse_date(d.get("invoice_date"), DB_DATE_FORMAT, DISPLAY_DATE_FORMAT))
        self.payment_terms_entry.insert(0, d.get("payment_terms") or "")
        self.notes_text.insert("1.0", d.get("notes") or "")
        self.no_charge_var.set(bool(d.get("is_complimentary")))
        self.show_qr_var.set(bool(d.get("show_qr_code", 1)))
        self.gst_var.set(bool(d.get("is_gst_invoice", 1)))
        self.tax_rate_entry.delete(0, "end"); self.tax_rate_entry.insert(0, str(d.get("tax_rate") or 0))
        self.gst_type_var.set(d.get("gst_type") or "CGST/SGST")
        self.global_discount_entry.delete(0, "end"); self.global_discount_entry.insert(0, str(d.get("global_discount_value") or 0))
        self.global_discount_type_var.set(d.get("global_discount_type") or "amount")

        for item in json.loads(d.get("items") or "[]"):
            self._add_item_row(item.get("title", ""), item.get("description", ""),
                               safe_float(item.get("amount")), safe_float(item.get("discount")))
        if not self.item_rows:
            self._add_item_row()

        for p in self.db_manager.get_payments(invoice_id):
            self._add_payment_row(safe_float(p.get("amount")), p.get("payment_date"),
                                  p.get("mode") or "Cash", p.get("note") or "", payment_id=p.get("id"))

        self._apply_field_states()
        self._recalc_totals()

    # ------------------------------------------------------------------ #
    #  Validation & Save
    # ------------------------------------------------------------------ #
    def _save(self):
        inv_num = self.invoice_number_entry.get().strip()
        bill_to = self.client_combo.get().strip()
        your_name = self.your_name_entry.get().strip()
        if not inv_num or not bill_to or not your_name:
            return Messagebox.show_error("Validation", "Invoice Number, Client Name, and Your Name are all required.")
        if INVALID_FILENAME_CHARS.search(inv_num):
            return Messagebox.show_error(
                "Validation",
                'Invoice Number cannot contain any of:  \\  /  :  * ?  "  <  >  |\n'
                "(these characters break the PDF file that gets generated from it).",
            )

        date_txt = self.invoice_date_entry.entry.get().strip()
        if not date_txt:
            return Messagebox.show_error("Validation", "Invoice Date is required.")
        try:
            inv_date = datetime.datetime.strptime(date_txt, DISPLAY_DATE_FORMAT).strftime(DB_DATE_FORMAT)
        except ValueError:
            return Messagebox.show_error("Validation", "Invoice Date is not a valid date.")

        no_charge = self.no_charge_var.get()
        tax_rate  = 0.0
        if not no_charge and self.gst_var.get():
            tax_rate = safe_float(self.tax_rate_entry.get())
            if not (0 <= tax_rate <= 100):
                return Messagebox.show_error("Validation", "GST / Tax Rate must be between 0 and 100.")

        g_disc = safe_float(self.global_discount_entry.get())
        if g_disc < 0:
            return Messagebox.show_error("Validation", "Overall discount cannot be negative.")
        if self.global_discount_type_var.get() == "%" and g_disc > 100:
            return Messagebox.show_error("Validation", "A percentage discount cannot exceed 100%.")

        items = self._collect_items()
        items = [i for i in items if i["title"] or i["description"] or i["amount"]]
        if not no_charge and not items:
            return Messagebox.show_error("Validation", "Add at least one item to the invoice.")
        for i in items:
            if i["amount"] < 0 or i["discount"] < 0:
                return Messagebox.show_error("Validation", "Item amounts and discounts cannot be negative.")

        payments = []
        if not no_charge:
            for p in self._collect_payments():
                if p["amount"] <= 0:
                    continue
                if not p["payment_date"]:
                    return Messagebox.show_error("Validation", "Every payment row needs a valid date.")
                payments.append(p)

        data = {
            "invoice_number": inv_num, "bill_to": bill_to,
            "client_contact": self.client_contact_text.get("1.0", "end-1c").strip(),
            "client_gstin": self.client_gstin_entry.get().strip(),
            "your_name": your_name,
            "your_contact": self.your_contact_text.get("1.0", "end-1c").strip(),
            "items": json.dumps(items), "amount_paid": 0.0, "tax_rate": tax_rate,
            "invoice_date": inv_date, "payment_date": None,
            "status": InvoiceStatus.DUE.value, "payment_terms": self.payment_terms_entry.get().strip(),
            "payment_mode": None, "show_qr_code": int(self.show_qr_var.get()),
            "is_complimentary": int(no_charge), "is_gst_invoice": int(self.gst_var.get()),
            "global_discount_value": g_disc, "global_discount_type": self.global_discount_type_var.get(),
            "advance_amount": 0.0, "advance_date": None, "gst_type": self.gst_type_var.get(),
            "notes": self.notes_text.get("1.0", "end-1c").strip(),
        }

        is_update = self.editing_id is not None
        ok, err = self.db_manager.save_invoice(data, is_update, original_id=self.editing_id,
                                                payments=None if no_charge else payments)
        if not ok:
            return Messagebox.show_error("Save Failed", err)

        if bill_to:
            self.db_manager.save_client(bill_to, data["client_contact"], data["client_gstin"])
        for item in items:
            if item["title"]:
                self.db_manager.save_catalog_item(item["title"], item["description"], item["amount"])

        Messagebox.show_info("Saved", f"Invoice #{inv_num} saved successfully.")
        self._initial_snapshot = self._snapshot() 
        self.controller.show_frame(DashboardFrame)

    # ------------------------------------------------------------------ #
    #  Public API used by application-level keyboard shortcuts
    # ------------------------------------------------------------------ #
    def save(self):
        """Ctrl+S: validates and saves the invoice, exactly like the
        Save button."""
        self._save()

    def _focused_item_row(self) -> Optional[Dict]:
        """Walks up from the current keyboard focus to find which item
        row (if any) owns it, so shortcuts can act on "whichever row the
        user is in" the same way F2/Ctrl+Shift+Delete do in a spreadsheet."""
        widget = self.focus_get()
        while widget is not None:
            for row in self.item_rows:
                if widget is row.get("frame"):
                    return row
            widget = getattr(widget, "master", None)
        return None

    def add_item(self):
        """Ctrl+Shift+A / Ctrl+Plus: adds a new blank item row and moves
        focus straight into it, ready to type."""
        row = self._add_item_row()
        row["title_combo"].focus_set()

    def remove_selected_item(self):
        """Ctrl+Shift+Delete / Ctrl+Minus: removes whichever item row
        currently has focus, or the last row if none does -- the same
        row a row's own "✕" button would remove."""
        row = self._focused_item_row() or (self.item_rows[-1] if self.item_rows else None)
        if row is not None:
            self._remove_item_row(row)

    def edit_selected_item(self):
        """F2: jumps focus into the title field of the currently
        focused/last item row, ready to edit."""
        row = self._focused_item_row() or (self.item_rows[-1] if self.item_rows else None)
        if row is not None:
            row["title_combo"].focus_set()


# =============================================================================
#  APPLICATION SHELL  (sidebar navigation + page container)
# =============================================================================
class Sidebar(ttk.Frame):
    NAV_ITEMS = [
        ("📊  Dashboard", "DashboardFrame"),
        ("👥  Clients",   "ClientsFrame"),
        ("📦  Catalog",   "CatalogFrame"),
    ]

    def __init__(self, parent, controller):
        # Pure ttk.Frame automatically listens to LIVE theme changes
        super().__init__(parent, width=220) 
        self.controller = controller
        self.pack_propagate(False)
        self._nav_btns: Dict[str, ttk.Button] = {}
        self._build()

    def _build(self):
        top = ttk.Frame(self, padding=(18, 22, 18, 12))
        top.pack(fill="x")
        
        # Centering the logo and text looks much cleaner
        if logo := load_app_logo_image(self):
            lbl = ttk.Label(top, image=logo)
            lbl.image = logo
            lbl.pack(anchor="center", pady=(0, 8))
            
        ttk.Label(top, text="Vyaprix", font=(FONT_FAMILY, 16, "bold")).pack(anchor="center")
        ttk.Label(top, text="by EllowLabs", font=(FONT_FAMILY, 8), bootstyle="secondary").pack(anchor="center")

        ttk.Separator(self, bootstyle="secondary").pack(fill="x", padx=18, pady=(4, 12))

        nav = ttk.Frame(self, padding=(10, 0))
        nav.pack(fill="x")
        
        for label, frame_name in self.NAV_ITEMS:
            # Outline buttons adapt flawlessly to both light and dark backgrounds
            btn = ttk.Button(nav, text=label, bootstyle="secondary-outline", padding=(10, 10),
                             command=lambda n=frame_name: self.controller.show_frame_by_name(n))
            btn.pack(fill="x", pady=4)
            self._nav_btns[frame_name] = btn

        bottom = ttk.Frame(self, padding=(10, 14))
        bottom.pack(side="bottom", fill="x")
        
        ttk.Button(bottom, text="⚙️  Settings", bootstyle="info-outline", padding=(10, 8),
                  command=self.controller.open_settings).pack(fill="x", pady=4)
        ttk.Button(bottom, text="⌨️  Shortcuts", bootstyle="info-outline", padding=(10, 8),
                  command=self.controller.open_shortcuts).pack(fill="x", pady=4)
        ttk.Button(bottom, text="ℹ️  About", bootstyle="info-outline", padding=(10, 8),
                  command=self.controller.open_about).pack(fill="x", pady=4)
                  
        ttk.Label(bottom, text=APP_VERSION, font=(FONT_FAMILY, 8), bootstyle="secondary",
                 anchor="center").pack(fill="x", padx=10, pady=(6, 0))

    def set_active(self, frame_name: str):
        for name, btn in self._nav_btns.items():
            # Solid primary blue for active tab, outline for inactive tabs
            btn.configure(bootstyle="primary" if name == frame_name else "secondary-outline")

# =============================================================================
#  SPLASH SCREEN (MODERNIZED)
# =============================================================================
class SplashScreen(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        # Remove standard window borders and controls
        self.overrideredirect(True)
        # Keep splash screen on top of other windows
        self.attributes("-topmost", True)
        
        # Start fully transparent for the fade-in effect
        self.attributes("-alpha", 0.0) 
        
        # --- Size and Centering ---
        width, height = 420, 420
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - width) // 2
        y = (sh - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # --- Modern Pseudo-Border ---
        # Because we removed the OS window, we lose drop shadows. 
        # A 1-2px border frame keeps the app from bleeding into white backgrounds.
        try:
            border_color = ttk.Style().colors.primary
        except Exception:
            border_color = "#3b82f6"  # fallback if the style isn't resolvable yet
        self.border_frame = tk.Frame(self, bg=border_color, bd=0)
        self.border_frame.pack(fill="both", expand=True)
        
        # --- Main UI Container ---
        # Using the default theme background makes it look cleaner than a solid block of color
        self.main_frame = ttk.Frame(self.border_frame, padding=2)
        self.main_frame.pack(fill="both", expand=True, padx=2, pady=2)
        
        # --- Center Content (Logo & Branding) ---
        center_content = ttk.Frame(self.main_frame)
        # Place it slightly above absolute center for better visual balance
        center_content.place(relx=0.5, rely=0.45, anchor="center")
        
        logo = load_app_logo_image(self)
        if logo:
            lbl = ttk.Label(center_content, image=logo)
            lbl.image = logo
            lbl.pack(pady=(0, 20))
            
        ttk.Label(center_content, text=APP_NAME, font=(FONT_FAMILY, 28, "bold")).pack()
        ttk.Label(center_content, text=APP_TAGLINE, font=(FONT_FAMILY, 10), bootstyle="secondary").pack(pady=(5, 0))
        
        # --- Bottom Bar (Loading Status) ---
        bottom_bar = ttk.Frame(self.main_frame, padding=(20, 15))
        bottom_bar.pack(side="bottom", fill="x")
        
        ttk.Label(bottom_bar, text="Starting up...", font=(FONT_FAMILY, 8, "bold"), bootstyle="secondary").pack(side="left")
        ttk.Label(bottom_bar, text=APP_VERSION, font=(FONT_FAMILY, 8), bootstyle="secondary").pack(side="right")
        
        pb = ttk.Progressbar(self.main_frame, mode="indeterminate", bootstyle="primary")
        pb.pack(side="bottom", fill="x")
        pb.start(15)

        # Trigger the fade-in animation
        self._fade_in()

    def _fade_in(self):
        """Creates a smooth fade-in effect."""
        current_alpha = self.attributes("-alpha")
        if current_alpha < 0.97: # 97% gives it a very subtle, modern transparency
            current_alpha += 0.05
            self.attributes("-alpha", current_alpha)
            self.after(15, self._fade_in)

# =============================================================================
#  CENTRALIZED KEYBOARD SHORTCUT MANAGER
# =============================================================================
def _humanize_sequence(sequence: str) -> str:
    """Turns a Tk binding string like '<Control-Shift-A>' into a
    human-readable label like 'Ctrl+Shift+A', for the Shortcuts
    reference dialog. Only needs to handle the common cases well --
    callers with an unusual or aliased combo (Ctrl+Plus, Ctrl+Minus)
    pass their own `display` string instead of relying on this."""
    parts = sequence.strip("<>").split("-")
    words = {"Control": "Ctrl", "Shift": "Shift", "Delete": "Del", "Escape": "Esc"}
    return "+".join(words.get(p, p.upper() if len(p) == 1 else p) for p in parts)


class ShortcutManager:
    """Single source of truth for every keyboard shortcut in the app --
    what key combo does what, and in which context.

    Two kinds of registrations:
      - **global**: the action is identical no matter which page is
        showing (Ctrl+Q, Ctrl+O, Ctrl+B, F1, ...).
      - **contextual**: the action depends on which frame is currently
        active -- this is what makes Ctrl+N mean "New Invoice" on the
        Dashboard, "New Client" on Clients, and "New Catalog Item" on
        Catalog, all from one registry instead of scattered isinstance
        checks.

    Conflict prevention: registering the same key sequence twice for
    the same scope (two globals, or the same sequence+frame combo
    twice) raises ValueError immediately, at startup -- so a copy-paste
    mistake while adding a shortcut fails loudly the moment the app
    launches, rather than silently shadowing an existing shortcut and
    only being noticed later as "that key stopped working".

    Every dispatch gets two safety behaviours for free:
      1. **No firing under an open dialog.** A modal dialog (Settings,
         PDF viewer, Add Client, ...) owns the keyboard while it's open;
         otherwise Ctrl+N while "Add Client" is open would pop a second
         dialog behind the first.
      2. **No fighting Tk's own bindings.** Ctrl+D normally deletes a
         character in Entry/Text widgets, and Ctrl+O inserts a newline in
         Text -- both collide with shortcuts this app wants. install_
         class_overrides() re-binds those two (widget class, sequence)
         pairs through this same dispatcher so typing never fights the
         shortcut sharing its key.
    """

    #: Tk class bindings that would otherwise silently compete with one
    #: of our shortcuts. Only entries that are actually registered below
    #: get overridden -- see install_class_overrides().
    _TK_COLLISIONS: Dict[str, Tuple[str, ...]] = {
        "<Control-d>": ("Entry", "TEntry", "TCombobox", "Text"),
        "<Control-o>": ("Text",),
    }

    def __init__(self, app: "InvoiceApp"):
        self.app = app
        self._owner: Dict[str, str] = {}                              # sequence -> "global" | "contextual"
        self._global_handlers: Dict[str, Callable[[], None]] = {}
        self._contextual_handlers: Dict[str, Dict[Type, Callable[[], None]]] = {}
        self._bound_sequences: set = set()
        self._reference: List[Dict[str, str]] = []                     # feeds ShortcutsDialog

    # ---- registration --------------------------------------------------- #
    def register_global(self, sequences, handler: Callable[[], None], description: str,
                         group: str, display: Optional[str] = None):
        """Registers a shortcut whose action never changes, regardless of
        which page or dialog is active. `sequences` is a single Tk
        binding string, or a tuple of equivalent ones (e.g. Ctrl+Plus is
        reachable as <Control-equal>, <Control-plus>, or <Control-KP_Add>
        depending on keyboard/platform -- they're one shortcut, not
        three, so they share one reference-card entry)."""
        seqs = (sequences,) if isinstance(sequences, str) else tuple(sequences)
        for s in seqs:
            if s in self._owner:
                raise ValueError(f"Shortcut conflict: {s} is already registered ({self._owner[s]}).")
        for s in seqs:
            self._owner[s] = "global"
            self._global_handlers[s] = handler
            self._bind(s)
        self._reference.append({"keys": display or _humanize_sequence(seqs[0]),
                                 "description": description, "group": group})

    def register_contextual(self, sequences, frame_class: Type, handler: Callable[[], None],
                             description: str, group: str, display: Optional[str] = None,
                             show_in_reference: bool = True):
        """Registers a shortcut whose action depends on which frame is
        active. The same sequence can be registered for multiple frame
        classes (that's the whole point -- Ctrl+N means something
        different on each page); it can NOT be registered twice for the
        *same* frame class, and it can not be mixed with a global
        registration of the same sequence -- both raise ValueError."""
        seqs = (sequences,) if isinstance(sequences, str) else tuple(sequences)
        for s in seqs:
            if self._owner.get(s, "contextual") != "contextual":
                raise ValueError(f"Shortcut conflict: {s} is already registered ({self._owner[s]}).")
            if frame_class in self._contextual_handlers.get(s, {}):
                raise ValueError(f"Shortcut conflict: {s} is already registered for {frame_class.__name__}.")
        for s in seqs:
            self._owner[s] = "contextual"
            self._contextual_handlers.setdefault(s, {})[frame_class] = handler
            self._bind(s)
        if show_in_reference:
            page = frame_class.__name__.replace("Frame", "") or frame_class.__name__
            self._reference.append({"keys": display or _humanize_sequence(seqs[0]),
                                     "description": f"{description} ({page})", "group": group})

    def register_info(self, display: str, description: str, group: str):
        """Adds a Shortcuts-reference line for real behaviour that isn't
        implemented as a single bind_all/bind_class registration --
        Tab/Shift+Tab field order and Enter-advances-focus are native Tk
        traversal, and "Esc closes this dialog" is bound locally on each
        individual dialog rather than centrally. Keeping them in the
        same reference list (instead of a second hardcoded list inside
        ShortcutsDialog) means there's still exactly one place that
        describes every documented shortcut, real bindings and native
        behaviour alike."""
        self._reference.append({"keys": display, "description": description, "group": group})

    def _bind(self, sequence: str):
        if sequence in self._bound_sequences:
            return
        self._bound_sequences.add(sequence)
        self.app.bind_all(sequence, lambda event, seq=sequence: self._dispatch(seq, event))

    def install_class_overrides(self):
        """Call once, after every shortcut is registered. Re-binds the
        small, known set of Tk widget-class bindings (see
        _TK_COLLISIONS) that would otherwise compete with a registered
        shortcut, routing them through the exact same dispatcher used
        everywhere else so behaviour can never diverge between "focus is
        in a text field" and "focus is anywhere else"."""
        for sequence, widget_classes in self._TK_COLLISIONS.items():
            if sequence not in self._owner:
                continue
            for widget_class in widget_classes:
                self.app.bind_class(widget_class, sequence,
                                     lambda event, seq=sequence: self._dispatch(seq, event))

    # ---- dispatch -------------------------------------------------------- #
    def _dispatch(self, sequence: str, event) -> str:
        # Safety rule 1: a modal dialog owns the keyboard while it's open.
        # event.widget is normally a live widget object, but Tk can hand
        # back something that fails to resolve cleanly in a few edge
        # cases -- a bare path string if the widget is mid-destroy
        # (TclError/AttributeError), or, as testing turned up, a
        # reference into a combobox's internal dropdown popup that
        # Python's binding can't map back to a widget object (KeyError).
        # None of that is worth crashing over: if we can't confidently
        # tell where this event came from, the safe answer is to do
        # nothing, the same as if a dialog were genuinely focused.
        try:
            if isinstance(event.widget, str) or event.widget.winfo_toplevel() is not self.app:
                return "break"
        except Exception:
            return "break"   # can't safely resolve this event's widget; nothing to act on

        handler = self._global_handlers.get(sequence)
        if handler is None:
            per_frame = self._contextual_handlers.get(sequence)
            if per_frame:
                handler = per_frame.get(type(self.app._active_frame()))
        if handler:
            handler()
        return "break"

    # ---- introspection, used by ShortcutsDialog --------------------------- #
    def grouped_reference(self) -> Dict[str, List[Tuple[str, str]]]:
        """Returns {group_name: [(key_label, description), ...]} built
        straight from the live registry, in registration order, so the
        Shortcuts reference dialog can never drift out of sync with what
        the keys actually do."""
        groups: Dict[str, List[Tuple[str, str]]] = {}
        for entry in self._reference:
            groups.setdefault(entry["group"], []).append((entry["keys"], entry["description"]))
        return groups


class InvoiceApp(ttk.Window):
    def __init__(self):
        cfg = load_config()
        super().__init__(themename=cfg.get("theme", "litera"))
        configure_windows_app_identity()
        
        # --- BULLETPROOF SCROLL FIX ---
        # Reaches directly into the Tcl interpreter to permanently delete 
        # the native bindings that cause Comboboxes to change values on scroll.
        self.tk.call('bind', 'TCombobox', '<MouseWheel>', '')
        self.tk.call('bind', 'TCombobox', '<Button-4>', '')
        self.tk.call('bind', 'TCombobox', '<Button-5>', '')
        
        # 1. HIDE MAIN WINDOW IMMEDIATELY
        self.withdraw()
        
        # 2. SHOW SPLASH SCREEN
        self.splash = SplashScreen(self)
        self.update() # Force Tkinter to draw the splash screen immediately

        self.title(f"{APP_NAME} • {APP_TAGLINE}")
        self.geometry("1280x800")
        self.minsize(1080, 680)
        try:
            if APP_LOGO_ICO.exists() and sys.platform == "win32":
                self.iconbitmap(str(APP_LOGO_ICO))
        except Exception:
            pass

        self.db_manager = DatabaseManager(DB_NAME)
        self.task_queue: queue.Queue = queue.Queue()

        self._current_frame_class: Optional[Type] = None
        self._build_shell()
        self._setup_keyboard_shortcuts()
        self.show_frame(DashboardFrame)
        self.after(200, self._poll_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # 3. SCHEDULE SPLASH SCREEN REMOVAL (2.5 seconds later)
        self.after(2500, self._close_splash_and_show_main)

    # 4. ADD THIS NEW METHOD TO HANDLE THE TRANSITION
    def _close_splash_and_show_main(self):
        if hasattr(self, "splash") and self.splash.winfo_exists():
            self.splash.destroy()
        self.deiconify() # Reveal the main application window
        self._warn_if_fallback_font()
    # ------------------------------------------------------------------ #
    def _warn_if_fallback_font(self):
        """One-time, non-blocking heads-up if PDF generation is running
        without the bundled NotoSans font (see register_fonts()). Without
        this, a missing/corrupt font install fails completely silently --
        every amount on every PDF would quietly show the wrong character
        in place of the currency symbol, with nothing telling the user or
        support anything is wrong. Only shown when the fallback is
        actually active, so it's invisible to the vast majority of
        installs where the fonts loaded correctly."""
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
    def _build_shell(self):
        self.sidebar = Sidebar(self, self)
        self.sidebar.pack(side="left", fill="y")
        ttk.Separator(self, orient="vertical").pack(side="left", fill="y")

        right = ttk.Frame(self)
        right.pack(side="left", fill="both", expand=True)

        self.progress_bar = ttk.Progressbar(right, mode="indeterminate", bootstyle="primary-striped")
        self.progress_lbl = ttk.Label(right, text="", font=(FONT_FAMILY, 9), bootstyle="secondary")

        self.container = ttk.Frame(right)
        self.container.pack(fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames: Dict[Type, ttk.Frame] = {}
        for FrameClass in (DashboardFrame, ClientsFrame, CatalogFrame, InvoiceFormFrame):
            frame = FrameClass(self.container, self)
            self.frames[FrameClass] = frame
            frame.grid(row=0, column=0, sticky="nsew")

    # ------------------------------------------------------------------ #
    #  Keyboard shortcuts
    # ------------------------------------------------------------------ #
    def _setup_keyboard_shortcuts(self):
        """Builds the app's ShortcutManager and registers every keyboard
        shortcut through it -- this method is the single place that
        defines what every key combo does, in every context. See
        ShortcutManager's class docstring for the safety behaviours
        (no firing under an open dialog, no fighting Tk's own Ctrl+D /
        Ctrl+O bindings) every registration below gets automatically.

        Because each entry goes through register_global/register_contextual,
        two shortcuts can never silently collide: registering the same
        key sequence twice for the same scope raises immediately, so a
        copy-paste mistake is caught the moment the app starts, not
        discovered later as "that key stopped working".
        """
        sc = self.shortcuts = ShortcutManager(self)

        # ---- Ctrl+N: context-aware "create new thing for this page" --
        # Registered once per page, each with its own handler and its
        # own reference-card line -- this is what makes one physical key
        # combo mean three different (correct) things depending on
        # where you are, instead of an if/elif chain buried in a handler.
        sc.register_contextual("<Control-n>", DashboardFrame,   self._sc_new_invoice,      "New Invoice",        "Navigation", display="Ctrl+N")
        sc.register_contextual("<Control-n>", ClientsFrame,     self._sc_new_client,       "New Client",         "Navigation", display="Ctrl+N")
        sc.register_contextual("<Control-n>", CatalogFrame,     self._sc_new_catalog_item, "New Catalog Item",   "Navigation", display="Ctrl+N")
        # Also meaningful while already on the invoice form (start
        # another blank one); kept out of the reference card to avoid
        # a confusing near-duplicate line right under "New Invoice
        # (Dashboard)" -- it's the same action, just reachable from a
        # second place.
        sc.register_contextual("<Control-n>", InvoiceFormFrame, self._sc_new_invoice,      "New Invoice",        "Navigation", show_in_reference=False)

        # ---- Global: identical action wherever you are ----------------
        sc.register_global("<Control-o>", self._sc_open_invoice,    "Open Invoice",    "Navigation", display="Ctrl+O")
        sc.register_global("<Control-q>", self._sc_quit,            "Quit",            "Navigation", display="Ctrl+Q")
        sc.register_global("<Control-b>", self._sc_backup_database, "Backup Database", "Navigation", display="Ctrl+B")

        # ---- Contextual: Dashboard --------------------------------------
        sc.register_contextual("<Control-f>", DashboardFrame, lambda: self.frames[DashboardFrame].focus_search(), "Search",            "Navigation", display="Ctrl+F")
        sc.register_contextual("<Control-d>", DashboardFrame, self._sc_duplicate_invoice,                          "Duplicate Invoice", "Navigation", display="Ctrl+D")
        sc.register_contextual("<Control-p>", DashboardFrame, self._sc_print_pdf,                                  "Print PDF",         "Navigation", display="Ctrl+P")
        sc.register_contextual("<Control-e>", DashboardFrame, self._sc_export_pdf,                                 "Export PDF",        "Navigation", display="Ctrl+E")
        sc.register_contextual("<Control-r>", DashboardFrame, self._sc_refresh_dashboard,                          "Refresh List",      "Navigation", display="Ctrl+R")
        sc.register_contextual("<Control-a>", DashboardFrame, self._sc_select_all_invoices,                        "Select All",         "Navigation", display="Ctrl+A")

        # ---- Contextual: Clients / Catalog search boxes ------------------
        sc.register_contextual("<Control-f>", ClientsFrame, lambda: self.frames[ClientsFrame].focus_search(), "Search", "Navigation", display="Ctrl+F")
        sc.register_contextual("<Control-f>", CatalogFrame, lambda: self.frames[CatalogFrame].focus_search(), "Search", "Navigation", display="Ctrl+F")

        # ---- Contextual: Invoice Form -------------------------------------
        sc.register_contextual("<Control-s>", InvoiceFormFrame, self._sc_save_invoice, "Save Invoice", "Invoice Form", display="Ctrl+S")
        sc.register_contextual("<Control-Shift-A>", InvoiceFormFrame, self._sc_add_item, "Add Item", "Invoice Form", display="Ctrl+Shift+A")
        sc.register_contextual(("<Control-equal>", "<Control-plus>", "<Control-KP_Add>"),      # Ctrl+Plus: physical =/+ key
                                InvoiceFormFrame, self._sc_add_item, "Add Item (Row)", "Invoice Form", display="Ctrl+Plus")
        sc.register_contextual("<Control-Shift-Delete>", InvoiceFormFrame, self._sc_remove_item, "Delete Selected Item", "Invoice Form", display="Ctrl+Shift+Del")
        sc.register_contextual(("<Control-minus>", "<Control-KP_Subtract>"),
                                InvoiceFormFrame, self._sc_remove_item, "Remove Item (Row)", "Invoice Form", display="Ctrl+Minus")
        sc.register_contextual("<F2>", InvoiceFormFrame, self._sc_edit_item, "Edit Selected Item", "Invoice Form", display="F2")
        # Esc on the invoice form == clicking "Back to Dashboard" -- same
        # method, same unsaved-changes confirmation, just from the keyboard.
        sc.register_contextual("<Escape>", InvoiceFormFrame, self._sc_back_to_dashboard, "Back to Dashboard", "Invoice Form", display="Esc")
        # Native Tk field traversal / per-widget Enter binding, not a
        # single central registration -- documented here anyway so the
        # reference card stays the one complete list of shortcuts.
        sc.register_info("Tab / Shift+Tab", "Move Between Fields", "Invoice Form")
        sc.register_info("↑ / ↓  •  Page Up / Down", "Scroll the Form", "Invoice Form")
        sc.register_info("Enter", "Next Field", "Invoice Form")

        # ---- Anywhere -------------------------------------------------------
        sc.register_global("<F1>", self._sc_show_shortcuts, "Show This Help", "Anywhere", display="F1")
        # Every dialog (Settings, About, PDF Viewer, Add Client, ...)
        # binds its own local Esc -> close/cancel; documented here too.
        sc.register_info("Esc", "Close Dialog / PDF Viewer", "Anywhere")

        # Must run after every shortcut above is registered -- see
        # ShortcutManager.install_class_overrides() for why Ctrl+D and
        # Ctrl+O specifically need this extra step.
        sc.install_class_overrides()

    def _active_frame(self) -> Optional[ttk.Frame]:
        return self.frames.get(self._current_frame_class)

    # Handlers below are only ever invoked once ShortcutManager has
    # already confirmed the right context (or none, for globals) -- so
    # unlike the old dict-dispatch version, they no longer need their own
    # isinstance/hasattr guards. That check now lives in exactly one
    # place: the registrations above.
    def _sc_new_invoice(self):
        """Ctrl+N on the Dashboard, or on the invoice form itself: start
        a fresh invoice. If the invoice form is already open with
        unsaved edits, ask before discarding them first -- the same
        confirmation navigating away from it any other way would give."""
        frame = self._active_frame()
        if isinstance(frame, InvoiceFormFrame) and not frame.confirm_navigate_away():
            return
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
        """Ctrl+O: always jumps to the invoice list and into its search
        box, regardless of which screen you're currently on -- unlike
        Ctrl+F, which searches whatever screen is already active."""
        self.show_frame(DashboardFrame)
        if self._current_frame_class is DashboardFrame:
            self.frames[DashboardFrame].focus_search()

    def _sc_quit(self):
        self._on_close()

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
        """Esc on the invoice form: identical to the "Back to Dashboard"
        button -- show_frame() already asks about unsaved changes on its
        own when leaving the invoice form for a different page."""
        self.show_frame(DashboardFrame)

    def _sc_show_shortcuts(self):
        self.open_shortcuts()

    # ------------------------------------------------------------------ #
    #  Navigation
    # ------------------------------------------------------------------ #
    def show_frame(self, frame_class: Type, *args):
        # Give the currently-visible frame a chance to block navigation
        # (e.g. an invoice form with unsaved edits).
        if self._current_frame_class is not None and self._current_frame_class is not frame_class:
            current_frame = self.frames.get(self._current_frame_class)
            if current_frame is not None and hasattr(current_frame, "confirm_navigate_away"):
                if not current_frame.confirm_navigate_away():
                    return

        frame = self.frames[frame_class]
        frame.tkraise()
        self.sidebar.set_active(frame_class.__name__ if frame_class in
                                (DashboardFrame, ClientsFrame, CatalogFrame) else "DashboardFrame")
        self._current_frame_class = frame_class
        if hasattr(frame, "on_show"):
            frame.on_show(*args)

    def show_frame_by_name(self, name: str):
        mapping = {"DashboardFrame": DashboardFrame, "ClientsFrame": ClientsFrame, "CatalogFrame": CatalogFrame}
        self.show_frame(mapping[name])

    def open_settings(self):
        SettingsDialog(self, self._on_settings_saved)

    def open_about(self):
        AboutDialog(self)

    def open_shortcuts(self):
        ShortcutsDialog(self)

    def _on_settings_saved(self, new_theme=None):
        # 1. LIVE THEME SWITCHING ENGINE
        if new_theme:
            self.style.theme_use(new_theme)
            # ttkbootstrap's theme_use() only re-colors ttk.* widgets --
            # plain tk.Text boxes (Notes, Contact, Payment Notes, Footer,
            # ...) are invisible to that and would otherwise be stuck
            # showing whatever theme was active when they were first
            # built until the app restarts. Walk every already-built
            # screen and fix that up right now instead.
            self._retheme_text_widgets(self)

        # 2. Refresh globals so the currency symbol updates live
        global CONFIG, CURRENCY_SYMBOL
        CONFIG = load_config()
        CURRENCY_SYMBOL = CONFIG.get("currency_symbol", "₹")

        # 3. Force all active screens to redraw with the new colors.
        #    InvoiceFormFrame is special-cased: on_show() resets/reloads the
        #    whole form, which would silently wipe out an in-progress invoice
        #    if the user opened Settings mid-edit without leaving this screen
        #    first. It gets a lightweight refresh instead that only updates
        #    currency/theme-dependent display, never form data.
        for frame in self.frames.values():
            if isinstance(frame, InvoiceFormFrame):
                frame.refresh_after_settings_change()
            elif hasattr(frame, "on_show"):
                try:
                    frame.on_show()
                except TypeError:
                    pass

    def _retheme_text_widgets(self, root_widget: tk.Misc):
        """Recursively finds every tk.Text descendant of root_widget
        (PlaceholderText or otherwise -- e.g. a BulkProgressDialog log,
        if one happens to be open) and re-applies the now-current
        theme's colors to it. See apply_theme_to_text_widget() for why
        this is needed at all."""
        for child in root_widget.winfo_children():
            if isinstance(child, tk.Text):
                apply_theme_to_text_widget(child)
            self._retheme_text_widgets(child)

    # ------------------------------------------------------------------ #
    #  Progress / background task plumbing
    # ------------------------------------------------------------------ #
    def show_progress(self, active: bool, message: str = ""):
        if active:
            self.progress_lbl.config(text=message)
            self.progress_lbl.pack(fill="x", padx=16, pady=(8, 0))
            self.progress_bar.pack(fill="x", padx=16, pady=(2, 8))
            self.progress_bar.start(12)
        else:
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            self.progress_lbl.pack_forget()

    def _poll_queue(self):
        """Drains messages from the single-PDF-view worker (_pdf_worker).
        Bulk operations (delete/export/print) run through their own
        BulkProgressDialog, each with its own dedicated queue, rather
        than this shared one -- keeping a bulk operation's much higher
        message volume from ever competing with, or being confused for,
        an unrelated single-PDF-view result."""
        try:
            while True:
                msg = self.task_queue.get_nowait()
                self.show_progress(False)
                if msg["status"] == "success":
                    PDFViewer(self, msg["path"])
                    self.frames[DashboardFrame].on_show()
                else:
                    Messagebox.show_error("Error", msg.get("message", "An unknown error occurred."))
        except queue.Empty:
            pass
        finally:
            self.after(200, self._poll_queue)

    def _on_close(self):
        frame = self._active_frame()
        if isinstance(frame, InvoiceFormFrame) and not frame.confirm_navigate_away():
            return
        self.destroy()

# =============================================================================
#  BACKGROUND WORKER
# =============================================================================
def startup_background_task(app_instance):
    """
    Put whatever heavy lifting you need to do here.
    Because this runs in a thread, it won't freeze the splash screen or GUI.
    """
    logger.info("Background thread started successfully.")
    
    # Example: Simulating a heavy startup task (like syncing a DB or checking for updates)
    # time.sleep(2) 
    
    # If this thread needs to update the GUI, it MUST use the app's queue!
    # app_instance.task_queue.put({"status": "info", "message": "Background sync complete"})

# =============================================================================
#  ENTRY POINT
# =============================================================================
def main():
    try:
        app = InvoiceApp()
        
        # --- ADD THREADING HERE ---
        # Start your background task as a daemon thread before the mainloop starts
        bg_thread = threading.Thread(
            target=startup_background_task, 
            args=(app,), 
            daemon=True  # Crucial: Ensures the thread dies when the app is closed
        )
        bg_thread.start()
        # --------------------------

        # The mainloop MUST remain on the main thread
        app.mainloop()
        
    except Exception:
        logger.critical("Fatal error during application startup", exc_info=True)
        try:
            import tkinter.messagebox as mb
            mb.showerror("Vyaprix — Fatal Error",
                         "Vyaprix could not start. See the log file for details:\n"
                         f"{LOG_FILE}")
        except Exception:
            pass
        raise

if __name__ == "__main__":
    main()