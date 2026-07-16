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
    "COMPLIMENTARY": {"foreground": "#3B82F6", "background": ""}, # Blue
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
    try:
        # POINT TO THE NEW PROGRAMDATA FOLDER
        reg  = DATA_FONTS_DIR / "NotoSans-Regular.ttf"
        bold = DATA_FONTS_DIR / "NotoSans-Bold.ttf"
        
        if reg.exists() and bold.exists():
            # Remember to keep the str() fix!
            pdfmetrics.registerFont(TTFont("NotoSans",      str(reg)))
            pdfmetrics.registerFont(TTFont("NotoSans-Bold", str(bold)))
            addMapping("NotoSans", 0, 0, "NotoSans")
            addMapping("NotoSans", 1, 0, "NotoSans-Bold")
            return "NotoSans"
    except Exception as exc:
        logger.warning(f"Could not register custom fonts: {exc}")
    return "Helvetica"

PDF_FONT = register_fonts()

# =============================================================================
#  CONFIGURATION
# =============================================================================
def get_default_config() -> Dict[str, Any]:
    return {
        "theme":           "litera",
        "user_details":    {"name": "Your Name/Company",      "contact_info": "email@example.com\n+91 XXXXX XXXXX"},
        "company_details": {"name": APP_NAME, "description": APP_TAGLINE, "gstin": "22AAAAA0000A1Z5"},
        "payment_details": {"upi_id": "yourupi@bank",         "payment_terms": "Due upon receipt"},
        "asset_paths":     {"logo": "assets/app_logo.png",    "signature": "assets/sign.png", "upi_qr_code": "assets/upi_qr.png"},
        "pdf_settings":    {"footer_message": "Generated with Vyaprix • Business Made Smarter."},
        "currency_symbol": "₹",
    }

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
        return f"{CURRENCY_SYMBOL}{float(value or 0):,.2f}"
    except (TypeError, ValueError):
        return f"{CURRENCY_SYMBOL}0.00"

def safe_float(value: Any, fallback: float = 0.0) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return fallback

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
        tk.Label(win, text=self.text, background="#FFFDE7", relief="flat",
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
    _PLACEHOLDER_COLOR = "#9CA3AF"  # muted grey; readable on every bundled theme

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
            self.configure(foreground=self._PLACEHOLDER_COLOR)
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
        return "" if self._showing_placeholder else ttk.Combobox.get(self)

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

    _PLACEHOLDER_COLOR = "#9CA3AF"

    def __init__(self, master=None, placeholder: str = "", **kw):
        tk.Text.__init__(self, master, **kw)
        self._placeholder = placeholder
        self._showing_placeholder = False
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
        self.configure(foreground=self._PLACEHOLDER_COLOR)
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
                c.execute("CREATE INDEX IF NOT EXISTS idx_invoices_status  ON invoices(status)")
                c.execute("CREATE INDEX IF NOT EXISTS idx_invoices_date    ON invoices(invoice_date)")
                c.execute("CREATE INDEX IF NOT EXISTS idx_invoices_billto  ON invoices(bill_to)")
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
        subtotal    = sum(i.get("amount", 0) for i in items)
        item_disc   = sum(i.get("discount", 0) for i in items)
        after_item  = max(0.0, subtotal - item_disc)
        global_disc = after_item * global_disc_val / 100.0 if global_disc_type == "%" else global_disc_val
        global_disc = max(0.0, min(global_disc, after_item))
        total_disc  = item_disc + global_disc
        taxable     = max(0.0, subtotal - total_disc)
        cgst = sgst = taxable * tax_rate / 200.0
        return {
            "subtotal": subtotal, "item_discount": item_disc, "global_discount": global_disc,
            "total_discount": total_disc, "taxable": taxable,
            "cgst": cgst, "sgst": sgst, "grand_total": taxable + cgst + sgst,
        }

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
        balance = grand_total - total_paid
        if grand_total > 0 and total_paid >= grand_total - 0.005:
            status = InvoiceStatus.PAID.value
        elif total_paid > 0:
            status = InvoiceStatus.PARTIAL.value
        else:
            status = InvoiceStatus.DUE.value
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

    def delete_invoices(self, inv_ids: List[str]) -> int:
        with closing(self._conn()) as c:
            with c:
                c.executemany("DELETE FROM invoices WHERE invoice_number=?", [(i,) for i in inv_ids])
        return len(inv_ids)

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
        left  = [Paragraph(safe_markup(comp.get("name", "Your Company")), self.styles["CompanyName"])]
        if d := comp.get("description"):
            left += [Spacer(1, 2), Paragraph(safe_markup(d), self.styles["Tagline"])]

        status    = self.data.get("status")
        doc_title = ("RECEIPT"       if status == InvoiceStatus.PAID.value else
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
        totals     = DatabaseManager.calculate_totals(items_list, tax_rate, g_disc, self.data.get("global_discount_type") or "amount")
        total_paid = sum(safe_float(p.get("amount")) for p in self.payments)
        balance    = totals["grand_total"] - total_paid

        rows = [[Paragraph("Subtotal", self.styles["Right"]),
                 Paragraph(f"{CURRENCY_SYMBOL}{totals['subtotal']:,.2f}", self.styles["Right"])]]
        if totals["total_discount"] > 0:
            rows.append([Paragraph("Discount", self.styles["Right"]),
                         Paragraph(f"- {CURRENCY_SYMBOL}{totals['total_discount']:,.2f}", self.styles["Right"])])

        # --- IGST vs CGST/SGST PDF Rendering Logic ---
        if self.is_gst and tax_rate > 0:
            rows.append([Paragraph("Taxable Amount", self.styles["Right"]),
                         Paragraph(f"{CURRENCY_SYMBOL}{totals['taxable']:,.2f}", self.styles["Right"])])

            if self.data.get("gst_type") == "IGST":
                rows.append([Paragraph(f"IGST ({tax_rate}%)", self.styles["Right"]),
                             Paragraph(f"{CURRENCY_SYMBOL}{totals['cgst']*2:,.2f}", self.styles["Right"])])
            else:
                rows += [
                    [Paragraph(f"CGST ({tax_rate/2}%)", self.styles["Right"]),
                     Paragraph(f"{CURRENCY_SYMBOL}{totals['cgst']:,.2f}", self.styles["Right"])],
                    [Paragraph(f"SGST ({tax_rate/2}%)", self.styles["Right"]),
                     Paragraph(f"{CURRENCY_SYMBOL}{totals['sgst']:,.2f}", self.styles["Right"])],
                ]
        # ---------------------------------------------

        rows.append([Paragraph("<b>Invoice Total</b>", self.styles["Right"]),
                     Paragraph(f"<b>{CURRENCY_SYMBOL}{totals['grand_total']:,.2f}</b>", self.styles["Right"])])
        if total_paid > 0:
            rows.append([Paragraph("Total Paid", self.styles["Right"]),
                         Paragraph(f"- {CURRENCY_SYMBOL}{total_paid:,.2f}", self.styles["Right"])])
        bc = "#059669" if balance <= 0 else "#DC2626"
        rows.append([Paragraph("<b>Balance Due</b>", self.styles["Right"]),
                     Paragraph(f"<font color='{bc}'><b>{CURRENCY_SYMBOL}{balance:,.2f}</b></font>", self.styles["Right"])])
        return Table(rows, colWidths=["75%","25%"], style=TableStyle([
            ("ALIGN",(0,0),(-1,-1),"RIGHT"), ("LEFTPADDING",(0,0),(-1,-1),10),
            ("TOPPADDING",(0,0),(-1,-1),3),  ("BOTTOMPADDING",(0,0),(-1,-1),3),
            ("BACKGROUND",(0,-1),(1,-1),colors.HexColor("#FEF2F2" if balance > 0 else "#ECFDF5")),
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

    def _build_footer_section(self) -> Table:
        qr_cell = sig_cell = Spacer(0, 0)
        if self.data.get("show_qr_code") and not self.data.get("is_complimentary"):
            if qr := self._asset("upi_qr_code"):
                upi = safe_markup(self.config.get("payment_details", {}).get("upi_id", ""))
                detail = Paragraph(f"<b>PAYMENT DETAILS</b><br/>Scan to pay.<br/>UPI ID: {upi}", self.styles["NC"])
                qr_cell = Table([[Image(str(qr), 2.2*cm, 2.2*cm, kind="proportional"), detail]],
                                colWidths=[2.5*cm,"100%"], style=[("VALIGN",(0,0),(-1,-1),"TOP")])
                               

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
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback  = callback
        self.title(f"{APP_NAME} Settings")
        self.geometry("640x760")
        self.minsize(580, 680)
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self._cfg = load_config()
        self._build()
        self._load()
        self.focus_force()

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

    def _file_picker(self, parent, label: str, var: tk.StringVar, dest: str):
        ttk.Label(parent, text=label, font=(FONT_FAMILY, 10, "bold")).pack(anchor="w", pady=(8, 2))
        row = ttk.Frame(parent)
        row.pack(fill="x")
        ttk.Entry(row, textvariable=var, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row, text="Browse…", bootstyle="secondary-outline",
                   command=lambda v=var, d=dest: self._browse(v, d)).pack(side="right")

    def _browse(self, var: tk.StringVar, dest: str):
        if path := filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")]):
            try:
                shutil.copy(path, ASSETS_DIR / dest)
                var.set(f"assets/{dest}")
                Messagebox.show_info("Done", "Asset saved successfully.")
            except Exception as exc:
                Messagebox.show_error("Error", f"Could not save file: {exc}")

    def _build(self):
        main = ttk.Frame(self, padding=20)
        main.pack(fill="both", expand=True)

        nb = ttk.Notebook(main)
        nb.pack(fill="both", expand=True, pady=(0, 16))

        t1 = ttk.Frame(nb, padding=16)
        nb.add(t1, text="  General  ")
        self._section(t1, "Interface")
        self.theme_var = tk.StringVar()
        self._labeled_combo(t1, "App Theme:", self.theme_var,
                            ["litera","cosmo","flatly","darkly","cyborg","superhero","morph","pulse"])
        ttk.Separator(t1, bootstyle="secondary").pack(fill="x", pady=12)
        self._section(t1, "Company Details")
        self.co_name = self._labeled_entry(t1, "Company Name:",    "Your company name")
        self.co_desc = self._labeled_entry(t1, "Company Tagline:", "A short tagline, e.g. Business Made Smarter.")
        self.co_gst  = self._labeled_entry(t1, "Company GSTIN:",   "22ABCDE1234F1Z5")
        ttk.Separator(t1, bootstyle="secondary").pack(fill="x", pady=12)
        self._section(t1, "Your / User Details")
        self.user_name = self._labeled_entry(t1, "Your Name / Firm:", "Your name or firm")
        ttk.Label(t1, text="Contact Info:", font=(FONT_FAMILY, 9)).pack(anchor="w", pady=(5, 2))
        self.contact_text = PlaceholderText(t1, placeholder="+91 98765 43210, name@email.com, City",
                                             height=4, relief="flat", highlightthickness=1, wrap="word")
        self.contact_text.pack(fill="x")

        t2 = ttk.Frame(nb, padding=16)
        nb.add(t2, text="  Payment & PDF  ")
        self._section(t2, "Payment Settings")
        self.currency   = self._labeled_entry(t2, "Currency Symbol:", "₹")
        self.upi_id     = self._labeled_entry(t2, "UPI ID:",          "company@upi")
        self.pay_terms  = self._labeled_entry(t2, "Default Terms:",   "Due on Receipt")
        ttk.Separator(t2, bootstyle="secondary").pack(fill="x", pady=12)
        self._section(t2, "PDF Footer Note")
        self.footer_text = PlaceholderText(t2, placeholder="Thank you for your business!",
                                            height=3, relief="flat", highlightthickness=1, wrap="word")
        self.footer_text.pack(fill="x")

        t3 = ttk.Frame(nb, padding=16)
        nb.add(t3, text="  Images & Assets  ")
        self._section(t3, "Upload Assets (PNG / JPG)")
        ttk.Label(
            t3,
            text=f"Files are copied to {ASSETS_DIR} and embedded in PDFs.",
            font=(FONT_FAMILY, 9),
            bootstyle="secondary",
            wraplength=540,
        ).pack(anchor="w", pady=(0, 12))
        self.logo_var = tk.StringVar(); self._file_picker(t3, "Company Logo:",         self.logo_var, "app_logo.png")
        self.sig_var  = tk.StringVar(); self._file_picker(t3, "Authorized Signature:", self.sig_var,  "sign.png")
        self.qr_var   = tk.StringVar(); self._file_picker(t3, "UPI QR Code:",          self.qr_var,   "upi_qr.png")

        t4 = ttk.Frame(nb, padding=16)
        nb.add(t4, text="  Data  ")
        self._section(t4, "Database Location")
        ttk.Label(t4, text=str(DB_NAME), font=(FONT_FAMILY, 9), bootstyle="secondary", wraplength=540).pack(anchor="w")
        ttk.Separator(t4, bootstyle="secondary").pack(fill="x", pady=12)
        self._section(t4, "Backup")
        ttk.Button(t4, text="💾  Save a Backup Copy…", bootstyle="secondary-outline",
                   command=self._backup_db).pack(anchor="w")

        btn_row = ttk.Frame(main)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Cancel",             bootstyle="secondary", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(btn_row, text="Save Configuration", bootstyle="success",   command=self._save).pack(side="right")

    def _backup_db(self):
        if path := filedialog.asksaveasfilename(
            initialfile=f"Vyaprix-backup-{datetime.date.today().isoformat()}.db",
            defaultextension=".db", filetypes=[("SQLite Database", "*.db")],
        ):
            try:
                shutil.copy(DB_NAME, path)
                Messagebox.show_info("Backup Saved", f"Database backed up to:\n{path}")
            except Exception as exc:
                Messagebox.show_error("Backup Failed", str(exc))

    def _load(self):
        c = self._cfg
        self.theme_var.set(c.get("theme", "litera"))
        self.co_name.insert(0, c.get("company_details", {}).get("name", ""))
        self.co_desc.insert(0, c.get("company_details", {}).get("description", ""))
        self.co_gst.insert(0, c.get("company_details", {}).get("gstin", ""))
        self.user_name.insert(0, c.get("user_details", {}).get("name", ""))
        self.contact_text.insert("1.0", c.get("user_details", {}).get("contact_info", ""))
        self.currency.insert(0, c.get("currency_symbol", "₹"))
        self.upi_id.insert(0, c.get("payment_details", {}).get("upi_id", ""))
        self.pay_terms.insert(0, c.get("payment_details", {}).get("payment_terms", ""))
        self.footer_text.insert("1.0", c.get("pdf_settings", {}).get("footer_message", ""))
        self.logo_var.set(c.get("asset_paths", {}).get("logo", ""))
        self.sig_var.set(c.get("asset_paths", {}).get("signature", ""))
        self.qr_var.set(c.get("asset_paths", {}).get("upi_qr_code", ""))

    def _save(self):
        save_config({
            "theme": self.theme_var.get(),
            "company_details": {
                "name": self.co_name.get().strip(),
                "description": self.co_desc.get().strip(),
                "gstin": self.co_gst.get().strip(),
            },
            "user_details": {
                "name": self.user_name.get().strip(),
                "contact_info": self.contact_text.get("1.0", "end-1c").strip(),
            },
            "payment_details": {
                "upi_id": self.upi_id.get().strip(),
                "payment_terms": self.pay_terms.get().strip(),
            },
            "pdf_settings": {"footer_message": self.footer_text.get("1.0", "end-1c").strip()},
            "asset_paths": {
                "logo": self.logo_var.get(),
                "signature": self.sig_var.get(),
                "upi_qr_code": self.qr_var.get(),
            },
            "currency_symbol": self.currency.get().strip() or "₹",
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
        self.protocol("WM_DELETE_WINDOW", self._close)
        self.transient(parent)
        self.grab_set()
        self.after(120, self._render)

    def _build(self):
        tb = ttk.Frame(self, padding=(10, 7))
        tb.pack(side="top", fill="x")

        ttk.Button(tb, text="◀  Prev",  command=self._prev,         width=9).pack(side="left", padx=(0, 4))
        ttk.Button(tb, text="Next  ▶",  command=self._next,         width=9).pack(side="left", padx=(0, 14))
        self.page_lbl = ttk.Label(tb, text="", font=(FONT_FAMILY, 10))
        self.page_lbl.pack(side="left")

        ttk.Button(tb, text="✕  Close",  bootstyle="danger",    command=self._close,      width=9).pack(side="right", padx=(4, 0))
        ttk.Button(tb, text="💾  Save",   bootstyle="secondary", command=self._save_copy,  width=9).pack(side="right", padx=4)
        ttk.Button(tb, text="🖨  Print",  bootstyle="success",   command=self._print,      width=9).pack(side="right", padx=4)

        ttk.Separator(tb, orient="vertical").pack(side="right", fill="y", padx=10)
        self.zoom_lbl = ttk.Label(tb, text="100%", font=(FONT_FAMILY, 9), width=6, anchor="center")
        self.zoom_lbl.pack(side="right")
        ttk.Button(tb, text="⊕", command=lambda: self._zoom(+0.15), width=3, bootstyle="secondary-outline").pack(side="right", padx=2)
        ttk.Button(tb, text="⊖", command=lambda: self._zoom(-0.15), width=3, bootstyle="secondary-outline").pack(side="right", padx=2)
        ttk.Button(tb, text="Fit",command=self._reset_zoom,        width=4, bootstyle="secondary-outline").pack(side="right", padx=2)

        ttk.Separator(self, bootstyle="secondary").pack(fill="x")
        self.canvas = tk.Canvas(self, bg="#555555", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        ttk.Label(self, text="← → or Page Up/Down to navigate  •  Ctrl + / − zoom  •  Ctrl+0 fit",
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
#  DASHBOARD  /  INVOICE LIST (IMPROVED)
# =============================================================================
class DashboardFrame(ttk.Frame):
    STATUS_FILTERS = ["All", "Paid", "Due", "Partial", "Complimentary"]
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

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller     = controller
        self.db_manager     = controller.db_manager
        self._status_filter = "All"
        self._sort_col      = "date"
        self._sort_rev      = True
        self._filter_btns: Dict[str, ttk.Button] = {}
        self._row_index: Dict[str, Dict] = {}   
        self._build()
        self._debounce_search = Debouncer(self, 200, self._populate)

    def _build(self):
        outer = ttk.Frame(self, padding=(24, 16, 24, 14))
        outer.pack(fill="both", expand=True)
        self._header(outer)
        self._stat_cards(outer)
        self._filter_bar(outer)
        self._table(outer)
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
        ttk.Button(right, text="📊  Export CSV", bootstyle="info-outline", command=self._export_csv).pack(side="left", padx=4)
        ttk.Button(right, text="＋  New Invoice", bootstyle="primary", command=lambda: self.controller.show_frame(InvoiceFormFrame)).pack(side="left", padx=(4, 0))

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
            btn = ttk.Button(
                filter_group, text=status, width=13,
                bootstyle="primary" if status == "All" else "secondary-outline",
                command=lambda s=status: self._set_filter(s),
            )
            btn.pack(side="left", padx=2)
            self._filter_btns[status] = btn

        srch = ttk.Frame(row)
        srch.pack(side="right")
        ttk.Label(srch, text="🔍", font=(FONT_FAMILY, 12)).pack(side="left", padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._debounce_search())
        
        entry = ttk.Entry(srch, textvariable=self.search_var, width=28)
        entry.pack(side="left")
        ToolTip(entry, "Search by invoice number, client name, or note")
        ttk.Button(srch, text="✕", width=3, bootstyle="secondary-outline", command=lambda: self.search_var.set("")).pack(side="left", padx=(4, 0))

    def _set_filter(self, status: str):
        self._status_filter = status
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
            
        # Update column headers to show the correct arrow and clear others
        for cid, text, _, _ in self.COL_DEFS:
            if cid == col:
                arrow = self.SORT_DESC if self._sort_rev else self.SORT_ASC
                self.tree.heading(cid, text=f"{text}{arrow}")
            else:
                self.tree.heading(cid, text=text) 
                
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

        # -- Bulk-selection actions --
        self.bulk_del_btn    = ttk.Button(row, text="🗑  Bulk Delete", bootstyle="danger-outline",    command=self._act_delete)
        self.bulk_export_btn = ttk.Button(row, text="📊  Bulk Export", bootstyle="secondary-outline",  command=self._act_bulk_export)
        self.bulk_print_btn  = ttk.Button(row, text="🖨  Bulk Print",  bootstyle="primary",            command=self._act_bulk_print)

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
        
        rows = self.db_manager.get_all_invoices(
            search_query=self.search_var.get().strip(),
            status_filter=self._status_filter,
        )
        
        entries = []
        for inv in rows:
            row = dict(inv)
            is_comp    = bool(row.get("is_complimentary"))
            status_val = (row["status"] or "Due").upper()
            date_val   = parse_date(row["invoice_date"], DB_DATE_FORMAT, DISPLAY_DATE_FORMAT)
            total_txt  = "No Charge" if is_comp else fmt_money(row.get("grand_total"))
            paid_txt   = "-"         if is_comp else fmt_money(row.get("total_paid"))
            bal_txt    = "-"         if is_comp else fmt_money(row.get("balance_due"))
            note_icon  = "📝" if (row.get("notes") or "").strip() else ""
            
            values = (row["invoice_number"], row["bill_to"], date_val, status_val,
                      total_txt, paid_txt, bal_txt, note_icon)
            
            # Robust Sort Keys - ensuring safe fallbacks for missing data
            sort_keys = {
                "id":      row.get("invoice_number", ""),
                "client":  row.get("bill_to", ""),
                "date":    row.get("invoice_date") or "0000-00-00", 
                "status":  status_val,
                "total":   float(row.get("grand_total") or 0.0), 
                "paid":    float(row.get("total_paid") or 0.0),
                "balance": float(row.get("balance_due") or 0.0), 
                "note":    note_icon,
            }
            entries.append((values, status_val, sort_keys, row))

        # Perform the actual sort based on the chosen column
        entries.sort(key=lambda e: e[2].get(self._sort_col, ""), reverse=self._sort_rev)

        # Insert rows back into the tree
        for values, tag, _keys, raw in entries:
            badge = tag if tag in STATUS_BADGE else ""
            iid = raw["invoice_number"]
            self.tree.insert("", "end", iid=iid, values=values, tags=(badge,) if badge else ())
            self._row_index[iid] = raw
            
        # Restore selection
        for item in current_selection:
            if self.tree.exists(item):
                self.tree.selection_add(item)
                
        self._on_select() # Re-evaluate button states

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
        ids = self._selected_ids()
        if not ids:
            return
        if len(ids) == 1:
            msg = f"Permanently delete Invoice #{ids[0]}?\nThis cannot be undone."
        else:
            msg = f"Permanently delete {len(ids)} invoices?\nThis cannot be undone."
            
        if Messagebox.okcancel("Delete Invoice" if len(ids) == 1 else "Bulk Delete", msg):
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
    #  Bulk actions
    # ------------------------------------------------------------------ #
    def _act_bulk_export(self):
        ids = self._selected_ids()
        if not ids:
            return
        if path := filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")],
                                                 title="Export Selected Invoices"):
            ok, msg = self.db_manager.export_to_csv(path, invoice_ids=ids)
            Messagebox.show_info("Exported", f"Saved {len(ids)} invoice(s) to:\n{path}") if ok else Messagebox.show_error("Failed", msg)

    def _act_bulk_print(self):
        ids = self._selected_ids()
        if not ids:
            return
        self.controller.show_progress(True, f"Generating {len(ids)} PDFs…")
        threading.Thread(target=self._bulk_pdf_worker, args=(ids, self.controller.task_queue), daemon=True).start()

    def _show_ctx_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row and row not in self.tree.selection():
            self.tree.selection_set(row)
        if len(self.tree.selection()) >= 2:
            self._ctx_bulk.post(event.x_root, event.y_root)
        elif row:
            self._ctx_single.post(event.x_root, event.y_root)

    def _export_csv(self):
        if path := filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")], title="Export All Invoices"):
            ok, msg = self.db_manager.export_to_csv(path)
            Messagebox.show_info("Exported", f"Saved to:\n{path}") if ok else Messagebox.show_error("Failed", msg)

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

    def _bulk_pdf_worker(self, invoice_ids: List[str], q: queue.Queue):
        try:
            cfg = load_config()
            paths = []
            for inv_id in invoice_ids:
                row = dict(self.db_manager.get_invoice_by_id(inv_id))
                payments = self.db_manager.get_payments(inv_id)
                paths.append(PDFGenerator(row, cfg, payments=payments).generate())
            stamp  = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
            merged = merge_pdfs(paths, BATCH_DIR / f"Batch-Print-{stamp}.pdf")
            q.put({"status": "success", "kind": "batch", "path": merged, "count": len(paths)})
        except Exception as exc:
            logger.error(f"Bulk PDF generation failed: {exc}", exc_info=True)
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
        ttk.Button(row, text="＋  Add Client", bootstyle="primary",
                   command=lambda: self._open_editor(None)).pack(side="right")

        srch = ttk.Frame(outer); srch.pack(fill="x", pady=(0, 8))
        ttk.Label(srch, text="🔍", font=(FONT_FAMILY, 12)).pack(side="left", padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._debounce_search())
        ttk.Entry(srch, textvariable=self.search_var, width=32).pack(side="left")

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
        cid = int(sel[0])
        for c in self.db_manager.get_all_clients():
            if c["id"] == cid:
                return c
        return None

    def _open_editor(self, client: Optional[Dict]):
        ClientDialog(self.controller, self.db_manager, client, self._populate)

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
        ttk.Button(row, text="＋  Add Item", bootstyle="primary",
                   command=lambda: self._open_editor(None)).pack(side="right")

        srch = ttk.Frame(outer); srch.pack(fill="x", pady=(0, 8))
        ttk.Label(srch, text="🔍", font=(FONT_FAMILY, 12)).pack(side="left", padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._debounce_search())
        ttk.Entry(srch, textvariable=self.search_var, width=32).pack(side="left")

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
        iid = int(sel[0])
        for it in self.db_manager.get_all_catalog_items():
            if it["id"] == iid:
                return it
        return None

    def _open_editor(self, item: Optional[Dict]):
        CatalogItemDialog(self.controller, self.db_manager, item, self._populate)

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
        ttk.Button(header, text="←  Back to Dashboard", bootstyle="secondary-outline",
                   command=lambda: self.controller.show_frame(DashboardFrame)).pack(side="left")
        self.title_lbl = ttk.Label(header, text="New Invoice", font=(FONT_FAMILY, 20, "bold"))
        self.title_lbl.pack(side="left", padx=20)
        self.status_badge = ttk.Label(header, text="", font=(FONT_FAMILY, 10, "bold"), padding=(10, 3))
        self.status_badge.pack(side="left")

        self.scroller = ScrolledFrame(self, autohide=True)
        self.scroller.pack(fill="both", expand=True, padx=24, pady=(0, 10))
        body = self.scroller

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
        
        ttk.Label(left, text="Contact Info", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.client_contact_text = PlaceholderText(left, placeholder="Phone, email, address...",
                                                     height=3, relief="flat", highlightthickness=1, wrap="word")
        self.client_contact_text.pack(fill="x", pady=(2, 8))
        ttk.Label(left, text="GSTIN", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.client_gstin_entry = PlaceholderEntry(left, placeholder="22ABCDE1234F1Z5 (optional)")
        self.client_gstin_entry.pack(fill="x", pady=(2, 0))

        right = ttk.Frame(wrap, padding=16)
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        ttk.Label(right, text="From", font=(FONT_FAMILY, 12, "bold")).pack(anchor="w", pady=(0, 10))
        ttk.Label(right, text="Your Name / Firm *", font=(FONT_FAMILY, 9)).pack(anchor="w")
        self.your_name_entry = PlaceholderEntry(right, placeholder="Your business / firm name")
        self.your_name_entry.pack(fill="x", pady=(2, 8))
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

        r2 = ttk.Frame(left); r2.pack(fill="x", pady=4)
        ttk.Label(r2, text="Invoice Date *", width=20, anchor="w").pack(side="left")
        self.invoice_date_entry = DateEntry(r2, dateformat=DISPLAY_DATE_FORMAT)
        self.invoice_date_entry.pack(side="left", fill="x", expand=True)

        r3 = ttk.Frame(left); r3.pack(fill="x", pady=4)
        ttk.Label(r3, text="Payment Terms", width=20, anchor="w").pack(side="left")
        self.payment_terms_entry = PlaceholderEntry(r3, placeholder="e.g. Due on Receipt, Net 15")
        self.payment_terms_entry.pack(side="left", fill="x", expand=True)

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
        self.items_container = ttk.Frame(wrap)
        self.items_container.pack(fill="x")
        ttk.Button(wrap, text="＋  Add Item", bootstyle="secondary-outline",
                  command=lambda: self._add_item_row()).pack(anchor="w", pady=(8, 0))

    def _add_item_row(self, title: str = "", description: str = "", amount: float = 0.0, discount: float = 0.0):
        row_frame = ttk.Frame(self.items_container, padding=(0, 6))
        row_frame.pack(fill="x")

        top = ttk.Frame(row_frame); top.pack(fill="x")
        title_var = tk.StringVar(value=title)
        title_combo = PlaceholderCombobox(top, placeholder="Item / service name", textvariable=title_var,
                                           values=[c["title"] for c in self._catalog_cache], height=6)
        title_combo.pack(side="left", fill="x", expand=True)
        self._prevent_scroll_hijack(title_combo)
        title_combo.bind("<<ComboboxSelected>>", lambda e, tc=title_combo: self._on_catalog_picked(tc))
        title_combo.bind("<KeyRelease>", lambda e, tc=title_combo: self._filter_catalog_suggestions(tc))

        amount_entry = ttk.Entry(top, width=12)
        amount_entry.insert(0, f"{amount:.2f}")
        amount_entry.pack(side="left", padx=6)
        amount_entry.bind("<KeyRelease>", lambda _: self._recalc_totals())

        disc_entry = ttk.Entry(top, width=10)
        disc_entry.insert(0, f"{discount:.2f}")
        disc_entry.pack(side="left", padx=(0, 6))
        disc_entry.bind("<KeyRelease>", lambda _: self._recalc_totals())
        ToolTip(disc_entry, "Per-item discount amount")

        remove_btn = ttk.Button(top, text="✕", width=3, bootstyle="danger-outline")
        remove_btn.pack(side="left")

        desc_entry = PlaceholderEntry(row_frame, placeholder="Description (optional)")
        desc_entry.insert(0, description)
        desc_entry.pack(fill="x", pady=(4, 0))

        ttk.Separator(row_frame, bootstyle="secondary").pack(fill="x", pady=(8, 0))

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
        self.payments_container = ttk.Frame(wrap)
        self.payments_container.pack(fill="x")
        self.add_payment_btn = ttk.Button(wrap, text="＋  Add Payment", bootstyle="secondary-outline",
                                          command=lambda: self._add_payment_row())
        self.add_payment_btn.pack(anchor="w", pady=(8, 0))

    def _add_payment_row(self, amount: float = 0.0, payment_date: Optional[str] = None,
                          mode: str = "Cash", note: str = "", payment_id: Optional[int] = None):
        row_frame = ttk.Frame(self.payments_container, padding=(0, 5))
        row_frame.pack(fill="x")

        date_entry = DateEntry(row_frame, dateformat=DISPLAY_DATE_FORMAT, width=12)
        if payment_date:
            try:
                date_entry.entry.delete(0, "end")
                date_entry.entry.insert(0, parse_date(payment_date, DB_DATE_FORMAT, DISPLAY_DATE_FORMAT))
            except Exception:
                pass
        date_entry.pack(side="left", padx=(0, 6))

        amount_entry = ttk.Entry(row_frame, width=12)
        amount_entry.insert(0, f"{amount:.2f}")
        amount_entry.pack(side="left", padx=(0, 6))
        amount_entry.bind("<KeyRelease>", lambda _: self._recalc_totals())

        mode_combo = ttk.Combobox(row_frame, values=PAYMENT_MODES, state="readonly", width=13)
        mode_combo.set(mode if mode in PAYMENT_MODES else "Cash")
        mode_combo.pack(side="left", padx=(0, 6))
        self._prevent_scroll_hijack(mode_combo)

        note_entry = PlaceholderEntry(row_frame, placeholder="Note (optional)")
        note_entry.insert(0, note)
        note_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))

        remove_btn = ttk.Button(row_frame, text="✕", width=3, bootstyle="danger-outline")
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
        self.var_tax_lbl     = tk.StringVar(value="Tax:")
        self.var_tax_val     = tk.StringVar(value="")
        self.var_grand_total = tk.StringVar(value="0.00")
        self.var_paid        = tk.StringVar(value="")
        self.var_balance     = tk.StringVar(value="0.00")

        # Helper to create consistent rows
        def add_row(lbl_text, var_val, is_bold=False, is_color=None):
            row = ttk.Frame(self.summary_box)
            font_style = (FONT_FAMILY, 11, "bold") if is_bold else (FONT_FAMILY, 10)
            
            lbl = ttk.Label(row, text=lbl_text, width=25, anchor="e", font=font_style)
            lbl.pack(side="left", padx=(0, 15))
            
            val_lbl = ttk.Label(row, textvariable=var_val, width=15, anchor="e", font=font_style)
            val_lbl.pack(side="right")
            if is_color: val_lbl.configure(bootstyle=is_color)
            return row

        # Build Rows and save explicit references
        self.row_subtotal = add_row("Subtotal:", self.var_subtotal)
        self.row_disc = add_row("Discount:", self.var_discount)
        
        self.row_tax = ttk.Frame(self.summary_box)
        ttk.Label(self.row_tax, textvariable=self.var_tax_lbl, width=25, anchor="e", font=(FONT_FAMILY, 10)).pack(side="left", padx=(0, 15))
        ttk.Label(self.row_tax, textvariable=self.var_tax_val, width=15, anchor="e", font=(FONT_FAMILY, 10)).pack(side="right")

        self.row_grand_total = add_row("Grand Total:", self.var_grand_total, is_bold=True)
        self.row_paid = add_row("Total Paid:", self.var_paid)
        self.row_balance = add_row("Balance Due:", self.var_balance, is_bold=True, is_color="danger")

    def _section_save_bar(self, parent):
        row = ttk.Frame(parent, padding=(0, 4, 0, 24))
        row.pack(fill="x")
        ttk.Button(row, text="Cancel", bootstyle="secondary",
                  command=lambda: self.controller.show_frame(DashboardFrame)).pack(side="right", padx=(8, 0))
        ttk.Button(row, text="💾  Save Invoice", bootstyle="success", command=self._save).pack(side="right")

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
        """Calculates live totals and instantly updates the UI grid."""
        if self.no_charge_var.get():
            self.var_subtotal.set("0.00")
            self.var_grand_total.set("0.00")
            self.var_balance.set("0.00")
            
            # Repack strictly the relevant rows
            self.row_subtotal.pack(fill="x", pady=2)
            self.row_disc.pack_forget()
            self.row_tax.pack_forget()
            self.row_grand_total.pack(fill="x", pady=2)
            self.row_paid.pack_forget()
            self.row_balance.pack(fill="x", pady=2)
            
            self.status_badge.config(text="COMPLIMENTARY", bootstyle="info-inverse")
            return
            
        items    = self._collect_items()
        tax_rate = safe_float(self.tax_rate_entry.get()) if self.gst_var.get() else 0.0
        g_disc   = safe_float(self.global_discount_entry.get())
        
        totals   = DatabaseManager.calculate_totals(items, tax_rate, g_disc, self.global_discount_type_var.get())
        paid     = sum(safe_float(r["amount_entry"].get()) for r in self.payment_rows)
        balance  = totals["grand_total"] - paid

        self.var_subtotal.set(fmt_money(totals['subtotal']))
        self.var_grand_total.set(fmt_money(totals['grand_total']))
        self.var_balance.set(fmt_money(balance))

        # Pack in chronological order to avoid layout bugs
        self.row_subtotal.pack(fill="x", pady=2)

        # 1. Handle Discount Row
        if totals["total_discount"] > 0:
            self.var_discount.set(f"− {fmt_money(totals['total_discount'])}")
            self.row_disc.pack(fill="x", pady=2)
        else:
            self.row_disc.pack_forget()

        # 2. Handle GST Row
        if self.gst_var.get() and tax_rate > 0:
            if self.gst_type_var.get() == "IGST":
                self.var_tax_lbl.set(f"IGST ({tax_rate:g}%):")
                self.var_tax_val.set(fmt_money(totals['cgst'] * 2))
            else:
                self.var_tax_lbl.set(f"GST ({tax_rate:g}% - CGST/SGST):")
                self.var_tax_val.set(fmt_money(totals['cgst'] + totals['sgst']))
            self.row_tax.pack(fill="x", pady=2)
        else:
            self.row_tax.pack_forget()

        self.row_grand_total.pack(fill="x", pady=2)

        # 3. Handle Paid Row
        if paid > 0:
            self.var_paid.set(f"− {fmt_money(paid)}")
            self.row_paid.pack(fill="x", pady=2)
        else:
            self.row_paid.pack_forget()

        self.row_balance.pack(fill="x", pady=2)

        # Status Logic mapping safely to ttkbootstrap styles
        if totals["grand_total"] > 0 and paid >= totals["grand_total"] - 0.005:
            status = "PAID"
            b_style = "success-inverse"
        elif paid > 0:
            status = "PARTIAL"
            b_style = "warning-inverse"
        else:
            status = "DUE"
            b_style = "danger-inverse"
            
        self.status_badge.config(text=status, bootstyle=b_style)

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
            self.invoice_number_entry.insert(0, self.db_manager.get_next_invoice_number())
            self.invoice_date_entry.entry.delete(0, "end")
            self.invoice_date_entry.entry.insert(0, datetime.date.today().strftime(DISPLAY_DATE_FORMAT))
            self.payment_terms_entry.insert(0, CONFIG.get("payment_details", {}).get("payment_terms", ""))
            self.your_name_entry.insert(0, CONFIG.get("user_details", {}).get("name", ""))
            self.your_contact_text.insert("1.0", CONFIG.get("user_details", {}).get("contact_info", ""))
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
        self.border_frame = tk.Frame(self, bg="#3b82f6", bd=0) # Match to your primary theme color
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

class InvoiceApp(ttk.Window):
    def __init__(self):
        cfg = load_config()
        super().__init__(themename=cfg.get("theme", "litera"))
        configure_windows_app_identity()
        
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

    def _on_settings_saved(self, new_theme=None):
        # 1. LIVE THEME SWITCHING ENGINE
        if new_theme:
            self.style.theme_use(new_theme)
        
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
        try:
            while True:
                msg = self.task_queue.get_nowait()
                self.show_progress(False)
                if msg["status"] == "success":
                    if msg.get("kind") == "batch":
                        Messagebox.show_info("Batch Ready", f"Generated {msg['count']} invoice PDF(s).")
                    PDFViewer(self, msg["path"])
                    self.frames[DashboardFrame].on_show()
                else:
                    Messagebox.show_error("Error", msg.get("message", "An unknown error occurred."))
        except queue.Empty:
            pass
        finally:
            self.after(200, self._poll_queue)

    def _on_close(self):
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