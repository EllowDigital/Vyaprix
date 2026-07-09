# =============================================================================
#  BillFlow By EllowLabs  –  Professional Edition v2.5
#  Complete rewrite with enhanced UI/UX, performance, and code quality.
# =============================================================================
import tkinter as tk
from tkinter import filedialog
import ttkbootstrap as ttk
from ttkbootstrap.dialogs import Messagebox
from ttkbootstrap.widgets.scrolled import ScrolledFrame
from ttkbootstrap.widgets import DateEntry
import sqlite3
import json
import os
import sys
import subprocess
import datetime
import threading
import queue
import logging
import csv
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple, Type
import shutil
from enum import Enum

# --- Windows printing support ---
try:
    if sys.platform == "win32":
        import win32api
        import win32print
except ImportError:
    win32api = win32print = None

# --- PyMuPDF ---
try:
    import fitz
except ImportError:
    print("PyMuPDF not installed.  Run:  pip install PyMuPDF")
    sys.exit(1)

# --- ReportLab ---
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.fonts import addMapping


# =============================================================================
#  CONSTANTS
# =============================================================================
BASE_DIR      = Path(__file__).resolve().parent
DB_NAME       = BASE_DIR / "invoices.db"
CONFIG_FILE   = BASE_DIR / "config.json"
LOG_FILE      = BASE_DIR / "invoice_log.txt"
ASSETS_DIR    = BASE_DIR / "assets"
INVOICES_DIR  = BASE_DIR / "invoices"

for _dir in [ASSETS_DIR, INVOICES_DIR]:
    _dir.mkdir(exist_ok=True)

FONT_FAMILY          = "Segoe UI"
APP_NAME          = "BillFlow by EllowLabs"
APP_VERSION       = "v2.5"
APP_BUILD_VERSION = "2.5.0"
APP_RELEASE_DATE  = "22 June 2026"

APP_TAGLINE = "Invoices That Work for You"

APP_ATTRIBUTION = (
    "Designed and Developed by Sarwan Yadav\n"
    "at EllowLabs\n"
    "© 2026 EllowLabs. All Rights Reserved."
)
DB_DATE_FORMAT       = "%Y-%m-%d"
DISPLAY_DATE_FORMAT  = "%d %b %Y"
PAYMENT_MODES        = ['Cash', 'UPI', 'Bank Transfer', 'Card', 'Cheque']
APP_LOGO_PNG         = ASSETS_DIR / "app_logo.png"
APP_LOGO_ICO         = ASSETS_DIR / "app_logo.ico"

# Treeview row badge colours per status
STATUS_BADGE = {
    "PAID":          {"foreground": "#ffffff", "background": "#059669"},
    "DUE":           {"foreground": "#ffffff", "background": "#DC2626"},
    "PARTIAL":       {"foreground": "#ffffff", "background": "#D97706"},
    "COMPLIMENTARY": {"foreground": "#ffffff", "background": "#2563EB"},
}


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
        logging.FileHandler(LOG_FILE, mode='a', encoding='utf-8'),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)


# =============================================================================
#  FONTS
# =============================================================================
def register_fonts() -> str:
    try:
        fonts_dir = BASE_DIR / "fonts"
        reg  = fonts_dir / "NotoSans-Regular.ttf"
        bold = fonts_dir / "NotoSans-Bold.ttf"
        if reg.exists() and bold.exists():
            pdfmetrics.registerFont(TTFont("NotoSans",      reg))
            pdfmetrics.registerFont(TTFont("NotoSans-Bold", bold))
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
        "asset_paths":     {"logo": "assets/logo.png",        "signature": "assets/sign.png", "upi_qr_code": "assets/upi_qr.png"},
        "pdf_settings":    {"footer_message": "Thank you for your business!"},
        "currency_symbol": "₹",
    }


def load_config() -> Dict[str, Any]:
    if not CONFIG_FILE.exists():
        _cfg = get_default_config()
        _save_config_raw(_cfg)
        return _cfg
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
            return json.load(fh)
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


def load_app_logo_image(master: tk.Misc) -> Optional[tk.PhotoImage]:
    if APP_LOGO_PNG.exists():
        try:
            return tk.PhotoImage(master=master, file=str(APP_LOGO_PNG))
        except Exception as exc:
            logger.warning(f"Could not load app logo image: {exc}")
    return None


class ToolTip:
    """Lightweight hover tooltip."""

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


# =============================================================================
#  DATABASE
# =============================================================================
class DatabaseManager:

    def __init__(self, db_path: Path):
        self.db_path = db_path
        self._init_db()

    # ------------------------------------------------------------------  connection
    def _conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    # ------------------------------------------------------------------  init
    def _init_db(self):
        with self._conn() as c:
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
            existing = {r["name"] for r in c.execute("PRAGMA table_info(invoices)").fetchall()}
            for col, ctype in {"advance_amount": "REAL DEFAULT 0", "advance_date": "TEXT"}.items():
                if col not in existing:
                    c.execute(f"ALTER TABLE invoices ADD COLUMN {col} {ctype}")
            c.commit()

    # ------------------------------------------------------------------  queries
    def get_all_invoices(self, search_query: str = "", status_filter: str = "All") -> List[sqlite3.Row]:
        clauses, params = [], []
        if search_query:
            clauses.append("(invoice_number LIKE ? OR bill_to LIKE ?)")
            q = f"%{search_query}%"
            params += [q, q]
        if status_filter and status_filter != "All":
            clauses.append("UPPER(status) = ?")
            params.append(status_filter.upper())
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        with self._conn() as c:
            return c.execute(
                f"SELECT * FROM invoices {where} ORDER BY invoice_date DESC, invoice_number DESC",
                params,
            ).fetchall()

    def get_invoice_by_id(self, inv_id: str) -> Optional[sqlite3.Row]:
        with self._conn() as c:
            return c.execute("SELECT * FROM invoices WHERE invoice_number=?", (inv_id,)).fetchone()

    def save_invoice(self, data: Dict[str, Any], is_update: bool,
                     original_id: Optional[str] = None) -> Tuple[bool, Optional[str]]:
        try:
            with self._conn() as c:
                if is_update:
                    nid = data["invoice_number"]
                    if nid != original_id and c.execute(
                        "SELECT 1 FROM invoices WHERE invoice_number=?", (nid,)
                    ).fetchone():
                        return False, f"Invoice number '{nid}' already exists."
                    qd = {**data, "original_id": original_id}
                    c.execute(
                        f"UPDATE invoices SET {', '.join(f'{k}=:{k}' for k in data)} WHERE invoice_number=:original_id",
                        qd,
                    )
                else:
                    if c.execute(
                        "SELECT 1 FROM invoices WHERE invoice_number=?", (data["invoice_number"],)
                    ).fetchone():
                        return False, f"Invoice number '{data['invoice_number']}' already exists."
                    c.execute(
                        f"INSERT INTO invoices ({', '.join(data)}) VALUES (:{', :'.join(data)})",
                        data,
                    )
                c.commit()
            return True, None
        except sqlite3.Error as exc:
            return False, str(exc)

    def delete_invoice(self, inv_id: str) -> bool:
        with self._conn() as c:
            c.execute("DELETE FROM invoices WHERE invoice_number=?", (inv_id,))
            c.commit()
        return True

    def get_next_invoice_number(self) -> str:
        with self._conn() as c:
            last = c.execute("SELECT invoice_number FROM invoices ORDER BY rowid DESC LIMIT 1").fetchone()
        if not last:
            return "INV-001"
        parts = last["invoice_number"].rsplit("-", 1)
        if len(parts) == 2 and parts[1].isdigit():
            return f"{parts[0]}-{int(parts[1]) + 1:03d}"
        return f"{last['invoice_number']}-1"

    def export_to_csv(self, file_path: str) -> Tuple[bool, Optional[str]]:
        try:
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

    # ------------------------------------------------------------------  stats (single-pass)
    def get_dashboard_stats(self) -> Dict[str, Any]:
        """Compute all dashboard stats in one DB fetch + one data pass."""
        total_billed = total_paid = 0.0
        count = {"total": 0, "paid": 0, "due": 0, "partial": 0}
        for inv in self.get_all_invoices():
            row = dict(inv)
            count["total"] += 1
            status = (row["status"] or "Due").upper()
            if not row["is_complimentary"]:
                try:
                    items = json.loads(row["items"] or "[]")
                except json.JSONDecodeError:
                    items = []
                t = self.calculate_totals(
                    items,
                    float(row["tax_rate"] or 0) if row.get("is_gst_invoice", 1) else 0.0,
                    float(row.get("global_discount_value") or 0),
                    row.get("global_discount_type") or "amount",
                )
                total_billed += t["grand_total"]
                total_paid   += float(row.get("amount_paid") or 0) + float(row.get("advance_amount") or 0)
            if status == "PAID":      count["paid"]    += 1
            elif status == "DUE":     count["due"]     += 1
            elif status == "PARTIAL": count["partial"] += 1
        return {
            "total_billed": total_billed,
            "total_paid":   total_paid,
            "balance_due":  total_billed - total_paid,
            **count,
        }

    # ------------------------------------------------------------------  static calc
    @staticmethod
    def calculate_totals(items: List[Dict], tax_rate: float,
                         global_disc_val: float = 0.0,
                         global_disc_type: str = "amount") -> Dict[str, float]:
        subtotal    = sum(i.get("amount", 0) for i in items)
        item_disc   = sum(i.get("discount", 0) for i in items)
        after_item  = max(0.0, subtotal - item_disc)
        global_disc = after_item * global_disc_val / 100.0 if global_disc_type == "%" else global_disc_val
        total_disc  = item_disc + global_disc
        taxable     = max(0.0, subtotal - total_disc)
        cgst = sgst = taxable * tax_rate / 200.0
        return {
            "subtotal": subtotal, "item_discount": item_disc, "global_discount": global_disc,
            "total_discount": total_disc, "taxable": taxable,
            "cgst": cgst, "sgst": sgst, "grand_total": taxable + cgst + sgst,
        }


# =============================================================================
#  PDF GENERATOR  (logic preserved; layout cleaned up)
# =============================================================================
class PDFGenerator:

    def __init__(self, data: Dict[str, Any], config: Dict[str, Any]):
        self.data     = data
        self.config   = config
        self.filename = INVOICES_DIR / f"Invoice-{data.get('invoice_number', 'NA')}.pdf"
        self.is_gst   = bool(data.get("is_gst_invoice", 1))
        self.styles   = self._setup_styles()

    def _asset(self, key: str) -> Optional[Path]:
        p = self.config.get("asset_paths", {}).get(key)
        if p and (path := BASE_DIR / p).exists():
            return path
        return None

    def _setup_styles(self) -> Dict[str, ParagraphStyle]:
        ss   = getSampleStyleSheet()
        base = PDF_FONT
        bold = f"{base}-Bold" if base != "Helvetica" else "Helvetica-Bold"
        ss.add(ParagraphStyle("NC",          fontName=base, fontSize=9,  parent=ss["Normal"],    leading=11))
        ss.add(ParagraphStyle("Right",       parent=ss["NC"], alignment=TA_RIGHT))
        ss.add(ParagraphStyle("Center",      parent=ss["NC"], alignment=TA_CENTER))
        ss.add(ParagraphStyle("YourContact", parent=ss["NC"], alignment=TA_RIGHT, leading=11))
        ss.add(ParagraphStyle("BillTo",      parent=ss["NC"], leading=11))
        ss.add(ParagraphStyle("ItemDesc",    parent=ss["NC"], textColor=colors.HexColor("#6B7280"), fontSize=8, leading=10))
        ss.add(ParagraphStyle("CompanyName", fontName=bold,   fontSize=16, textColor=colors.HexColor("#1F2937"), leading=18))
        ss.add(ParagraphStyle("Tagline",     parent=ss["NC"], fontSize=8.5, textColor=colors.HexColor("#6B7280")))
        ss.add(ParagraphStyle("InvoiceTitle",fontName=bold,   fontSize=18, alignment=TA_RIGHT, textColor=colors.HexColor("#111827")))
        return ss

    # --  page elements  -------------------------------------------------------
    def _build_header(self) -> Table:
        comp  = self.config.get("company_details", {})
        left  = [Paragraph(comp.get("name", "Your Company"), self.styles["CompanyName"])]
        if d := comp.get("description"):
            left += [Spacer(1, 2), Paragraph(d, self.styles["Tagline"])]

        status    = self.data.get("status")
        doc_title = ("RECEIPT"       if status == InvoiceStatus.PAID.value else
                     "COMPLIMENTARY" if status == InvoiceStatus.COMPLIMENTARY.value else "INVOICE")
        title_p   = Paragraph(doc_title, self.styles["InvoiceTitle"])

        if logo := self._asset("logo"):
            try:
                img       = Image(logo, 2.0*cm, 2.0*cm, kind="proportional")
                left_side = Table([[img, left]], colWidths=[2.5*cm, 6.5*cm],
                                  style=[("VALIGN",(0,0),(-1,-1),"MIDDLE"), ("LEFTPADDING",(0,0),(-1,-1),0)])
            except Exception:
                left_side = left
        else:
            left_side = left

        return Table([[left_side, title_p]], colWidths=[9.0*cm, 9.0*cm], style=[
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"), ("ALIGN",(1,0),(1,-1),"RIGHT"),
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
                cg = f"<br/>GSTIN: {cgstin}"
            if ygstin := self.config.get("company_details", {}).get("gstin"):
                yg = f"<br/>GSTIN: {ygstin}"
        your_c   = (self.data.get("your_contact")   or "").replace("\n", "<br/>")
        client_c = (self.data.get("client_contact") or "").replace("\n", "<br/>")
        your_p   = Paragraph(f"<font color='#6B7280'>FROM:</font><br/><b>{self.data.get('your_name') or ''}</b><br/>{your_c}{yg}",     self.styles["YourContact"])
        bill_p   = Paragraph(f"<font color='#6B7280'>BILL TO:</font><br/><b>{self.data.get('bill_to') or ''}</b><br/>{client_c}{cg}", self.styles["BillTo"])
        return Table([[bill_p, your_p]], colWidths=["50%","50%"], style=[
            ("VALIGN",(0,0),(-1,-1),"TOP"),
            ("LEFTPADDING",(0,0),(-1,-1),0), ("RIGHTPADDING",(0,0),(-1,-1),0),
        ])

    def _build_meta(self) -> Table:
        date_str = parse_date(self.data.get("invoice_date"), DB_DATE_FORMAT, DISPLAY_DATE_FORMAT)
        meta     = f"<b>Invoice #:</b> {self.data.get('invoice_number') or ''}<br/><b>Invoice Date:</b> {date_str}"
        if not self.data.get("is_complimentary"):
            meta += f"<br/><b>Payment Terms:</b> {self.data.get('payment_terms') or 'N/A'}"
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
                cell = [Paragraph(f"<b>{title}</b>", self.styles["NC"])]
                if desc.strip():
                    cell.append(Paragraph(desc, self.styles["ItemDesc"]))
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
        advance    = float(self.data.get("advance_amount") or 0)
        amt_paid   = float(self.data.get("amount_paid")   or 0)
        balance    = totals["grand_total"] - advance - amt_paid

        rows = [[Paragraph("Subtotal", self.styles["Right"]),
                 Paragraph(f"{CURRENCY_SYMBOL}{totals['subtotal']:,.2f}", self.styles["Right"])]]
        if totals["total_discount"] > 0:
            rows.append([Paragraph("Discount", self.styles["Right"]),
                         Paragraph(f"- {CURRENCY_SYMBOL}{totals['total_discount']:,.2f}", self.styles["Right"])])
        if self.is_gst and tax_rate > 0:
            rows += [
                [Paragraph("Taxable Amount", self.styles["Right"]),
                 Paragraph(f"{CURRENCY_SYMBOL}{totals['taxable']:,.2f}", self.styles["Right"])],
                [Paragraph(f"CGST ({tax_rate/2}%)", self.styles["Right"]),
                 Paragraph(f"{CURRENCY_SYMBOL}{totals['cgst']:,.2f}", self.styles["Right"])],
                [Paragraph(f"SGST ({tax_rate/2}%)", self.styles["Right"]),
                 Paragraph(f"{CURRENCY_SYMBOL}{totals['sgst']:,.2f}", self.styles["Right"])],
            ]
        rows.append([Paragraph("<b>Invoice Total</b>", self.styles["Right"]),
                     Paragraph(f"<b>{CURRENCY_SYMBOL}{totals['grand_total']:,.2f}</b>", self.styles["Right"])])
        if advance > 0:
            ad_label = f"Advance Paid ({parse_date(self.data.get('advance_date'), DB_DATE_FORMAT, DISPLAY_DATE_FORMAT)})"
            rows.append([Paragraph(ad_label, self.styles["Right"]),
                         Paragraph(f"- {CURRENCY_SYMBOL}{advance:,.2f}", self.styles["Right"])])
        if amt_paid > 0:
            pd_label = f"Amount Paid ({parse_date(self.data.get('payment_date'), DB_DATE_FORMAT, DISPLAY_DATE_FORMAT)})"
            rows.append([Paragraph(pd_label, self.styles["Right"]),
                         Paragraph(f"- {CURRENCY_SYMBOL}{amt_paid:,.2f}", self.styles["Right"])])
        bc = "#059669" if balance <= 0 else "#DC2626"
        rows.append([Paragraph("<b>Balance Due</b>", self.styles["Right"]),
                     Paragraph(f"<font color='{bc}'><b>{CURRENCY_SYMBOL}{balance:,.2f}</b></font>", self.styles["Right"])])
        return Table(rows, colWidths=["75%","25%"], style=TableStyle([
            ("ALIGN",(0,0),(-1,-1),"RIGHT"), ("LEFTPADDING",(0,0),(-1,-1),10),
            ("TOPPADDING",(0,0),(-1,-1),3),  ("BOTTOMPADDING",(0,0),(-1,-1),3),
            ("BACKGROUND",(0,-1),(1,-1),colors.HexColor("#FEF2F2" if balance > 0 else "#ECFDF5")),
            ("GRID",(0,-1),(1,-1),1,colors.HexColor("#E5E7EB")),
        ]))

    def _build_footer_section(self) -> Table:
        qr_cell = sig_cell = Spacer(0, 0)
        if self.data.get("show_qr_code") and not self.data.get("is_complimentary"):
            if qr := self._asset("upi_qr_code"):
                upi = self.config.get("payment_details", {}).get("upi_id", "")
                detail = Paragraph(f"<b>PAYMENT DETAILS</b><br/>Scan to pay.<br/>UPI ID: {upi}", self.styles["NC"])
                qr_cell = Table([[Image(qr, 2.2*cm, 2.2*cm, kind="proportional"), detail]],
                                colWidths=[2.5*cm,"100%"], style=[("VALIGN",(0,0),(-1,-1),"TOP")])
        if sig := self._asset("signature"):
            name = self.data.get("your_name") or ""
            sig_cell = Table(
                [[Image(sig, 4.0*cm, 2.0*cm, kind="proportional")],
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
            self._build_totals(),         Spacer(1, 0.8*cm),
            self._build_footer_section(),
        ]
        doc.build(story, onFirstPage=self.on_each_page, onLaterPages=self.on_each_page)
        return self.filename


# =============================================================================
#  SETTINGS DIALOG
# =============================================================================
class SettingsDialog(tk.Toplevel):

    def __init__(self, parent, callback):
        super().__init__(parent)
        self.callback  = callback
        self.title("Application Settings")
        self.geometry("640x760")
        self.minsize(580, 680)
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self._cfg = load_config()
        self._build()
        self._load()
        self.focus_force()

    # --  builder helpers  -----------------------------------------------------
    def _section(self, parent, title: str):
        ttk.Label(parent, text=title, font=(FONT_FAMILY, 11, "bold")).pack(anchor="w", pady=(0, 8))

    def _labeled_entry(self, parent, label: str, var: tk.StringVar):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)
        ttk.Label(row, text=label, width=24, anchor="w").pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)

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

    # --  layout  --------------------------------------------------------------
    def _build(self):
        main = ttk.Frame(self, padding=20)
        main.pack(fill="both", expand=True)

        nb = ttk.Notebook(main)
        nb.pack(fill="both", expand=True, pady=(0, 16))

        # Tab 1 – General
        t1 = ttk.Frame(nb, padding=16)
        nb.add(t1, text="  General  ")
        self._section(t1, "Interface")
        self.theme_var = tk.StringVar()
        self._labeled_combo(t1, "App Theme:", self.theme_var,
                            ["litera","cosmo","flatly","darkly","cyborg","superhero","morph","pulse"])
        ttk.Separator(t1, bootstyle="secondary").pack(fill="x", pady=12)
        self._section(t1, "Company Details")
        self.co_name = tk.StringVar();  self._labeled_entry(t1, "Company Name:",    self.co_name)
        self.co_desc = tk.StringVar();  self._labeled_entry(t1, "Company Tagline:", self.co_desc)
        self.co_gst  = tk.StringVar();  self._labeled_entry(t1, "Company GSTIN:",   self.co_gst)
        ttk.Separator(t1, bootstyle="secondary").pack(fill="x", pady=12)
        self._section(t1, "Your / User Details")
        self.user_name = tk.StringVar(); self._labeled_entry(t1, "Your Name / Firm:", self.user_name)
        ttk.Label(t1, text="Contact Info:", font=(FONT_FAMILY, 9)).pack(anchor="w", pady=(5, 2))
        self.contact_text = tk.Text(t1, height=4, relief="flat", highlightthickness=1, wrap="word")
        self.contact_text.pack(fill="x")

        # Tab 2 – Payment & PDF
        t2 = ttk.Frame(nb, padding=16)
        nb.add(t2, text="  Payment & PDF  ")
        self._section(t2, "Payment Settings")
        self.currency   = tk.StringVar(); self._labeled_entry(t2, "Currency Symbol:", self.currency)
        self.upi_id     = tk.StringVar(); self._labeled_entry(t2, "UPI ID:",          self.upi_id)
        self.pay_terms  = tk.StringVar(); self._labeled_entry(t2, "Default Terms:",   self.pay_terms)
        ttk.Separator(t2, bootstyle="secondary").pack(fill="x", pady=12)
        self._section(t2, "PDF Footer Note")
        self.footer_text = tk.Text(t2, height=3, relief="flat", highlightthickness=1, wrap="word")
        self.footer_text.pack(fill="x")

        # Tab 3 – Assets
        t3 = ttk.Frame(nb, padding=16)
        nb.add(t3, text="  Images & Assets  ")
        self._section(t3, "Upload Assets (PNG / JPG)")
        ttk.Label(t3, text="Files are copied into the 'assets' folder and embedded in PDFs.",
                  font=(FONT_FAMILY, 9), bootstyle="secondary").pack(anchor="w", pady=(0, 12))
        self.logo_var = tk.StringVar(); self._file_picker(t3, "Company Logo:",         self.logo_var, "logo.png")
        self.sig_var  = tk.StringVar(); self._file_picker(t3, "Authorized Signature:", self.sig_var,  "sign.png")
        self.qr_var   = tk.StringVar(); self._file_picker(t3, "UPI QR Code:",          self.qr_var,   "upi_qr.png")

        # Button row
        btn_row = ttk.Frame(main)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="Cancel",             bootstyle="secondary", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(btn_row, text="Save Configuration", bootstyle="success",   command=self._save).pack(side="right")

    def _load(self):
        c = self._cfg
        self.theme_var.set(c.get("theme", "litera"))
        self.co_name.set(c.get("company_details", {}).get("name", ""))
        self.co_desc.set(c.get("company_details", {}).get("description", ""))
        self.co_gst.set(c.get("company_details", {}).get("gstin", ""))
        self.user_name.set(c.get("user_details", {}).get("name", ""))
        self.contact_text.insert("1.0", c.get("user_details", {}).get("contact_info", ""))
        self.currency.set(c.get("currency_symbol", "₹"))
        self.upi_id.set(c.get("payment_details", {}).get("upi_id", ""))
        self.pay_terms.set(c.get("payment_details", {}).get("payment_terms", ""))
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
        self.callback()
        self.destroy()


# =============================================================================
#  ABOUT DIALOG
# =============================================================================

class AboutDialog(tk.Toplevel):

    def __init__(self, parent):
        super().__init__(parent)

        self.title(f"About {APP_NAME}")
        self.geometry("400x400")
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

        # =====================================================
        # HEADER
        # =====================================================

        hero = ttk.Frame(
            container,
            padding=15,
            bootstyle="primary"
        )
        hero.pack(fill="x")

        ttk.Label(
            hero,
            text=APP_NAME,
            font=(FONT_FAMILY, 18, "bold"),
            bootstyle="inverse-primary",
            anchor="center",
            justify="center"
        ).pack(fill="x")

        ttk.Label(
            hero,
            text=APP_TAGLINE,
            font=(FONT_FAMILY, 9),
            bootstyle="inverse-primary",
            anchor="center",
            justify="center"
        ).pack(fill="x", pady=(4, 0))

        # =====================================================
        # VERSION
        # =====================================================

        ttk.Label(
            container,
            text=APP_VERSION,
            font=(FONT_FAMILY, 28, "bold"),
            bootstyle="info"
        ).pack(pady=(20, 5))

        ttk.Label(
            container,
            text=f"Released on {APP_RELEASE_DATE}",
            font=(FONT_FAMILY, 9),
            bootstyle="secondary"
        ).pack()

        ttk.Separator(container).pack(fill="x", pady=15)

        # =====================================================
        # ATTRIBUTION
        # =====================================================

        lines = APP_ATTRIBUTION.split("\n")

        ttk.Label(
            container,
            text=lines[0],
            font=(FONT_FAMILY, 11, "bold"),
            anchor="center",
            justify="center"
        ).pack(fill="x")

        for line in lines[1:]:
            ttk.Label(
                container,
                text=line,
                font=(FONT_FAMILY, 9),
                bootstyle="secondary",
                anchor="center",
                justify="center"
            ).pack(fill="x", pady=1)

        # =====================================================
        # CLOSE BUTTON
        # =====================================================

        ttk.Button(
            container,
            text="Close",
            width=14,
            bootstyle="primary",
            command=self.destroy
        ).pack(side="bottom", pady=(20, 0))

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
        ttk.Button(tb, text="⊕", command=lambda: self._zoom(+0.15), width=3,
                   bootstyle="secondary-outline").pack(side="right", padx=2)
        ttk.Button(tb, text="⊖", command=lambda: self._zoom(-0.15), width=3,
                   bootstyle="secondary-outline").pack(side="right", padx=2)
        ttk.Button(tb, text="Fit",command=self._reset_zoom,        width=4,
                   bootstyle="secondary-outline").pack(side="right", padx=2)

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
        if self.page_idx < self.total - 1: self.page_idx += 1; self._render()
    def _prev(self):
        if self.page_idx > 0:             self.page_idx -= 1; self._render()
    def _zoom(self, delta: float):
        self.zoom = max(self.MIN_ZOOM, min(self.MAX_ZOOM, self.zoom + delta)); self._render()
    def _reset_zoom(self):
        self.zoom = 1.0; self._render()

    def _print(self):
        try:
            if sys.platform == "win32":
                try:
                    os.startfile(self.pdf_path, "print")
                except OSError:
                    Messagebox.show_warning("Manual Print Required",
                        "A PDF reader is not set as default.\n"
                        "The file will open — press Ctrl+P to print.", parent=self)
                    os.startfile(self.pdf_path)
            else:
                subprocess.run(["lp", str(self.pdf_path)], check=True)
        except Exception as exc:
            Messagebox.show_error("Print Error", str(exc), parent=self)

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
        self.doc.close(); self.destroy()


# =============================================================================
#  DASHBOARD
# =============================================================================
class DashboardFrame(ttk.Frame):
    STATUS_FILTERS = ["All", "Paid", "Due", "Partial", "Complimentary"]
    SORT_ASC  = " ▲"
    SORT_DESC = " ▼"
    COL_DEFS  = [
        ("id",     "Invoice #",       118, "center"),
        ("client", "Client / Company", 290, "w"),
        ("date",   "Date",             112, "center"),
        ("status", "Status",            98, "center"),
        ("total",  "Total Amount",     148, "e"),
    ]

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller     = controller
        self.db_manager     = controller.db_manager
        self._status_filter = "All"
        self._sort_col      = "date"
        self._sort_rev      = True
        self._debounce_id   = None
        self._filter_btns: Dict[str, ttk.Button] = {}
        self._build()

    # --  layout  --------------------------------------------------------------
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
        ttk.Label(left, text=APP_NAME,
                  font=(FONT_FAMILY, 24, "bold")).pack(anchor="w")
        ttk.Label(left, text=APP_TAGLINE,
                  font=(FONT_FAMILY, 10), bootstyle="secondary").pack(anchor="w")
        ttk.Label(left, text="Dashboard",
                  font=(FONT_FAMILY, 10, "bold"), bootstyle="secondary").pack(anchor="w", pady=(2, 0))

        right = ttk.Frame(row); right.pack(side="right")
        ttk.Button(right, text="⚙  Settings", bootstyle="secondary-outline",
                   command=self._open_settings).pack(side="left", padx=4)
        ttk.Button(right, text="ℹ  About", bootstyle="info-outline",
               command=self._open_about).pack(side="left", padx=4)
        ttk.Button(right, text="📊  Export CSV", bootstyle="info-outline",
                   command=self._export_csv).pack(side="left", padx=4)
        ttk.Button(right, text="＋  New Invoice", bootstyle="primary",
                   command=lambda: self.controller.show_frame(InvoiceFormFrame)).pack(side="left", padx=(4, 0))

    def _stat_cards(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(0, 18))
        for i in range(4):
            row.columnconfigure(i, weight=1)
        self.s_billed   = self._card(row, "💰  Total Billed",  f"{CURRENCY_SYMBOL}0.00", "primary",   0)
        self.s_received = self._card(row, "✅  Total Received", f"{CURRENCY_SYMBOL}0.00", "success",   1)
        self.s_due      = self._card(row, "⚠️  Balance Due",   f"{CURRENCY_SYMBOL}0.00", "danger",    2)
        self.s_count    = self._card(row, "📄  Total Invoices", "0",                      "secondary", 3)

    def _card(self, parent, title: str, value: str, style: str, col: int) -> ttk.Label:
        card = ttk.Frame(parent, padding=(18, 14), bootstyle=style)
        card.grid(row=0, column=col, sticky="ew", padx=6)
        ttk.Label(card, text=title, font=(FONT_FAMILY, 10),
                  bootstyle=f"inverse-{style}").pack(anchor="w")
        lbl = ttk.Label(card, text=value, font=(FONT_FAMILY, 20, "bold"),
                        bootstyle=f"inverse-{style}")
        lbl.pack(anchor="w", pady=(4, 0))
        return lbl

    def _filter_bar(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(0, 8))

        # Status filter pills
        filter_group = ttk.Frame(row)
        filter_group.pack(side="left")
        ttk.Label(filter_group, text="Filter:", font=(FONT_FAMILY, 9),
                  bootstyle="secondary").pack(side="left", padx=(0, 8))
        for status in self.STATUS_FILTERS:
            btn = ttk.Button(
                filter_group, text=status, width=13,
                bootstyle="primary" if status == "All" else "secondary-outline",
                command=lambda s=status: self._set_filter(s),
            )
            btn.pack(side="left", padx=2)
            self._filter_btns[status] = btn

        # Search
        srch = ttk.Frame(row)
        srch.pack(side="right")
        ttk.Label(srch, text="🔍", font=(FONT_FAMILY, 12)).pack(side="left", padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search)
        entry = ttk.Entry(srch, textvariable=self.search_var, width=28)
        entry.pack(side="left")
        ToolTip(entry, "Search by invoice number or client name")
        ttk.Button(srch, text="✕", width=3, bootstyle="secondary-outline",
                   command=lambda: self.search_var.set("")).pack(side="left", padx=(4, 0))

    def _set_filter(self, status: str):
        self._status_filter = status
        for s, btn in self._filter_btns.items():
            btn.configure(bootstyle="primary" if s == status else "secondary-outline")
        self._populate()

    def _on_search(self, *_):
        if self._debounce_id:
            self.after_cancel(self._debounce_id)
        self._debounce_id = self.after(200, self._populate)

    def _table(self, parent):
        wrap = ttk.Frame(parent)
        wrap.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(wrap, columns=[c[0] for c in self.COL_DEFS],
                                 show="headings", bootstyle="primary",
                                 selectmode="browse", height=18)
        for cid, text, width, anchor in self.COL_DEFS:
            self.tree.heading(cid, text=text, anchor=anchor,
                              command=lambda c=cid: self._sort_by(c))
            self.tree.column(cid, width=width, anchor=anchor, minwidth=60)

        for tag, colors_map in STATUS_BADGE.items():
            self.tree.tag_configure(tag, foreground=colors_map["foreground"],
                                    background=colors_map["background"])

        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        self.tree.bind("<Double-1>",          lambda _: self._act_edit())
        self.tree.bind("<<TreeviewSelect>>",  self._on_select)
        self.tree.bind("<Button-3>",          self._show_ctx_menu)
        self.tree.bind("<Delete>",            lambda _: self._act_delete())

        # Context menu
        self._ctx = tk.Menu(self, tearoff=0)
        self._ctx.add_command(label="📄  View PDF",   command=self._act_view_pdf)
        self._ctx.add_command(label="✏️  Edit",       command=self._act_edit)
        self._ctx.add_command(label="📋  Duplicate",  command=self._act_duplicate)
        self._ctx.add_separator()
        self._ctx.add_command(label="🗑️  Delete",     command=self._act_delete)

    def _sort_by(self, col: str):
        if self._sort_col == col:
            self._sort_rev = not self._sort_rev
        else:
            self._sort_col = col; self._sort_rev = True
        col_texts = {c[0]: c[1] for c in self.COL_DEFS}
        for c, t in col_texts.items():
            arrow = (self.SORT_DESC if self._sort_rev else self.SORT_ASC) if c == col else ""
            self.tree.heading(c, text=t + arrow)
        self._populate()

    def _action_bar(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(10, 0))

        self.hint_lbl = ttk.Label(row, text="Select an invoice to take action",
                                  bootstyle="secondary", font=(FONT_FAMILY, 9))
        self.hint_lbl.pack(side="left")

        self.del_btn  = ttk.Button(row, text="🗑  Delete",    bootstyle="danger-outline",   state="disabled", command=self._act_delete)
        self.del_btn.pack(side="right", padx=4)
        self.dup_btn  = ttk.Button(row, text="📋  Duplicate",  bootstyle="secondary-outline", state="disabled", command=self._act_duplicate)
        self.dup_btn.pack(side="right", padx=4)
        self.edit_btn = ttk.Button(row, text="✏  Edit",      bootstyle="secondary",          state="disabled", command=self._act_edit)
        self.edit_btn.pack(side="right", padx=4)
        self.pdf_btn  = ttk.Button(row, text="📄  View PDF",  bootstyle="primary",            state="disabled", command=self._act_view_pdf)
        self.pdf_btn.pack(side="right", padx=4)

    # --  lifecycle  -----------------------------------------------------------
    def on_show(self, _=None):
        self._populate()
        self._update_stats()

    def _on_select(self, _=None):
        has   = bool(self.tree.selection())
        state = "normal" if has else "disabled"
        for btn in [self.pdf_btn, self.edit_btn, self.dup_btn, self.del_btn]:
            btn.config(state=state)
        self.hint_lbl.config(
            text="Double-click to edit  •  Right-click for more options" if has
                 else "Select an invoice to take action"
        )

    def _selected_id(self) -> Optional[str]:
        sel = self.tree.selection()
        return self.tree.item(sel[0], "values")[0] if sel else None

    # --  data  ----------------------------------------------------------------
    def _populate(self):
        self.tree.delete(*self.tree.get_children())
        rows = self.db_manager.get_all_invoices(
            search_query=self.search_var.get().strip(),
            status_filter=self._status_filter,
        )
        col_idx = {c[0]: i for i, c in enumerate(self.COL_DEFS)}
        entries = []
        for inv in rows:
            row = dict(inv)
            try:   items = json.loads(inv["items"] or "[]")
            except json.JSONDecodeError: items = []
            t = DatabaseManager.calculate_totals(
                items,
                float(row["tax_rate"] or 0) if row.get("is_gst_invoice", 1) else 0.0,
                float(row.get("global_discount_value") or 0),
                row.get("global_discount_type") or "amount",
            )
            total_txt  = "No Charge" if row["is_complimentary"] else f"{CURRENCY_SYMBOL}{t['grand_total']:,.2f}"
            status_val = (row["status"] or "Due").upper()
            date_val   = parse_date(row["invoice_date"], DB_DATE_FORMAT, DISPLAY_DATE_FORMAT)
            values     = (row["invoice_number"], row["bill_to"], date_val, status_val, total_txt)
            entries.append((values, status_val))

        # Sort client-side
        si = col_idx.get(self._sort_col, 2)
        def _key(e):
            v = e[0][si]
            if self._sort_col == "total" and v != "No Charge":
                try: return float(v.replace(CURRENCY_SYMBOL, "").replace(",", ""))
                except: return 0.0
            return v
        entries.sort(key=_key, reverse=self._sort_rev)

        for values, tag in entries:
            badge = tag if tag in STATUS_BADGE else ""
            self.tree.insert("", "end", values=values, tags=(badge,) if badge else ())

    def _update_stats(self):
        s = self.db_manager.get_dashboard_stats()
        self.s_billed.config(text=f"{CURRENCY_SYMBOL}{s['total_billed']:,.2f}")
        self.s_received.config(text=f"{CURRENCY_SYMBOL}{s['total_paid']:,.2f}")
        self.s_due.config(text=f"{CURRENCY_SYMBOL}{s['balance_due']:,.2f}")
        self.s_count.config(text=str(s["total"]))

    # --  actions  -------------------------------------------------------------
    def _act_view_pdf(self):
        if inv_id := self._selected_id():
            self.controller.show_progress(True)
            threading.Thread(target=self._pdf_worker, args=(inv_id, self.controller.task_queue), daemon=True).start()

    def _act_edit(self):
        if inv_id := self._selected_id():
            self.controller.show_frame(InvoiceFormFrame, inv_id)

    def _act_delete(self):
        if not (inv_id := self._selected_id()): return
        if Messagebox.okcancel("Delete Invoice", f"Permanently delete Invoice #{inv_id}?\nThis cannot be undone."):
            self.db_manager.delete_invoice(inv_id)
            self.on_show()

    def _act_duplicate(self):
        if not (inv_id := self._selected_id()): return
        row = self.db_manager.get_invoice_by_id(inv_id)
        if not row: return
        new_data = dict(row)
        new_data["invoice_number"] = self.db_manager.get_next_invoice_number()
        new_data["invoice_date"]   = datetime.date.today().strftime(DB_DATE_FORMAT)
        new_data["status"]         = InvoiceStatus.DUE.value
        new_data["amount_paid"]    = 0.0
        new_data["advance_amount"] = 0.0
        new_data["payment_date"]   = None
        new_data["advance_date"]   = None
        ok, err = self.db_manager.save_invoice(new_data, is_update=False)
        if ok:
            self.on_show()
            Messagebox.show_info("Duplicated", f"Created {new_data['invoice_number']} as a copy.")
        else:
            Messagebox.show_error("Error", f"Could not duplicate:\n{err}")

    def _show_ctx_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            self._ctx.post(event.x_root, event.y_root)

    def _export_csv(self):
        if path := filedialog.asksaveasfilename(defaultextension=".csv",
                                                filetypes=[("CSV Files", "*.csv")],
                                                title="Export Invoices"):
            ok, msg = self.db_manager.export_to_csv(path)
            Messagebox.show_info("Exported", f"Saved to:\n{path}") if ok else Messagebox.show_error("Failed", msg)

    def _open_settings(self):
        def _apply():
            ttk.Style().theme_use(CONFIG.get("theme", "litera"))
            self.on_show()
        SettingsDialog(self.controller, _apply)

    def _open_about(self):
        AboutDialog(self.controller)

    def _pdf_worker(self, invoice_id: str, q: queue.Queue):
        try:
            path = PDFGenerator(dict(self.db_manager.get_invoice_by_id(invoice_id)), load_config()).generate()
            q.put({"status": "success", "path": path})
        except Exception as exc:
            logger.error(f"PDF generation failed ({invoice_id}): {exc}", exc_info=True)
            q.put({"status": "error", "message": str(exc)})


# =============================================================================
#  INVOICE FORM
# =============================================================================
class InvoiceFormFrame(ttk.Frame):

    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.db_manager = controller.db_manager
        self.item_widgets: List[Dict] = []
        self.is_editing = False
        self.current_id: Optional[str] = None

        self.show_qr_var  = tk.IntVar(value=1)
        self.is_gst_var   = tk.BooleanVar(value=True)
        self.is_no_charge = tk.BooleanVar(value=False)
        self.is_no_charge.trace_add("write", self._toggle_no_charge)

        self._build()
        self.bind_all("<Control-s>", lambda _: self.save_invoice())
        self.bind_all("<Control-n>", lambda _: self.add_item())
        self.bind_all("<Escape>",    lambda _: self.controller.show_frame(DashboardFrame))

    # --  master layout  -------------------------------------------------------
    def _build(self):
        scroll = ScrolledFrame(self, padding=(28, 14, 44, 14), autohide=True)
        scroll.pack(fill="both", expand=True)

        # Page header
        hdr = ttk.Frame(scroll); hdr.pack(fill="x", pady=(0, 18))
        self.form_title = ttk.Label(hdr, text="New Invoice", font=(FONT_FAMILY, 26, "bold"))
        self.form_title.pack(side="left", anchor="w")
        ttk.Label(hdr, text="Ctrl+S  save  •  Esc  back  •  Ctrl+N  add item",
                  font=(FONT_FAMILY, 9), bootstyle="secondary").pack(side="right", anchor="se", pady=(12, 0))

        # 2-column body
        body = ttk.Frame(scroll); body.pack(fill="x", expand=True)
        body.columnconfigure((0, 1), weight=1, uniform="col")
        left  = ttk.Frame(body); left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        right = ttk.Frame(body); right.grid(row=0, column=1, sticky="nsew", padx=(12, 0))

        self._sec_billing(left)
        self._sec_from(left)
        self._sec_invoice_details(right)
        self._sec_payment(right)

        self._sec_items(scroll)
        self._sec_totals(scroll)

        # Action buttons
        act = ttk.Frame(scroll); act.pack(fill="x", pady=(22, 4))
        ttk.Button(act, text="← Back", bootstyle="secondary-outline",
                   command=lambda: self.controller.show_frame(DashboardFrame)).pack(side="left")
        self.save_btn = ttk.Button(act, text="💾  Save Invoice",
                                   bootstyle="success", command=self.save_invoice)
        self.save_btn.pack(side="right")
        ToolTip(self.save_btn, "Save invoice  (Ctrl+S)")

    # --  section helpers  -----------------------------------------------------
    def _lf(self, parent, title: str) -> ttk.Labelframe:
        lf = ttk.Labelframe(parent, text=f"  {title}  ", padding=(16, 12))
        lf.pack(fill="x", pady=(0, 14))
        return lf

    def _field_entry(self, parent, label: str) -> ttk.Entry:
        ttk.Label(parent, text=label, font=(FONT_FAMILY, 9)).pack(anchor="w", pady=(6, 2))
        w = ttk.Entry(parent); w.pack(fill="x"); return w

    def _field_text(self, parent, label: str, height: int = 4) -> tk.Text:
        ttk.Label(parent, text=label, font=(FONT_FAMILY, 9)).pack(anchor="w", pady=(6, 2))
        w = tk.Text(parent, height=height, relief="flat", highlightthickness=1, wrap="word", bd=0)
        w.pack(fill="x"); return w

    def _field_date(self, parent, label: str) -> DateEntry:
        ttk.Label(parent, text=label, font=(FONT_FAMILY, 9)).pack(anchor="w", pady=(6, 2))
        w = DateEntry(parent, dateformat=DISPLAY_DATE_FORMAT, bootstyle="primary")
        w.pack(fill="x"); return w

    def _field_combo(self, parent, label: str, values: list) -> ttk.Combobox:
        ttk.Label(parent, text=label, font=(FONT_FAMILY, 9)).pack(anchor="w", pady=(6, 2))
        w = ttk.Combobox(parent, values=values, state="readonly"); w.pack(fill="x"); return w

    # --  sections  ------------------------------------------------------------
    def _sec_billing(self, parent):
        s = self._lf(parent, "Bill To")
        self.bill_to_entry       = self._field_entry(s, "Client Name / Company *")
        self.client_gstin_entry  = self._field_entry(s, "Client GSTIN (Optional)")
        self.client_contact_text = self._field_text(s,  "Client Contact Info")
        ToolTip(self.bill_to_entry, "Required — the client or company you're billing")

    def _sec_from(self, parent):
        s = self._lf(parent, "From (Your Details)")
        self.your_name_entry   = self._field_entry(s, "Your Name / Company *")
        self.your_contact_text = self._field_text(s,  "Your Contact Info")

    def _sec_invoice_details(self, parent):
        s = self._lf(parent, "Invoice Details")
        self.invoice_number_entry = self._field_entry(s, "Invoice Number *")
        self.invoice_date_entry   = self._field_date(s,  "Invoice Date")
        self.payment_terms_entry  = self._field_entry(s, "Payment Terms")

    def _sec_payment(self, parent):
        s = self._lf(parent, "Tax & Payment")

        # Discount row (value + type side-by-side)
        ttk.Label(s, text="Global Discount:", font=(FONT_FAMILY, 9)).pack(anchor="w", pady=(6, 2))
        disc_row = ttk.Frame(s); disc_row.pack(fill="x")
        self.global_discount_entry = ttk.Entry(disc_row)
        self.global_discount_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.global_discount_type_combo = ttk.Combobox(disc_row, values=[CURRENCY_SYMBOL, "%"],
                                                       state="readonly", width=5)
        self.global_discount_type_combo.pack(side="left")
        self.global_discount_type_combo.set(CURRENCY_SYMBOL)

        self.tax_rate_entry       = self._field_entry(s, "GST / Tax Rate (%)")
        self.advance_amount_entry = self._field_entry(s, "Advance Amount Paid")
        self.advance_date_entry   = self._field_date(s,  "Advance Payment Date")
        self.amount_paid_entry    = self._field_entry(s, "Subsequent Amount Paid")
        self.payment_date_entry   = self._field_date(s,  "Date of Payment")
        self.payment_mode_combo   = self._field_combo(s, "Payment Mode", PAYMENT_MODES)

        # Defaults
        self.global_discount_entry.insert(0, "0")
        self.advance_amount_entry.insert(0,  "0")
        self.amount_paid_entry.insert(0,     "0")
        self.tax_rate_entry.insert(0,        "18.0")

        # Toggle switches
        tgl = ttk.Frame(s); tgl.pack(fill="x", pady=(14, 2))
        self.is_gst_chk = ttk.Checkbutton(tgl, text="GST Invoice",
                                           variable=self.is_gst_var, bootstyle="round-toggle",
                                           command=self._toggle_gst)
        self.is_gst_chk.grid(row=0, column=0, sticky="w", pady=3)
        self.show_qr_chk = ttk.Checkbutton(tgl, text="Include QR Code in PDF",
                                            variable=self.show_qr_var, bootstyle="round-toggle")
        self.show_qr_chk.grid(row=1, column=0, sticky="w", pady=3)
        ttk.Checkbutton(tgl, text="Mark as 'No Charge' / Complimentary",
                        variable=self.is_no_charge, bootstyle="round-toggle").grid(row=2, column=0, sticky="w", pady=3)

        # Live-update bindings
        for w in [self.tax_rate_entry, self.amount_paid_entry,
                  self.global_discount_entry, self.advance_amount_entry]:
            w.bind("<KeyRelease>", self._update_totals)
        self.global_discount_type_combo.bind("<<ComboboxSelected>>", self._update_totals)

    def _sec_items(self, parent):
        lf = ttk.Labelframe(parent, text="  Items  ", padding=(16, 12))
        lf.pack(fill="x", expand=True, pady=(0, 8))
        # Header row
        hdr = ttk.Frame(lf); hdr.pack(fill="x", pady=(0, 6))
        for i, (txt, w) in enumerate([("Item / Title", 3), ("Description", 5), ("Amount", 2), ("Discount", 2), ("", 0)]):
            hdr.columnconfigure(i, weight=w)
            ttk.Label(hdr, text=txt, font=(FONT_FAMILY, 9, "bold"),
                      bootstyle="secondary").grid(row=0, column=i, sticky="w", padx=2)
        self.items_container = ScrolledFrame(lf, autohide=True, height=230)
        self.items_container.pack(fill="x", expand=True, pady=4)
        ttk.Button(lf, text="＋  Add Item  (Ctrl+N)", command=self.add_item,
                   bootstyle="success-outline").pack(pady=(8, 0), anchor="w")

    def _sec_totals(self, parent):
        outer = ttk.Frame(parent); outer.pack(fill="x", pady=(0, 8))
        ttk.Frame(outer).pack(side="left", expand=True, fill="x")  # spacer

        totals_lf = ttk.Labelframe(outer, text="  Summary  ", padding=(18, 12))
        totals_lf.pack(side="right")

        row_idx = [0]
        def add_row(label: str, attr: str, bold: bool = False, style: str = "default"):
            r = row_idx[0]; row_idx[0] += 1
            ttk.Label(totals_lf, text=label,
                      font=(FONT_FAMILY, 10, "bold" if bold else "normal"),
                      anchor="e", width=20).grid(row=r, column=0, sticky="e", padx=(0, 10), pady=2)
            lbl = ttk.Label(totals_lf, text=f"{CURRENCY_SYMBOL}0.00",
                            font=(FONT_FAMILY, 10, "bold" if bold else "normal"),
                            bootstyle=style, anchor="e", width=18)
            lbl.grid(row=r, column=1, sticky="e", pady=2)
            setattr(self, attr, lbl)

        add_row("Subtotal:",       "subtotal_label")
        add_row("Total Discount:", "discount_label",  style="danger")
        add_row("Taxable Amount:", "taxable_label")
        add_row("CGST:",           "cgst_label",      style="secondary")
        add_row("SGST:",           "sgst_label",      style="secondary")
        ttk.Separator(totals_lf).grid(row=row_idx[0], column=0, columnspan=2, sticky="ew", pady=6)
        row_idx[0] += 1
        add_row("Grand Total:",    "total_label",     bold=True)
        add_row("Balance Due:",    "balance_label",   bold=True, style="danger")

    # --  item rows  -----------------------------------------------------------
    def add_item(self, event=None, item: Optional[Dict] = None):
        fr = ttk.Frame(self.items_container)
        fr.pack(fill="x", pady=(0, 4), padx=2)
        fr.columnconfigure(0, weight=3)
        fr.columnconfigure(1, weight=5)
        fr.columnconfigure(2, weight=2)
        fr.columnconfigure(3, weight=2)

        defaults  = {"title": "", "description": "", "amount": "0.00", "discount": "0.00"}
        item_data = item if isinstance(item, dict) else {}
        record    = {"frame": fr, "vars": {}}

        for i, key in enumerate(defaults):
            var = tk.StringVar(value=str(item_data.get(key, defaults[key])))
            ent = ttk.Entry(fr, textvariable=var)
            ent.grid(row=0, column=i, sticky="ew", padx=(0, 6))
            record[key]         = ent
            record["vars"][key] = var
            if key in ("amount", "discount"):
                var.trace_add("write", self._update_totals)

        ttk.Button(fr, text="✕", width=3, bootstyle="danger-outline",
                   command=lambda f=fr: self._remove_item(f)).grid(row=0, column=4, sticky="e")
        self.item_widgets.append(record)
        self._update_totals()

    def _remove_item(self, fr: ttk.Frame):
        self.item_widgets = [w for w in self.item_widgets if w["frame"] is not fr]
        fr.destroy()
        self._update_totals()

    # --  toggles  -------------------------------------------------------------
    def _toggle_gst(self, *_):
        if self.is_no_charge.get(): return
        st = "normal" if self.is_gst_var.get() else "disabled"
        self.tax_rate_entry.config(state=st)
        self.client_gstin_entry.config(state=st)
        self._update_totals()

    def _toggle_no_charge(self, *_):
        nc       = self.is_no_charge.get()
        st, cst  = ("disabled", "disabled") if nc else ("normal", "readonly")
        for w in [self.amount_paid_entry, self.payment_date_entry.entry,
                  self.payment_terms_entry, self.show_qr_chk,
                  self.is_gst_chk, self.global_discount_entry,
                  self.advance_amount_entry, self.advance_date_entry.entry]:
            w.config(state=st)
        self.payment_mode_combo.config(state=cst)
        self.global_discount_type_combo.config(state=cst)
        gst_st = "normal" if not nc and self.is_gst_var.get() else "disabled"
        self.tax_rate_entry.config(state=gst_st)
        self.client_gstin_entry.config(state=gst_st)
        self._update_totals()

    def _update_totals(self, *_):
        if self.is_no_charge.get():
            for attr in ["subtotal_label","discount_label","taxable_label",
                         "cgst_label","sgst_label","total_label","balance_label"]:
                getattr(self, attr).config(text="No Charge")
            return

        items = []
        for w in self.item_widgets:
            try:
                items.append({"amount":   float(w["vars"]["amount"].get()   or 0),
                              "discount": float(w["vars"]["discount"].get()  or 0)})
            except ValueError:
                pass

        def _float(widget, fallback=0.0) -> float:
            try: return float(widget.get() or 0)
            except ValueError: return fallback

        tax  = _float(self.tax_rate_entry)       if self.is_gst_var.get() else 0.0
        gd   = _float(self.global_discount_entry)
        adv  = _float(self.advance_amount_entry)
        paid = _float(self.amount_paid_entry)

        t = DatabaseManager.calculate_totals(items, tax, gd, self.global_discount_type_combo.get())

        self.subtotal_label.config(text=f"{CURRENCY_SYMBOL}{t['subtotal']:,.2f}")
        self.discount_label.config(text=f"- {CURRENCY_SYMBOL}{t['total_discount']:,.2f}")
        self.taxable_label.config(text=f"{CURRENCY_SYMBOL}{t['taxable']:,.2f}" if self.is_gst_var.get() else "N/A")
        self.cgst_label.config(text=f"{CURRENCY_SYMBOL}{t['cgst']:,.2f} ({tax/2}%)" if self.is_gst_var.get() else "N/A")
        self.sgst_label.config(text=f"{CURRENCY_SYMBOL}{t['sgst']:,.2f} ({tax/2}%)" if self.is_gst_var.get() else "N/A")
        self.total_label.config(text=f"{CURRENCY_SYMBOL}{t['grand_total']:,.2f}")
        self.balance_label.config(text=f"{CURRENCY_SYMBOL}{t['grand_total'] - adv - paid:,.2f}")

    # --  form lifecycle  ------------------------------------------------------
    def on_show(self, invoice_id: Optional[str] = None):
        if invoice_id: self._load_for_edit(invoice_id)
        else:          self._reset_new()

    def _reset_new(self):
        self._reset(defaults=True)
        self.is_editing = False
        self.current_id = None
        self.invoice_number_entry.insert(0, self.db_manager.get_next_invoice_number())
        self.form_title.config(text="New Invoice")
        self.save_btn.config(text="💾  Save Invoice")

    def _reset(self, defaults: bool = True):
        for w in [self.invoice_number_entry, self.bill_to_entry, self.client_gstin_entry,
                  self.your_name_entry, self.amount_paid_entry, self.advance_amount_entry,
                  self.payment_terms_entry, self.tax_rate_entry, self.global_discount_entry]:
            w.delete(0, "end")
        for w in [self.client_contact_text, self.your_contact_text]:
            w.delete("1.0", "end")
        for de in [self.invoice_date_entry, self.payment_date_entry, self.advance_date_entry]:
            de.entry.delete(0, "end")
        self.payment_mode_combo.set("")
        for iw in self.item_widgets: iw["frame"].destroy()
        self.item_widgets.clear()

        if defaults:
            self.your_name_entry.insert(0,      CONFIG.get("user_details", {}).get("name", ""))
            self.your_contact_text.insert("1.0", CONFIG.get("user_details", {}).get("contact_info", ""))
            self.payment_terms_entry.insert(0,   CONFIG.get("payment_details", {}).get("payment_terms", ""))
            self.invoice_date_entry.entry.insert(0, datetime.date.today().strftime(DISPLAY_DATE_FORMAT))
            self.tax_rate_entry.insert(0,        "18.0")
            self.global_discount_entry.insert(0, "0")
            self.global_discount_type_combo.set(CURRENCY_SYMBOL)
            self.advance_amount_entry.insert(0,  "0")
            self.amount_paid_entry.insert(0,     "0")
            self.show_qr_var.set(1)
            self.is_gst_var.set(True)
            self.is_no_charge.set(False)
            self._toggle_no_charge()
            self.add_item()

    def _load_for_edit(self, invoice_id: str):
        self._reset(defaults=False)
        self.is_editing = True
        self.current_id = invoice_id
        row = self.db_manager.get_invoice_by_id(invoice_id)
        if not row:
            self.controller.show_frame(DashboardFrame); return
        d = dict(row)

        self.invoice_number_entry.insert(0,    d.get("invoice_number") or "")
        self.bill_to_entry.insert(0,            d.get("bill_to")       or "")
        self.client_gstin_entry.insert(0,       d.get("client_gstin")  or "")
        self.client_contact_text.insert("1.0",  d.get("client_contact")or "")
        self.your_name_entry.insert(0,          d.get("your_name")     or "")
        self.your_contact_text.insert("1.0",    d.get("your_contact")  or "")
        self.tax_rate_entry.delete(0, "end"); self.tax_rate_entry.insert(0, str(d.get("tax_rate") or 0))

        for widget, key in [(self.invoice_date_entry, "invoice_date"),
                            (self.advance_date_entry,  "advance_date"),
                            (self.payment_date_entry,  "payment_date")]:
            val = parse_date(d.get(key), DB_DATE_FORMAT, DISPLAY_DATE_FORMAT)
            if val and val != "N/A":
                widget.entry.delete(0, "end"); widget.entry.insert(0, val)

        self.payment_terms_entry.insert(0,         d.get("payment_terms") or "")
        self.advance_amount_entry.insert(0,         str(d.get("advance_amount") or "0"))
        self.amount_paid_entry.insert(0,            str(d.get("amount_paid")    or "0"))
        self.payment_mode_combo.set(                d.get("payment_mode")    or "")
        self.global_discount_entry.insert(0,        str(d.get("global_discount_value") or "0"))
        self.global_discount_type_combo.set(        d.get("global_discount_type") or CURRENCY_SYMBOL)
        self.show_qr_var.set(                       d.get("show_qr_code", 1))
        self.is_gst_var.set(                    bool(d.get("is_gst_invoice", 1)))
        self.is_no_charge.set(                  bool(d.get("is_complimentary", 0)))

        try:    items_list = json.loads(d.get("items") or "[]")
        except json.JSONDecodeError: items_list = []
        for it in items_list: self.add_item(item=it)

        self.form_title.config(text=f"Edit Invoice  —  #{invoice_id}")
        self.save_btn.config(text="✔  Update Invoice")
        self._toggle_no_charge()

    # --  save  ----------------------------------------------------------------
    def save_invoice(self):
        # ---- Validate items
        items = []
        for w in self.item_widgets:
            try:
                amt  = float(w["vars"]["amount"].get()   or 0)
                disc = float(w["vars"]["discount"].get()  or 0)
            except ValueError:
                return Messagebox.show_error("Validation", "All amount/discount fields must be numbers.")
            if amt < 0 or disc < 0:
                return Messagebox.show_error("Validation", "Amounts cannot be negative.")
            if disc > amt:
                return Messagebox.show_error("Validation", "Discount cannot exceed amount.")
            items.append({"title": w["vars"]["title"].get(),
                          "description": w["vars"]["description"].get(),
                          "amount": amt, "discount": disc})

        def _flt(widget, name: str) -> Tuple[Optional[float], bool]:
            try:   return float(widget.get() or 0), True
            except ValueError:
                Messagebox.show_error("Validation", f"{name} must be a number.")
                return None, False

        tax, ok = _flt(self.tax_rate_entry, "Tax Rate");       
        if not ok: return
        gd,  ok = _flt(self.global_discount_entry, "Global Discount"); 
        if not ok: return
        tax = tax if self.is_gst_var.get() else 0.0

        is_comp = self.is_no_charge.get()
        totals  = DatabaseManager.calculate_totals(items, tax, gd, self.global_discount_type_combo.get())

        if is_comp:
            advance_amt = amount_paid = 0.0
            status = InvoiceStatus.COMPLIMENTARY.value
        else:
            advance_amt, ok = _flt(self.advance_amount_entry, "Advance Amount"); 
            if not ok: return
            amount_paid, ok = _flt(self.amount_paid_entry,    "Amount Paid");    
            if not ok: return
            total_paid = advance_amt + amount_paid
            if totals["grand_total"] > 0 and total_paid >= totals["grand_total"]:
                status = InvoiceStatus.PAID.value
            elif total_paid > 0:
                status = InvoiceStatus.PARTIAL.value
            else:
                status = InvoiceStatus.DUE.value

        # ---- Parse dates
        def _parse_date(entry, label: str) -> Tuple[Optional[str], bool]:
            val = entry.entry.get()
            if not val: return None, True
            try:   return datetime.datetime.strptime(val, DISPLAY_DATE_FORMAT).strftime(DB_DATE_FORMAT), True
            except ValueError:
                Messagebox.show_error("Validation", f"{label} is invalid.")
                return None, False

        inv_date, ok = _parse_date(self.invoice_date_entry, "Invoice Date");
        if not ok or not inv_date:
            return Messagebox.show_error("Validation", "Invoice Date is required.")

        adv_date = pay_date = None
        if not is_comp:
            if self.advance_date_entry.entry.get() and advance_amt > 0:
                adv_date, ok = _parse_date(self.advance_date_entry, "Advance Date")
                if not ok: return
            if status in (InvoiceStatus.PAID.value, InvoiceStatus.PARTIAL.value):
                pay_date, ok = _parse_date(self.payment_date_entry, "Payment Date")
                if not ok: return

        # ---- Build payload
        invoice_data = {
            "invoice_number":        self.invoice_number_entry.get().strip(),
            "bill_to":               self.bill_to_entry.get().strip(),
            "client_contact":        self.client_contact_text.get("1.0","end-1c").strip(),
            "client_gstin":          self.client_gstin_entry.get().strip() if self.is_gst_var.get() else "",
            "your_name":             self.your_name_entry.get().strip(),
            "your_contact":          self.your_contact_text.get("1.0","end-1c").strip(),
            "items":                 json.dumps(items),
            "amount_paid":           amount_paid,
            "tax_rate":              tax,
            "invoice_date":          inv_date,
            "status":                status,
            "payment_date":          pay_date,
            "payment_terms":         "" if is_comp else self.payment_terms_entry.get().strip(),
            "payment_mode":          None if is_comp or status not in (
                                        InvoiceStatus.PAID.value, InvoiceStatus.PARTIAL.value
                                     ) else self.payment_mode_combo.get(),
            "show_qr_code":          0 if is_comp else self.show_qr_var.get(),
            "is_complimentary":      1 if is_comp else 0,
            "is_gst_invoice":        1 if self.is_gst_var.get() else 0,
            "global_discount_value": gd,
            "global_discount_type":  self.global_discount_type_combo.get(),
            "advance_amount":        advance_amt,
            "advance_date":          adv_date,
        }

        if not all(invoice_data.get(k) for k in ["invoice_number","bill_to","your_name"]):
            return Messagebox.show_error("Required Fields",
                                         "Invoice #, Client Name, and Your Name are required.")

        ok, err = self.db_manager.save_invoice(invoice_data, self.is_editing,
                                               original_id=self.current_id if self.is_editing else None)
        if ok:
            Messagebox.show_info("Saved", f"Invoice {invoice_data['invoice_number']} saved successfully!")
            self.controller.show_frame(DashboardFrame)
        else:
            Messagebox.show_error("Database Error", f"Could not save invoice.\n{err}")


# =============================================================================
#  MAIN APP WINDOW
# =============================================================================
class InvoiceApp(ttk.Window):

    def __init__(self, db_manager: DatabaseManager):
        theme = load_config().get("theme", "litera")
        super().__init__(themename=theme)
        self.title(APP_NAME)
        self.geometry("1300x880")
        self.minsize(1100, 760)

        self.db_manager = db_manager
        self.app_logo_image = load_app_logo_image(self)
        self._apply_window_branding()
        self.task_queue = queue.Queue()
        self._prog_win  = None

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        self.frames: Dict[Type, ttk.Frame] = {
            F: F(container, self) for F in (DashboardFrame, InvoiceFormFrame)
        }
        for f in self.frames.values():
            f.grid(row=0, column=0, sticky="nsew")

        # Status bar
        ttk.Separator(self, bootstyle="secondary").pack(side="bottom", fill="x")
        self.status_bar = ttk.Label(self, text="Ready",
                                    font=(FONT_FAMILY, 9), bootstyle="secondary",
                                    padding=(14, 4), anchor="w")
        self.status_bar.pack(side="bottom", fill="x")

        self.after(100, self._poll_queue)
        self.show_frame(DashboardFrame)

    def _apply_window_branding(self):
        if APP_LOGO_ICO.exists():
            try:
                self.iconbitmap(default=str(APP_LOGO_ICO))
            except Exception as exc:
                logger.warning(f"Could not set window icon bitmap: {exc}")
        if APP_LOGO_PNG.exists() and self.app_logo_image is not None:
            try:
                self.iconphoto(True, self.app_logo_image)
            except Exception as exc:
                logger.warning(f"Could not set window icon image: {exc}")

    # --  queue + progress  ----------------------------------------------------
    def _poll_queue(self):
        try:
            msg = self.task_queue.get_nowait()
            self.show_progress(False)
            if msg["status"] == "success":
                self.set_status(f"PDF ready  —  {msg['path'].name}")
                PDFViewer(self, msg["path"])
            else:
                self.set_status("PDF generation failed.")
                Messagebox.show_error("PDF Error", f"Could not generate PDF.\n{msg['message']}")
        except queue.Empty:
            pass
        finally:
            self.after(100, self._poll_queue)

    def show_frame(self, frame_class: Type, invoice_id: Optional[str] = None):
        frame = self.frames[frame_class]
        if hasattr(frame, "on_show"):
            frame.on_show(invoice_id)
        frame.tkraise()
        self.set_status({DashboardFrame: "Dashboard", InvoiceFormFrame: "Invoice Form"}.get(frame_class, ""))

    def set_status(self, msg: str):
        self.status_bar.config(text=msg)

    def show_progress(self, show: bool):
        if show and not self._prog_win:
            pw = tk.Toplevel(self)
            pw.title("Please wait…")
            pw.attributes("-toolwindow", 1)
            pw.resizable(False, False)
            self._prog_win = pw
            ttk.Label(pw, text="Generating PDF…", padding=(30, 12)).pack()
            pb = ttk.Progressbar(pw, bootstyle="success-striped",
                                 mode="indeterminate", length=240)
            pb.pack(padx=30, pady=(0, 22))
            pb.start(10)
            pw.transient(self); pw.grab_set(); pw.update_idletasks()
            x = self.winfo_x() + (self.winfo_width()  - pw.winfo_width())  // 2
            y = self.winfo_y() + (self.winfo_height() - pw.winfo_height()) // 2
            pw.geometry(f"+{x}+{y}")
        elif not show and self._prog_win:
            self._prog_win.destroy()
            self._prog_win = None


# =============================================================================
#  ENTRY POINT
# =============================================================================
if __name__ == "__main__":
    try:
        InvoiceApp(DatabaseManager(DB_NAME)).mainloop()
    except Exception as exc:
        logger.critical(f"Unhandled exception: {exc}", exc_info=True)