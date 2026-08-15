import os
import re
import sys
import shutil
import smtplib
import subprocess
import random
from pathlib import Path
from datetime import datetime, timedelta
from email.message import EmailMessage
from tkinter import Tk, Toplevel, StringVar, BooleanVar, END, W, filedialog, messagebox
from tkinter import ttk
from fpdf import FPDF
import pandas as pd
from io import BytesIO
from contextlib import contextmanager
import hashlib
import secrets
import threading
import hmac
import json
import time
import socket
import getpass
from dateutil.relativedelta import relativedelta

try:
    from reportlab.pdfgen import canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfbase.ttfonts import TTFont
    from pypdf import PdfReader, PdfWriter
except ImportError:
    canvas = None
    stringWidth = None
    PdfReader = None
    PdfWriter = None


def app_dir():
    """Folder containing the program. Works when frozen by PyInstaller."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _bundled_dir(name):
    """Find a data folder next to the program, or inside a PyInstaller bundle."""
    meipass = getattr(sys, "_MEIPASS", None)
    for base in ([Path(meipass)] if meipass else []) + [app_dir()]:
        candidate = base / name
        if candidate.exists():
            return candidate
    return app_dir() / name


# Unicode fonts for PDFs. Without them, non-Latin characters in customer
# names would be printed as "?" (the built-in PDF fonts are Latin-1 only).
FONTS_DIR = _bundled_dir("fonts")
_UNICODE_TTF = {
    "": FONTS_DIR / "DejaVuSans.ttf",
    "B": FONTS_DIR / "DejaVuSans-Bold.ttf",
    "I": FONTS_DIR / "DejaVuSans-Oblique.ttf",
}
UNICODE_FONTS = all(p.exists() for p in _UNICODE_TTF.values())

LETTER_FONT = "Helvetica"
if UNICODE_FONTS and canvas is not None:
    try:
        pdfmetrics.registerFont(TTFont("DejaVuSans", str(_UNICODE_TTF[""])))
        LETTER_FONT = "DejaVuSans"
    except Exception:
        pass


# ============================================================
# Time & Tune — Ernest Money Tracker
# ============================================================
# Multi-user version:
# - Point every computer at the same shared/network data folder.
# - Email credentials are never hard-coded in the program.
# - Use the Email Settings screen or environment variables for SMTP settings.
# ============================================================

APP_NAME = "Time & Tune — Ernest Money Tracker"
DATA_DIR = Path.home() / "TimeAndTune_Ernest_Data"
RECORDS_DIR = DATA_DIR / "Records"
RECORDS_XLSX = DATA_DIR / "records.xlsx"
RETURNED_XLSX = DATA_DIR / "returned_records.xlsx"
EMAILED_FILE = DATA_DIR / "emailed_customers.txt"

# Persistent sequence for generated letter references: 01-EM, 02-EM, ...
REFERENCE_COUNTER_FILE = DATA_DIR / "letter_reference_counter.txt"

TEMPLATE_FILENAME = "blank temp 1.pdf"
TEMPLATE_CONFIG = DATA_DIR / "template_path.txt"
TEMPLATE_PATH = app_dir() / TEMPLATE_FILENAME


def resolve_template_path(interactive=True):
    """Find the TIME & TUNE template reliably, even if its filename was URL-encoded.

    interactive=False skips the file-picker fallback and raises instead, so
    background email threads never open a dialog.
    """
    global TEMPLATE_PATH

    # 1) A previously selected/saved template always wins if it still exists.
    if TEMPLATE_CONFIG.exists():
        try:
            saved_text = TEMPLATE_CONFIG.read_text(encoding="utf-8").strip()
            if saved_text:
                saved = Path(saved_text)
                if saved.exists() and saved.is_file():
                    TEMPLATE_PATH = saved
                    return saved
        except Exception:
            pass

    script_dir = app_dir()

    # 2) Try the normal filename plus the URL-encoded filename that can be
    # produced when a PDF is downloaded/exported from a web source.
    filenames = [
        "blank temp 1.pdf",
        "blank%20temp%201.pdf",
    ]

    bases = [script_dir, Path.cwd(), DATA_DIR,
             Path.home() / "Downloads",
             Path.home() / "Desktop"]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        # Template bundled inside a PyInstaller exe.
        bases.insert(1, Path(meipass))

    candidates = []
    for base in bases:
        for name in filenames:
            candidates.append(base / name)

    for candidate in candidates:
        if candidate.exists() and candidate.is_file():
            TEMPLATE_PATH = candidate
            try:
                DATA_DIR.mkdir(parents=True, exist_ok=True)
                TEMPLATE_CONFIG.write_text(str(candidate), encoding="utf-8")
            except Exception:
                pass
            return candidate

    # 3) Last automatic attempt: find a PDF in the program folder whose
    # normalized filename contains TIME, TUNE and TEMPLATE.
    try:
        for candidate in script_dir.glob("*.pdf"):
            normalized = candidate.name.lower().replace("%20", " ")
            normalized = re.sub(r"[^a-z0-9]+", " ", normalized)
            if all(word in normalized.split() for word in ("time", "tune", "template")):
                TEMPLATE_PATH = candidate
                try:
                    DATA_DIR.mkdir(parents=True, exist_ok=True)
                    TEMPLATE_CONFIG.write_text(str(candidate), encoding="utf-8")
                except Exception:
                    pass
                return candidate
    except Exception:
        pass

    # 4) If it still cannot be found, let the user choose it once. The chosen
    # path is remembered in template_path.txt for future runs.
    if not interactive:
        raise FileNotFoundError(
            "The TIME & TUNE letter template could not be located. "
            "Open the app and create or send any letter once to select it."
        )
    selected = filedialog.askopenfilename(
        title="Locate the TIME & TUNE letter template",
        filetypes=[
            ("PDF template", "*.pdf"),
            ("All files", "*.*"),
        ],
    )

    if selected:
        selected_path = Path(selected)
        TEMPLATE_PATH = selected_path
        try:
            DATA_DIR.mkdir(parents=True, exist_ok=True)
            TEMPLATE_CONFIG.write_text(str(selected_path), encoding="utf-8")
        except Exception:
            pass
        return selected_path

    raise FileNotFoundError(
        "TIME AND TUNE TEMPLATE.pdf could not be located.\n\n"
        "Put the template PDF beside the Python program, or select it "
        "when prompted."
    )

ACTIVE_COLUMNS = [
    "Customer", "Bank", "Quotation", "Payorder",
    "Date", "Initial Amount", "Ernest", "Item",
    # Letter/template information. Existing records remain compatible because
    # read_excel() adds any missing columns automatically.
    "Reference No", "Letter Date", "Recipient", "Organization", "P.O. Box",
    "City", "Phone", "Extension", "Fax", "Subject",
    "Tender No", "CDR/Pay Order No", "Letter Amount"
]
RETURNED_COLUMNS = [
    "Customer", "Return Bank", "Payorder", "Date", "Amount"
]

# ----------------------------- Paths / storage -----------------------------

def ensure_storage():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    RECORDS_DIR.mkdir(parents=True, exist_ok=True)

    if not RECORDS_XLSX.exists():
        pd.DataFrame(columns=ACTIVE_COLUMNS).to_excel(RECORDS_XLSX, index=False)

    if not RETURNED_XLSX.exists():
        pd.DataFrame(columns=RETURNED_COLUMNS).to_excel(RETURNED_XLSX, index=False)


def _highest_existing_reference():
    """Find the highest existing NN-EM reference in the active database."""
    highest = 0
    try:
        df = read_excel(RECORDS_XLSX, ACTIVE_COLUMNS)
        if "Reference No" in df.columns:
            for raw in df["Reference No"].fillna("").astype(str):
                match = re.fullmatch(r"\s*(\d+)\s*-\s*EM\s*", raw, re.IGNORECASE)
                if match:
                    highest = max(highest, int(match.group(1)))
    except Exception:
        pass
    return highest


def next_letter_reference():
    """
    Generate a permanent reference such as 01-EM, 02-EM, 03-EM.
    The counter survives closing/reopening the program.
    """
    highest = _highest_existing_reference()

    try:
        if REFERENCE_COUNTER_FILE.exists():
            raw = REFERENCE_COUNTER_FILE.read_text(encoding="utf-8").strip()
            if raw.isdigit():
                highest = max(highest, int(raw))
    except Exception:
        pass

    next_number = highest + 1

    try:
        REFERENCE_COUNTER_FILE.parent.mkdir(parents=True, exist_ok=True)
        REFERENCE_COUNTER_FILE.write_text(str(next_number), encoding="utf-8")
    except Exception as exc:
        raise RuntimeError(
            "Could not save the letter reference counter.\n\n"
            f"{REFERENCE_COUNTER_FILE}\n\n{exc}"
        )

    return f"{next_number:02d}-EM"


def safe_name(value):
    value = str(value).strip()
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    value = value.rstrip(". ")
    return value or "Unnamed"


def customer_folder(customer):
    return RECORDS_DIR / safe_name(customer)


def open_path(path):
    path = Path(path)
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as exc:
        messagebox.showerror("Open failed", f"Could not open:\n{path}\n\n{exc}")


def read_excel(path, columns, silent=False):
    # silent=True is required on background-thread and lock-holding paths:
    # a messagebox there either crashes Tk or stalls every other PC waiting
    # on the shared-data lock.
    path = Path(path)
    try:
        if not path.exists():
            return pd.DataFrame(columns=columns)
        df = pd.read_excel(path)
        for col in columns:
            if col not in df.columns:
                df[col] = ""
        return df[columns]
    except Exception as exc:
        if silent:
            print(f"[read_excel] {path}: {exc}", file=sys.stderr)
        else:
            messagebox.showerror(
                "Could not read data",
                f"Could not read:\n{path}\n\n{exc}"
            )
        return pd.DataFrame(columns=columns)


def write_excel(path, df, silent=False):
    path = Path(path)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_name(path.stem + "_tmp.xlsx")
        df.to_excel(tmp, index=False)
        os.replace(tmp, path)
        return True
    except PermissionError:
        if silent:
            print(f"[write_excel] file open elsewhere: {path}", file=sys.stderr)
        else:
            messagebox.showerror(
                "Excel file is open",
                f"Close this file and try again:\n\n{path}"
            )
    except Exception as exc:
        if silent:
            print(f"[write_excel] {path}: {exc}", file=sys.stderr)
        else:
            messagebox.showerror(
                "Save failed",
                f"Could not save:\n{path}\n\n{exc}"
            )
    return False


# ----------------------------- PDF generation -----------------------------

def pdf_safe(value):
    # With the bundled DejaVu fonts, any text is fine. Without them, the
    # built-in PDF fonts are Latin-1 only: replace instead of crashing.
    if UNICODE_FONTS:
        return str(value)
    return str(value).encode("latin-1", "replace").decode("latin-1")


def create_record_pdf(data, folder):
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / "Record.pdf"

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()

    if UNICODE_FONTS:
        family = "DejaVu"
        for style, font_path in _UNICODE_TTF.items():
            pdf.add_font(family, style, str(font_path))
    else:
        family = "Arial"

    pdf.set_fill_color(15, 23, 42)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(family, "B", 20)
    pdf.cell(0, 16, pdf_safe("TIME & TUNE"), ln=True, align="C", fill=True)

    pdf.set_text_color(15, 23, 42)
    pdf.set_font(family, "B", 15)
    pdf.cell(0, 12, pdf_safe("Ernest Money Record"), ln=True)
    pdf.ln(3)

    for key, value in data.items():
        pdf.set_font(family, "B", 10)
        pdf.cell(52, 8, pdf_safe(key))
        pdf.set_font(family, "", 10)
        # fpdf2: reset x to the left margin after each row, otherwise the
        # cursor drifts right and eventually runs out of horizontal space.
        pdf.multi_cell(0, 8, pdf_safe(value), new_x="LMARGIN", new_y="NEXT")

    pdf.ln(8)
    pdf.set_font(family, "I", 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 8, pdf_safe(f"Generated: {datetime.now():%d %b %Y, %I:%M %p}"),
             ln=True)

    pdf.output(str(path))
    return path


def create_letter_pdf(data, folder):
    """
    Create Letter.pdf using the blank TIME & TUNE template as the
    background. All variable letter text and table contents are drawn
    by Python.
    """

    if canvas is None or PdfReader is None or PdfWriter is None:
        raise RuntimeError(
            "PDF letter support is unavailable. Install reportlab and pypdf."
        )

    # ============================================================
    # BLANK TEMPLATE
    # ============================================================
    # The file-picker fallback is only allowed on the Tk main thread.
    template = resolve_template_path(
        interactive=threading.current_thread() is threading.main_thread()
    )

    output = Path(folder) / "Letter.pdf"

    reader = PdfReader(str(template))

    if not reader.pages:
        raise ValueError("The blank template PDF has no pages.")

    template_page = reader.pages[0]

    width = float(template_page.mediabox.width)
    height = float(template_page.mediabox.height)

    # ============================================================
    # GLOBAL ALIGNMENT ADJUSTMENT
    #
    # Everything is shifted:
    #   RIGHT = +8 points
    #   DOWN  = -10 points
    #
    # PDF coordinates start from the bottom-left.
    # ============================================================

    X_SHIFT = 8
    Y_SHIFT = 10

    def X(value):
        return value + X_SHIFT

    def Y(value):
        return height - value - Y_SHIFT

    # ============================================================
    # CREATE OVERLAY
    # ============================================================

    buffer = BytesIO()

    c = canvas.Canvas(
        buffer,
        pagesize=(width, height)
    )

    c.setFillColorRGB(0, 0, 0)

    # ============================================================
    # REFERENCE + DATE
    # ============================================================

    c.setFont(LETTER_FONT, 10.5)

    c.drawString(
        X(60),
        Y(92),
        "Ref:"
    )

    c.drawString(
        X(440),
        Y(92),
        "Date:"
    )

    reference = str(
        data.get("Reference No", "")
    ).strip()

    c.drawString(
        X(115),
        Y(92),
        reference
    )

    # Letter date
    letter_date = str(
        data.get("Letter Date", "")
    ).strip()

    if not letter_date:
        raw_date = str(
            data.get("Date", "")
        ).strip()

        try:
            letter_date = datetime.strptime(
                raw_date,
                "%Y-%m-%d"
            ).strftime("%d.%m.%Y")

        except ValueError:
            letter_date = raw_date

    c.drawString(
        X(500),
        Y(92),
        letter_date
    )

    # ============================================================
    # ADDRESS BLOCK
    # ============================================================

    # Slightly larger as requested
    c.setFont(LETTER_FONT, 11.5)

    recipient = str(
        data.get("Recipient", "")
    ).strip()

    organization = str(
        data.get("Organization", "")
    ).strip()

    po_box = str(
        data.get("P.O. Box", "")
    ).strip()

    city = str(
        data.get("City", "")
    ).strip()

    address_lines = []

    if recipient:
        address_lines.append(recipient)

    if organization:
        address_lines.append(organization)

    if po_box:
        address_lines.append(
            f"P.O. Box {po_box}"
        )

    if city:
        address_lines.append(city)

    # Address starts below Ref/Date
    address_x = 60
    address_y = 130

    for line in address_lines:
        c.drawString(
            X(address_x),
            Y(address_y),
            line
        )

        # Slightly more breathing room for 11.5 pt font
        address_y += 15

    # ============================================================
    # PHONE / EXTENSION / FAX
    # ============================================================

    phone = str(
        data.get("Phone", "")
    ).strip()

    extension = str(
        data.get("Extension", "")
    ).strip()

    fax = str(
        data.get("Fax", "")
    ).strip()

    # Fax on its own line
    fax_line = f"Fax: {fax}"

    # Phone on the line below
    phone_line = (
        f"Ph: {phone}, "
        f"ext: {extension}"
    )

    c.setFont(LETTER_FONT, 10.5)

    # Fax
    c.drawString(
        X(60),
        Y(205),
        fax_line
    )

    # Phone
    c.drawString(
        X(60),
        Y(219),
        phone_line
    )

    # ============================================================
    # SUBJECT
    # ============================================================

    # ============================================================
    # SUBJECT
    # ============================================================

    subject = "Return of Earnest money"

    c.setFont(LETTER_FONT, 10.5)

    c.drawString(
        X(60),
        Y(260),
        "Subject:"
    )

    c.drawString(
        X(118),
        Y(260),
        subject
    )
    # ============================================================
    # GREETING
    # ============================================================

    c.setFont(LETTER_FONT, 10.5)

    c.drawString(
        X(60),
        Y(305),
        "Dear Sir,"
    )

    # ============================================================
    # MAIN PARAGRAPH
    # ============================================================

    c.setFont(LETTER_FONT, 10.5)

    paragraph = (
        "We draw your kind attention for our outstanding "
        "earnest money. Detail:"
    )

    c.drawString(
        X(60),
        Y(360),
        paragraph
    )

    # ============================================================
    # TABLE
    # ============================================================

    table_left = 48
    table_right = 542

    table_top = 395
    table_bottom = 450

    table_width = table_right - table_left
    table_height = table_bottom - table_top

    # Column positions
    col1 = table_left
    col2 = 140
    col3 = 360
    col4 = 480
    col5 = table_right

    header_bottom = table_top + 27

    # ------------------------------------------------------------
    # Outer border
    # ------------------------------------------------------------

    c.setLineWidth(1.2)

    c.rect(
        X(table_left),
        Y(table_bottom),
        table_width,
        table_height,
        stroke=1,
        fill=0
    )

    # ------------------------------------------------------------
    # Header separator
    # ------------------------------------------------------------

    c.line(
        X(table_left),
        Y(header_bottom),
        X(table_right),
        Y(header_bottom)
    )

    # ------------------------------------------------------------
    # Column separators
    # ------------------------------------------------------------

    c.line(
        X(col2),
        Y(table_bottom),
        X(col2),
        Y(table_top)
    )

    c.line(
        X(col3),
        Y(table_bottom),
        X(col3),
        Y(table_top)
    )

    c.line(
        X(col4),
        Y(table_bottom),
        X(col4),
        Y(table_top)
    )

    # ------------------------------------------------------------
    # Header text
    # ------------------------------------------------------------

    c.setFont(LETTER_FONT, 9.5)

    c.drawString(
        X(col1 + 6),
        Y(table_top + 18),
        "S. NO"
    )

    c.drawString(
        X(col2 + 7),
        Y(table_top + 18),
        "TENDER NO"
    )

    c.drawString(
        X(col3 + 7),
        Y(table_top + 18),
        "CDR/PAY ORDER NO"
    )

    c.drawString(
        X(col4 + 7),
        Y(table_top + 18),
        "AMOUNT"
    )

    # ============================================================
    # TABLE DATA
    # ============================================================

    tender_no = str(
        data.get(
            "Tender No",
            data.get("Quotation", "")
        )
    ).strip()

    pay_order = str(
        data.get(
            "CDR/Pay Order No",
            data.get("Payorder", "")
        )
    ).strip()

    amount = str(
        data.get(
            "Letter Amount",
            data.get("Ernest", "")
        )
    ).strip()

    # Format amount
    try:
        numeric_amount = float(
            amount.replace(",", "")
        )

        if numeric_amount.is_integer():
            amount = f"{int(numeric_amount):,}"
        else:
            amount = f"{numeric_amount:,.2f}"

    except (ValueError, TypeError):
        pass

    # ------------------------------------------------------------
    # Table row
    # ------------------------------------------------------------

    c.setFont(LETTER_FONT, 10)

    row_y = table_bottom - 10

    # S. No
    c.drawCentredString(
        X((col1 + col2) / 2),
        Y(row_y),
        "1"
    )

    # Tender
    c.drawString(
        X(col2 + 8),
        Y(row_y),
        tender_no
    )

    # Pay order
    c.drawString(
        X(col3 + 8),
        Y(row_y),
        pay_order
    )

    # Amount
    c.drawString(
        X(col4 + 8),
        Y(row_y),
        amount
    )

    # ============================================================
    # ============================================================
    # REQUEST FOR RELEASE
    # ============================================================

    c.setFont(LETTER_FONT, 10.5)

    paragraph = (
        "We have completed supply of stores and have received payment. "
        "The warranty period of the subject stores has also expired. "
        "Kindly release our earnest money at the earliest."
    )

    text = c.beginText()
    text.setTextOrigin(X(60), Y(490))
    text.setFont(LETTER_FONT, 10.5)
    text.setLeading(14)

    words = paragraph.split()
    line = ""
    max_width = 470

    for word in words:
        test_line = f"{line} {word}".strip()

        if stringWidth(test_line, LETTER_FONT, 10.5) <= max_width:
            line = test_line
        else:
            text.textLine(line)
            line = word

    if line:
        text.textLine(line)

    c.drawText(text)
    # ============================================================
    # CLOSING
    # ============================================================

    c.setFont(LETTER_FONT, 10.5)

    c.drawString(
        X(60),
        Y(555),
        "Yours truly,"
    )

    # ============================================================
    # NAME / COMPANY
    # ============================================================

    c.drawString(
        X(60),
        Y(620),
        "Accounts"
    )

    c.drawString(
        X(60),
        Y(636),
        "TIME & TUNE"
    )

    # ============================================================
    # FINISH
    # ============================================================

    c.save()

    buffer.seek(0)

    overlay = PdfReader(buffer)

    template_page.merge_page(
        overlay.pages[0]
    )

    writer = PdfWriter()
    writer.add_page(template_page)

    output.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    with open(output, "wb") as f:
        writer.write(f)

    return output
# ----------------------------- Multi-user application -----------------------------

CLIENT_CONFIG_DIR = Path.home() / "TimeAndTune_Ernest_Client"
CLIENT_CONFIG_FILE = CLIENT_CONFIG_DIR / "shared_data_path.txt"
DEFAULT_LOCAL_DATA_DIR = Path.home() / "TimeAndTune_Ernest_Data"

USERS_XLSX_NAME = "users.xlsx"
AUDIT_XLSX_NAME = "audit_log.xlsx"
EMAIL_HISTORY_XLSX_NAME = "email_history.xlsx"
SETTINGS_JSON_NAME = "settings.json"
SMTP_PASSWORD_FILE_NAME = "smtp_password.txt"
LOCK_FILE_NAME = ".ernest_data.lock"
BACKUP_DIR_NAME = "Backups"

USERS_COLUMNS = [
    "User ID", "Full Name", "Role", "Password Hash",
    "Active", "Created At", "Last Login"
]
AUDIT_COLUMNS = [
    "Timestamp", "User ID", "Action", "Record Reference",
    "Customer", "Details", "Computer"
]
EMAIL_HISTORY_COLUMNS = [
    "Timestamp", "Reference No", "Customer", "Recipient",
    "Subject", "Status", "Attempts", "Error", "User ID"
]

ROLE_ADMIN = "Admin"
ROLE_EMPLOYEE = "Employee"
ROLE_VIEWER = "Viewer"

REMINDER_MONTHS = 15
LOCK_TIMEOUT_SECONDS = 30


def auto_check_ms():
    """Automatic reminder-check interval, from the auto_check_minutes setting."""
    try:
        minutes = int(load_settings().get("auto_check_minutes", 5))
    except Exception:
        minutes = 5
    return max(1, minutes) * 60 * 1000


def _load_shared_path():
    try:
        raw = CLIENT_CONFIG_FILE.read_text(encoding="utf-8").strip()
        if raw:
            p = Path(raw)
            if p.exists() and p.is_dir():
                return p
    except Exception:
        pass
    return None


def _choose_shared_data_folder(parent=None):
    CLIENT_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    selected = filedialog.askdirectory(
        title="Choose the shared Time & Tune data folder",
        parent=parent,
    )
    if not selected:
        return None
    p = Path(selected)
    p.mkdir(parents=True, exist_ok=True)
    CLIENT_CONFIG_FILE.write_text(str(p), encoding="utf-8")
    return p


def configure_data_paths():
    """Set the module's shared data paths before any storage function is used."""
    global DATA_DIR, RECORDS_DIR, RECORDS_XLSX, RETURNED_XLSX, EMAILED_FILE
    global REFERENCE_COUNTER_FILE, TEMPLATE_CONFIG

    configured = _load_shared_path()
    if configured is None:
        # Use a normal local folder until the user chooses a shared folder.
        # The login window offers the shared-folder configuration button.
        configured = DEFAULT_LOCAL_DATA_DIR

    DATA_DIR = configured
    RECORDS_DIR = DATA_DIR / "Records"
    RECORDS_XLSX = DATA_DIR / "records.xlsx"
    RETURNED_XLSX = DATA_DIR / "returned_records.xlsx"
    EMAILED_FILE = DATA_DIR / "emailed_customers.txt"
    REFERENCE_COUNTER_FILE = DATA_DIR / "letter_reference_counter.txt"
    TEMPLATE_CONFIG = DATA_DIR / "template_path.txt"


configure_data_paths()


def users_path():
    return DATA_DIR / USERS_XLSX_NAME


def audit_path():
    return DATA_DIR / AUDIT_XLSX_NAME


def email_history_path():
    return DATA_DIR / EMAIL_HISTORY_XLSX_NAME


def settings_path():
    return DATA_DIR / SETTINGS_JSON_NAME


def smtp_password_path():
    # Per-machine, NOT in the shared folder: everyone on the network could
    # read a plaintext app password stored beside the shared workbooks.
    return CLIENT_CONFIG_DIR / SMTP_PASSWORD_FILE_NAME


def backup_dir():
    return DATA_DIR / BACKUP_DIR_NAME


@contextmanager
def data_lock(timeout=LOCK_TIMEOUT_SECONDS):
    """
    Cross-process lock based on exclusive lock-file creation.
    This prevents two employee PCs from writing the same Excel file at once.
    """
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    lock = DATA_DIR / LOCK_FILE_NAME
    start = time.time()
    fd = None
    while True:
        try:
            fd = os.open(str(lock), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, f"{socket.gethostname()}|{os.getpid()}|{datetime.now().isoformat()}".encode())
            os.close(fd)
            fd = None
            break
        except FileExistsError:
            # Recover a stale lock left behind by a crashed client.
            try:
                age = time.time() - lock.stat().st_mtime
                if age > timeout + 10:
                    lock.unlink()
                    continue
            except Exception:
                pass
            if time.time() - start >= timeout:
                raise TimeoutError(
                    "The shared data folder is busy. Another employee is saving data.\n"
                    "Please wait a few seconds and try again."
                )
            time.sleep(0.25)
    try:
        yield
    finally:
        try:
            lock.unlink()
        except FileNotFoundError:
            pass
        except Exception:
            pass


def safe_write_excel(path, df):
    """Write an Excel file atomically while holding the shared-data lock."""
    # Write silently inside the lock and show the error dialog only after the
    # lock is released, so a user staring at a dialog cannot block other PCs.
    with data_lock():
        ok = write_excel(path, df, silent=True)
    if not ok:
        messagebox.showerror(
            "Save failed",
            f"Could not save:\n\n{path}\n\n"
            "Close the file if it is open in Excel and try again."
        )
    return ok


def ensure_workbook(path, columns):
    if not path.exists():
        write_excel(path, pd.DataFrame(columns=columns))


def default_settings():
    return {
        "smtp_host": os.getenv("ERNEST_SMTP_HOST", "smtp.gmail.com"),
        "smtp_port": int(os.getenv("ERNEST_SMTP_PORT", "465")),
        "smtp_security": os.getenv("ERNEST_SMTP_SECURITY", "SSL"),
        "sender_email": os.getenv("ERNEST_SENDER_EMAIL", ""),
        "recipient_email": os.getenv("ERNEST_RECIPIENT_EMAIL", ""),
        "smtp_password": "",
        "auto_check_minutes": 5,
        "reminder_months": REMINDER_MONTHS,
    }


def load_settings():
    data = default_settings()
    try:
        if settings_path().exists():
            saved = json.loads(settings_path().read_text(encoding="utf-8"))
            if isinstance(saved, dict):
                data.update(saved)
    except Exception:
        pass

    # Password is persisted separately. This also repairs older settings files
    # that were created before password persistence was added.
    try:
        password_file = smtp_password_path()
        # One-time migration: older versions kept the password in the shared
        # data folder. Move it to this machine and remove the shared copy.
        legacy = DATA_DIR / SMTP_PASSWORD_FILE_NAME
        if legacy != password_file and legacy.exists():
            if not password_file.exists():
                password_file.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(legacy, password_file)
                try:
                    os.chmod(password_file, 0o600)
                except Exception:
                    pass
            legacy.unlink()
        if password_file.exists():
            saved_password = password_file.read_text(encoding="utf-8").strip()
            if saved_password:
                data["smtp_password"] = saved_password
    except Exception:
        pass

    # Environment variables override shared settings for deployment/config.
    if os.getenv("ERNEST_SENDER_EMAIL"):
        data["sender_email"] = os.getenv("ERNEST_SENDER_EMAIL")
    if os.getenv("ERNEST_SMTP_HOST"):
        data["smtp_host"] = os.getenv("ERNEST_SMTP_HOST")
    if os.getenv("ERNEST_RECIPIENT_EMAIL"):
        data["recipient_email"] = os.getenv("ERNEST_RECIPIENT_EMAIL")
    if os.getenv("ERNEST_APP_PASSWORD"):
        data["smtp_password"] = os.getenv("ERNEST_APP_PASSWORD").strip()

    return data


def save_settings(settings):
    clean = dict(settings)
    clean["reminder_months"] = REMINDER_MONTHS
    password = str(clean.pop("smtp_password", "")).strip()

    with data_lock():
        tmp = settings_path().with_suffix(".tmp")
        tmp.write_text(json.dumps(clean, indent=2), encoding="utf-8")
        os.replace(tmp, settings_path())

        # Persist the sender password separately (per-machine, owner-only).
        # The application reads this file on the next launch, so Settings
        # never loses it.
        if password:
            smtp_password_path().parent.mkdir(parents=True, exist_ok=True)
            password_tmp = smtp_password_path().with_suffix(".tmp")
            password_tmp.write_text(password, encoding="utf-8")
            os.replace(password_tmp, smtp_password_path())
            try:
                os.chmod(smtp_password_path(), 0o600)
            except Exception:
                pass


def hash_password(password, salt=None):
    if salt is None:
        salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt, 260_000
    )
    return f"pbkdf2_sha256$260000${salt.hex()}${digest.hex()}"


def verify_password(password, stored):
    try:
        scheme, rounds, salt_hex, digest_hex = stored.split("$")
        if scheme != "pbkdf2_sha256":
            return False
        test = hashlib.pbkdf2_hmac(
            "sha256", password.encode("utf-8"),
            bytes.fromhex(salt_hex), int(rounds)
        ).hex()
        return hmac.compare_digest(test, digest_hex)
    except Exception:
        return False


def initialize_enterprise_storage():
    """Create the shared Excel workbooks without destroying existing data."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    RECORDS_DIR.mkdir(parents=True, exist_ok=True)
    ensure_workbook(RECORDS_XLSX, ACTIVE_COLUMNS)
    ensure_workbook(RETURNED_XLSX, RETURNED_COLUMNS)
    ensure_workbook(users_path(), USERS_COLUMNS)
    ensure_workbook(audit_path(), AUDIT_COLUMNS)
    ensure_workbook(email_history_path(), EMAIL_HISTORY_COLUMNS)
    if not settings_path().exists():
        save_settings(default_settings())


def add_audit(user, action, reference="", customer="", details=""):
    try:
        # Fully silent: audit runs on background email threads, where a
        # messagebox is not allowed.
        df = read_excel(audit_path(), AUDIT_COLUMNS, silent=True)
        row = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "User ID": user.get("User ID", "SYSTEM") if user else "SYSTEM",
            "Action": action,
            "Record Reference": reference,
            "Customer": customer,
            "Details": details,
            "Computer": socket.gethostname(),
        }
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        with data_lock():
            write_excel(audit_path(), df, silent=True)
    except Exception:
        pass


def record_email_history(reference, customer, recipient, subject, status,
                         attempts=1, error="", user_id="SYSTEM"):
    """Write an email result without silently discarding diagnostic errors."""
    row = {
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Reference No": str(reference),
        "Customer": str(customer),
        "Recipient": str(recipient),
        "Subject": str(subject),
        "Status": str(status),
        "Attempts": int(attempts or 0),
        "Error": str(error or ""),
        "User ID": str(user_id),
    }
    try:
        with data_lock():
            df = read_excel(email_history_path(), EMAIL_HISTORY_COLUMNS, silent=True)
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            write_excel(email_history_path(), df, silent=True)
        return True
    except Exception as exc:
        try:
            with (DATA_DIR / "email_errors.log").open("a", encoding="utf-8") as f:
                f.write(
                    f"{row['Timestamp']} | {row['Reference No']} | "
                    f"{row['Customer']} | {row['Status']} | "
                    f"{row['Error']} | HISTORY WRITE ERROR: {exc}\n"
                )
        except Exception:
            pass
        return False


def calendar_due_date(value, months=REMINDER_MONTHS):
    """Return the exact calendar date that is `months` after a record date."""
    try:
        dt = pd.to_datetime(value, errors="raise").to_pydatetime()
        return dt.date() + relativedelta(months=months)
    except Exception:
        return None


def is_record_due(value):
    due = calendar_due_date(value, REMINDER_MONTHS)
    return bool(due and datetime.now().date() >= due)


def due_date_text(value):
    due = calendar_due_date(value, REMINDER_MONTHS)
    return due.strftime("%Y-%m-%d") if due else ""


def get_smtp_password():
    """Return the saved sender SMTP/App Password, with environment fallback."""
    try:
        saved = str(load_settings().get("smtp_password", "")).strip()
        if saved:
            return saved
    except Exception:
        pass
    return os.getenv("ERNEST_APP_PASSWORD", "").strip()


def smtp_send(msg, settings, password):
    host = str(settings.get("smtp_host", "smtp.gmail.com"))
    port = int(settings.get("smtp_port", 465))
    security = str(settings.get("smtp_security", "SSL")).upper()
    if security == "STARTTLS":
        with smtplib.SMTP(host, port, timeout=20) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(settings["sender_email"], password)
            smtp.send_message(msg)
    else:
        with smtplib.SMTP_SSL(host, port, timeout=20) as smtp:
            smtp.ehlo()
            smtp.login(settings["sender_email"], password)
            smtp.send_message(msg)



def claim_email_send(reference, customer, recipient, user_id):
    """
    Atomically claim a reminder so two PCs cannot both send it.
    Returns True only for the process that acquired the claim.
    """
    with data_lock():
        df = read_excel(email_history_path(), EMAIL_HISTORY_COLUMNS, silent=True)
        existing = df[df["Reference No"].astype(str).eq(str(reference))]
        if not existing.empty:
            statuses = existing["Status"].astype(str).tolist()
            if "SENT" in statuses:
                return False
            sending = existing[existing["Status"].astype(str).eq("SENDING")]
            if not sending.empty:
                try:
                    latest = pd.to_datetime(sending["Timestamp"], errors="coerce").max()
                    if pd.notna(latest):
                        age = (datetime.now() - latest.to_pydatetime()).total_seconds()
                        if age < 2 * 60:
                            return False
                    # A crashed/closed client can leave SENDING forever.
                    df.loc[sending.index, "Status"] = "RETRY"
                    df.loc[sending.index, "Error"] = "Previous automatic email attempt timed out; retrying now."
                except Exception:
                    return False
        row = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Reference No": reference,
            "Customer": customer,
            "Recipient": recipient,
            "Subject": f"Ernest Money Reminder — {customer}",
            "Status": "SENDING",
            "Attempts": 0,
            "Error": "",
            "User ID": user_id,
        }
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        return write_excel(email_history_path(), df, silent=True)


def finalize_email_claim(reference, status, attempts, error="", user_id="SYSTEM"):
    with data_lock():
        df = read_excel(email_history_path(), EMAIL_HISTORY_COLUMNS, silent=True)
        mask = (
            df["Reference No"].astype(str).eq(str(reference)) &
            df["Status"].astype(str).eq("SENDING")
        )
        if mask.any():
            df.loc[mask, "Status"] = status
            df.loc[mask, "Attempts"] = attempts
            df.loc[mask, "Error"] = error
            df.loc[mask, "Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            df.loc[mask, "User ID"] = user_id
            write_excel(email_history_path(), df, silent=True)
        else:
            # Preserve an auditable result even if another process repaired the row.
            row = {
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Reference No": reference,
                "Customer": "",
                "Recipient": "",
                "Subject": "Ernest Money Reminder",
                "Status": status,
                "Attempts": attempts,
                "Error": error,
                "User ID": user_id,
            }
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            write_excel(email_history_path(), df, silent=True)

def send_email(customer, letter_path, record_path, reference="",
               recipient=None, user_id="SYSTEM", silent=False):
    settings = load_settings()
    sender = str(settings.get("sender_email", "")).strip()
    recipient = str(recipient or settings.get("recipient_email", "")).strip()
    password = get_smtp_password()

    if not sender or not recipient or not password:
        err = (
            "Email is not configured. Open Settings and make sure Sender Email, "
            "Default Recipient and SMTP/App Password are saved."
        )
        record_email_history(
            reference, customer, recipient,
            f"Ernest Money Reminder — {customer}",
            "FAILED", 0, err, user_id
        )
        if not silent:
            messagebox.showwarning("Email not configured", err)
        return False

    subject = f"Ernest Money Reminder — {customer}"

    # If no reference was supplied, use a stable fallback so duplicate
    # protection still works for manual sends.
    if not reference:
        reference = f"MANUAL-{customer}"

    if not claim_email_send(reference, customer, recipient, user_id):
        if not silent:
            messagebox.showinfo(
                "Already handled",
                "This reminder has already been sent or is currently being "
                "handled by another computer."
            )
        return False

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = recipient
    msg.set_content(
        f"Dear Sir/Madam,\n\n"
        f"This is an automatic Ernest Money reminder from Time & Tune.\n\n"
        f"Customer/Company: {customer}\n"
        f"Reference: {reference}\n\n"
        f"The configured 15-month reminder period has been reached. "
        f"Please review the attached documents and take the necessary action.\n\n"
        f"Time & Tune Ernest Money Tracker"
    )

    try:
        for path in (letter_path, record_path):
            path = Path(path)
            if not path.exists():
                raise FileNotFoundError(f"Attachment not found: {path}")
            msg.add_attachment(
                path.read_bytes(),
                maintype="application",
                subtype="pdf",
                filename=path.name
            )

        last_error = ""
        for attempt in range(1, 3):
            try:
                smtp_send(msg, settings, password)
                finalize_email_claim(reference, "SENT", attempt, "", user_id)
                add_audit(
                    {"User ID": user_id},
                    "EMAIL SENT", reference, customer,
                    f"Reminder sent to {recipient}"
                )
                if not silent:
                    messagebox.showinfo(
                        "Email sent",
                        f"Reminder sent successfully for {customer}."
                    )
                return True
            except Exception as exc:
                last_error = f"{type(exc).__name__}: {exc}"
                if attempt < 2:
                    time.sleep(2)

        finalize_email_claim(reference, "FAILED", 2, last_error, user_id)
        add_audit(
            {"User ID": user_id},
            "EMAIL FAILED", reference, customer, last_error
        )
        if not silent:
            messagebox.showerror(
                "Email failed",
                "The email could not be sent after 2 attempts.\n\n"
                + last_error
            )
        return False

    except Exception as exc:
        err = f"{type(exc).__name__}: {exc}"
        finalize_email_claim(reference, "FAILED", 1, err, user_id)
        if not silent:
            messagebox.showerror("Email error", err)
        return False



def reminder_already_sent(reference):
    try:
        df = read_excel(email_history_path(), EMAIL_HISTORY_COLUMNS, silent=True)
        if df.empty:
            return False
        rows = df[
            (df["Reference No"].astype(str).eq(str(reference))) &
            (df["Status"].astype(str).eq("SENT"))
        ]
        return not rows.empty
    except Exception:
        return False


def backup_all_data():
    """Create a timestamped backup of all important shared files."""
    backup_dir().mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = backup_dir() / stamp
    target.mkdir(parents=True, exist_ok=True)

    for name in [
        RECORDS_XLSX.name, RETURNED_XLSX.name, users_path().name,
        audit_path().name, email_history_path().name, settings_path().name,
        REFERENCE_COUNTER_FILE.name,
    ]:
        source = DATA_DIR / name
        if source.exists():
            shutil.copy2(source, target / source.name)

    # Customer folders and generated PDFs are important too.
    if RECORDS_DIR.exists():
        shutil.copytree(RECORDS_DIR, target / RECORDS_DIR.name)

    # Keep only the latest 30 backup folders.
    backups = sorted(
        [p for p in backup_dir().iterdir() if p.is_dir()],
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    for old in backups[30:]:
        shutil.rmtree(old, ignore_errors=True)

    return target


class ErnestMoneyApp:
    def __init__(self, root, user):
        self.root = root
        self.user = user
        self._email_worker_busy = False
        self._last_auto_check = ""
        self.root.title(APP_NAME)
        self.root.geometry("1000x620")
        self.root.minsize(900, 560)
        self.root.configure(bg="#0b1120")

        initialize_enterprise_storage()
        self.configure_styles()
        self.build_main_ui()
        self.refresh_dashboard()
        add_audit(
            self.user, "LOGIN", "", "",
            f"Logged in from {socket.gethostname()}"
        )
        self.root.after(auto_check_ms(), self.auto_check)
        threading.Thread(target=self._daily_backup, daemon=True).start()

    def _daily_backup(self):
        """Once per day, on the first login of any PC: back up the shared data."""
        try:
            today = datetime.now().strftime("%Y%m%d")
            if backup_dir().exists() and any(
                p.is_dir() and p.name.startswith(today)
                for p in backup_dir().iterdir()
            ):
                return
            backup_all_data()
            add_audit(self.user, "AUTO BACKUP", "", "", "Daily backup created")
        except Exception as exc:
            print(f"[daily backup] {exc}", file=sys.stderr)

    @property
    def is_admin(self):
        return self.user.get("Role") == ROLE_ADMIN

    @property
    def can_edit(self):
        return self.user.get("Role") in (ROLE_ADMIN, ROLE_EMPLOYEE)

    def configure_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Treeview", background="#111827", foreground="#e5e7eb",
            fieldbackground="#111827", rowheight=34,
            font=("Segoe UI", 10)
        )
        style.configure(
            "Treeview.Heading", background="#1e293b", foreground="#ffffff",
            font=("Segoe UI Semibold", 10)
        )
        style.map(
            "Treeview",
            background=[("selected", "#2563eb")],
            foreground=[("selected", "#ffffff")]
        )
        style.configure(
            "TButton", font=("Segoe UI Semibold", 10), padding=(14, 9),
            background="#2563eb", foreground="white"
        )
        style.map("TButton", background=[("active", "#1d4ed8")])
        style.configure(
            "TEntry", padding=8, fieldbackground="#f8fafc",
            foreground="#0f172a"
        )

    def build_main_ui(self):
        header = ttk.Frame(self.root, padding=(30, 24))
        header.pack(fill="x")

        title = ttk.Label(
            header, text="TIME & TUNE",
            font=("Segoe UI", 25, "bold"),
            foreground="#f8fafc", background="#0b1120"
        )
        title.pack(side="left")

        subtitle = ttk.Label(
            header, text="  |  Ernest Money Tracker",
            font=("Segoe UI", 13),
            foreground="#94a3b8", background="#0b1120"
        )
        subtitle.pack(side="left", pady=(8, 0))

        ttk.Label(
            header,
            text=f"  {self.user['Full Name']} • {self.user['Role']}",
            foreground="#60a5fa", background="#0b1120",
            font=("Segoe UI Semibold", 10)
        ).pack(side="right", pady=8)

        ttk.Button(
            header, text="Settings",
            command=self.settings_window
        ).pack(side="right", padx=8)

        ttk.Button(
            header, text="Open Data Folder",
            command=lambda: open_path(DATA_DIR)
        ).pack(side="right", pady=4)

        cards = ttk.Frame(self.root, padding=(30, 0, 30, 20))
        cards.pack(fill="x")
        self.total_var = StringVar(value="0")
        self.due_var = StringVar(value="0")
        self.upcoming_var = StringVar(value="0")
        self.returned_var = StringVar(value="0")
        self.sent_var = StringVar(value="0")
        self.failed_var = StringVar(value="0")

        self.make_card(cards, "ACTIVE RECORDS", self.total_var, "#2563eb", 0)
        self.make_card(cards, "DUE / 15 MONTHS+", self.due_var, "#d97706", 1)
        self.make_card(cards, "DUE IN 30 DAYS", self.upcoming_var, "#eab308", 2)
        self.make_card(cards, "RETURNED", self.returned_var, "#059669", 3)
        self.make_card(cards, "EMAILS SENT", self.sent_var, "#8b5cf6", 4)
        self.make_card(cards, "EMAIL FAILURES", self.failed_var, "#dc2626", 5)

        actions = ttk.Frame(self.root, padding=(30, 0))
        actions.pack(fill="x")
        action_buttons = [
            ("＋  Add Record", self.add_record_window),
            ("▣  Records", self.records_window),
            ("⌕  Search", self.search_window),
            ("◷  Due Orders", self.due_orders_window),
            ("↻  Check Reminders", self.check_reminders),
            ("📧 Email History", self.email_history_window),
            ("📝 Audit Log", self.audit_window),
            ("🗑 Remove Records", self.admin_delete_menu),
        ]

        if self.is_admin:
            action_buttons += [
                ("👥 Users", self.users_window),
                ("💾 Backup", self.manual_backup),
                ("↺ Reset App", self.reset_application),
            ]
        # Keep all main-page action buttons in exactly two rows.
        # Nothing else in the main page is changed.
        for i, (button_text, command) in enumerate(action_buttons):
            row = 0 if i < 6 else 1
            col = i if i < 6 else i - 6
            btn = ttk.Button(
                actions, text=button_text, command=command, width=16
            )
            btn.grid(
                row=row, column=col,
                sticky="ew", padx=(0, 8), pady=(0, 8)
            )

        for col in range(6):
            actions.columnconfigure(col, weight=1)

        body = ttk.Frame(self.root, padding=(30, 22))
        body.pack(fill="both", expand=True)

        ttk.Label(
            body, text="Recent Records",
            font=("Segoe UI Semibold", 16),
            foreground="#f8fafc", background="#0b1120"
        ).pack(anchor="w", pady=(0, 10))

        table_frame = ttk.Frame(body)
        table_frame.pack(fill="both", expand=True)
        columns = (
            "Customer", "Bank", "Quotation", "Payorder",
            "Date", "Initial Amount", "Ernest", "Item"
        )
        self.main_tree = ttk.Treeview(
            table_frame, columns=columns, show="headings",
            selectmode="browse"
        )
        widths = {
            "Customer": 180, "Bank": 130, "Quotation": 115,
            "Payorder": 115, "Date": 105, "Initial Amount": 120,
            "Ernest": 110, "Item": 190
        }
        for col in columns:
            self.main_tree.heading(col, text=col)
            self.main_tree.column(col, width=widths[col], minwidth=80)

        yscroll = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.main_tree.yview
        )
        xscroll = ttk.Scrollbar(
            table_frame, orient="horizontal", command=self.main_tree.xview
        )
        self.main_tree.configure(
            yscrollcommand=yscroll.set, xscrollcommand=xscroll.set
        )
        self.main_tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.status_var = StringVar(
            value=f"Shared data: {DATA_DIR}"
        )
        ttk.Label(
            self.root, textvariable=self.status_var,
            foreground="#64748b", background="#0b1120",
            font=("Segoe UI", 9)
        ).pack(fill="x", padx=30, pady=(0, 12))

        self.main_tree.bind("<Double-1>", lambda _e: self.edit_selected_from_main())

    def make_card(self, parent, title, variable, color, column):
        frame = ttk.Frame(parent, padding=14)
        frame.grid(row=0, column=column, sticky="ew", padx=(0, 10))
        parent.columnconfigure(column, weight=1)
        ttk.Label(
            frame, text=title, foreground="#94a3b8",
            background="#111827", font=("Segoe UI Semibold", 9)
        ).pack(anchor="w")
        ttk.Label(
            frame, textvariable=variable, foreground=color,
            background="#111827", font=("Segoe UI", 22, "bold")
        ).pack(anchor="w", pady=(4, 0))

    def refresh_dashboard(self):
        # Re-read the shared files every time; this is what keeps all PCs current.
        df = read_excel(RECORDS_XLSX, ACTIVE_COLUMNS)
        returned = read_excel(RETURNED_XLSX, RETURNED_COLUMNS)
        self.total_var.set(str(len(df)))
        self.returned_var.set(str(len(returned)))

        if df.empty:
            due = upcoming = 0
        else:
            due = sum(is_record_due(v) for v in df["Date"])
            upcoming = len(self.upcoming_dataframe(df))
        self.due_var.set(str(due))
        self.upcoming_var.set(str(upcoming))

        try:
            hist = read_excel(email_history_path(), EMAIL_HISTORY_COLUMNS)
            sent = int((hist["Status"].astype(str) == "SENT").sum()) if not hist.empty else 0
            failed = int((hist["Status"].astype(str) == "FAILED").sum()) if not hist.empty else 0
        except Exception:
            sent = failed = 0
        self.sent_var.set(str(sent))
        self.failed_var.set(str(failed))

        if hasattr(self, "main_tree"):
            for item in self.main_tree.get_children():
                self.main_tree.delete(item)
            for _, row in df.tail(50).iloc[::-1].iterrows():
                self.main_tree.insert(
                    "", "end",
                    values=[str(row[c]) for c in ACTIVE_COLUMNS[:8]]
                )

        self.refresh_status()

    def refresh_status(self):
        """One-line health strip: catches silent misconfiguration early."""
        try:
            settings = load_settings()
            email_ok = bool(
                str(settings.get("sender_email", "")).strip()
                and str(settings.get("recipient_email", "")).strip()
                and get_smtp_password()
            )
        except Exception:
            email_ok = False
        try:
            resolve_template_path(interactive=False)
            template_ok = True
        except Exception:
            template_ok = False
        folder_ok = DATA_DIR.exists()
        self.status_var.set(
            f"Shared data: {DATA_DIR} ({'OK' if folder_ok else 'NOT FOUND'})"
            f"   |   Email: {'configured' if email_ok else 'NOT CONFIGURED'}"
            f"   |   Template: {'found' if template_ok else 'MISSING'}"
            f"   |   Last auto-check: {self._last_auto_check or 'not yet'}"
        )

    def current_record_from_values(self, values):
        df = read_excel(RECORDS_XLSX, ACTIVE_COLUMNS)
        if df.empty:
            return None
        customer, bank, quotation, payorder, date = [str(v) for v in values[:5]]
        mask = (
            df["Customer"].astype(str).eq(customer) &
            df["Bank"].astype(str).eq(bank) &
            df["Quotation"].astype(str).eq(quotation) &
            df["Payorder"].astype(str).eq(payorder) &
            df["Date"].astype(str).eq(date)
        )
        matches = df[mask]
        if matches.empty:
            return None
        return matches.iloc[0].copy()

    def edit_selected_from_main(self):
        sel = self.main_tree.selection()
        if not sel:
            return
        values = self.main_tree.item(sel[0], "values")
        row = self.current_record_from_values(values)
        if row is not None:
            self.edit_record_window(row)

    def add_record_window(self):
        self.record_editor_window(None)

    def edit_record_window(self, row):
        self.record_editor_window(row)

    def record_editor_window(self, row=None):
        editing = row is not None
        if editing and not self.can_edit:
            messagebox.showwarning(
                "Permission denied",
                "Your account is not allowed to edit records."
            )
            return

        win = Toplevel(self.root)
        win.title("Edit Ernest Money Record" if editing else "Add Ernest Money Record")
        # Compact window: the form itself is scrollable, while the action bar
        # stays visible at the bottom so Save/Update is always accessible.
        win.geometry("700x600")
        win.minsize(700, 560)
        win.configure(bg="#0f172a")
        win.resizable(True, True)
        win.transient(self.root)
        win.grab_set()

        ttk.Label(
            win, text="Edit Record" if editing else "Add Record",
            font=("Segoe UI", 22, "bold"),
            foreground="#f8fafc", background="#0f172a"
        ).pack(anchor="w", padx=30, pady=(25, 4))

        save_status = StringVar(value="Ready to save")
        ttk.Label(
            win, textvariable=save_status,
            foreground="#94a3b8", background="#0f172a"
        ).pack(anchor="w", padx=30, pady=(0, 14))

        outer = ttk.Frame(win, padding=(30, 0))
        outer.pack(fill="both", expand=True)
        canvas_frame = ttk.Frame(outer)
        canvas_frame.pack(fill="both", expand=True)
        form_canvas = __import__("tkinter").Canvas(
            canvas_frame, bg="#0f172a", highlightthickness=0, borderwidth=0
        )
        scrollbar = ttk.Scrollbar(
            canvas_frame, orient="vertical", command=form_canvas.yview
        )
        form = ttk.Frame(canvas_frame, padding=(0, 0, 12, 10))
        form_window = form_canvas.create_window((0, 0), window=form, anchor="nw")
        form_canvas.configure(yscrollcommand=scrollbar.set)
        form_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def update_scroll(_event=None):
            form_canvas.configure(scrollregion=form_canvas.bbox("all"))
            form_canvas.itemconfigure(form_window, width=form_canvas.winfo_width())
        form.bind("<Configure>", update_scroll)
        form_canvas.bind("<Configure>", update_scroll)

        entries = {}
        uploaded = {"Quotation": None, "Payorder": None}

        existing_fields = list(ACTIVE_COLUMNS[:8])
        letter_fields = [f for f in ACTIVE_COLUMNS[8:] if f != "Reference No"]

        ttk.Label(
            form, text="Record Information",
            font=("Segoe UI Semibold", 12),
            foreground="#60a5fa", background="#0f172a"
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(4, 8))

        row_no = 1
        for field in existing_fields:
            ttk.Label(form, text=field).grid(
                row=row_no, column=0, sticky="w",
                padx=(0, 18), pady=7
            )
            entry = ttk.Entry(form, width=55)
            entry.grid(row=row_no, column=1, sticky="ew", pady=7)
            entries[field] = entry
            row_no += 1

        if editing:
            reference = str(row["Reference No"])
            old_customer = str(row["Customer"])
            for field in existing_fields:
                entries[field].insert(0, str(row[field]))
        else:
            reference = next_letter_reference()
            entries["Date"].insert(0, datetime.now().strftime("%Y-%m-%d"))

        ttk.Label(
            form, text="Letter Information",
            font=("Segoe UI Semibold", 12),
            foreground="#60a5fa", background="#0f172a"
        ).grid(row=row_no, column=0, columnspan=2, sticky="w", pady=(18, 8))
        row_no += 1

        ttk.Label(form, text="Reference No").grid(
            row=row_no, column=0, sticky="w", padx=(0, 18), pady=7
        )
        ref_var = StringVar(value=reference)
        ttk.Entry(
            form, textvariable=ref_var, width=55, state="readonly"
        ).grid(row=row_no, column=1, sticky="ew", pady=7)
        row_no += 1

        for field in letter_fields:
            ttk.Label(form, text=field).grid(
                row=row_no, column=0, sticky="w",
                padx=(0, 18), pady=7
            )
            entry = ttk.Entry(form, width=55)
            entry.grid(row=row_no, column=1, sticky="ew", pady=7)
            entries[field] = entry
            if editing:
                entry.insert(0, str(row[field]))
            row_no += 1

        if not editing:
            entries["Letter Date"].insert(0, datetime.now().strftime("%d.%m.%Y"))
            entries["Subject"].insert(0, "Earnest money")

        form.columnconfigure(1, weight=1)

        ttk.Label(
            form,
            text=(
                f"Reminder deadline: {due_date_text(row['Date'])}"
                if editing else
                "Reminder deadline will be 15 months after the Record Date."
            ),
            foreground="#f59e0b", background="#0f172a"
        ).grid(row=row_no, column=0, columnspan=2, sticky="w", pady=(5, 10))
        row_no += 1

        # ---------------- Random test record ----------------
        # Kept for testing/demo use. It never creates or attaches fake files.
        def generate_random_test_record():
            import random

            companies = [
                "HP Pakistan", "Dell Pakistan", "Lenovo Pakistan",
                "ASUS Pakistan", "TCL Pakistan", "Intel Pakistan",
                "Huawei Pakistan", "Epson Pakistan"
            ]
            banks = [
                "HBL Islamabad", "UBL Islamabad", "MCB Rawalpindi",
                "Meezan Bank Islamabad", "Allied Bank Islamabad"
            ]
            cities = ["Islamabad", "Rawalpindi", "Lahore", "Karachi"]
            items = [
                "Laptop Computers", "Desktop Computers",
                "Network Equipment", "Printers", "IT Equipment"
            ]

            customer = random.choice(companies)
            quotation = f"QTN-{random.randint(1000, 9999)}"
            payorder = str(random.randint(100000, 999999))
            initial = random.randint(100000, 900000)
            earnest = random.randint(10000, max(10000, initial // 2))
            city = random.choice(cities)
            phone = f"03{random.randint(10000000, 99999999)}"
            extension = str(random.randint(1000, 3999))
            fax = f"{random.randint(3000000, 3999999)}"

            test_data = {
                "Customer": customer,
                "Bank": random.choice(banks),
                "Quotation": quotation,
                "Payorder": payorder,
                "Date": datetime.now().strftime("%Y-%m-%d"),
                "Initial Amount": str(initial),
                "Ernest": str(earnest),
                "Item": random.choice(items),
                "Letter Date": datetime.now().strftime("%d.%m.%Y"),
                "Recipient": "Incharge Procurement",
                "Organization": customer,
                "P.O. Box": str(random.randint(100, 9999)),
                "City": city,
                "Phone": phone,
                "Extension": extension,
                "Fax": fax,
                "Subject": "Earnest money",
                "Tender No": quotation,
                "CDR/Pay Order No": payorder,
                "Letter Amount": str(earnest),
            }

            for field, value in test_data.items():
                if field in entries:
                    entries[field].delete(0, END)
                    entries[field].insert(0, value)

            # Never invent document attachments.
            uploaded["Quotation"] = None
            uploaded["Payorder"] = None
            q_label.set("Quotation: not attached")
            p_label.set("Payorder: not attached")

            save_status.set("Random test record loaded — review before saving.")

        test_button = ttk.Button(
            form,
            text="⚡ Fill Random Test Data",
            command=generate_random_test_record
        )
        test_button.grid(
            row=row_no, column=0, columnspan=2,
            sticky="w", pady=(0, 12)
        )
        row_no += 1

        upload_frame = ttk.Frame(form)
        upload_frame.grid(row=row_no, column=0, columnspan=2, sticky="ew", pady=10)
        q_label = StringVar(value="Quotation: not changed")
        p_label = StringVar(value="Payorder: not changed")

        def upload(kind, label_var):
            path = filedialog.askopenfilename(
                title=f"Choose {kind}",
                filetypes=[
                    ("PDF files", "*.pdf"),
                    ("Images", "*.png *.jpg *.jpeg"),
                    ("All files", "*.*")
                ],
                parent=win
            )
            if path:
                uploaded[kind] = path
                label_var.set(f"{kind}: {Path(path).name}")

        ttk.Button(
            upload_frame, text="Attach Quotation",
            command=lambda: upload("Quotation", q_label)
        ).pack(side="left", padx=(0, 10))
        ttk.Label(
            upload_frame, textvariable=q_label,
            foreground="#94a3b8"
        ).pack(side="left")
        ttk.Button(
            upload_frame, text="Attach Payorder",
            command=lambda: upload("Payorder", p_label)
        ).pack(side="left", padx=(25, 10))
        ttk.Label(
            upload_frame, textvariable=p_label,
            foreground="#94a3b8"
        ).pack(side="left")

        def save():
            save_status.set("Saving...")
            win.update_idletasks()
            try:
                data = {
                    field: entry.get().strip()
                    for field, entry in entries.items()
                }
                data["Reference No"] = reference

                missing = [f for f in existing_fields if not data.get(f)]
                if missing:
                    messagebox.showerror(
                        "Missing information",
                        "Please fill these fields:\n\n" + "\n".join(missing),
                        parent=win
                    )
                    return

                if not data.get("Tender No"):
                    data["Tender No"] = data.get("Quotation", "")
                if not data.get("CDR/Pay Order No"):
                    data["CDR/Pay Order No"] = data.get("Payorder", "")
                if not data.get("Letter Amount"):
                    data["Letter Amount"] = data.get("Ernest", "")

                missing_letter = [
                    f for f in letter_fields if not data.get(f)
                ]
                if missing_letter:
                    messagebox.showerror(
                        "Missing letter information",
                        "Please fill these fields:\n\n" + "\n".join(missing_letter),
                        parent=win
                    )
                    return

                datetime.strptime(data["Date"], "%Y-%m-%d")

                df = read_excel(RECORDS_XLSX, ACTIVE_COLUMNS)

                if editing:
                    # Identify the exact original row before changing any fields.
                    mask = (
                        df["Reference No"].astype(str).eq(reference)
                    )
                    if not mask.any():
                        messagebox.showerror(
                            "Record not found",
                            "This record was changed or removed by another user. "
                            "Refresh the records and try again.",
                            parent=win
                        )
                        return
                    old = df.loc[mask].iloc[0].copy()
                    for col in ACTIVE_COLUMNS:
                        df.loc[mask, col] = data.get(col, "")
                    action = "EDIT RECORD"
                else:
                    new_row = pd.DataFrame(
                        [{col: data.get(col, "") for col in ACTIVE_COLUMNS}],
                        columns=ACTIVE_COLUMNS
                    )
                    df = pd.concat([df, new_row], ignore_index=True)
                    old = None
                    action = "ADD RECORD"

                if not safe_write_excel(RECORDS_XLSX, df):
                    return

                # Customer folders follow the current customer name.
                folder = customer_folder(data["Customer"])
                folder.mkdir(parents=True, exist_ok=True)

                # Copy newly selected attachments.
                for kind, source in uploaded.items():
                    if source:
                        source_path = Path(source)
                        shutil.copy2(
                            source_path,
                            folder / f"{kind}_{source_path.name}"
                        )

                # Regenerate PDFs every time the record is saved/edited.
                pdf_errors = []
                try:
                    create_record_pdf(data, folder)
                except Exception as exc:
                    pdf_errors.append(f"Record.pdf: {type(exc).__name__}: {exc}")
                try:
                    create_letter_pdf(data, folder)
                except Exception as exc:
                    pdf_errors.append(f"Letter.pdf: {type(exc).__name__}: {exc}")

                # If the date changed, a previously sent reminder is no longer
                # automatically valid. Do not delete history; flag it in audit.
                if editing and old is not None:
                    if str(old["Date"]) != str(data["Date"]):
                        add_audit(
                            self.user, "REMINDER DATE RECALCULATED",
                            reference, data["Customer"],
                            f"{old['Date']} -> {data['Date']}; "
                            f"new deadline {due_date_text(data['Date'])}"
                        )

                add_audit(
                    self.user, action, reference, data["Customer"],
                    "Record saved and PDFs regenerated"
                )

                # If a customer name changed, preserve old folder instead of
                # deleting it automatically; this avoids accidental data loss.
                self.refresh_dashboard()
                self.status_var.set(
                    f"{'Updated' if editing else 'Saved'}: "
                    f"{data['Customer']} • {reference}"
                )

                if pdf_errors:
                    messagebox.showwarning(
                        "Record saved — PDF issue",
                        "The record was saved, but PDF generation had issues:\n\n"
                        + "\n".join(pdf_errors),
                        parent=win
                    )
                else:
                    messagebox.showinfo(
                        "Record Updated" if editing else "Record Saved",
                        (
                            "Record updated successfully.\n\n"
                            if editing else "Record saved successfully.\n\n"
                        )
                        + f"Reference: {reference}\n"
                        + f"15-month reminder date: {due_date_text(data['Date'])}",
                        parent=win
                    )
                win.destroy()

            except Exception as exc:
                save_status.set("Save failed")
                messagebox.showerror(
                    "Could not save record",
                    f"{type(exc).__name__}: {exc}",
                    parent=win
                )

        bottom = ttk.Frame(win, padding=(30, 15))
        bottom.pack(fill="x")
        ttk.Button(bottom, text="Cancel", command=win.destroy).pack(side="right")
        ttk.Button(
            bottom,
            text="Update Record" if editing else "Save Record",
            command=save
        ).pack(side="right", padx=(0, 10))
        win.bind("<Return>", lambda _event: save())

    def view_records(self):
        self.records_window()

    def records_window(self):
        win = Toplevel(self.root)
        win.title("All Records")
        win.geometry("1000x600")
        win.configure(bg="#0b1120")

        top = ttk.Frame(win, padding=22)
        top.pack(fill="x")
        ttk.Label(
            top, text="All Active Records",
            font=("Segoe UI", 20, "bold"),
            foreground="#f8fafc", background="#0b1120"
        ).pack(side="left")

        ttk.Button(
            top, text="Open Records Folder",
            command=lambda: open_path(RECORDS_DIR)
        ).pack(side="right")

        frame = ttk.Frame(win, padding=(22, 0, 22, 22))
        frame.pack(fill="both", expand=True)
        cols = ACTIVE_COLUMNS[:8]
        tree = ttk.Treeview(frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=145, minwidth=80)
        sy = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        sx = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        def populate():
            for item in tree.get_children():
                tree.delete(item)
            df = read_excel(RECORDS_XLSX, ACTIVE_COLUMNS)
            for _, r in df.iterrows():
                tree.insert("", "end", values=[str(r[c]) for c in cols])
        populate()

        def selected_row():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Select a record", "Select a record first.", parent=win)
                return None
            values = tree.item(sel[0], "values")
            return self.current_record_from_values(values)

        buttons = ttk.Frame(win, padding=(22, 0, 22, 18))
        buttons.pack(fill="x")
        if self.can_edit:
            ttk.Button(
                buttons, text="Edit Selected",
                command=lambda: (
                    self.edit_record_window(selected_row())
                    if selected_row() is not None else None
                )
            ).pack(side="left", padx=(0, 8))
        ttk.Button(
            buttons, text="Open Selected Customer Folder",
            command=lambda: (
                open_path(customer_folder(str(selected_row()["Customer"])))
                if selected_row() is not None else None
            )
        ).pack(side="left")

        def refresh():
            populate()
            self.refresh_dashboard()
        ttk.Button(buttons, text="Refresh", command=refresh).pack(side="right")

    def search_window(self):
        win = Toplevel(self.root)
        win.title("Search Records")
        win.geometry("950x580")
        win.configure(bg="#0b1120")

        ttk.Label(
            win, text="Search",
            font=("Segoe UI", 22, "bold"),
            foreground="#f8fafc", background="#0b1120"
        ).pack(anchor="w", padx=25, pady=(22, 10))
        search_var = StringVar()
        filters = ttk.Frame(win, padding=(25, 0, 25, 10))
        filters.pack(fill="x")
        entry = ttk.Entry(filters, textvariable=search_var, width=55)
        entry.pack(side="left", fill="x", expand=True)
        entry.focus()
        status_filter = StringVar(value="ALL")
        status_box = ttk.Combobox(
            filters, textvariable=status_filter,
            values=["ALL", "ACTIVE", "RETURNED", "DUE 15M+"],
            state="readonly", width=14
        )
        status_box.pack(side="left", padx=8)

        display_cols = ("Status",) + tuple(ACTIVE_COLUMNS[:8])
        frame = ttk.Frame(win, padding=(25, 0, 25, 20))
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=display_cols, show="headings")
        widths = {
            "Status": 90, "Customer": 180, "Bank": 130,
            "Quotation": 115, "Payorder": 115, "Date": 105,
            "Initial Amount": 120, "Ernest": 110, "Item": 190
        }
        for col in display_cols:
            tree.heading(col, text=col)
            tree.column(col, width=widths.get(col, 130), minwidth=80)
        sy = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sy.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        def run_search(*_):
            for item in tree.get_children():
                tree.delete(item)
            keyword = search_var.get().strip().lower()
            selected_status = status_filter.get()
            active = read_excel(RECORDS_XLSX, ACTIVE_COLUMNS)
            returned = read_excel(RETURNED_XLSX, RETURNED_COLUMNS)

            if not active.empty and selected_status in ("ALL", "ACTIVE", "DUE 15M+"):
                for _, r in active.iterrows():
                    blob = " ".join(str(r[c]) for c in ACTIVE_COLUMNS).lower()
                    if keyword and keyword not in blob:
                        continue
                    if selected_status == "DUE 15M+" and not is_record_due(r["Date"]):
                        continue
                    tree.insert(
                        "", "end",
                        values=["ACTIVE"] + [str(r[c]) for c in ACTIVE_COLUMNS[:8]]
                    )

            if not returned.empty and selected_status in ("ALL", "RETURNED"):
                for _, r in returned.iterrows():
                    blob = " ".join(str(r[c]) for c in RETURNED_COLUMNS).lower()
                    if keyword and keyword not in blob:
                        continue
                    values = [
                        "RETURNED",
                        str(r["Customer"]), "", "", str(r["Payorder"]),
                        str(r["Date"]), str(r["Amount"]), "", ""
                    ]
                    tree.insert("", "end", values=values)

        search_var.trace_add("write", run_search)
        status_filter.trace_add("write", run_search)
        run_search()

    def due_dataframe(self):
        df = read_excel(RECORDS_XLSX, ACTIVE_COLUMNS)
        if df.empty:
            return df
        return df[df["Date"].apply(is_record_due)].copy()

    def upcoming_dataframe(self, df=None, days=30):
        """Records that reach the 15-month mark within the next `days` days."""
        if df is None:
            df = read_excel(RECORDS_XLSX, ACTIVE_COLUMNS)
        if df.empty:
            return df
        today = datetime.now().date()
        horizon = today + timedelta(days=days)

        def soon(value):
            due = calendar_due_date(value)
            return bool(due and today < due <= horizon)

        return df[df["Date"].apply(soon)].copy()

    def due_orders_window(self):
        win = Toplevel(self.root)
        win.title("Due Orders — 15 Months+")
        win.geometry("920x560")
        win.configure(bg="#0b1120")

        ttk.Label(
            win, text="Due Orders — 15 Months+",
            font=("Segoe UI", 21, "bold"),
            foreground="#f8fafc", background="#0b1120"
        ).pack(anchor="w", padx=25, pady=(22, 5))
        ttk.Label(
            win,
            text="A record becomes due exactly 15 calendar months after its Record Date.",
            foreground="#94a3b8", background="#0b1120"
        ).pack(anchor="w", padx=25, pady=(0, 15))

        frame = ttk.Frame(win, padding=(25, 0, 25, 15))
        frame.pack(fill="both", expand=True)
        cols = ("Reference No", "Customer", "Bank", "Date", "Due Date", "Ernest", "Item")
        tree = ttk.Treeview(frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=145, minwidth=90)
        sy = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sy.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        due = self.due_dataframe()
        for idx, r in due.iterrows():
            tree.insert(
                "", "end", iid=str(idx),
                values=[
                    str(r["Reference No"]), str(r["Customer"]),
                    str(r["Bank"]), str(r["Date"]),
                    due_date_text(r["Date"]), str(r["Ernest"]),
                    str(r["Item"])
                ]
            )

        def process_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning(
                    "Select a record", "Select a due record first.", parent=win
                )
                return
            idx = int(selected[0])
            row = due.loc[idx]
            self.process_return(row, win)

        ttk.Button(
            win, text="Enter Return Details",
            command=process_selected
        ).pack(pady=(0, 12))

        # Records approaching the 15-month mark within 30 days.
        upcoming = self.upcoming_dataframe()
        ttk.Label(
            win,
            text=f"Due within the next 30 days ({len(upcoming)})",
            font=("Segoe UI Semibold", 13),
            foreground="#eab308", background="#0b1120"
        ).pack(anchor="w", padx=25, pady=(0, 8))
        up_frame = ttk.Frame(win, padding=(25, 0, 25, 20))
        up_frame.pack(fill="x")
        up_tree = ttk.Treeview(
            up_frame, columns=cols, show="headings", height=5
        )
        for col in cols:
            up_tree.heading(col, text=col)
            up_tree.column(col, width=145, minwidth=90)
        up_sy = ttk.Scrollbar(up_frame, orient="vertical", command=up_tree.yview)
        up_tree.configure(yscrollcommand=up_sy.set)
        up_tree.grid(row=0, column=0, sticky="ew")
        up_sy.grid(row=0, column=1, sticky="ns")
        up_frame.columnconfigure(0, weight=1)
        for _, r in upcoming.iterrows():
            up_tree.insert(
                "", "end",
                values=[
                    str(r["Reference No"]), str(r["Customer"]),
                    str(r["Bank"]), str(r["Date"]),
                    due_date_text(r["Date"]), str(r["Ernest"]),
                    str(r["Item"])
                ]
            )

    def process_return(self, row, parent=None):
        win = Toplevel(self.root)
        win.title("Return Details")
        win.geometry("480x450")
        win.configure(bg="#0f172a")
        win.transient(parent or self.root)
        win.grab_set()

        ttk.Label(
            win, text="Record Returned",
            font=("Segoe UI", 20, "bold"),
            foreground="#f8fafc", background="#0f172a"
        ).pack(anchor="w", padx=25, pady=(22, 4))
        ttk.Label(
            win, text=f"Customer: {row['Customer']}",
            foreground="#60a5fa", background="#0f172a"
        ).pack(anchor="w", padx=25, pady=(0, 18))

        form = ttk.Frame(win, padding=(25, 0))
        form.pack(fill="both", expand=True)
        fields = RETURNED_COLUMNS
        entries = {}
        for i, field in enumerate(fields):
            ttk.Label(form, text=field).grid(
                row=i, column=0, sticky="w", padx=(0, 15), pady=8
            )
            entry = ttk.Entry(form, width=38)
            entry.grid(row=i, column=1, sticky="ew", pady=8)
            entries[field] = entry

        entries["Customer"].insert(0, str(row["Customer"]))
        entries["Customer"].configure(state="readonly")
        entries["Date"].insert(0, datetime.now().strftime("%Y-%m-%d"))
        entries["Payorder"].insert(0, str(row["Payorder"]))
        form.columnconfigure(1, weight=1)

        def save_return():
            if not self.can_edit:
                messagebox.showwarning(
                    "Permission denied",
                    "Your account is not allowed to process returns.",
                    parent=win
                )
                return
            data = {f: entries[f].get().strip() for f in fields}
            data["Customer"] = str(row["Customer"])
            missing = [
                f for f, value in data.items()
                if not value and f != "Payorder"
            ]
            if missing:
                messagebox.showerror(
                    "Missing information",
                    "Please fill:\n\n" + "\n".join(missing),
                    parent=win
                )
                return
            try:
                datetime.strptime(data["Date"], "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Invalid date", "Use YYYY-MM-DD.", parent=win)
                return

            try:
                returned = read_excel(RETURNED_XLSX, RETURNED_COLUMNS)
                returned = pd.concat(
                    [returned, pd.DataFrame([data], columns=RETURNED_COLUMNS)],
                    ignore_index=True
                )
                if not safe_write_excel(RETURNED_XLSX, returned):
                    return

                active = read_excel(RECORDS_XLSX, ACTIVE_COLUMNS)
                reference = str(row["Reference No"])
                mask = active["Reference No"].astype(str).eq(reference)
                if not mask.any():
                    messagebox.showerror(
                        "Record changed",
                        "This record no longer exists in the active records.",
                        parent=win
                    )
                    return
                active = active.loc[~mask].copy()
                if not safe_write_excel(RECORDS_XLSX, active):
                    return

                add_audit(
                    self.user, "RETURNED RECORD", reference,
                    str(row["Customer"]),
                    f"Return details saved; active record removed"
                )
                self.refresh_dashboard()
                messagebox.showinfo(
                    "Return saved",
                    "Return details saved and the active record was moved.",
                    parent=win
                )
                win.destroy()
                if parent and parent.winfo_exists():
                    parent.destroy()
            except Exception as exc:
                messagebox.showerror("Could not save return", str(exc), parent=win)

        ttk.Button(
            win, text="Save Return", command=save_return
        ).pack(pady=(0, 25))

    def check_reminders(self):
        due = self.due_dataframe()
        if due.empty:
            self.refresh_dashboard()
            messagebox.showinfo(
                "Reminders", "There are no records at or beyond 15 months."
            )
            return
        if self._email_worker_busy:
            messagebox.showinfo(
                "Reminders", "A reminder check is already running."
            )
            return

        # Resolve (and if needed, pick) the template on the main thread so the
        # background worker never has to open a dialog.
        try:
            resolve_template_path()
        except Exception as exc:
            messagebox.showerror("Template missing", str(exc))
            return

        rows = [{c: row[c] for c in ACTIVE_COLUMNS} for _, row in due.iterrows()]
        self._email_worker_busy = True
        threading.Thread(
            target=self._send_reminders_worker, args=(rows, True), daemon=True
        ).start()

    def auto_check(self):
        try:
            # Any PC running the app can perform the automatic check.
            if not self._email_worker_busy:
                due = self.due_dataframe()
                rows = [
                    {c: row[c] for c in ACTIVE_COLUMNS}
                    for _, row in due.iterrows()
                ]
                if rows:
                    self._email_worker_busy = True
                    threading.Thread(
                        target=self._send_reminders_worker,
                        args=(rows, False), daemon=True
                    ).start()
        except Exception:
            pass
        finally:
            self._last_auto_check = datetime.now().strftime("%H:%M")
            self.refresh_status()
            self.root.after(auto_check_ms(), self.auto_check)

    def _send_reminders_worker(self, rows, report):
        """Send due reminders off the UI thread. No tkinter calls allowed here;
        results are marshalled back with root.after()."""
        sent = failed = skipped = 0
        try:
            for row in rows:
                reference = str(row["Reference No"])
                customer = str(row["Customer"])
                if reminder_already_sent(reference):
                    skipped += 1
                    continue

                folder = customer_folder(customer)
                folder.mkdir(parents=True, exist_ok=True)

                try:
                    # Regenerate every time before a reminder is sent, so edited
                    # data is guaranteed to be reflected in the attachments.
                    create_letter_pdf(row, folder)
                    create_record_pdf(row, folder)
                    if send_email(
                        customer, folder / "Letter.pdf", folder / "Record.pdf",
                        reference=reference,
                        user_id=self.user["User ID"],
                        silent=True
                    ):
                        sent += 1
                    else:
                        failed += 1
                except Exception as exc:
                    failed += 1
                    error_text = f"{type(exc).__name__}: {exc}"
                    finalize_email_claim(
                        reference, "FAILED", 0, error_text,
                        self.user["User ID"]
                    )
                    add_audit(
                        self.user, "EMAIL FAILED", reference, customer, error_text
                    )
        finally:
            self._email_worker_busy = False
            if report:
                total = len(rows)
                self.root.after(0, lambda: self._reminder_summary(
                    total, sent, failed, skipped
                ))

    def _reminder_summary(self, total, sent, failed, skipped):
        self.refresh_dashboard()
        messagebox.showinfo(
            "Reminder check",
            f"Due records: {total}\n"
            f"Emails sent: {sent}\n"
            f"Failed: {failed}\n"
            f"Already sent/skipped: {skipped}"
        )

    def email_history_window(self):
        win = Toplevel(self.root)
        win.title("Email History")
        win.geometry("950x540")
        win.configure(bg="#0b1120")
        ttk.Label(
            win, text="Email History",
            font=("Segoe UI", 21, "bold"),
            foreground="#f8fafc", background="#0b1120"
        ).pack(anchor="w", padx=25, pady=(22, 12))
        frame = ttk.Frame(win, padding=(25, 0, 25, 20))
        frame.pack(fill="both", expand=True)
        cols = EMAIL_HISTORY_COLUMNS
        tree = ttk.Treeview(frame, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=125, minwidth=80)
        sy = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        sx = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        df = read_excel(email_history_path(), EMAIL_HISTORY_COLUMNS)
        for _, r in df.iloc[::-1].iterrows():
            tree.insert("", "end", values=[str(r[c]) for c in cols])

        def show_selected_error(event=None):
            sel = tree.selection()
            if not sel:
                messagebox.showwarning(
                    "Select email",
                    "Select an email history entry first.",
                    parent=win
                )
                return
            vals = tree.item(sel[0], "values")
            timestamp = str(vals[0]) if len(vals) > 0 else ""
            status = str(vals[5]) if len(vals) > 5 else ""
            attempts = str(vals[6]) if len(vals) > 6 else ""
            error = str(vals[7]) if len(vals) > 7 else ""
            if status == "SENDING":
                details = (
                    f"Status: SENDING\n"
                    f"Started: {timestamp}\n"
                    f"Attempts: {attempts or '0'}\n\n"
                    "The automatic SMTP send is still in progress. "
                    "The connection has a 20-second timeout; if it fails, "
                    "the entry will change to FAILED and show the exact error."
                )
            else:
                details = (
                    f"Status: {status}\n"
                    f"Time: {timestamp}\n"
                    f"Attempts: {attempts or '0'}\n\n"
                    f"Error / Details:\n{error or 'No error was recorded for this email.'}"
                )
            messagebox.showinfo(
                f"Email Details — {status}",
                details,
                parent=win
            )

        tree.bind("<Double-1>", show_selected_error)

        def resend_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Select email", "Select an email history entry first.", parent=win)
                return
            vals = tree.item(sel[0], "values")
            reference = str(vals[1])
            customer = str(vals[2])
            recipient = str(vals[3])
            active = read_excel(RECORDS_XLSX, ACTIVE_COLUMNS)
            matches = active[active["Reference No"].astype(str).eq(reference)]
            if matches.empty:
                messagebox.showerror(
                    "Record unavailable",
                    "The active record for this email no longer exists.",
                    parent=win
                )
                return
            row_data = {c: matches.iloc[0][c] for c in ACTIVE_COLUMNS}
            folder = customer_folder(customer)
            folder.mkdir(parents=True, exist_ok=True)
            try:
                resolve_template_path()
            except Exception as exc:
                messagebox.showerror("Template missing", str(exc), parent=win)
                return

            def resend_worker():
                try:
                    create_letter_pdf(row_data, folder)
                    create_record_pdf(row_data, folder)
                    ok = send_email(
                        customer, folder / "Letter.pdf", folder / "Record.pdf",
                        reference=reference, recipient=recipient,
                        user_id=self.user["User ID"], silent=True
                    )
                    if ok:
                        self.root.after(0, lambda: messagebox.showinfo(
                            "Email sent",
                            f"Reminder sent successfully for {customer}."
                        ))
                    else:
                        self.root.after(0, lambda: messagebox.showwarning(
                            "Email not sent",
                            "The email was not sent. It may already have been "
                            "sent or is being handled by another computer — "
                            "check Email History for details."
                        ))
                except Exception as exc:
                    err = str(exc)
                    self.root.after(0, lambda: messagebox.showerror(
                        "Resend failed", err
                    ))

            threading.Thread(target=resend_worker, daemon=True).start()

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=25, pady=(0, 18))

        def preview_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("Select email", "Select an email history entry first.", parent=win)
                return
            vals = tree.item(sel[0], "values")
            reference, customer, recipient, subject, status = (
                str(vals[1]), str(vals[2]), str(vals[3]), str(vals[4]), str(vals[5])
            )
            messagebox.showinfo(
                "Email Preview",
                f"To: {recipient}\n"
                f"Subject: {subject}\n"
                f"Status: {status}\n\n"
                f"Dear Sir/Madam,\n\n"
                f"Please find attached the Ernest Money Letter and Record for {customer}.\n\n"
                f"Reference: {reference}\n\nTime & Tune",
                parent=win
            )

        ttk.Button(
            btns, text="View Error / Details", command=show_selected_error
        ).pack(side="left")
        ttk.Button(
            btns, text="Preview Selected", command=preview_selected
        ).pack(side="left", padx=8)
        ttk.Button(
            btns, text="Send Again", command=resend_selected
        ).pack(side="left", padx=8)

    def audit_window(self):
        win = Toplevel(self.root)
        win.title("Audit Log")
        win.geometry("980x560")
        win.configure(bg="#0b1120")
        ttk.Label(
            win, text="Audit Log",
            font=("Segoe UI", 21, "bold"),
            foreground="#f8fafc", background="#0b1120"
        ).pack(anchor="w", padx=25, pady=(22, 12))
        frame = ttk.Frame(win, padding=(25, 0, 25, 20))
        frame.pack(fill="both", expand=True)
        cols = AUDIT_COLUMNS
        tree = ttk.Treeview(frame, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=140, minwidth=90)
        sy = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        sx = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        df = read_excel(audit_path(), AUDIT_COLUMNS)
        for _, r in df.iloc[::-1].iterrows():
            tree.insert("", "end", values=[str(r[c]) for c in cols])

    def export_report(self):
        """Export current active/returned/audit/email data to a user-selected folder."""
        target = filedialog.askdirectory(
            title="Choose a folder for the Excel report",
            parent=self.root
        )
        if not target:
            return
        try:
            target = Path(target)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report = target / f"TimeAndTune_Report_{stamp}.xlsx"
            with pd.ExcelWriter(report, engine="openpyxl") as writer:
                read_excel(RECORDS_XLSX, ACTIVE_COLUMNS).to_excel(
                    writer, sheet_name="Active Records", index=False
                )
                read_excel(RETURNED_XLSX, RETURNED_COLUMNS).to_excel(
                    writer, sheet_name="Returned Records", index=False
                )
                read_excel(email_history_path(), EMAIL_HISTORY_COLUMNS).to_excel(
                    writer, sheet_name="Email History", index=False
                )
                read_excel(audit_path(), AUDIT_COLUMNS).to_excel(
                    writer, sheet_name="Audit Log", index=False
                )
            add_audit(
                self.user, "REPORT EXPORTED", "", "",
                str(report)
            )
            messagebox.showinfo("Report exported", f"Report saved to:\n{report}")
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))

    def restore_backup_window(self):
        if not self.admin_reauth(
            "Restoring a backup changes the shared system data. Administrator credentials are required."
        ):
            return
        backup_dir_path = backup_dir()
        backup_dir_path.mkdir(parents=True, exist_ok=True)
        choices = sorted(
            [p for p in backup_dir_path.iterdir() if p.is_dir()],
            key=lambda p: p.name,
            reverse=True
        )
        if not choices:
            messagebox.showinfo("No backups", "No backups are available yet.")
            return

        win = Toplevel(self.root)
        win.title("Restore Backup")
        win.geometry("600x400")
        win.configure(bg="#0f172a")
        ttk.Label(
            win, text="Restore Backup",
            font=("Segoe UI", 20, "bold"),
            foreground="#f8fafc", background="#0f172a"
        ).pack(anchor="w", padx=25, pady=(25, 10))
        ttk.Label(
            win,
            text="The current shared data will first be backed up automatically.",
            foreground="#f59e0b", background="#0f172a"
        ).pack(anchor="w", padx=25, pady=(0, 15))
        lb = __import__("tkinter").Listbox(
            win, bg="#111827", fg="#e5e7eb",
            selectbackground="#2563eb", height=12
        )
        lb.pack(fill="both", expand=True, padx=25)
        for p in choices:
            lb.insert("end", p.name)

        def restore():
            idx = lb.curselection()
            if not idx:
                messagebox.showwarning("Select backup", "Select a backup first.", parent=win)
                return
            source = choices[idx[0]]
            if not messagebox.askyesno(
                "Final restore confirmation",
                f"Restore backup {source.name}?\n\n"
                "The current system will be backed up before restoration.",
                parent=win
            ):
                return
            try:
                current_backup = backup_all_data()
                with data_lock():
                    for name in [
                        RECORDS_XLSX.name, RETURNED_XLSX.name,
                        users_path().name, audit_path().name,
                        email_history_path().name, settings_path().name,
                        REFERENCE_COUNTER_FILE.name,
                    ]:
                        src = source / name
                        if src.exists():
                            shutil.copy2(src, DATA_DIR / name)
                    src_records = source / RECORDS_DIR.name
                    if src_records.exists():
                        if RECORDS_DIR.exists():
                            shutil.rmtree(RECORDS_DIR)
                        shutil.copytree(src_records, RECORDS_DIR)
                add_audit(
                    self.user, "BACKUP RESTORED", "", "",
                    f"{source}; pre-restore backup: {current_backup}"
                )
                self.refresh_dashboard()
                messagebox.showinfo(
                    "Restore complete",
                    "Backup restored successfully.\n\nRestart the application after restoration.",
                    parent=win
                )
                win.destroy()
            except Exception as exc:
                messagebox.showerror("Restore failed", str(exc), parent=win)

        ttk.Button(win, text="Restore Selected Backup", command=restore).pack(
            pady=18
        )

    def manual_backup(self):
        if not self.is_admin:
            return
        try:
            target = backup_all_data()
            add_audit(self.user, "BACKUP CREATED", "", "", str(target))
            messagebox.showinfo("Backup created", f"Backup saved to:\n{target}")
        except Exception as exc:
            messagebox.showerror("Backup failed", str(exc))

    def admin_confirm(self, title="Admin Confirmation"):
        if self.is_admin:
            return True
        return False

    def admin_reauth(self, action_text):
        """Require an admin password for destructive actions."""
        win = Toplevel(self.root)
        win.title("Administrator Verification")
        win.geometry("480x260")
        win.configure(bg="#0f172a")
        win.transient(self.root)
        win.grab_set()

        ttk.Label(
            win, text="Administrator Verification",
            font=("Segoe UI", 18, "bold"),
            foreground="#f8fafc", background="#0f172a"
        ).pack(anchor="w", padx=25, pady=(25, 8))
        ttk.Label(
            win, text=action_text,
            foreground="#f59e0b", background="#0f172a"
        ).pack(anchor="w", padx=25, pady=(0, 18))
        ttk.Label(win, text="Admin User ID").pack(anchor="w", padx=25)
        uid = ttk.Entry(win, width=40)
        uid.pack(fill="x", padx=25, pady=(4, 10))
        ttk.Label(win, text="Admin Password").pack(anchor="w", padx=25)
        pw = ttk.Entry(win, width=40, show="•")
        pw.pack(fill="x", padx=25, pady=(4, 15))
        result = {"ok": False}

        def verify():
            users = read_excel(users_path(), USERS_COLUMNS)
            row = users[
                (users["User ID"].astype(str).str.lower() == uid.get().strip().lower()) &
                (users["Role"].astype(str) == ROLE_ADMIN) &
                (users["Active"].astype(str).str.lower() == "true")
            ]
            if row.empty or not verify_password(pw.get(), str(row.iloc[0]["Password Hash"])):
                messagebox.showerror(
                    "Verification failed",
                    "Invalid administrator credentials.",
                    parent=win
                )
                return
            result["ok"] = True
            win.destroy()

        ttk.Button(win, text="Verify", command=verify).pack(pady=(0, 20))
        win.wait_window()
        return result["ok"]

    def admin_delete_menu(self):
        if not self.admin_reauth(
            "Deleting records is permanent. Administrator credentials are required."
        ):
            return
        self.delete_record_window()

    def delete_record_window(self):
        win = Toplevel(self.root)
        win.title("Remove Records")
        win.geometry("900x540")
        win.configure(bg="#0b1120")

        ttk.Label(
            win, text="Remove Records",
            font=("Segoe UI", 21, "bold"),
            foreground="#f8fafc", background="#0b1120"
        ).pack(anchor="w", padx=25, pady=(22, 6))

        ttk.Label(
            win,
            text="Select one or more records to permanently remove. Hold Ctrl or Shift to select multiple.",
            foreground="#94a3b8", background="#0b1120"
        ).pack(anchor="w", padx=25, pady=(0, 12))

        frame = ttk.Frame(win, padding=(25, 0, 25, 20))
        frame.pack(fill="both", expand=True)

        cols = ("Reference No", "Customer", "Bank", "Quotation", "Payorder", "Date")
        tree = ttk.Treeview(
            frame, columns=cols, show="headings", selectmode="extended"
        )
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=145, minwidth=90)

        yscroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        df = read_excel(RECORDS_XLSX, ACTIVE_COLUMNS)
        for _, r in df.iterrows():
            tree.insert(
                "", "end",
                values=[
                    str(r["Reference No"]), str(r["Customer"]),
                    str(r["Bank"]), str(r["Quotation"]),
                    str(r["Payorder"]), str(r["Date"])
                ]
            )

        def delete_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning(
                    "Select records",
                    "Select at least one record to remove.",
                    parent=win
                )
                return

            selected_refs = []
            selected_customers = {}
            for item in sel:
                vals = tree.item(item, "values")
                reference = str(vals[0])
                customer = str(vals[1])
                selected_refs.append(reference)
                selected_customers[reference] = customer

            count = len(selected_refs)
            if not messagebox.askyesno(
                "Confirm permanent deletion",
                f"Are you sure you want to permanently remove {count} record"
                f"{'s' if count != 1 else ''}?\n\n"
                "This will also remove their stored record files.",
                parent=win
            ):
                return

            current = read_excel(RECORDS_XLSX, ACTIVE_COLUMNS)
            ref_series = current["Reference No"].astype(str)
            mask = ref_series.isin(selected_refs)

            if not mask.any():
                messagebox.showerror(
                    "Not found",
                    "The selected records were already changed or removed.",
                    parent=win
                )
                return

            removed = current.loc[mask].copy()
            remaining = current.loc[~mask].copy()

            if not safe_write_excel(RECORDS_XLSX, remaining):
                return

            # Only remove a customer's folder when that customer has no
            # remaining active records. This avoids deleting files belonging
            # to another record for the same customer.
            remaining_customers = set(
                remaining["Customer"].astype(str).str.strip().tolist()
            )

            for _, row in removed.iterrows():
                reference = str(row["Reference No"])
                customer = str(row["Customer"])
                if customer.strip() not in remaining_customers:
                    folder = customer_folder(customer)
                    if folder.exists():
                        shutil.rmtree(folder, ignore_errors=True)

                add_audit(
                    self.user,
                    "DELETE RECORD",
                    reference,
                    customer,
                    "Permanent deletion after administrator verification"
                )

            self.refresh_dashboard()
            messagebox.showinfo(
                "Records Removed",
                f"{len(removed)} record{'s' if len(removed) != 1 else ''} "
                "were permanently removed.",
                parent=win
            )
            win.destroy()

        button_frame = ttk.Frame(win)
        button_frame.pack(fill="x", padx=25, pady=(0, 20))

        ttk.Button(
            button_frame,
            text="Remove Selected Permanently",
            command=delete_selected
        ).pack(side="right")

        ttk.Button(
            button_frame,
            text="Cancel",
            command=win.destroy
        ).pack(side="right", padx=(0, 10))

    def reset_application(self):
        """Reset application data/counters to a fresh-install state.

        Users and saved Settings/email configuration are preserved so the
        shared application can be used again immediately. All records,
        generated customer files, returned records, email history, sent
        markers, and the letter reference counter are reset.
        """
        if not self.admin_reauth(
            "This resets the application data and letter numbering to the original state."
        ):
            return

        if not messagebox.askyesno(
            "Reset Application",
            "This will permanently remove ALL active and returned records, "
            "their PDFs/attachments, email history and sent markers, and "
            "reset the letter reference number to 01-EM.\n\n"
            "Users and Settings will NOT be deleted.\n\n"
            "This cannot be undone. Continue?",
            parent=self.root
        ):
            return

        try:
            # Reset the two record databases.
            safe_write_excel(
                RECORDS_XLSX, pd.DataFrame(columns=ACTIVE_COLUMNS)
            )
            safe_write_excel(
                RETURNED_XLSX, pd.DataFrame(columns=RETURNED_COLUMNS)
            )

            # Remove all customer record folders/files.
            if RECORDS_DIR.exists():
                shutil.rmtree(RECORDS_DIR)
            RECORDS_DIR.mkdir(parents=True, exist_ok=True)

            # Reset email history and automatic-send markers so the dashboard
            # starts at zero and old reminders can never block new records.
            safe_write_excel(
                email_history_path(),
                pd.DataFrame(columns=EMAIL_HISTORY_COLUMNS)
            )
            if EMAILED_FILE.exists():
                EMAILED_FILE.unlink()

            # Reset the persistent letter reference counter. The next letter
            # generated will therefore be 01-EM.
            REFERENCE_COUNTER_FILE.parent.mkdir(parents=True, exist_ok=True)
            REFERENCE_COUNTER_FILE.write_text("1", encoding="utf-8")

            # Keep an audit entry showing who performed the reset.
            add_audit(
                self.user, "APPLICATION RESET", "", "",
                "Records, returned records, email history, sent markers and "
                "letter reference counter reset to original state"
            )

            self.refresh_dashboard()
            messagebox.showinfo(
                "Reset Complete",
                "The application has been reset to its original data state.\n\n"
                "The next letter reference will be 01-EM.",
                parent=self.root
            )
        except Exception as exc:
            messagebox.showerror(
                "Reset failed",
                f"The application could not be completely reset.\n\n{exc}",
                parent=self.root
            )

    def clear_all_records(self):
        if not self.admin_reauth(
            "This permanently deletes ALL active/returned records and files."
        ):
            return
        if not messagebox.askyesno(
            "Clear All Records",
            "This will permanently remove ALL records, PDFs and attachments.\n\nContinue?",
            parent=self.root
        ):
            return
        try:
            safe_write_excel(
                RECORDS_XLSX, pd.DataFrame(columns=ACTIVE_COLUMNS)
            )
            safe_write_excel(
                RETURNED_XLSX, pd.DataFrame(columns=RETURNED_COLUMNS)
            )
            if RECORDS_DIR.exists():
                shutil.rmtree(RECORDS_DIR)
            RECORDS_DIR.mkdir(parents=True, exist_ok=True)
            if EMAILED_FILE.exists():
                EMAILED_FILE.unlink()
            add_audit(
                self.user, "CLEAR ALL RECORDS", "", "",
                "All active/returned records and customer folders cleared"
            )
            self.refresh_dashboard()
            messagebox.showinfo(
                "Records Cleared",
                "All records were cleared. The reference counter was preserved."
            )
        except Exception as exc:
            messagebox.showerror("Could not clear records", str(exc))

    def users_window(self):
        if not self.is_admin:
            return
        win = Toplevel(self.root)
        win.title("User Management")
        win.geometry("900x560")
        win.configure(bg="#0b1120")
        ttk.Label(
            win, text="User Management",
            font=("Segoe UI", 21, "bold"),
            foreground="#f8fafc", background="#0b1120"
        ).pack(anchor="w", padx=25, pady=(22, 12))

        frame = ttk.Frame(win, padding=(25, 0, 25, 15))
        frame.pack(fill="both", expand=True)
        cols = USERS_COLUMNS[:4] + ["Active", "Created At", "Last Login"]
        tree = ttk.Treeview(frame, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=145)
        tree.pack(fill="both", expand=True)

        def refresh():
            for i in tree.get_children():
                tree.delete(i)
            users = read_excel(users_path(), USERS_COLUMNS)
            for _, r in users.iterrows():
                tree.insert("", "end", values=[str(r[c]) for c in cols])
        refresh()

        def add_user():
            self.create_user_window(parent=win, force_role=None)
            refresh()

        def toggle_active():
            sel = tree.selection()
            if not sel:
                return
            vals = tree.item(sel[0], "values")
            uid = str(vals[0])
            if uid == self.user["User ID"]:
                messagebox.showwarning("Not allowed", "You cannot disable your own account.")
                return
            users = read_excel(users_path(), USERS_COLUMNS)
            mask = users["User ID"].astype(str).eq(uid)
            if mask.any():
                old = str(users.loc[mask, "Active"].iloc[0]).lower() == "true"
                users.loc[mask, "Active"] = "False" if old else "True"
                safe_write_excel(users_path(), users)
                add_audit(
                    self.user, "USER STATUS CHANGED", "", "",
                    f"{uid}: {'disabled' if old else 'enabled'}"
                )
                refresh()

        ttk.Button(win, text="Add User", command=add_user).pack(
            side="left", padx=25, pady=(0, 20)
        )
        ttk.Button(win, text="Enable / Disable Selected", command=toggle_active).pack(
            side="left", padx=5, pady=(0, 20)
        )
        ttk.Button(win, text="Refresh", command=refresh).pack(
            side="right", padx=25, pady=(0, 20)
        )

    def create_user_window(self, parent=None, force_role=None):
        win = Toplevel(parent or self.root)
        win.title("Create User Account")
        win.geometry("480x400")
        win.configure(bg="#0f172a")
        win.transient(parent or self.root)
        win.grab_set()

        ttk.Label(
            win, text="Create User Account",
            font=("Segoe UI", 20, "bold"),
            foreground="#f8fafc", background="#0f172a"
        ).pack(anchor="w", padx=25, pady=(25, 20))
        form = ttk.Frame(win, padding=(25, 0))
        form.pack(fill="both", expand=True)
        fields = ["User ID", "Full Name", "Password", "Confirm Password"]
        ent = {}
        for i, f in enumerate(fields):
            ttk.Label(form, text=f).grid(row=i, column=0, sticky="w", pady=9)
            e = ttk.Entry(form, width=34, show="•" if "Password" in f else "")
            e.grid(row=i, column=1, sticky="ew", pady=9)
            ent[f] = e
        ttk.Label(form, text="Role").grid(row=4, column=0, sticky="w", pady=9)
        role_var = StringVar(value=force_role or ROLE_EMPLOYEE)
        role_box = ttk.Combobox(
            form, textvariable=role_var,
            values=[ROLE_EMPLOYEE, ROLE_VIEWER, ROLE_ADMIN],
            state="readonly", width=31
        )
        role_box.grid(row=4, column=1, sticky="ew", pady=9)
        if force_role:
            role_box.configure(state="disabled")
        form.columnconfigure(1, weight=1)

        def save():
            uid = ent["User ID"].get().strip().lower()
            name = ent["Full Name"].get().strip()
            pw = ent["Password"].get()
            cpw = ent["Confirm Password"].get()
            role = role_var.get()
            if not uid or not name or not pw:
                messagebox.showerror("Missing information", "Fill all required fields.", parent=win)
                return
            if len(pw) < 8:
                messagebox.showerror("Weak password", "Use at least 8 characters.", parent=win)
                return
            if pw != cpw:
                messagebox.showerror("Password mismatch", "Passwords do not match.", parent=win)
                return
            users = read_excel(users_path(), USERS_COLUMNS)
            if (users["User ID"].astype(str).str.lower() == uid).any():
                messagebox.showerror("User exists", "That User ID already exists.", parent=win)
                return
            row = {
                "User ID": uid,
                "Full Name": name,
                "Role": role,
                "Password Hash": hash_password(pw),
                "Active": "True",
                "Created At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Last Login": "",
            }
            users = pd.concat([users, pd.DataFrame([row])], ignore_index=True)
            if safe_write_excel(users_path(), users):
                add_audit(
                    self.user if hasattr(self, "user") else {"User ID": "SYSTEM"},
                    "USER CREATED", "", "",
                    f"{uid} / {role}"
                )
                messagebox.showinfo("User created", f"Account {uid} created.", parent=win)
                win.destroy()

        ttk.Button(win, text="Create Account", command=save).pack(pady=(0, 25))

    def settings_window(self):
        win = Toplevel(self.root)
        win.title("Settings")
        win.geometry("650x600")
        win.configure(bg="#0f172a")
        ttk.Label(
            win, text="Settings",
            font=("Segoe UI", 22, "bold"),
            foreground="#f8fafc", background="#0f172a"
        ).pack(anchor="w", padx=30, pady=(25, 15))

        settings = load_settings()
        form = ttk.Frame(win, padding=(30, 0))
        form.pack(fill="both", expand=True)

        entries = {}
        fields = [
            ("SMTP Host", "smtp_host"),
            ("SMTP Port", "smtp_port"),
            ("Sender Email", "sender_email"),
            ("Default Recipient", "recipient_email"),
        ]
        for i, (label, key) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=i, column=0, sticky="w", pady=9)
            e = ttk.Entry(form, width=50)
            e.insert(0, str(settings.get(key, "")))
            e.grid(row=i, column=1, sticky="ew", pady=9)
            entries[key] = e

        ttk.Label(form, text="SMTP Security").grid(row=4, column=0, sticky="w", pady=9)
        security_var = StringVar(value=str(settings.get("smtp_security", "SSL")))
        security = ttk.Combobox(
            form, textvariable=security_var,
            values=["SSL", "STARTTLS"], state="readonly", width=47
        )
        security.grid(row=4, column=1, sticky="ew", pady=9)

        ttk.Label(
            form,
            text="Reminder period",
        ).grid(row=5, column=0, sticky="w", pady=9)
        ttk.Label(
            form, text="15 calendar months (fixed)"
        ).grid(row=5, column=1, sticky="w", pady=9)

        ttk.Label(
            form,
            text="SMTP / App Password",
        ).grid(row=6, column=0, sticky="w", pady=9)
        pw = ttk.Entry(form, width=50, show="•")
        pw.insert(0, str(settings.get("smtp_password", "")))
        pw.grid(row=6, column=1, sticky="ew", pady=9)
        ttk.Label(
            form,
            text="This is the password/app-password for the SENDER email account.\n"
                 "For Gmail, use a Google App Password — not your normal Google password.\n"
                 "Save this setting so automatic reminders can send in the background.\n"
                 "ERNEST_APP_PASSWORD is also supported as a deployment fallback.",
            foreground="#94a3b8", background="#0f172a"
        ).grid(row=7, column=0, columnspan=2, sticky="w", pady=(3, 14))

        form.columnconfigure(1, weight=1)

        def save():
            try:
                new = dict(settings)
                for key, entry in entries.items():
                    new[key] = entry.get().strip()
                # Save the sender's SMTP/app password so the automatic
                # background reminder can authenticate without a manual test.
                new["smtp_password"] = pw.get().strip()
                new["smtp_port"] = int(entries["smtp_port"].get().strip())
                new["smtp_security"] = security_var.get()
                new["reminder_months"] = REMINDER_MONTHS
                save_settings(new)

                # Verify the password survives the save before closing the
                # window. This prevents the old "looks saved but comes back
                # empty" problem.
                saved_check = load_settings()
                if new.get("smtp_password", "").strip() and not str(
                    saved_check.get("smtp_password", "")
                ).strip():
                    raise RuntimeError(
                        "The SMTP/App Password could not be persisted. "
                        "The settings file may not be writable."
                    )

                add_audit(self.user, "SETTINGS UPDATED", "", "", "SMTP settings")
                messagebox.showinfo(
                    "Settings saved",
                    "Email settings and the sender App Password were saved.",
                    parent=win
                )
                win.destroy()
            except Exception as exc:
                messagebox.showerror("Could not save settings", str(exc), parent=win)

        ttk.Button(
            win, text="Test Email",
            command=lambda: self.test_email(settings, entries, security_var, pw, win)
        ).pack(side="left", padx=30, pady=20)
        ttk.Button(win, text="Save", command=save).pack(side="right", padx=30, pady=20)

    def test_email(self, settings, entries, security_var, pw, parent):
        test_settings = dict(settings)
        for key, entry in entries.items():
            test_settings[key] = entry.get().strip()
        test_settings["smtp_security"] = security_var.get()
        password = pw.get().strip() or get_smtp_password()
        if not test_settings.get("sender_email") or not test_settings.get("recipient_email") or not password:
            messagebox.showwarning(
                "Missing email configuration",
                "Provide sender, recipient and an SMTP password.",
                parent=parent
            )
            return
        msg = EmailMessage()
        msg["Subject"] = "Time & Tune — Test Email"
        msg["From"] = test_settings["sender_email"]
        msg["To"] = test_settings["recipient_email"]
        msg.set_content(
            "This is a test email from Time & Tune Ernest Money Tracker.\n\n"
            "If you received this, SMTP is working."
        )
        try:
            smtp_send(msg, test_settings, password)
            messagebox.showinfo("Email test successful", "The test email was sent successfully.", parent=parent)
        except Exception as exc:
            messagebox.showerror(
                "Email test failed",
                f"{type(exc).__name__}: {exc}",
                parent=parent
            )

    def change_shared_folder(self):
        selected = _choose_shared_data_folder(self.root)
        if selected:
            messagebox.showinfo(
                "Shared folder changed",
                f"Restart the application to use:\n{selected}",
                parent=self.root
            )

    def manual_refresh(self):
        self.refresh_dashboard()

# ----------------------------- Login / first-run setup -----------------------------

def create_first_admin(parent=None):
    win = Toplevel(parent)
    win.title("First-Time Administrator Setup")
    win.geometry("520x420")
    win.configure(bg="#0f172a")
    win.transient(parent)
    win.grab_set()
    win.lift()
    win.focus_force()
    try:
        win.attributes("-topmost", True)
        win.after(300, lambda: win.attributes("-topmost", False))
    except Exception:
        pass

    ttk.Label(
        win, text="Create Administrator Account",
        font=("Segoe UI", 21, "bold"),
        foreground="#f8fafc", background="#0f172a"
    ).pack(anchor="w", padx=30, pady=(28, 8))
    ttk.Label(
        win,
        text="This is the first account for this shared Ernest Money system.",
        foreground="#94a3b8", background="#0f172a"
    ).pack(anchor="w", padx=30, pady=(0, 18))

    form = ttk.Frame(win, padding=(30, 0))
    form.pack(fill="both", expand=True)
    fields = ["Admin User ID", "Full Name", "Password", "Confirm Password"]
    ent = {}
    for i, f in enumerate(fields):
        ttk.Label(form, text=f).grid(row=i, column=0, sticky="w", pady=10)
        e = ttk.Entry(form, width=35, show="•" if "Password" in f else "")
        e.grid(row=i, column=1, sticky="ew", pady=10)
        ent[f] = e
    form.columnconfigure(1, weight=1)
    result = {"user": None}

    def save():
        uid = ent["Admin User ID"].get().strip().lower()
        name = ent["Full Name"].get().strip()
        pw = ent["Password"].get()
        cpw = ent["Confirm Password"].get()
        if not uid or not name or len(pw) < 8:
            messagebox.showerror(
                "Invalid setup",
                "User ID/name are required and the password must be at least 8 characters.",
                parent=win
            )
            return
        if pw != cpw:
            messagebox.showerror("Password mismatch", "Passwords do not match.", parent=win)
            return

        users = read_excel(users_path(), USERS_COLUMNS)
        if not users.empty:
            messagebox.showerror(
                "Setup already completed",
                "Users already exist in this shared folder.",
                parent=win
            )
            win.destroy()
            return

        row = {
            "User ID": uid,
            "Full Name": name,
            "Role": ROLE_ADMIN,
            "Password Hash": hash_password(pw),
            "Active": "True",
            "Created At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Last Login": "",
        }
        users = pd.DataFrame([row], columns=USERS_COLUMNS)
        if safe_write_excel(users_path(), users):
            result["user"] = row
            add_audit(row, "FIRST ADMIN CREATED", "", "", "Initial system setup")
            win.destroy()

    ttk.Button(win, text="Create Administrator", command=save).pack(pady=(0, 28))
    win.wait_window()
    return result["user"]


def login_screen(root):
    """Return the authenticated user or None."""
    initialize_enterprise_storage()
    users = read_excel(users_path(), USERS_COLUMNS)

    if users.empty:
        admin = create_first_admin(root)
        if admin is None:
            return None
        users = read_excel(users_path(), USERS_COLUMNS)

    result = {"user": None, "closed": False}
    win = Toplevel(root)
    win.title(APP_NAME + " — Login")
    win.geometry("480x400")
    win.configure(bg="#0b1120")
    win.transient(root)
    win.grab_set()
    win.lift()
    win.focus_force()
    try:
        win.attributes("-topmost", True)
        win.after(300, lambda: win.attributes("-topmost", False))
    except Exception:
        pass

    ttk.Label(
        win, text="TIME & TUNE",
        font=("Segoe UI", 28, "bold"),
        foreground="#f8fafc", background="#0b1120"
    ).pack(pady=(40, 4))
    ttk.Label(
        win, text="Ernest Money Tracker",
        font=("Segoe UI", 14),
        foreground="#94a3b8", background="#0b1120"
    ).pack(pady=(0, 28))

    form = ttk.Frame(win, padding=(45, 0))
    form.pack(fill="both", expand=True)
    ttk.Label(form, text="User ID").pack(anchor="w")
    uid = ttk.Entry(form, width=45)
    uid.pack(fill="x", pady=(5, 14))
    ttk.Label(form, text="Password").pack(anchor="w")
    pw = ttk.Entry(form, width=45, show="•")
    pw.pack(fill="x", pady=(5, 15))

    status_label = "Shared data" if _load_shared_path() is not None else "Local data (first setup)"
    status = StringVar(value=f"{status_label}: {DATA_DIR}")
    ttk.Label(
        form, textvariable=status,
        foreground="#64748b", background="#0b1120"
    ).pack(anchor="w", pady=(0, 14))

    def login():
        users_now = read_excel(users_path(), USERS_COLUMNS)
        u = uid.get().strip().lower()
        p = pw.get()
        rows = users_now[
            (users_now["User ID"].astype(str).str.lower() == u) &
            (users_now["Active"].astype(str).str.lower() == "true")
        ]
        if rows.empty or not verify_password(p, str(rows.iloc[0]["Password Hash"])):
            messagebox.showerror(
                "Login failed", "Invalid User ID or password.", parent=win
            )
            return
        user = {c: rows.iloc[0][c] for c in USERS_COLUMNS}
        # Update last login in the central users workbook.
        # Pandas can infer an all-empty Excel column such as "Last Login" as
        # float64 (NaN). Assigning a timestamp string to that dtype raises
        # LossySetitemError/TypeError on newer pandas versions, so explicitly
        # make this column string/object before assigning.
        mask = users_now["User ID"].astype(str).str.lower().eq(u)
        if "Last Login" not in users_now.columns:
            users_now["Last Login"] = ""
        users_now["Last Login"] = users_now["Last Login"].astype("object")
        users_now.loc[mask, "Last Login"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            safe_write_excel(users_path(), users_now)
        except Exception:
            # A login should never fail just because the optional audit field
            # could not be written.
            pass
        result["user"] = user
        win.destroy()

    def register():
        # Self-registration always creates an Employee account. Admin promotion
        # remains an administrator-only action.
        app_stub = type("Stub", (), {"user": {"User ID": "SYSTEM-REG"}})()
        # Reuse a small registration form without exposing role selection.
        create_registration_window(win)

    def choose_shared():
        selected = _choose_shared_data_folder(win)
        if selected:
            messagebox.showinfo(
                "Restart required",
                "The shared folder has been saved.\n\n"
                "Close and reopen the application so it initializes the new shared system.",
                parent=win
            )

    buttons = ttk.Frame(form)
    buttons.pack(fill="x", pady=8)
    ttk.Button(buttons, text="Login", command=login).pack(side="left")
    ttk.Button(buttons, text="Create Employee Account", command=register).pack(side="left", padx=8)
    ttk.Button(buttons, text="Choose Shared Folder", command=choose_shared).pack(side="right")

    win.bind("<Return>", lambda _e: login())
    win.protocol("WM_DELETE_WINDOW", lambda: (result.update(closed=True), win.destroy()))
    win.wait_window()
    return result["user"]


def create_registration_window(parent):
    win = Toplevel(parent)
    win.title("Create Employee Account")
    win.geometry("480x360")
    win.configure(bg="#0f172a")
    win.transient(parent)
    win.grab_set()
    ttk.Label(
        win, text="Create Employee Account",
        font=("Segoe UI", 20, "bold"),
        foreground="#f8fafc", background="#0f172a"
    ).pack(anchor="w", padx=25, pady=(25, 20))
    form = ttk.Frame(win, padding=(25, 0))
    form.pack(fill="both", expand=True)
    fields = ["User ID", "Full Name", "Password", "Confirm Password"]
    ent = {}
    for i, f in enumerate(fields):
        ttk.Label(form, text=f).grid(row=i, column=0, sticky="w", pady=9)
        e = ttk.Entry(form, width=34, show="•" if "Password" in f else "")
        e.grid(row=i, column=1, sticky="ew", pady=9)
        ent[f] = e
    form.columnconfigure(1, weight=1)

    def save():
        uid = ent["User ID"].get().strip().lower()
        name = ent["Full Name"].get().strip()
        pw = ent["Password"].get()
        cpw = ent["Confirm Password"].get()
        if not uid or not name or len(pw) < 8:
            messagebox.showerror(
                "Invalid account",
                "User ID/name are required and password must be at least 8 characters.",
                parent=win
            )
            return
        if pw != cpw:
            messagebox.showerror("Password mismatch", "Passwords do not match.", parent=win)
            return
        users = read_excel(users_path(), USERS_COLUMNS)
        if (users["User ID"].astype(str).str.lower() == uid).any():
            messagebox.showerror("User exists", "That User ID already exists.", parent=win)
            return
        row = {
            "User ID": uid,
            "Full Name": name,
            "Role": ROLE_EMPLOYEE,
            "Password Hash": hash_password(pw),
            "Active": "True",
            "Created At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Last Login": "",
        }
        users = pd.concat([users, pd.DataFrame([row])], ignore_index=True)
        if safe_write_excel(users_path(), users):
            add_audit(row, "EMPLOYEE SELF-REGISTERED", "", "", "New employee account")
            messagebox.showinfo(
                "Account created",
                "Your employee account was created. An administrator can manage your permissions.",
                parent=win
            )
            win.destroy()

    ttk.Button(win, text="Create Account", command=save).pack(pady=(0, 25))


def main():
    root = Tk()

    # Keep the Tk root alive and visible while the login Toplevel is displayed.
    # A withdrawn root can cause its Toplevel to be hidden behind the Windows
    # console on some Windows/Python configurations, which looks like a hang.
    # Use a tiny temporary root; ErnestMoneyApp will resize it to the full GUI
    # immediately after successful login.
    root.geometry("2x2+0+0")
    root.configure(bg="#0b1120")
    root.overrideredirect(True)

    configure_data_paths()
    initialize_enterprise_storage()

    user = login_screen(root)
    if user is None:
        root.destroy()
        return

    root.overrideredirect(False)
    app = ErnestMoneyApp(root, user)
    root.deiconify()
    root.mainloop()


if __name__ == "__main__":
    main()
