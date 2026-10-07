"""
Front Office
=================

The office's app, in tabs that share staff, PINs, settings and backups but keep their
money and logs completely separate:

  Laundry tokens   point of sale for the laundry block's tokens (this file)
  Drawer           the token cash drawer: cash in it now, no sale / change, count, petty cash,
                   banking (DrawerView, this file)
  Rent & deposits  rent, arrears and deposits from receipt to bank (rent.py)
  Keys             keys lent to contractors, from issue to return (keys.py, on logbook.py)
  Post             special post in (CEO/DCEO letters, handed over) and out (post.py, on logbook.py)
  Visitors         the visitor book: signed in and out by staff (visitors.py, on logbook.py)

Laundry tokens:
  * Washing tokens (default £1.00) and dryer tokens (default £0.50)
  * Tenants are picked from tenants.csv (flat number + name)
  * Every transaction needs a staff PIN; the staff member is recorded
  * Opens a cash drawer through a USB trigger box (virtual COM port)
  * Petty cash: money taken out of the drawer, logged with amount and reason
  * Append-only audit log in sales_log.csv (voids are new rows, never edits)
  * Daily totals, reports by day / by flat / petty cash, CSV export
  * Daily backups of the sales log in the backups folder

Folder layout (the folder holding the .exe; when run from source, the folder above source/):
  Front Office.exe
  data/
    config.json      settings, prices, staff (PINs are stored hashed)
    tenants.csv      Laundry block flats: Flat, Tenant Name
    sales_log.csv    every sale, void, petty cash payment and no-sale drawer opening
                     (petty cash rows have a negative total: cash leaving the drawer),
                     the drawer emptied for banking (BANKED, negative total),
                     plus token stock received and counted (no money: total 0)
    (rent files: see rent.py; key files: keys.py; post_log.csv: post.py; visitor_log.csv: visitors.py)
  receipts/          rent receipts as issued (PDF)
  banking/           banking sheet for each rent deposit (PDF)
  keys/              key slips printed for contractors to sign (PDF)
  backups/           one copy of each log per day, last 60 kept
  source/            this file, rent.py, keys.py, post.py, visitors.py, logbook.py, logo, icon and build_exe.bat
"""

import csv
import ctypes
import glob
import hashlib
import hmac
import json
import os
import re
import secrets
import shutil
import subprocess
import sys
import tempfile
import traceback
import tkinter as tk
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from tkinter import filedialog, messagebox, ttk

try:
    import serial
    from serial.tools import list_ports
except ImportError:  # the app still runs (and logs sales) without pyserial
    serial = None
    list_ports = None

try:
    from PIL import Image, ImageTk
except ImportError:  # the logo is then scaled with Tk's own (coarser) PNG support
    Image = ImageTk = None


# --------------------------------------------------------------------------
# Constants
# --------------------------------------------------------------------------

APP_TITLE = "Front Office"
DEFAULT_FLAT_COUNT = 29
DEFAULT_PIN = "1234"


def app_dir():
    """The folder staff open: the .exe, or the folder above source/ when run as a script."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(SOURCE_DIR)


SOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = app_dir()
DATA_DIR = os.path.join(BASE_DIR, "data")
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")
TENANTS_PATH = os.path.join(DATA_DIR, "tenants.csv")
SALES_PATH = os.path.join(DATA_DIR, "sales_log.csv")
BACKUP_DIR = os.path.join(BASE_DIR, "backups")
ERROR_LOG_PATH = os.path.join(DATA_DIR, "error.log")
# Files that older versions kept next to the .exe; moved into data/ on start-up.
DATA_FILES = ("config.json", "tenants.csv", "sales_log.csv", "rent_log.csv", "rent_tenants.csv", "error.log")
LOGO_FILE = "logo.png"
LOGO_MARK_WIDTH = 134  # the house mark is the left 134 px of the logo; used as the window icon
BACKUPS_TO_KEEP = 60

SALES_HEADERS = [
    "Receipt No", "Date & Time", "Type", "Flat", "Tenant Name",
    "Washing Qty", "Dryer Qty", "Total (£)", "Staff", "Ref Receipt", "Notes", "Check",
]
# Each row's Check is a hash of the row plus the previous row's Check, so a row changed,
# added or removed outside the app (e.g. in Excel) no longer matches. This catches edits
# by hand; it can't stop someone who reads this source code and recomputes the hashes.
CHAIN_SALT = b"token-shop/log-chain/v1"
TENANT_HEADERS = ["Flat", "Tenant Name"]

TYPE_SALE = "SALE"
TYPE_VOID = "VOID"
TYPE_NOSALE = "NO SALE"
TYPE_COUNT = "DRAWER COUNT"       # cash counted: Total = counted minus expected, so the drawer figure matches
TYPE_PETTY = "PETTY CASH"
TYPE_ACCEPTED = "LOG ACCEPTED"  # an admin accepted a log that was changed outside the app
TYPE_STOCK_IN = "STOCK IN"        # tokens received into the office (Washing/Dryer Qty = tokens added)
TYPE_STOCK_COUNT = "STOCK COUNT"  # tokens counted by hand (Qty = counted minus expected)
STOCK_TYPES = (TYPE_STOCK_IN, TYPE_STOCK_COUNT)
# The drawer emptied and paid into the bank (token takings account). Total = minus the cash
# taken out. The drawer is always emptied fully, so what should be in it is everything since.
TYPE_BANKED = "BANKED"
DEFAULT_LOW_STOCK = 50
DEFAULT_BANK_LIMITS = {"drawer": 200, "rent": 1500}  # £; Home highlights money waiting to be banked above this
DEFAULT_IDLE_MINUTES = 3

# ESC/POS "kick drawer" command. The trigger box fires on any byte, and this
# also works if a receipt printer is ever put in between.
DRAWER_KICK = b"\x1b\x70\x00\x19\xfa"
BAUD_RATES = ["9600", "19200", "38400", "57600", "115200"]

CSV_DATE_FMT = "%Y-%m-%d %H:%M:%S"
CSV_DATE_FALLBACKS = ["%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"]
UI_DATE_FMT = "%d/%m/%Y"

PIN_ITERATIONS = 120_000
MAX_PIN_ATTEMPTS = 3

# Themes. The names below are the dark theme; apply_theme() swaps them all before any window is
# built (and before rent.py, keys.py and post.py import them), so a theme change needs a restart.
THEMES = {
    "dark": dict(
        BG="#121214", PANEL="#1e1e24", PANEL_2="#26262e", BORDER="#34343f", STATUS_BG="#18181c",
        TEXT="#f5f5f7", MUTED="#9a9aa5", DIM="#5c5c66", TEAL="#00adb5", TEAL_HOVER="#14c6cf",
        ORANGE="#ff5722", ORANGE_HOVER="#ff7449", GREEN="#22a45d", GREEN_HOVER="#2bbd6d", RED="#e5484d",
        RED_HOVER="#ef6a6e", YELLOW="#f5c451", AMBER="#f0a13c", AMBER_HOVER="#f5b35c", PURPLE="#8b6cf0",
        PURPLE_HOVER="#a18af5", BLUE="#4f7cf3", BLUE_HOVER="#6c93f6", PINK="#e0529c", PINK_HOVER="#e873b0",
        TAMPERED_BG="#4a1f24",
        TAMPERED_FG="#ffc9cb", GREY_BTN="#34343e", GREY_HOVER="#44444f", DISABLED_BG="#2a2a31",
        ROW_A="#1e1e24", ROW_B="#23232a"),
    "light": dict(
        BG="#eef0f4", PANEL="#ffffff", PANEL_2="#f1f3f7", BORDER="#d3d7e0", STATUS_BG="#e2e5ec",
        TEXT="#1b1c21", MUTED="#5d6070", DIM="#9a9dab", TEAL="#00949b", TEAL_HOVER="#00aab2",
        ORANGE="#e8491a", ORANGE_HOVER="#f25f33", GREEN="#1b8f4f", GREEN_HOVER="#22a45d", RED="#d33a40",
        RED_HOVER="#e0555a", YELLOW="#a86b00", AMBER="#c46a00", AMBER_HOVER="#d97c0f", PURPLE="#6e4fe0",
        PURPLE_HOVER="#8466ea", BLUE="#2f63e0", BLUE_HOVER="#4a78e8", PINK="#c2367f", PINK_HOVER="#d04e93",
        TAMPERED_BG="#fde1e2",
        TAMPERED_FG="#9b1c22", GREY_BTN="#e1e4ea", GREY_HOVER="#d2d6de", DISABLED_BG="#eceef2",
        ROW_A="#ffffff", ROW_B="#f6f7fa"),
}
globals().update(THEMES["dark"])


def apply_theme(name):
    globals().update(THEMES.get(name, THEMES["dark"]))
FONT = "Segoe UI"


# --------------------------------------------------------------------------
# Helpers: money, text, PINs
# --------------------------------------------------------------------------

def fmt_money(pence):
    sign = "-" if pence < 0 else ""
    return f"{sign}£{abs(pence) / 100:,.2f}"


def csv_money(pence):
    return f"{pence / 100:.2f}"


def parse_pence(text):
    """'£1.50' / '1.5' / '-2' -> pence as int, or None if not a number."""
    s = str(text).replace("£", "").replace(",", "").strip()
    if not s:
        return 0
    try:
        return int((Decimal(s) * 100).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    except InvalidOperation:
        return None


def parse_int(text, default=0):
    try:
        return int(Decimal(str(text).strip() or "0"))
    except (InvalidOperation, ValueError):
        return default


def natural_key(text):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", str(text))]


def describe_items(wash, dry):
    parts = []
    if wash:
        parts.append(f"{abs(wash)} × washing")
    if dry:
        parts.append(f"{abs(dry)} × dryer")
    return ", ".join(parts) if parts else "no tokens"


def hash_pin(pin, salt=None):
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", pin.encode(), bytes.fromhex(salt), PIN_ITERATIONS)
    return salt, digest.hex()


def pin_matches(pin, staff):
    _, digest = hash_pin(pin, staff["salt"])
    return hmac.compare_digest(digest, staff["pin_hash"])


def valid_pin_format(pin):
    return pin.isdigit() and 4 <= len(pin) <= 8


def parse_ui_date(text):
    text = text.strip()
    for fmt in (UI_DATE_FMT, "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


# --------------------------------------------------------------------------
# Storage: config, tenants, sales log, backups
# --------------------------------------------------------------------------

def atomic_write_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    os.replace(tmp, path)


def make_staff(name, pin, admin, checker=False):
    salt, digest = hash_pin(pin)
    return {"name": name, "salt": salt, "pin_hash": digest, "admin": bool(admin), "checker": bool(checker)}


def load_config():
    """Returns (config, warning_or_None)."""
    cfg = {
        "com_port": "COM3",
        "baud_rate": 9600,
        "drawer_enabled": True,
        "price_washing_pence": 100,
        "price_dryer_pence": 50,
        "staff": [],
    }
    warning = None
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                cfg.update(json.load(f))
        except (OSError, ValueError) as e:
            broken = f"{CONFIG_PATH}.broken-{datetime.now():%Y%m%d-%H%M%S}"
            try:
                shutil.move(CONFIG_PATH, broken)
            except OSError:
                pass
            warning = (f"config.json could not be read ({e}).\n\nIt has been moved to:\n{broken}\n\n"
                       f"Default settings are being used and the admin PIN is {DEFAULT_PIN}.")
    if not cfg["staff"]:
        cfg["staff"] = [make_staff("Admin", DEFAULT_PIN, True)]
    save_config(cfg)
    return cfg, warning


def save_config(cfg):
    atomic_write_json(CONFIG_PATH, cfg)


def load_tenants():
    if not os.path.exists(TENANTS_PATH):
        tenants = [{"flat": str(i), "name": ""} for i in range(1, DEFAULT_FLAT_COUNT + 1)]
        save_tenants(tenants)
        return tenants
    tenants, seen = [], set()
    with open(TENANTS_PATH, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            flat = (row.get("Flat") or "").strip()
            if flat and flat.lower() not in seen:
                seen.add(flat.lower())
                name = (row.get("Tenant Name") or row.get("Name") or "").strip()
                tenants.append({"flat": flat, "name": name})
    tenants.sort(key=lambda t: natural_key(t["flat"]))
    return tenants


def save_tenants(tenants):
    tmp = TENANTS_PATH + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(TENANT_HEADERS)
        for t in sorted(tenants, key=lambda t: natural_key(t["flat"])):
            w.writerow([t["flat"], t["name"]])
    os.replace(tmp, TENANTS_PATH)


def ensure_sales_log():
    if not os.path.exists(SALES_PATH) or os.path.getsize(SALES_PATH) == 0:
        with open(SALES_PATH, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow(SALES_HEADERS)


def parse_csv_datetime(text):
    text = (text or "").strip()
    for fmt in [CSV_DATE_FMT] + CSV_DATE_FALLBACKS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def row_to_record(row):
    receipt = parse_int(row.get("Receipt No"), default=None)
    dt = parse_csv_datetime(row.get("Date & Time"))
    total = parse_pence(row.get("Total (£)", ""))
    if receipt is None or dt is None or total is None:
        return None
    rtype = (row.get("Type") or TYPE_SALE).strip().upper()
    if rtype not in (TYPE_SALE, TYPE_VOID, TYPE_NOSALE, TYPE_COUNT, TYPE_PETTY, TYPE_ACCEPTED, TYPE_BANKED) + STOCK_TYPES:
        return None
    return {
        "receipt": receipt,
        "dt": dt,
        "type": rtype,
        "flat": (row.get("Flat") or "").strip(),
        "name": (row.get("Tenant Name") or "").strip(),
        "wash": parse_int(row.get("Washing Qty")),
        "dry": parse_int(row.get("Dryer Qty")),
        "total": total,
        "staff": (row.get("Staff") or "").strip(),
        "ref": parse_int(row.get("Ref Receipt"), default=None) if (row.get("Ref Receipt") or "").strip() else None,
        "notes": (row.get("Notes") or "").strip(),
        "check": (row.get("Check") or "").strip(),
        "tampered": False,
    }


def record_to_row(rec):
    return [
        rec["receipt"], rec["dt"].strftime(CSV_DATE_FMT), rec["type"], rec["flat"], rec["name"],
        rec["wash"], rec["dry"], csv_money(rec["total"]), rec["staff"],
        rec["ref"] if rec["ref"] is not None else "", rec["notes"], rec.get("check", ""),
    ]


def row_check(prev_check, rec):
    """Hash of the row's values (not its text), so Excel re-formatting dates or dropping
    '.00' from amounts on save doesn't count as a change. Times are compared to the minute
    because Excel drops the seconds. The 'C' prefix stops Excel turning it into a number."""
    return chain_hash(CHAIN_SALT, prev_check, [
        rec["receipt"], rec["dt"].strftime("%Y-%m-%d %H:%M"), rec["type"], rec["flat"], rec["name"],
        rec["wash"], rec["dry"], rec["total"], rec["staff"], "" if rec["ref"] is None else rec["ref"], rec["notes"]])


def chain_hash(salt, prev_check, fields):
    data = salt + prev_check.encode() + b"\x1e" + "\x1f".join(map(str, fields)).encode("utf-8")
    return "C" + hashlib.sha256(data).hexdigest()[:16].upper()


def read_csv_rows(path):
    """Returns (headers, non-blank rows as dicts)."""
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = [row for row in reader
                if any((v or "").strip() for v in row.values() if isinstance(v, str))]
        return reader.fieldnames or [], rows


def read_sales_rows():
    return read_csv_rows(SALES_PATH)


def upgrade_sales_log():
    """Adds the Check column to a log written by an older version, sealing the rows as they
    are now. A copy of the old file is kept in backups. Returns True if it upgraded."""
    if not os.path.exists(SALES_PATH):
        return False
    if "Check" in read_sales_rows()[0]:
        return False
    reseal_sales_log("before_check_column")
    return True


def reseal_sales_log(backup_label):
    """Recalculates every Check, accepting the log as it is now. The file as it was is copied
    to backups first. Returns the backup's path."""
    headers, rows = read_sales_rows()
    os.makedirs(BACKUP_DIR, exist_ok=True)
    backup = os.path.join(BACKUP_DIR, f"sales_log_{backup_label}_{datetime.now():%Y-%m-%d_%H%M%S}.csv")
    shutil.copy2(SALES_PATH, backup)
    prev, out = "", []
    for row in rows:
        rec = row_to_record(row)
        if rec is None:  # unreadable: keep it as it was, unsealed
            out.append([row.get(h) or "" for h in SALES_HEADERS])
            continue
        rec["check"] = prev = row_check(prev, rec)
        out.append(record_to_row(rec))
    tmp = SALES_PATH + ".tmp"
    write_csv_export(tmp, SALES_HEADERS, out)
    os.replace(tmp, SALES_PATH)
    return backup


def load_records():
    """Returns (records sorted oldest first, number of unreadable rows, Check of the last row).
    Rows whose Check doesn't match are marked rec["tampered"] = True."""
    records, skipped, prev = [], 0, ""
    if not os.path.exists(SALES_PATH):
        return records, skipped, prev
    for row in read_sales_rows()[1]:
        stored = (row.get("Check") or "").strip()
        rec = row_to_record(row)
        if rec is None:
            skipped += 1
        else:
            rec["tampered"] = stored != row_check(prev, rec)
            records.append(rec)
        prev = stored or prev  # a row added by hand (no Check) flags only itself
    records.sort(key=lambda r: r["receipt"])
    return records, skipped, prev


def append_record(rec):
    """Raises OSError (e.g. PermissionError when Excel has the file open)."""
    append_csv_rows(SALES_PATH, SALES_HEADERS, [record_to_row(rec)])


def append_csv_rows(path, headers, rows):
    """Appends rows to a CSV log, creating it with headers if needed. Raises OSError."""
    if not os.path.exists(path) or os.path.getsize(path) == 0:
        write_csv_export(path, headers, [])
    with open(path, "rb") as f:
        f.seek(-1, os.SEEK_END)
        needs_newline = f.read(1) not in (b"\n", b"\r")
    with open(path, "a", newline="", encoding="utf-8") as f:
        if needs_newline:
            f.write("\r\n")
        csv.writer(f).writerows(rows)


def backup_log(path, prefix):
    """One dated copy per day in backups/, keeping the last BACKUPS_TO_KEEP."""
    if not os.path.exists(path):
        return
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        dst = os.path.join(BACKUP_DIR, f"{prefix}_{date.today():%Y-%m-%d}.csv")
        if not os.path.exists(dst):
            shutil.copy2(path, dst)
        old = sorted(glob.glob(os.path.join(BACKUP_DIR, f"{prefix}_????-??-??.csv")))
        for old_path in old[:-BACKUPS_TO_KEEP]:
            os.remove(old_path)
    except OSError:
        pass  # a failed backup must never stop the shop opening


def prepare_data_dir():
    """Creates data/ and moves in any data files an older version left next to the .exe."""
    os.makedirs(DATA_DIR, exist_ok=True)
    for name in DATA_FILES:
        old, new = os.path.join(BASE_DIR, name), os.path.join(DATA_DIR, name)
        if os.path.exists(old) and not os.path.exists(new):
            shutil.move(old, new)


def find_logo():
    """A logo next to the .exe wins (so it can be swapped); otherwise the one in source/ or
    the copy bundled into the .exe."""
    for folder in (BASE_DIR, SOURCE_DIR, getattr(sys, "_MEIPASS", None)):
        if folder and os.path.exists(os.path.join(folder, LOGO_FILE)):
            return os.path.join(folder, LOGO_FILE)
    return None


def write_csv_export(path, headers, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerows(rows)


# --------------------------------------------------------------------------
# Cash drawer
# --------------------------------------------------------------------------

def available_ports():
    if list_ports is None:
        return []
    try:
        return sorted(list_ports.comports(), key=lambda p: natural_key(p.device))
    except Exception:
        return []


def fire_drawer(port, baud, enabled=True):
    """Returns (ok, message). Never raises."""
    if not enabled:
        return True, "Drawer is switched off in Settings"
    if serial is None:
        return False, "The serial library (pyserial) is not available."
    if not port:
        return False, "No COM port is set. Choose one in Settings."
    try:
        with serial.Serial(port=port, baudrate=int(baud), timeout=1, write_timeout=2) as conn:
            conn.write(DRAWER_KICK)
            conn.flush()
        return True, f"Drawer opened ({port})"
    except Exception as e:  # unplugged, wrong port, port busy, driver error...
        return False, f"{port}: {e}"


# --------------------------------------------------------------------------
# Widgets
# --------------------------------------------------------------------------

def entry_opts(**kw):
    opts = dict(bg=PANEL_2, fg=TEXT, insertbackground=TEXT, relief="flat", highlightthickness=1,
                highlightbackground=BORDER, highlightcolor=TEAL, font=(FONT, 11),
                disabledbackground=PANEL, disabledforeground=MUTED)
    opts.update(kw)
    return opts


def center_window(win, parent):
    win.update_idletasks()
    w, h = win.winfo_reqwidth(), win.winfo_reqheight()
    px, py = parent.winfo_rootx(), parent.winfo_rooty()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    x = max(0, px + (pw - w) // 2)
    y = max(0, py + (ph - h) // 3)
    win.geometry(f"+{x}+{y}")


class HoverButton(tk.Button):
    def __init__(self, master, bg=None, hover=None, fg=None, font=None, **kw):
        # colours are looked up now, not as defaults, so they follow the theme
        bg, hover, fg = bg or GREY_BTN, hover or GREY_HOVER, fg or TEXT
        opts = dict(bg=bg, fg=fg, activebackground=hover, activeforeground=fg, relief="flat", bd=0,
                    highlightthickness=0, cursor="hand2", font=font or (FONT, 11, "bold"),
                    disabledforeground=DIM, padx=14, pady=8)
        opts.update(kw)
        super().__init__(master, **opts)
        self._bg, self._hover, self._fg = bg, hover, fg
        self.bind("<Enter>", self._on_enter, add="+")
        self.bind("<Leave>", self._on_leave, add="+")

    def _on_enter(self, _):
        if str(self["state"]) != "disabled":
            self.config(bg=self._hover)

    def _on_leave(self, _):
        if str(self["state"]) != "disabled":
            self.config(bg=self._bg)

    def set_colors(self, bg, hover, fg=None):
        fg = fg or TEXT
        self._bg, self._hover, self._fg = bg, hover, fg
        if str(self["state"]) != "disabled":
            self.config(bg=bg, activebackground=hover, fg=fg, activeforeground=fg)

    def set_enabled(self, enabled):
        if enabled:
            self.config(state="normal", bg=self._bg, fg=self._fg, activebackground=self._hover,
                        activeforeground=self._fg, cursor="hand2")
        else:
            self.config(state="disabled", bg=DISABLED_BG, cursor="arrow")


def _sort_key(text):
    """Sorts what a table shows: amounts as numbers, dd/mm/yyyy dates as dates, K0012 or 6201 in
    number order, anything else alphabetically. Blanks go last."""
    t = str(text).strip()
    if not t or t in ("—", "-"):
        return (3, ())
    m = re.fullmatch(r"(-?)£([\d,]+\.\d\d)", t.split("  ")[0])
    if m:
        return (0, (float(m.group(2).replace(",", "")) * (-1 if m.group(1) else 1),))
    m = re.match(r"(\d\d)/(\d\d)/(\d{4})(?: (\d\d):(\d\d))?", t)
    if m:
        d, mo, y, h, mi = m.groups()
        return (1, (int(y), int(mo), int(d), int(h or 0), int(mi or 0)))
    return (2, tuple(natural_key(t)))


def make_sortable(tree):
    """Click a column heading to sort by it; click again to reverse. The sort is kept when the
    table refreshes: call tree.resort() at the end of each refresh."""
    state = {"col": None, "rev": False}
    titles = {c: tree.heading(c, "text") for c in tree["columns"]}

    def resort():
        col = state["col"]
        if col is None:
            return
        rows = sorted(tree.get_children(""), key=lambda i: _sort_key(tree.set(i, col)), reverse=state["rev"])
        for n, iid in enumerate(rows):
            tree.move(iid, "", n)
            tags = [t for t in tree.item(iid, "tags") if t not in ("odd", "even")]
            tree.item(iid, tags=["odd" if n % 2 else "even"] + tags)

    def click(col):
        state["rev"] = not state["rev"] if state["col"] == col else False
        state["col"] = col
        for c, title in titles.items():
            tree.heading(c, text=title + ((" ▼" if state["rev"] else " ▲") if c == col else ""))
        resort()

    for c in titles:
        tree.heading(c, command=lambda c=c: click(c))
    tree.resort = resort


class PlaceholderEntry(tk.Entry):
    def __init__(self, master, placeholder, **kw):
        super().__init__(master, **entry_opts(**kw))
        self._placeholder = placeholder
        self._showing = False
        self.bind("<FocusIn>", self._focus_in, add="+")
        self.bind("<FocusOut>", self._focus_out, add="+")
        self._show_placeholder()

    def _show_placeholder(self):
        if not self.get():
            self._showing = True
            self.insert(0, self._placeholder)
            self.config(fg=DIM)

    def _focus_in(self, _):
        if self._showing:
            self.delete(0, "end")
            self.config(fg=TEXT)
            self._showing = False

    def _focus_out(self, _):
        self._show_placeholder()

    def value(self):
        return "" if self._showing else self.get().strip()

    def set(self, text):
        self.delete(0, "end")
        self._showing = False
        self.config(fg=TEXT)
        self.insert(0, text)
        if not text and self.focus_get() is not self:
            self._show_placeholder()

    def clear(self):
        self.delete(0, "end")
        self._showing = False
        self.config(fg=TEXT)
        if self.focus_get() is not self:
            self._show_placeholder()


class TokenStepper(tk.Frame):
    MAX_QTY = 99

    def __init__(self, master, title, color, hover, on_change):
        super().__init__(master, bg=PANEL_2)
        self.qty = 0
        self._on_change = on_change
        tk.Frame(self, bg=color, width=4).pack(side="left", fill="y")
        self.plus = HoverButton(self, text="+", bg=color, hover=hover, font=(FONT, 16, "bold"),
                                width=3, pady=2, command=lambda: self.change(1))
        self.plus.pack(side="right", padx=(4, 10), pady=10)
        self.qty_label = tk.Label(self, text="0", width=3, font=(FONT, 17, "bold"), fg=TEXT, bg=PANEL_2)
        self.qty_label.pack(side="right")
        self.minus = HoverButton(self, text="−", font=(FONT, 16, "bold"), width=3, pady=2,
                                 command=lambda: self.change(-1))
        self.minus.pack(side="right", padx=(10, 4), pady=10)
        info = tk.Frame(self, bg=PANEL_2)
        info.pack(side="left", fill="both", expand=True, padx=12, pady=8)
        tk.Label(info, text=title, font=(FONT, 12, "bold"), fg=TEXT, bg=PANEL_2, anchor="w").pack(fill="x")
        self.price_label = tk.Label(info, text="", font=(FONT, 10), fg=MUTED, bg=PANEL_2, anchor="w")
        self.price_label.pack(fill="x")
        self.stock_label = tk.Label(info, text="", font=(FONT, 9), fg=MUTED, bg=PANEL_2, anchor="w")

    def set_price(self, pence):
        self.price_label.config(text=f"{fmt_money(pence)} each")

    def set_stock(self, count, low):
        """count: tokens in stock, or None when stock isn't tracked."""
        if count is None:
            self.stock_label.pack_forget()
            return
        self.stock_label.config(text=f"{count} in stock" + ("  ·  low" if count < low else ""),
                                fg=AMBER if count < low else MUTED)
        self.stock_label.pack(fill="x")

    def change(self, delta):
        self.set(self.qty + delta)

    def set(self, qty):
        self.qty = max(0, min(self.MAX_QTY, qty))
        self.qty_label.config(text=str(self.qty), fg=TEXT if self.qty else MUTED)
        self.minus.set_enabled(self.qty > 0)
        self._on_change()


# --------------------------------------------------------------------------
# Dialogs
# --------------------------------------------------------------------------

class Dialog(tk.Toplevel):
    """Dark modal dialog base. Call show() to block until closed."""

    def __init__(self, parent, title, resizable=False):
        super().__init__(parent)
        self.parent = parent
        self.result = None
        self.withdraw()
        self.title(title)
        self.configure(bg=PANEL)
        self.resizable(resizable, resizable)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.bind("<Escape>", lambda e: self.cancel())
        self.body = tk.Frame(self, bg=PANEL)
        self.body.pack(fill="both", expand=True, padx=26, pady=22)
        self._initial_focus = None

    def heading(self, text, sub=None):
        tk.Label(self.body, text=text, font=(FONT, 15, "bold"), fg=TEXT, bg=PANEL,
                 anchor="w", justify="left").pack(fill="x")
        if sub:
            tk.Label(self.body, text=sub, font=(FONT, 10), fg=MUTED, bg=PANEL, anchor="w",
                     justify="left", wraplength=420).pack(fill="x", pady=(2, 0))

    def label(self, text, top=14):
        tk.Label(self.body, text=text, font=(FONT, 10, "bold"), fg=MUTED, bg=PANEL,
                 anchor="w").pack(fill="x", pady=(top, 4))

    def button_row(self, confirm_text, confirm_bg=None, confirm_hover=None, cancel_text="Cancel"):
        confirm_bg, confirm_hover = confirm_bg or TEAL, confirm_hover or TEAL_HOVER
        row = tk.Frame(self.body, bg=PANEL)
        row.pack(fill="x", pady=(18, 0))
        HoverButton(row, text=confirm_text, bg=confirm_bg, hover=confirm_hover,
                    command=self.confirm).pack(side="right")
        if cancel_text:
            HoverButton(row, text=cancel_text, command=self.cancel).pack(side="right", padx=(0, 8))
        self.bind("<Return>", lambda e: self.confirm())
        return row

    def show(self):
        center_window(self, self.parent)
        self.deiconify()
        self.lift()
        self.grab_set()
        (self._initial_focus or self).focus_force()
        self.wait_window()
        return self.result

    def confirm(self):
        self.close()

    def cancel(self):
        self.result = None
        self.close()

    def close(self):
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()
        if isinstance(self.parent, tk.Toplevel) and self.parent.winfo_exists():
            self.parent.grab_set()
            self.parent.focus_force()


class PinDialog(Dialog):
    """Asks for a staff PIN (optionally an amount and a reason). result = (staff, reason).
    When amount_label is given, the amount entered is in self.amount_pence.
    allow(staff) can return a reason to refuse that person, e.g. "not a checker".
    cash_due (pence) adds an optional "cash given" box that works out the change: self.given_pence
    and self.change_pence afterwards (None if left empty, i.e. the exact money)."""

    def __init__(self, parent, app, title, heading, lines, admin_only=False, reason_label=None,
                 confirm_text="Confirm", confirm_bg=None, confirm_hover=None, amount_label=None,
                 allow=None, reason_value="", cash_due=None):
        super().__init__(parent, title)
        confirm_bg, confirm_hover = confirm_bg or GREEN, confirm_hover or GREEN_HOVER
        self.app = app
        self.admin_only = admin_only
        self.allow = allow
        self.attempts = 0
        self.locked_out = False
        self.heading(heading)
        box = tk.Frame(self.body, bg=PANEL_2)
        box.pack(fill="x", pady=(12, 0))
        for i, line in enumerate(lines):
            tk.Label(box, text=line, font=(FONT, 11, "bold" if i == 0 else "normal"),
                     fg=TEXT if i == 0 else MUTED, bg=PANEL_2, anchor="w", justify="left",
                     wraplength=380).pack(fill="x", padx=14, pady=(10 if i == 0 else 0, 0))
        tk.Frame(box, bg=PANEL_2, height=10).pack()
        self.amount = None
        self.amount_pence = None
        if amount_label:
            self.label(amount_label)
            self.amount = tk.Entry(self.body, width=12, **entry_opts(font=(FONT, 16)))
            self.amount.pack(fill="x", ipady=3)
        self.cash_due = cash_due
        self.given = None
        self.given_pence = self.change_pence = None
        if cash_due is not None:
            self.label("Cash given (£)  ·  leave empty if it's the exact money")
            row = tk.Frame(self.body, bg=PANEL)
            row.pack(fill="x")
            self.given_var = tk.StringVar()
            self.given = tk.Entry(row, width=8, textvariable=self.given_var, **entry_opts(font=(FONT, 16)))
            self.given.pack(side="left", fill="x", expand=True, ipady=3)
            for pounds in (5, 10, 20, 50):
                if pounds * 100 > cash_due:
                    HoverButton(row, text=f"£{pounds}", font=(FONT, 12, "bold"),
                                command=lambda p=pounds: self.given_var.set(str(p))).pack(side="left", padx=(6, 0))
            self.change_label = tk.Label(self.body, text="", font=(FONT, 16, "bold"), fg=MUTED, bg=PANEL,
                                         anchor="w")
            self.change_label.pack(fill="x", pady=(6, 0))
            self.given_var.trace_add("write", lambda *a: self._show_change())
        self.reason = None
        if reason_label:
            self.label(reason_label)
            self.reason = tk.Entry(self.body, width=40, **entry_opts())
            self.reason.insert(0, reason_value)
            self.reason.pack(fill="x", ipady=5)
        self.label("Admin PIN" if admin_only else "Staff PIN")
        vcmd = (self.register(lambda p: p == "" or (p.isdigit() and len(p) <= 8)), "%P")
        self.pin = tk.Entry(self.body, show="•", width=10, justify="center", validate="key",
                            validatecommand=vcmd, **entry_opts(font=(FONT, 20)))
        self.pin.pack(fill="x", ipady=4)
        self.error = tk.Label(self.body, text="", font=(FONT, 10), fg=RED, bg=PANEL, anchor="w",
                              justify="left", wraplength=380)
        self.error.pack(fill="x", pady=(6, 0))
        self.button_row(confirm_text, confirm_bg, confirm_hover)
        self._initial_focus = next(w for w in (self.amount, self.given, self.reason, self.pin) if w is not None)

    def _show_change(self):
        given = parse_pence(self.given_var.get())
        if not self.given_var.get().strip():
            self.change_label.config(text="", fg=MUTED)
        elif given is None:
            self.change_label.config(text="Enter an amount, e.g. 50", fg=RED)
        elif given < self.cash_due:
            self.change_label.config(text=f"Not enough: {fmt_money(self.cash_due - given)} short", fg=RED)
        else:
            self.change_label.config(text=f"Change to give: {fmt_money(given - self.cash_due)}", fg=GREEN)

    def confirm(self):
        if self.given is not None and self.given_var.get().strip():
            given = parse_pence(self.given_var.get())
            if given is None or given < self.cash_due:
                self.error.config(text=f"Cash given must be at least {fmt_money(self.cash_due)}, "
                                       "or leave it empty for the exact money.")
                self.given.focus_set()
                return
            self.given_pence, self.change_pence = given, given - self.cash_due
        if self.amount is not None:
            pence = parse_pence(self.amount.get())
            if not pence or pence < 0:
                self.error.config(text="Please enter an amount, e.g. 5.00")
                self.amount.focus_set()
                return
            self.amount_pence = pence
        reason = self.reason.get().strip() if self.reason is not None else ""
        if self.reason is not None and not reason:
            self.error.config(text="Please enter a reason.")
            self.reason.focus_set()
            return
        staff = self.app.find_staff(self.pin.get())
        if staff is None:
            problem = "Incorrect PIN"
        elif self.admin_only and not staff["admin"]:
            problem = "That PIN does not have admin rights"
        else:
            problem = self.allow(staff) if self.allow else None
        if problem is None:
            self.result = (staff, reason)
            self.close()
            return
        self.attempts += 1
        if self.attempts >= MAX_PIN_ATTEMPTS:
            self.locked_out = True
            self.cancel()
            return
        left = MAX_PIN_ATTEMPTS - self.attempts
        self.error.config(text=f"{problem}. {left} attempt{'s' if left != 1 else ''} left.")
        self.pin.delete(0, "end")
        self.pin.focus_set()


def admin_save_path(app, parent, **kw):
    """A Save As box is a small file browser (it can open, rename and delete files, data\\ included),
    so on the front-desk PC only an admin may open one. Returns the path chosen, or ''."""
    dlg = PinDialog(parent, app, "Export", "Save to a file", ["An admin PIN is needed to save files."],
                    admin_only=True, confirm_text="Continue", confirm_bg=TEAL, confirm_hover=TEAL_HOVER)
    if not dlg.show():
        app._pin_locked_out(dlg)
        return ""
    return filedialog.asksaveasfilename(parent=parent, **kw)


class FieldsDialog(Dialog):
    """Small form. fields: dicts with key, label, kind ('text'|'pin'|'check'), value."""

    def __init__(self, parent, title, heading, fields, validate=None, confirm_text="Save", sub=None):
        super().__init__(parent, title)
        self.validate = validate
        self.fields = fields
        self.vars = {}
        self.heading(heading, sub)
        pin_vcmd = (self.register(lambda p: p == "" or (p.isdigit() and len(p) <= 8)), "%P")
        for f in fields:
            kind = f.get("kind", "text")
            if kind == "check":
                var = tk.BooleanVar(value=bool(f.get("value")))
                tk.Checkbutton(self.body, text=f["label"], variable=var, font=(FONT, 11), fg=TEXT,
                               bg=PANEL, selectcolor=PANEL_2, activebackground=PANEL,
                               activeforeground=TEXT, anchor="w", bd=0,
                               highlightthickness=0).pack(fill="x", pady=(14, 0))
            else:
                var = tk.StringVar(value=f.get("value", ""))
                self.label(f["label"])
                extra = dict(show="•", validate="key", validatecommand=pin_vcmd) if kind == "pin" else {}
                entry = tk.Entry(self.body, textvariable=var, width=36, **extra, **entry_opts())
                entry.pack(fill="x", ipady=5)
                if self._initial_focus is None:
                    self._initial_focus = entry
            self.vars[f["key"]] = var
        self.error = tk.Label(self.body, text="", font=(FONT, 10), fg=RED, bg=PANEL, anchor="w",
                              justify="left", wraplength=380)
        self.error.pack(fill="x", pady=(8, 0))
        self.button_row(confirm_text)

    def confirm(self):
        values = {}
        for key, var in self.vars.items():
            v = var.get()
            values[key] = v.strip() if isinstance(v, str) else v
        if self.validate:
            problem = self.validate(values)
            if problem:
                self.error.config(text=problem)
                return
        self.result = values
        self.close()


# --------------------------------------------------------------------------
# Reports window
# --------------------------------------------------------------------------

def summarise(records, petty_receipts=frozenset()):
    """Token takings and petty cash are kept apart. petty_receipts is the set of receipt
    numbers that are petty cash, so voids of petty cash count against petty cash, not takings.
    by_day values are [washing, dryer, takings, sales, petty cash]."""
    s = {"wash": 0, "dry": 0, "total": 0, "sales": 0, "voids": 0, "nosales": 0,
         "petty": 0, "petty_count": 0, "by_day": {}, "by_flat": {}}
    for r in records:
        if r["type"] == TYPE_NOSALE:
            s["nosales"] += 1
            continue
        if r["type"] in (TYPE_ACCEPTED, TYPE_BANKED, TYPE_COUNT) or r["type"] in STOCK_TYPES:
            continue  # a drawer count corrects the cash in the drawer; it isn't takings
        day = s["by_day"].setdefault(r["dt"].date(), [0, 0, 0, 0, 0])
        if r["type"] == TYPE_PETTY or (r["type"] == TYPE_VOID and r["ref"] in petty_receipts):
            s["petty"] -= r["total"]  # petty cash rows are negative, their voids positive
            day[4] -= r["total"]
            if r["type"] == TYPE_PETTY:
                s["petty_count"] += 1
            continue
        if r["type"] == TYPE_SALE:
            s["sales"] += 1
        else:
            s["voids"] += 1
        s["wash"] += r["wash"]
        s["dry"] += r["dry"]
        s["total"] += r["total"]
        day[0] += r["wash"]
        day[1] += r["dry"]
        day[2] += r["total"]
        day[3] += 1 if r["type"] == TYPE_SALE else 0
        flat = s["by_flat"].setdefault(r["flat"], {"name": r["name"], "wash": 0, "dry": 0, "total": 0})
        flat["wash"] += r["wash"]
        flat["dry"] += r["dry"]
        flat["total"] += r["total"]
        if r["name"]:
            flat["name"] = r["name"]
    return s


def token_stock(records):
    """Tokens in the office: received, plus count corrections, minus sold, plus voided sales.
    Counted from the first stock entry on, so sales before stock was set up don't count.
    Returns None until there is a stock entry, else {"wash": n, "dry": n}."""
    start = next((r["receipt"] for r in records if r["type"] in STOCK_TYPES), None)
    if start is None:
        return None
    stock = {"wash": 0, "dry": 0}
    for r in records:
        if r["receipt"] < start:
            continue
        if r["type"] in STOCK_TYPES:
            sign = 1
        elif r["type"] == TYPE_SALE or (r["type"] == TYPE_VOID and r["ref"] is not None and r["ref"] >= start):
            sign = -1  # a void's quantities are negative, so it puts the tokens back
        else:
            continue
        stock["wash"] += sign * r["wash"]
        stock["dry"] += sign * r["dry"]
    return stock


def drawer_expected(records):
    """Cash that should be in the drawer now: every sale, void, petty cash payment and count
    correction since it was last emptied for banking."""
    last = max((i for i, r in enumerate(records) if r["type"] == TYPE_BANKED), default=-1)
    return sum(r["total"] for r in records[last + 1:] if r["type"] != TYPE_BANKED)


def bankings(records):
    """Every time the drawer was emptied for banking, oldest first, with the transactions it
    covered: {"rec", "rows", "expected", "paid", "difference", "slip"}."""
    out, since = [], []
    for r in records:
        if r["type"] == TYPE_BANKED:
            expected = sum(x["total"] for x in since)
            out.append({"rec": r, "rows": since, "expected": expected, "paid": -r["total"],
                        "difference": -r["total"] - expected, "slip": banking_slip(r)})
            since = []
        else:
            since.append(r)
    return out


def banking_note(slip, expected, paid):
    diff = paid - expected
    text = f"Slip {slip}" if slip else "No slip ref"
    text += f" · expected {fmt_money(expected)}, counted {fmt_money(paid)}"
    if diff:
        text += f" ({'over' if diff > 0 else 'short'} {fmt_money(abs(diff))})"
    return text


def banking_slip(rec):
    first = rec["notes"].split(" · ")[0]
    return first[5:] if first.startswith("Slip ") else ""


class ReportsWindow(Dialog):
    def __init__(self, app):
        super().__init__(app, "Reports", resizable=True)
        self.app = app
        self.geometry(f"{app.S(1140)}x{app.S(680)}")
        self.minsize(app.S(1080), app.S(560))
        self.heading("Reports", "Totals for a date range. Voids are subtracted. Petty cash is kept separate "
                                "from takings; net cash is takings minus petty cash. Export to CSV opens in Excel.")

        bar = tk.Frame(self.body, bg=PANEL)
        bar.pack(fill="x", pady=(16, 0))
        tk.Label(bar, text="From", font=(FONT, 10, "bold"), fg=MUTED, bg=PANEL).pack(side="left")
        self.from_entry = tk.Entry(bar, width=11, **entry_opts())
        self.from_entry.pack(side="left", padx=(6, 12), ipady=4)
        tk.Label(bar, text="To", font=(FONT, 10, "bold"), fg=MUTED, bg=PANEL).pack(side="left")
        self.to_entry = tk.Entry(bar, width=11, **entry_opts())
        self.to_entry.pack(side="left", padx=(6, 8), ipady=4)
        HoverButton(bar, text="Show", bg=TEAL, hover=TEAL_HOVER, font=(FONT, 10, "bold"),
                    command=self.run).pack(side="left", padx=(0, 16))
        today = date.today()
        month_start = today.replace(day=1)
        last_month_end = month_start - timedelta(days=1)
        presets = [
            ("Today", today, today),
            ("This week", today - timedelta(days=today.weekday()), today),
            ("This month", month_start, today),
            ("Last month", last_month_end.replace(day=1), last_month_end),
            ("All time", None, None),
        ]
        for text, start, end in presets:
            HoverButton(bar, text=text, font=(FONT, 9, "bold"), padx=10, pady=4,
                        command=lambda s=start, e=end: self.set_range(s, e)).pack(side="left", padx=2)
        self.from_entry.bind("<Return>", lambda e: self.run())
        self.to_entry.bind("<Return>", lambda e: self.run())

        cards = tk.Frame(self.body, bg=PANEL)
        cards.pack(fill="x", pady=(16, 0))
        self.card_values = {}
        for i, (key, title, accent) in enumerate([
            ("wash", "Washing tokens", TEAL), ("dry", "Dryer tokens", ORANGE),
            ("total", "Takings", GREEN), ("petty", "Petty cash", PURPLE), ("sales", "Sales", MUTED),
            ("voids", "Voids", RED), ("nosales", "No-sale opens", YELLOW),
        ]):
            cards.columnconfigure(i, weight=1, uniform="card")
            card = tk.Frame(cards, bg=PANEL_2)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 6, 0))
            tk.Frame(card, bg=accent, height=3).pack(fill="x")
            tk.Label(card, text=title.upper(), font=(FONT, 8, "bold"), fg=MUTED, bg=PANEL_2,
                     anchor="w").pack(fill="x", padx=12, pady=(8, 0))
            val = tk.Label(card, text="0", font=(FONT, 17, "bold"), fg=TEXT, bg=PANEL_2, anchor="w")
            val.pack(fill="x", padx=12, pady=(0, 8))
            self.card_values[key] = val

        nb = ttk.Notebook(self.body)
        nb.pack(fill="both", expand=True, pady=(16, 0))
        self.day_tree = self._make_tree(nb, [("date", "Date", 140, "w"), ("wash", "Washing", 100, "e"),
                                             ("dry", "Dryer", 100, "e"), ("sales", "Sales", 100, "e"),
                                             ("total", "Takings", 120, "e"), ("petty", "Petty cash", 110, "e"),
                                             ("net", "Net cash", 110, "e")])
        self.flat_tree = self._make_tree(nb, [("flat", "Flat", 80, "w"), ("name", "Tenant", 240, "w"),
                                              ("wash", "Washing", 100, "e"), ("dry", "Dryer", 100, "e"),
                                              ("total", "Spent", 120, "e")])
        self.petty_tree = self._make_tree(nb, [("receipt", "#", 60, "w"), ("datetime", "Date & time", 140, "w"),
                                               ("amount", "Amount", 100, "e"), ("notes", "What for", 320, "w"),
                                               ("staff", "Staff", 120, "w"), ("status", "Status", 100, "w")])
        self.petty_tree.tag_configure("voided", foreground=DIM, font=(FONT, 10, "overstrike"))
        self.bank_tree = self._make_tree(nb, [("ref", "Ref", 70, "w"), ("date", "Paid in", 140, "w"),
                                              ("expected", "Should have been", 130, "e"),
                                              ("paid", "Paid in", 110, "e"), ("diff", "Difference", 110, "e"),
                                              ("slip", "Paying-in ref", 130, "w"), ("staff", "Banked by", 120, "w")])
        self.bank_tree.configure(selectmode="browse")
        self.bank_tree.tag_configure("off", foreground=AMBER)
        nb.add(self.day_tree.master, text="By day")
        nb.add(self.flat_tree.master, text="By flat  (double-click for history)")
        self.flat_tree.bind("<Double-1>", self._open_flat_history)
        nb.add(self.petty_tree.master, text="Petty cash")
        nb.add(self.bank_tree.master, text="Banking")

        foot = tk.Frame(self.body, bg=PANEL)
        foot.pack(fill="x", pady=(14, 0))
        HoverButton(foot, text="Close", command=self.cancel).pack(side="right")
        HoverButton(foot, text="Export by flat…", font=(FONT, 10, "bold"),
                    command=self.export_by_flat).pack(side="left")
        HoverButton(foot, text="Export by day…", font=(FONT, 10, "bold"),
                    command=self.export_by_day).pack(side="left", padx=6)
        HoverButton(foot, text="Export transactions…", font=(FONT, 10, "bold"),
                    command=self.export_transactions).pack(side="left")
        HoverButton(foot, text="Export petty cash…", font=(FONT, 10, "bold"),
                    command=self.export_petty).pack(side="left", padx=6)
        HoverButton(foot, text="Export banking…", font=(FONT, 10, "bold"),
                    command=self.export_banking).pack(side="left")
        HoverButton(foot, text="Banking sheet…", bg=TEAL, hover=TEAL_HOVER, font=(FONT, 10, "bold"),
                    command=self.banking_sheet).pack(side="left", padx=6)

        self.set_range(month_start, today)

    def _make_tree(self, parent, columns):
        frame = tk.Frame(parent, bg=PANEL)
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", style="Dark.Treeview")
        for key, text, width, anchor in columns:
            tree.heading(key, text=text, anchor=anchor)
            tree.column(key, width=self.app.S(width), anchor=anchor, stretch=True)
        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        tree.tag_configure("odd", background=ROW_B)
        tree.tag_configure("even", background=ROW_A)
        tree.tag_configure("total", font=(FONT, 10, "bold"))
        return tree

    def set_range(self, start, end):
        if start is None:
            recs = self.app.records
            start = recs[0]["dt"].date() if recs else date.today()
            end = date.today()
        for entry, d in ((self.from_entry, start), (self.to_entry, end)):
            entry.delete(0, "end")
            entry.insert(0, d.strftime(UI_DATE_FMT))
        self.run()

    def current_range(self):
        start, end = parse_ui_date(self.from_entry.get()), parse_ui_date(self.to_entry.get())
        if start is None or end is None:
            messagebox.showwarning("Reports", "Please enter dates as DD/MM/YYYY.", parent=self)
            return None
        if start > end:
            start, end = end, start
        return start, end

    def run(self):
        rng = self.current_range()
        if rng is None:
            return
        self.range = rng
        start, end = rng
        self.rows = [r for r in self.app.records if start <= r["dt"].date() <= end]
        self.summary = s = summarise(self.rows, self.app.petty_receipts())
        self.card_values["wash"].config(text=str(s["wash"]))
        self.card_values["dry"].config(text=str(s["dry"]))
        self.card_values["total"].config(text=fmt_money(s["total"]))
        self.card_values["petty"].config(text=fmt_money(s["petty"]))
        self.card_values["sales"].config(text=str(s["sales"]))
        self.card_values["voids"].config(text=str(s["voids"]))
        self.card_values["nosales"].config(text=str(s["nosales"]))

        self.day_tree.delete(*self.day_tree.get_children())
        for i, (d, (w, dr, t, n, pc)) in enumerate(sorted(s["by_day"].items(), reverse=True)):
            self.day_tree.insert("", "end", values=(d.strftime("%a %d/%m/%Y"), w, dr, n, fmt_money(t),
                                                    fmt_money(pc), fmt_money(t - pc)),
                                 tags=("odd" if i % 2 else "even",))

        voided = self.app.voided_receipts()
        self.petty_tree.delete(*self.petty_tree.get_children())
        petty = [r for r in self.rows if r["type"] == TYPE_PETTY]
        for i, r in enumerate(reversed(petty)):
            is_voided = r["receipt"] in voided
            self.petty_tree.insert("", "end", values=(r["receipt"], r["dt"].strftime("%d/%m/%Y %H:%M"),
                                                      fmt_money(-r["total"]), r["notes"], r["staff"],
                                                      "Voided" if is_voided else ""),
                                   tags=("odd" if i % 2 else "even",) + (("voided",) if is_voided else ()))

        self.bankings = [b for b in bankings(self.app.records) if start <= b["rec"]["dt"].date() <= end]
        self.bank_tree.delete(*self.bank_tree.get_children())
        for i, b in enumerate(reversed(self.bankings)):
            diff = b["difference"]
            self.bank_tree.insert("", "end", iid=str(b["rec"]["receipt"]), values=(
                f"T{b['rec']['receipt']}", b["rec"]["dt"].strftime("%d/%m/%Y %H:%M"), fmt_money(b["expected"]),
                fmt_money(b["paid"]), (f"{'over' if diff > 0 else 'short'} {fmt_money(abs(diff))}" if diff else "—"),
                b["slip"], b["rec"]["staff"]), tags=("odd" if i % 2 else "even",) + (("off",) if diff else ()))

        current_names = {t["flat"]: t["name"] for t in self.app.tenants}
        self.flat_tree.delete(*self.flat_tree.get_children())
        flats = sorted(s["by_flat"].items(), key=lambda kv: natural_key(kv[0]))
        for i, (flat, v) in enumerate(flats):
            name = current_names.get(flat) or v["name"]
            self.flat_tree.insert("", "end", iid=flat, values=(flat, name, v["wash"], v["dry"], fmt_money(v["total"])),
                                  tags=("odd" if i % 2 else "even",))

    def _open_flat_history(self, event):
        flat = self.flat_tree.identify_row(event.y)
        if flat:
            TenantHistoryWindow(self, self.app, flat).show()

    def _ask_path(self, kind):
        start, end = self.range
        name = f"laundry_tokens_{kind}_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}.csv"
        return admin_save_path(self.app, self, title="Export to CSV", initialfile=name,
                               defaultextension=".csv", filetypes=[("CSV (Excel)", "*.csv")])

    def _export(self, kind, headers, rows):
        path = self._ask_path(kind)
        if not path:
            return
        try:
            write_csv_export(path, headers, rows)
        except OSError as e:
            messagebox.showerror("Export failed", f"Could not write the file:\n\n{e}", parent=self)
            return
        messagebox.showinfo("Exported", f"Saved {len(rows)} rows to:\n{path}", parent=self)

    def export_transactions(self):
        self._export("transactions", SALES_HEADERS, [record_to_row(r) for r in self.rows])

    def export_by_day(self):
        s = self.summary
        rows = [[d.strftime("%Y-%m-%d"), w, dr, n, csv_money(t), csv_money(pc), csv_money(t - pc)]
                for d, (w, dr, t, n, pc) in sorted(s["by_day"].items())]
        rows.append(["TOTAL", s["wash"], s["dry"], s["sales"], csv_money(s["total"]),
                     csv_money(s["petty"]), csv_money(s["total"] - s["petty"])])
        self._export("by_day", ["Date", "Washing Qty", "Dryer Qty", "Sales", "Takings (£)",
                                "Petty Cash (£)", "Net Cash (£)"], rows)

    def export_petty(self):
        voided = self.app.voided_receipts()
        rows = [[r["receipt"], r["dt"].strftime(CSV_DATE_FMT), csv_money(-r["total"]), r["notes"], r["staff"],
                 "Voided" if r["receipt"] in voided else ""]
                for r in self.rows if r["type"] == TYPE_PETTY]
        rows.append(["TOTAL", "", csv_money(self.summary["petty"]), "after voids", "", ""])
        self._export("petty_cash", ["Receipt No", "Date & Time", "Amount (£)", "What For", "Staff", "Status"], rows)

    def export_banking(self):
        rows = [[f"T{b['rec']['receipt']}", b["rec"]["dt"].strftime(CSV_DATE_FMT), csv_money(b["expected"]),
                 csv_money(b["paid"]), csv_money(b["difference"]), b["slip"], b["rec"]["staff"]]
                for b in self.bankings]
        rows.append(["TOTAL", "", csv_money(sum(b["expected"] for b in self.bankings)),
                     csv_money(sum(b["paid"] for b in self.bankings)),
                     csv_money(sum(b["difference"] for b in self.bankings)), "", ""])
        self._export("banking", ["Banking Ref", "Date & Time", "Should Have Been (£)", "Paid In (£)",
                                 "Difference (£)", "Paying-in Ref", "Banked By"], rows)

    def banking_sheet(self):
        sel = self.bank_tree.selection()
        if not sel:
            messagebox.showinfo("Banking sheet", "Select a banking on the Banking tab first.", parent=self)
            return
        self.app.print_token_banking(int(sel[0]), keep=False, parent=self)

    def export_by_flat(self):
        current_names = {t["flat"]: t["name"] for t in self.app.tenants}
        rows = [[flat, current_names.get(flat) or v["name"], v["wash"], v["dry"], csv_money(v["total"])]
                for flat, v in sorted(self.summary["by_flat"].items(), key=lambda kv: natural_key(kv[0]))]
        rows.append(["TOTAL", "", self.summary["wash"], self.summary["dry"], csv_money(self.summary["total"])])
        self._export("by_flat", ["Flat", "Tenant Name", "Washing Qty", "Dryer Qty", "Spent (£)"], rows)


# --------------------------------------------------------------------------
# Tenant history window
# --------------------------------------------------------------------------

class TenantHistoryWindow(Dialog):
    """Everything bought for one flat. The flat can have had several tenants, so each
    row shows the name recorded at the time."""

    def __init__(self, parent, app, flat):
        super().__init__(parent, f"Flat {flat}", resizable=True)
        self.app = app
        self.flat = flat
        self.geometry(f"{app.S(960)}x{app.S(620)}")
        self.minsize(app.S(820), app.S(480))
        name = app.tenant_name(flat) or "(no name on file)"
        self.heading(f"Flat {flat}  ·  {name}", "All purchases for this flat, newest first. Voided sales are "
                                                 "struck through and not counted.")
        self.rows = [r for r in app.records if r["flat"] == flat and r["type"] in (TYPE_SALE, TYPE_VOID)]
        voided = app.voided_receipts()
        s = summarise(self.rows)
        month = date.today().replace(day=1)
        this_month = summarise([r for r in self.rows if r["dt"].date() >= month])
        sales = [r for r in self.rows if r["type"] == TYPE_SALE and r["receipt"] not in voided]
        last = sales[-1]["dt"].strftime("%d/%m/%Y") if sales else "never"

        cards = tk.Frame(self.body, bg=PANEL)
        cards.pack(fill="x", pady=(16, 0))
        for i, (title, value, accent) in enumerate([
            ("Spent (all time)", fmt_money(s["total"]), GREEN), ("This month", fmt_money(this_month["total"]), GREEN),
            ("Washing tokens", str(s["wash"]), TEAL), ("Dryer tokens", str(s["dry"]), ORANGE),
            ("Purchases", str(len(sales)), MUTED), ("Last purchase", last, MUTED),
        ]):
            cards.columnconfigure(i, weight=1, uniform="card")
            card = tk.Frame(cards, bg=PANEL_2)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 6, 0))
            tk.Frame(card, bg=accent, height=3).pack(fill="x")
            tk.Label(card, text=title.upper(), font=(FONT, 8, "bold"), fg=MUTED, bg=PANEL_2,
                     anchor="w").pack(fill="x", padx=12, pady=(8, 0))
            tk.Label(card, text=value, font=(FONT, 14, "bold"), fg=TEXT, bg=PANEL_2,
                     anchor="w").pack(fill="x", padx=12, pady=(0, 8))

        frame = tk.Frame(self.body, bg=PANEL)
        frame.pack(fill="both", expand=True, pady=(16, 0))
        columns = [("receipt", "#", 60, "w"), ("datetime", "Date & time", 140, "w"), ("type", "Type", 120, "w"),
                   ("name", "Tenant at the time", 190, "w"), ("wash", "Washing", 80, "e"), ("dry", "Dryer", 70, "e"),
                   ("total", "Total", 90, "e"), ("staff", "Staff", 110, "w"), ("notes", "Notes", 200, "w")]
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", style="Dark.Treeview")
        for key, text, width, anchor in columns:
            tree.heading(key, text=text, anchor=anchor)
            tree.column(key, width=app.S(width), anchor=anchor, stretch=key in ("name", "notes"))
        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        tree.tag_configure("even", background=ROW_A)
        tree.tag_configure("odd", background=ROW_B)
        tree.tag_configure("void", foreground=RED)
        tree.tag_configure("voided", foreground=DIM, font=(FONT, 10, "overstrike"))
        for i, r in enumerate(reversed(self.rows)):
            is_voided = r["type"] == TYPE_SALE and r["receipt"] in voided
            tags = ("odd" if i % 2 else "even",) + (("void",) if r["type"] == TYPE_VOID else
                                                    ("voided",) if is_voided else ())
            tree.insert("", "end", values=(r["receipt"], r["dt"].strftime("%d/%m/%Y %H:%M"),
                                           "SALE (voided)" if is_voided else r["type"], r["name"], r["wash"],
                                           r["dry"], fmt_money(r["total"]), r["staff"], r["notes"]), tags=tags)
        if not self.rows:
            tk.Label(frame, text="No purchases for this flat yet.", font=(FONT, 11), fg=MUTED,
                     bg=ROW_A).place(relx=0.5, rely=0.5, anchor="center")

        foot = tk.Frame(self.body, bg=PANEL)
        foot.pack(fill="x", pady=(14, 0))
        HoverButton(foot, text="Close", command=self.cancel).pack(side="right")
        HoverButton(foot, text="Export to CSV…", font=(FONT, 10, "bold"), command=self.export).pack(side="left")

    def export(self):
        path = admin_save_path(self.app, self, title="Export to CSV",
                               initialfile=f"laundry_flat_{self.flat}_history.csv",
                               defaultextension=".csv", filetypes=[("CSV (Excel)", "*.csv")])
        if not path:
            return
        try:
            write_csv_export(path, SALES_HEADERS, [record_to_row(r) for r in self.rows])
        except OSError as e:
            messagebox.showerror("Export failed", f"Could not write the file:\n\n{e}", parent=self)
            return
        messagebox.showinfo("Exported", f"Saved {len(self.rows)} rows to:\n{path}", parent=self)


# --------------------------------------------------------------------------
# Token stock window
# --------------------------------------------------------------------------

def parse_count(text):
    """A whole number of tokens, 0 or more, or None."""
    text = str(text).strip()
    return int(text) if text.isdigit() else None


class StockWindow(Dialog):
    """Tokens in the office. Deliveries and hand counts go in the sales log (with a PIN like
    everything else); the level is worked out from them and the sales."""

    def __init__(self, app):
        super().__init__(app, "Token stock", resizable=True)
        self.app = app
        self.geometry(f"{app.S(1040)}x{app.S(600)}")
        self.minsize(app.S(760), app.S(480))
        self.heading("Token stock", "Tokens in the office: tokens received, minus tokens sold. Count them now and "
                                    "then; the count corrects the figure and any difference is recorded.")
        cards = tk.Frame(self.body, bg=PANEL)
        cards.pack(fill="x", pady=(16, 0))
        self.cards = {}
        for i, (key, title, accent) in enumerate([("wash", "Washing tokens", TEAL), ("dry", "Dryer tokens", ORANGE),
                                                   ("count", "Last counted", MUTED)]):
            cards.columnconfigure(i, weight=1, uniform="card")
            card = tk.Frame(cards, bg=PANEL_2)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 6, 0))
            bar = tk.Frame(card, bg=accent, height=3)
            bar.pack(fill="x")
            tk.Label(card, text=title.upper(), font=(FONT, 8, "bold"), fg=MUTED, bg=PANEL_2,
                     anchor="w").pack(fill="x", padx=12, pady=(8, 0))
            value = tk.Label(card, text="", font=(FONT, 17, "bold"), fg=TEXT, bg=PANEL_2, anchor="w")
            value.pack(fill="x", padx=12)
            sub = tk.Label(card, text="", font=(FONT, 9), fg=MUTED, bg=PANEL_2, anchor="w")
            sub.pack(fill="x", padx=12, pady=(0, 8))
            self.cards[key] = (value, sub, bar, accent)

        foot = tk.Frame(self.body, bg=PANEL)
        foot.pack(side="bottom", fill="x", pady=(14, 0))  # before the table, so the table gives way
        HoverButton(foot, text="Close", command=self.cancel).pack(side="right")
        HoverButton(foot, text="Tokens received…", bg=TEAL, hover=TEAL_HOVER, font=(FONT, 10, "bold"),
                    command=self.received).pack(side="left")
        HoverButton(foot, text="Stock count…", bg=ORANGE, hover=ORANGE_HOVER, font=(FONT, 10, "bold"),
                    command=self.count).pack(side="left", padx=8)

        frame = tk.Frame(self.body, bg=PANEL)
        frame.pack(fill="both", expand=True, pady=(16, 0))
        columns = [("receipt", "#", 60, "w"), ("datetime", "Date & time", 140, "w"), ("type", "Type", 120, "w"),
                   ("wash", "Washing", 80, "e"), ("dry", "Dryer", 70, "e"), ("staff", "Staff", 110, "w"),
                   ("notes", "Notes", 280, "w")]
        self.tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", style="Dark.Treeview")
        for key, text, width, anchor in columns:
            self.tree.heading(key, text=text, anchor=anchor)
            self.tree.column(key, width=app.S(width), anchor=anchor, stretch=key == "notes")
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("even", background=ROW_A)
        self.tree.tag_configure("odd", background=ROW_B)
        self.empty_label = tk.Label(frame, text="Stock isn't being tracked yet. Start with a stock count of the "
                                                "tokens in the office.", font=(FONT, 11), fg=MUTED, bg=ROW_A)
        self.fill()

    def fill(self):
        stock = token_stock(self.app.records)
        low = self.app.low_stock_levels()
        for key in ("wash", "dry"):
            value, sub, bar, accent = self.cards[key]
            if stock is None:
                value.config(text="—")
                sub.config(text="Not tracked yet", fg=MUTED)
                continue
            is_low = stock[key] < low[key]
            value.config(text=str(stock[key]))
            sub.config(text=f"Low below {low[key]}" if not is_low else "Low: order more", fg=AMBER if is_low else MUTED)
            bar.config(bg=RED if is_low else accent)
        counts = [r for r in self.app.records if r["type"] == TYPE_STOCK_COUNT]
        value, sub, _, _ = self.cards["count"]
        value.config(text=counts[-1]["dt"].strftime("%d/%m/%Y") if counts else "Never")
        sub.config(text=f"by {counts[-1]['staff']}" if counts else "")

        self.tree.delete(*self.tree.get_children())
        rows = [r for r in self.app.records if r["type"] in STOCK_TYPES]
        for i, r in enumerate(reversed(rows)):
            signed = r["type"] == TYPE_STOCK_COUNT
            self.tree.insert("", "end", values=(
                r["receipt"], r["dt"].strftime("%d/%m/%Y %H:%M"), r["type"].title(),
                f"{r['wash']:+d}" if signed else r["wash"], f"{r['dry']:+d}" if signed else r["dry"],
                r["staff"], r["notes"]), tags=("odd" if i % 2 else "even",))
        if rows:
            self.empty_label.place_forget()
        else:
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")

    def _ask_quantities(self, title, heading, sub, note_label, allow_zero):
        def validate(v):
            wash, dry = parse_count(v["wash"]), parse_count(v["dry"])
            if wash is None or dry is None:
                return "Enter whole numbers of tokens, e.g. 200 (0 if none)."
            if not allow_zero and wash + dry == 0:
                return "Enter how many tokens were received."
            return None

        res = FieldsDialog(self, title, heading, [
            {"key": "wash", "label": "Washing tokens"},
            {"key": "dry", "label": "Dryer tokens"},
            {"key": "note", "label": note_label},
        ], validate=validate, confirm_text="Next", sub=sub).show()
        if not res:
            return None
        return parse_count(res["wash"]), parse_count(res["dry"]), res["note"]

    def _confirm(self, title, heading, lines, confirm_text, bg, hover):
        dlg = PinDialog(self, self.app, title, heading, lines, confirm_text=confirm_text, confirm_bg=bg,
                        confirm_hover=hover)
        result = dlg.show()
        if not result:
            self.app._pin_locked_out(dlg)
            return None
        return result[0]

    def received(self):
        res = self._ask_quantities("Tokens received", "Tokens received", "Tokens added to the office's stock.",
                                   "Note (optional), e.g. delivery from the supplier", allow_zero=False)
        if not res:
            return
        wash, dry, note = res
        staff = self._confirm("Tokens received", f"Add {describe_items(wash, dry)}", [
            "Added to the token stock.", note or "No note"], "Add to stock", TEAL, TEAL_HOVER)
        if staff and self.app.record_event(TYPE_STOCK_IN, "", "", wash, dry, 0, staff["name"], notes=note,
                                           parent=self):
            self.fill()
            self.app.set_message(f"✓ Stock: {describe_items(wash, dry)} received by {staff['name']}", GREEN)

    def count(self):
        stock = token_stock(self.app.records)
        sub = ("Count every washing and dryer token in the office." if stock is None else
               f"Count every token in the office. The app expects {stock['wash']} washing and "
               f"{stock['dry']} dryer tokens.")
        res = self._ask_quantities("Stock count", "Stock count", sub, "Note (optional)", allow_zero=True)
        if not res:
            return
        wash, dry, note = res
        if stock is None:
            lines = [f"Counted {wash} washing, {dry} dryer", "Opening count: stock is tracked from now on."]
            notes = f"Opening count: {wash} washing, {dry} dryer"
            diff_w, diff_d = wash, dry
        else:
            diff_w, diff_d = wash - stock["wash"], dry - stock["dry"]
            lines = [f"Counted {wash} washing, {dry} dryer",
                     f"Expected {stock['wash']} washing, {stock['dry']} dryer",
                     "Matches the expected stock." if not (diff_w or diff_d) else
                     f"Difference: washing {diff_w:+d}, dryer {diff_d:+d}. This is recorded."]
            notes = f"Counted {wash} washing, {dry} dryer; expected {stock['wash']}, {stock['dry']}"
        if note:
            notes += f". {note}"
        staff = self._confirm("Stock count", "Record stock count", lines, "Record count", ORANGE, ORANGE_HOVER)
        if staff and self.app.record_event(TYPE_STOCK_COUNT, "", "", diff_w, diff_d, 0, staff["name"], notes=notes,
                                           parent=self):
            self.fill()
            off = stock is not None and (diff_w or diff_d)
            self.app.set_message(f"Stock count recorded by {staff['name']}" +
                                 (f": washing {diff_w:+d}, dryer {diff_d:+d} against expected" if off else ""),
                                 YELLOW if off else GREEN)


# --------------------------------------------------------------------------
# Settings window
# --------------------------------------------------------------------------

class SettingsWindow(Dialog):
    def __init__(self, app, admin):
        super().__init__(app, "Settings", resizable=True)
        self.app = app
        self.admin = admin
        self.geometry(f"{app.S(1240)}x{app.S(660)}")
        self.minsize(app.S(1180), app.S(560))
        self.heading("Settings", f"Signed in as {admin['name']} (admin)")
        nb = ttk.Notebook(self.body)
        nb.pack(fill="both", expand=True, pady=(14, 0))
        rent, keys = app.rent_module, app.keys_module
        for text, builder in (("Staff & PINs", self._build_staff), ("General", self._build_general),
                              ("Rent tenants", lambda p: rent.build_tenants_tab(self, p)),
                              ("Receipts", lambda p: rent.build_receipts_tab(self, p)),
                              ("Properties", lambda p: keys.build_properties_tab(self, p)),
                              ("Contractors", lambda p: keys.build_contractors_tab(self, p)),
                              ("Laundry flats", self._build_tenants), ("Token prices", self._build_prices),
                              ("Cash drawer", self._build_drawer), ("Logs", self._build_log)):
            frame = tk.Frame(nb, bg=PANEL)
            inner = tk.Frame(frame, bg=PANEL)
            inner.pack(fill="both", expand=True, padx=18, pady=16)
            builder(inner)
            nb.add(frame, text=text)
        foot = tk.Frame(self.body, bg=PANEL)
        foot.pack(fill="x", pady=(14, 0))
        HoverButton(foot, text="Close", command=self.cancel).pack(side="right")

    # ---- small builders --------------------------------------------------
    @staticmethod
    def _label(parent, text, top=12):
        tk.Label(parent, text=text, font=(FONT, 10, "bold"), fg=MUTED, bg=PANEL,
                 anchor="w").pack(fill="x", pady=(top, 4))

    @staticmethod
    def _note(parent, text):
        tk.Label(parent, text=text, font=(FONT, 10), fg=MUTED, bg=PANEL, anchor="w", justify="left",
                 wraplength=860).pack(fill="x", pady=(10, 0))

    def _tree(self, parent, columns):
        frame = tk.Frame(parent, bg=PANEL)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings",
                            style="Dark.Treeview", selectmode="browse")
        for key, text, width in columns:
            tree.heading(key, text=text, anchor="w")
            tree.column(key, width=self.app.S(width), anchor="w")
        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        tree.tag_configure("odd", background=ROW_B)
        tree.tag_configure("even", background=ROW_A)
        return tree

    # ---- cash drawer -----------------------------------------------------
    def _build_drawer(self, p):
        cfg = self.app.cfg
        self._label(p, "COM port", top=0)
        row = tk.Frame(p, bg=PANEL)
        row.pack(fill="x")
        self.port_var = tk.StringVar(value=cfg["com_port"])
        self.port_box = ttk.Combobox(row, textvariable=self.port_var, width=48, font=(FONT, 11))
        self.port_box.pack(side="left", ipady=3)
        HoverButton(row, text="Refresh", font=(FONT, 10, "bold"),
                    command=self._refresh_ports).pack(side="left", padx=8)
        self._refresh_ports()

        self._label(p, "Baud rate")
        self.baud_var = tk.StringVar(value=str(cfg["baud_rate"]))
        ttk.Combobox(p, textvariable=self.baud_var, values=BAUD_RATES, width=12, state="readonly",
                     font=(FONT, 11)).pack(anchor="w", ipady=3)

        self.enabled_var = tk.BooleanVar(value=cfg["drawer_enabled"])
        tk.Checkbutton(p, text="Open the cash drawer on each sale", variable=self.enabled_var,
                       font=(FONT, 11), fg=TEXT, bg=PANEL, selectcolor=PANEL_2, activebackground=PANEL,
                       activeforeground=TEXT, anchor="w", bd=0, highlightthickness=0).pack(fill="x", pady=(14, 0))

        btns = tk.Frame(p, bg=PANEL)
        btns.pack(fill="x", pady=(18, 0))
        HoverButton(btns, text="Save", bg=TEAL, hover=TEAL_HOVER, command=self._save_drawer).pack(side="left")
        HoverButton(btns, text="Test drawer", bg=ORANGE, hover=ORANGE_HOVER,
                    command=self._test_drawer).pack(side="left", padx=8)
        self._note(p, "Plug in the USB trigger box, then press Refresh. In Windows Device Manager it appears "
                      "under 'Ports (COM & LPT)', usually as 'USB-SERIAL CH340 (COMx)'. Choose it and press "
                      "Test drawer. Tests are recorded in the log as a no-sale.\n\n"
                      "If the test reports success but the drawer doesn't move, check that the RJ12 cable "
                      "goes into the drawer's port and that the trigger box outputs 24V, which is what "
                      "the Tera drawer needs.")

    def _refresh_ports(self):
        ports = available_ports()
        self.port_box["values"] = [f"{p.device} — {p.description}" for p in ports]

    def _selected_port(self):
        return self.port_var.get().split("—")[0].strip().upper()

    def _save_drawer(self):
        port = self._selected_port()
        if not re.fullmatch(r"COM\d+", port):
            messagebox.showwarning("Cash drawer", "Please choose a COM port, e.g. COM3.", parent=self)
            return
        cfg = self.app.cfg
        cfg["com_port"] = port
        cfg["baud_rate"] = int(self.baud_var.get())
        cfg["drawer_enabled"] = bool(self.enabled_var.get())
        self.app.save_cfg()
        self.port_var.set(port)
        messagebox.showinfo("Cash drawer", "Drawer settings saved.", parent=self)

    def _test_drawer(self):
        port = self._selected_port()
        rec = self.app.record_event(TYPE_NOSALE, "", "", 0, 0, 0, self.admin["name"],
                                    notes=f"Drawer test from Settings ({port})", parent=self)
        if rec is None:
            return
        self.config(cursor="watch")
        self.update_idletasks()
        ok, msg = fire_drawer(port, self.baud_var.get(), enabled=True)
        self.config(cursor="")
        if ok:
            messagebox.showinfo("Test drawer", f"Command sent. {msg}.\n\nDid the drawer open? If yes, "
                                "press Save to keep this port.", parent=self)
        else:
            messagebox.showwarning("Test drawer", f"Could not reach the trigger box.\n\n{msg}", parent=self)

    # ---- general ---------------------------------------------------------
    def _build_general(self, p):
        cfg = self.app.cfg
        low = self.app.low_stock_levels()
        self.general_entries = {}
        for i, (key, label, value) in enumerate([
                ("idle", "Show the idle screen after this many minutes with nobody using the app (0 = never)",
                 self.app.idle_minutes()),
                ("low_wash", "Warn when washing tokens in stock fall below", low["wash"]),
                ("low_dry", "Warn when dryer tokens in stock fall below", low["dry"]),
                ("bank_rent", "Highlight when rent waiting to be banked (awaiting check + in safe) goes over £ "
                              "(0 = never)", self.app.bank_limit("rent")),
                ("bank_drawer", "Highlight when the token cash drawer goes over £ (0 = never)",
                 self.app.bank_limit("drawer"))]):
            self._label(p, label, top=0 if i == 0 else 12)
            entry = tk.Entry(p, width=8, **entry_opts())
            entry.insert(0, str(value))
            entry.pack(anchor="w", ipady=5)
            self.general_entries[key] = entry
        self.fullscreen_var = tk.BooleanVar(value=bool(cfg.get("fullscreen", True)))
        tk.Checkbutton(p, text="Full screen (hides the Windows taskbar; F11 switches it on or off for now)",
                       variable=self.fullscreen_var, font=(FONT, 11), fg=TEXT, bg=PANEL, selectcolor=PANEL_2,
                       activebackground=PANEL, activeforeground=TEXT, anchor="w", bd=0,
                       highlightthickness=0).pack(fill="x", pady=(14, 0))
        row = tk.Frame(p, bg=PANEL)
        row.pack(fill="x", pady=(12, 0))
        tk.Label(row, text="Colours:", font=(FONT, 11), fg=TEXT, bg=PANEL).pack(side="left", padx=(0, 10))
        self.theme_var = tk.StringVar(value=cfg.get("theme", "dark"))
        for value, text in (("dark", "Dark"), ("light", "Light (easier to read in a bright office)")):
            tk.Radiobutton(row, text=text, value=value, variable=self.theme_var, font=(FONT, 11), fg=TEXT,
                           bg=PANEL, selectcolor=PANEL_2, activebackground=PANEL, activeforeground=TEXT,
                           bd=0, highlightthickness=0).pack(side="left", padx=(0, 14))
        HoverButton(p, text="Save", bg=TEAL, hover=TEAL_HOVER,
                    command=self._save_general).pack(anchor="w", pady=(18, 0))
        self._note(p, "The idle screen covers the app with the time and what's waiting (keys out, post to hand "
                      "over), and never shows money, so takings aren't on display when the PC is left on. A tap or "
                      "key press goes back to Home. It also clears anything half-entered (a token sale or a rent "
                      "payment that wasn't taken), so the next person starts fresh. Nothing that was recorded is "
                      "affected.")

    def _save_general(self):
        values = {k: parse_count(e.get()) for k, e in self.general_entries.items()}
        if None in values.values():
            messagebox.showwarning("General", "Please enter whole numbers, e.g. 3 or 50.", parent=self)
            return
        cfg = self.app.cfg
        cfg["idle_minutes"] = values["idle"]
        cfg["low_stock_washing"] = values["low_wash"]
        cfg["low_stock_dryer"] = values["low_dry"]
        cfg["bank_limit_rent"] = values["bank_rent"]
        cfg["bank_limit_drawer"] = values["bank_drawer"]
        cfg["fullscreen"] = self.fullscreen_var.get()
        new_theme = self.theme_var.get() != cfg.get("theme", "dark")
        cfg["theme"] = self.theme_var.get()
        self.app.save_cfg()
        self.app.apply_fullscreen()
        self.lift()  # keep Settings in front of the main window after it changes size
        self.app.refresh_stats()
        if new_theme:
            if messagebox.askyesno("General", "Saved. The new colours need the app to restart.\n\n"
                                   "Restart now? (Nothing is lost: everything is already saved.)", parent=self):
                self.app.restart_app()
            return
        messagebox.showinfo("General", "Saved.", parent=self)

    # ---- prices ----------------------------------------------------------
    def _build_prices(self, p):
        cfg = self.app.cfg
        self._label(p, "Washing token price (£)", top=0)
        self.wash_price = tk.Entry(p, width=12, **entry_opts())
        self.wash_price.insert(0, csv_money(cfg["price_washing_pence"]))
        self.wash_price.pack(anchor="w", ipady=5)
        self._label(p, "Dryer token price (£)")
        self.dry_price = tk.Entry(p, width=12, **entry_opts())
        self.dry_price.insert(0, csv_money(cfg["price_dryer_pence"]))
        self.dry_price.pack(anchor="w", ipady=5)
        HoverButton(p, text="Save prices", bg=TEAL, hover=TEAL_HOVER,
                    command=self._save_prices).pack(anchor="w", pady=(18, 0))
        self._note(p, "New prices apply to new sales only. Past sales keep the price that was charged.")

    def _save_prices(self):
        wash, dry = parse_pence(self.wash_price.get()), parse_pence(self.dry_price.get())
        if not wash or not dry or wash < 0 or dry < 0:
            messagebox.showwarning("Prices", "Please enter prices greater than zero, e.g. 1.00", parent=self)
            return
        self.app.cfg["price_washing_pence"] = wash
        self.app.cfg["price_dryer_pence"] = dry
        self.app.save_cfg()
        self.app.apply_prices()
        messagebox.showinfo("Prices", f"Saved: washing {fmt_money(wash)}, dryer {fmt_money(dry)}.", parent=self)

    # ---- staff -----------------------------------------------------------
    def _build_staff(self, p):
        self.staff_tree = self._tree(p, [("name", "Name", 300), ("role", "Role", 260)])
        btns = tk.Frame(p, bg=PANEL)
        btns.pack(fill="x", pady=(12, 0))
        HoverButton(btns, text="Add staff", bg=TEAL, hover=TEAL_HOVER, font=(FONT, 10, "bold"),
                    command=self._add_staff).pack(side="left")
        for text, cmd in (("Change PIN", self._change_pin), ("Admin on/off", self._toggle_admin),
                          ("Checker on/off", self._toggle_checker), ("Remove", self._remove_staff)):
            HoverButton(btns, text=text, font=(FONT, 10, "bold"), command=cmd).pack(side="left", padx=(8, 0))
        self._note(p, "Each person needs their own PIN (4–8 digits). The PIN identifies who did each thing. "
                      "Admins can void token sales, cancel receipts and open Settings. Checkers (the office "
                      "manager) check rent payments and put them in the safe; nobody can check a payment they "
                      "took themselves.")
        self._fill_staff()

    def _fill_staff(self):
        self.staff_tree.delete(*self.staff_tree.get_children())
        for i, s in enumerate(self.app.cfg["staff"]):
            role = ", ".join(r for r, on in (("Admin", s["admin"]), ("Checker", s.get("checker"))) if on) or "Staff"
            self.staff_tree.insert("", "end", iid=str(i), values=(s["name"], role),
                                   tags=("odd" if i % 2 else "even",))

    def _selected_staff(self):
        sel = self.staff_tree.selection()
        if not sel:
            messagebox.showinfo("Staff", "Select a staff member first.", parent=self)
            return None
        return self.app.cfg["staff"][int(sel[0])]

    def _pin_problem(self, pin, pin2, owner=None):
        if not valid_pin_format(pin):
            return "The PIN must be 4 to 8 digits."
        if pin != pin2:
            return "The two PINs don't match."
        existing = self.app.find_staff(pin)
        if existing is not None and existing is not owner:
            return "Another staff member already uses this PIN. Please choose a different one."
        return None

    def _add_staff(self):
        names = {s["name"].lower() for s in self.app.cfg["staff"]}

        def validate(v):
            if not v["name"]:
                return "Please enter a name."
            if v["name"].lower() in names:
                return "A staff member with that name already exists."
            return self._pin_problem(v["pin"], v["pin2"])

        res = FieldsDialog(self, "Add staff", "Add staff member", [
            {"key": "name", "label": "Name"},
            {"key": "pin", "label": "PIN (4–8 digits)", "kind": "pin"},
            {"key": "pin2", "label": "Confirm PIN", "kind": "pin"},
            {"key": "admin", "label": "Admin (can void sales, cancel receipts and change settings)", "kind": "check"},
            {"key": "checker", "label": "Checker (checks rent payments and puts them in the safe)", "kind": "check"},
        ], validate=validate, confirm_text="Add").show()
        if res:
            self.app.cfg["staff"].append(make_staff(res["name"], res["pin"], res["admin"], res["checker"]))
            self.app.save_cfg()
            self._fill_staff()

    def _change_pin(self):
        staff = self._selected_staff()
        if staff is None:
            return
        res = FieldsDialog(self, "Change PIN", f"New PIN for {staff['name']}", [
            {"key": "pin", "label": "New PIN (4–8 digits)", "kind": "pin"},
            {"key": "pin2", "label": "Confirm new PIN", "kind": "pin"},
        ], validate=lambda v: self._pin_problem(v["pin"], v["pin2"], owner=staff)).show()
        if res:
            staff["salt"], staff["pin_hash"] = hash_pin(res["pin"])
            self.app.save_cfg()
            messagebox.showinfo("Staff", f"PIN changed for {staff['name']}.", parent=self)

    def _admin_count(self):
        return sum(1 for s in self.app.cfg["staff"] if s["admin"])

    def _toggle_admin(self):
        staff = self._selected_staff()
        if staff is None:
            return
        if staff["admin"] and self._admin_count() == 1:
            messagebox.showwarning("Staff", "There must always be at least one admin.", parent=self)
            return
        staff["admin"] = not staff["admin"]
        self.app.save_cfg()
        self._fill_staff()

    def _toggle_checker(self):
        staff = self._selected_staff()
        if staff is None:
            return
        staff["checker"] = not staff.get("checker")
        self.app.save_cfg()
        self._fill_staff()

    def _remove_staff(self):
        staff = self._selected_staff()
        if staff is None:
            return
        if staff["admin"] and self._admin_count() == 1:
            messagebox.showwarning("Staff", "You can't remove the only admin.", parent=self)
            return
        if not messagebox.askyesno("Remove staff", f"Remove {staff['name']}? Their past sales stay in the log.",
                                   parent=self):
            return
        self.app.cfg["staff"].remove(staff)
        self.app.save_cfg()
        self._fill_staff()

    # ---- logs ------------------------------------------------------------
    def _logs(self):
        """(title, problems, entry count, accept function) for each log."""
        app, rent, keys = self.app, self.app.rent, self.app.keys
        return [("Laundry token log (sales_log.csv)", app.log_problems, len(app.records), app.accept_log),
                ("Rent log (rent_log.csv)", rent.problems, len(rent.entries), rent.accept),
                ("Key log (key_log.csv)", keys.problems, len(keys.entries), keys.accept),
                ("Post log (post_log.csv)", app.post.problems, len(app.post.entries), app.post.accept),
                ("Visitor log (visitor_log.csv)", app.visitors.problems, len(app.visitors.entries),
                 app.visitors.accept)]

    def _build_log(self, p):
        self.log_rows = []
        for i, (title, _, _, accept) in enumerate(self._logs()):
            tk.Label(p, text=title, font=(FONT, 11, "bold"), fg=TEXT, bg=PANEL,
                     anchor="w").pack(fill="x", pady=(0 if i == 0 else 16, 4))
            row = tk.Frame(p, bg=PANEL)
            row.pack(fill="x")
            btn = HoverButton(row, text="Accept as it is now…", bg=RED, hover=RED_HOVER, font=(FONT, 9, "bold"),
                              padx=10, pady=4, command=lambda a=accept: self._accept_log(a))
            btn.pack(side="right", anchor="n")
            status = tk.Label(row, text="", font=(FONT, 10, "bold"), bg=PANEL, anchor="w", justify="left",
                              wraplength=700)
            status.pack(side="left", fill="x", expand=True)
            self.log_rows.append((status, btn))
        HoverButton(p, text="Open backups folder", font=(FONT, 10, "bold"),
                    command=self._open_backups).pack(anchor="w", pady=(18, 0))
        self._note(p, "Every entry in each log carries a check code, worked out from the entry and the one "
                      "before it. If a file is changed outside this app (a row edited, added or deleted in "
                      "Excel, say), the codes stop matching. The app then highlights those entries in red "
                      "and shows a warning.\n\n"
                      "To fix it, compare with the daily copies in the backups folder and put a good copy back "
                      "while the app is closed. Or, if the change was a genuine correction, accept the log as it "
                      "is now. That is recorded in the log with your name and reason, and a copy of the flagged "
                      "file is kept in backups.")
        self._fill_log_status()

    def _fill_log_status(self):
        for (status, btn), (_, problems, count, _) in zip(self.log_rows, self._logs()):
            if problems:
                status.config(text="⚠  " + "\n⚠  ".join(problems), fg=RED)
            else:
                status.config(text=f"✓  Intact: all {count} entries match their check codes.", fg=GREEN)
            btn.set_enabled(bool(problems))

    def _open_backups(self):
        try:
            os.makedirs(BACKUP_DIR, exist_ok=True)
            os.startfile(BACKUP_DIR)
        except OSError as e:
            messagebox.showerror("Backups", f"Could not open the backups folder:\n\n{e}", parent=self)

    def _accept_log(self, accept):
        res = FieldsDialog(self, "Accept log", "Accept the log as it is now?", [
            {"key": "reason", "label": "Reason (e.g. restored yesterday's backup after a spill)"},
        ], validate=lambda v: None if v["reason"] else "Please enter a reason.", confirm_text="Accept",
            sub="Only do this once you're sure the log is right. Your name and reason go in the log.").show()
        if res and accept(self.admin["name"], res["reason"], parent=self):
            self._fill_log_status()
            self.app.refresh_all()
            messagebox.showinfo("Logs", "Log accepted. The warning is cleared.", parent=self)

    # ---- tenants ---------------------------------------------------------
    def _build_tenants(self, p):
        self.tenant_tree = self._tree(p, [("flat", "Flat", 120), ("name", "Tenant name", 380)])
        self.tenant_tree.bind("<Double-1>", lambda e: self._edit_tenant())
        btns = tk.Frame(p, bg=PANEL)
        btns.pack(fill="x", pady=(12, 0))
        HoverButton(btns, text="Edit", bg=TEAL, hover=TEAL_HOVER, font=(FONT, 10, "bold"),
                    command=self._edit_tenant).pack(side="left")
        HoverButton(btns, text="Add flat", font=(FONT, 10, "bold"),
                    command=self._add_tenant).pack(side="left", padx=(8, 0))
        HoverButton(btns, text="Remove flat", font=(FONT, 10, "bold"),
                    command=self._remove_tenant).pack(side="left", padx=(8, 0))
        self._note(p, "Double-click a flat to change the tenant's name, e.g. when someone moves in. "
                      "Past sales keep the name that was recorded at the time.")
        self._fill_tenants()

    def _fill_tenants(self, select_flat=None):
        self.tenant_tree.delete(*self.tenant_tree.get_children())
        for i, t in enumerate(self.app.tenants):
            iid = self.tenant_tree.insert("", "end", values=(t["flat"], t["name"] or "—"),
                                          tags=("odd" if i % 2 else "even",))
            if t["flat"] == select_flat:
                self.tenant_tree.selection_set(iid)
                self.tenant_tree.see(iid)

    def _selected_tenant(self):
        sel = self.tenant_tree.selection()
        if not sel:
            messagebox.showinfo("Tenants", "Select a flat first.", parent=self)
            return None
        return self.app.tenants[self.tenant_tree.index(sel[0])]

    def _tenant_dialog(self, heading, tenant=None):
        others = {t["flat"].lower() for t in self.app.tenants if t is not tenant}

        def validate(v):
            if not v["flat"]:
                return "Please enter a flat number."
            if v["flat"].lower() in others:
                return "That flat is already in the list."
            return None

        return FieldsDialog(self, "Tenant", heading, [
            {"key": "flat", "label": "Flat", "value": tenant["flat"] if tenant else ""},
            {"key": "name", "label": "Tenant name", "value": tenant["name"] if tenant else ""},
        ], validate=validate).show()

    def _edit_tenant(self):
        tenant = self._selected_tenant()
        if tenant is None:
            return
        res = self._tenant_dialog(f"Flat {tenant['flat']}", tenant)
        if res:
            tenant["flat"], tenant["name"] = res["flat"], res["name"]
            self._save_tenants(select_flat=res["flat"])

    def _add_tenant(self):
        res = self._tenant_dialog("Add flat")
        if res:
            self.app.tenants.append({"flat": res["flat"], "name": res["name"]})
            self._save_tenants(select_flat=res["flat"])

    def _remove_tenant(self):
        tenant = self._selected_tenant()
        if tenant is None:
            return
        if not messagebox.askyesno("Remove flat", f"Remove flat {tenant['flat']} from the list?\n"
                                   "Past sales for this flat stay in the log.", parent=self):
            return
        self.app.tenants.remove(tenant)
        self._save_tenants()

    def _save_tenants(self, select_flat=None):
        self.app.tenants.sort(key=lambda t: natural_key(t["flat"]))
        try:
            save_tenants(self.app.tenants)
        except OSError as e:
            messagebox.showerror("Tenants", f"Could not save tenants.csv:\n\n{e}", parent=self)
        self._fill_tenants(select_flat)
        self.app.on_tenants_changed()


# --------------------------------------------------------------------------
# Cash drawer screen
# --------------------------------------------------------------------------

class DrawerView(tk.Frame):
    """The laundry-token cash drawer: what should be in it now, and everything that moved cash in
    or out since it was last emptied. Opening it (no sale / change), counting it, petty cash and
    banking are all done here. Rent money never goes in the drawer."""

    def __init__(self, master, app):
        super().__init__(master, bg=BG)
        self.app = app
        self.filter = "since"
        S = app.S
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        side = tk.Frame(self, bg=PANEL, width=S(420))
        side.grid(row=0, column=0, sticky="ns")
        side.pack_propagate(False)
        p = tk.Frame(side, bg=PANEL)
        p.pack(fill="both", expand=True, padx=S(22), pady=S(16))
        tk.Label(p, text="IN THE DRAWER NOW", font=(FONT, 10, "bold"), fg=MUTED, bg=PANEL,
                 anchor="w").pack(fill="x")
        self.total = tk.Label(p, text="", font=(FONT, 40, "bold"), fg=TEXT, bg=PANEL, anchor="w")
        self.total.pack(fill="x")
        self.since = tk.Label(p, text="", font=(FONT, 10), fg=MUTED, bg=PANEL, anchor="w", justify="left",
                              wraplength=S(370))
        self.since.pack(fill="x")
        tk.Frame(p, bg=BORDER, height=1).pack(fill="x", pady=S(14))
        self.lines = {}
        for key, label in (("takings", "Token takings"), ("petty", "Petty cash taken out"),
                           ("counts", "Count corrections"), ("last_count", "Last counted"),
                           ("last_banked", "Last emptied for banking")):
            row = tk.Frame(p, bg=PANEL)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label, font=(FONT, 11), fg=MUTED, bg=PANEL).pack(side="left")
            value = tk.Label(row, text="", font=(FONT, 11, "bold"), fg=TEXT, bg=PANEL)
            value.pack(side="right")
            self.lines[key] = value

        actions = tk.Frame(p, bg=PANEL)
        actions.pack(fill="x", side="bottom")
        for text, sub, bg, hover, command in (
                ("Open drawer  ·  no sale", "e.g. change for a note: the drawer total stays the same",
                 AMBER, AMBER_HOVER, app.no_sale),
                ("Count cash", "count it (e.g. end of day); any difference is recorded", None, None,
                 app.count_cash),
                ("Petty cash", "cash taken out for an office expense", None, None, app.petty_cash),
                ("Bank takings", "empty the drawer and take the cash to the bank", GREEN, GREEN_HOVER,
                 app.bank_takings)):
            HoverButton(actions, text=text, bg=bg, hover=hover, fg="#ffffff" if bg else None,
                        font=(FONT, 12, "bold"), pady=10, command=command).pack(fill="x", pady=(S(10), 0))
            tk.Label(actions, text=sub, font=(FONT, 9), fg=DIM, bg=PANEL, anchor="w").pack(fill="x")

        main = tk.Frame(self, bg=BG)
        main.grid(row=0, column=1, sticky="nsew", padx=S(24), pady=S(14))
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)
        bar = tk.Frame(main, bg=BG)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        tk.Label(bar, text="Cash in and out", font=(FONT, 14, "bold"), fg=TEXT, bg=BG).pack(side="left",
                                                                                          padx=(0, 16))
        self.filter_buttons = {}
        for key, text in (("since", "Since last emptied"), ("all", "Everything")):
            btn = HoverButton(bar, text=text, font=(FONT, 9, "bold"), padx=10, pady=4,
                              command=lambda k=key: self.set_filter(k))
            btn.pack(side="left", padx=(0, 4))
            self.filter_buttons[key] = btn
        table = tk.Frame(main, bg=PANEL)
        table.grid(row=1, column=0, sticky="nsew")
        columns = [("no", "#", 60, "w"), ("when", "Date & time", 140, "w"), ("type", "What", 130, "w"),
                   ("details", "Details", 360, "w"), ("amount", "Amount", 100, "e"),
                   ("balance", "In drawer after", 120, "e"), ("staff", "Staff", 130, "w")]
        self.tree = ttk.Treeview(table, columns=[c[0] for c in columns], show="headings", style="Dark.Treeview",
                                 selectmode="browse")
        for key, text, width, anchor in columns:
            self.tree.heading(key, text=text, anchor=anchor)
            self.tree.column(key, width=S(width), minwidth=S(40), anchor=anchor, stretch=key == "details")
        sb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("tampered", background=TAMPERED_BG, foreground=TAMPERED_FG)
        self.tree.tag_configure("even", background=ROW_A)
        self.tree.tag_configure("odd", background=ROW_B)
        for tag, color in (("out", PURPLE_HOVER), ("count", AMBER), ("banked", GREEN_HOVER), ("open", MUTED),
                           ("void", RED)):
            self.tree.tag_configure(tag, foreground=color)
        make_sortable(self.tree)
        self.empty = tk.Label(table, text="", font=(FONT, 11), fg=MUTED, bg=ROW_A)
        self.set_filter("since")

    def set_filter(self, key):
        self.filter = key
        for k, btn in self.filter_buttons.items():
            if k == key:
                btn.set_colors(AMBER, AMBER_HOVER, "#ffffff")
            else:
                btn.set_colors(GREY_BTN, GREY_HOVER, TEXT)
        self.refresh()

    def refresh(self):
        records = self.app.records
        last = max((i for i, r in enumerate(records) if r["type"] == TYPE_BANKED), default=-1)
        since = records[last + 1:]
        petty_refs = self.app.petty_receipts()
        takings = sum(r["total"] for r in since if r["type"] == TYPE_SALE
                      or (r["type"] == TYPE_VOID and r["ref"] not in petty_refs))
        petty = -sum(r["total"] for r in since if r["type"] == TYPE_PETTY
                     or (r["type"] == TYPE_VOID and r["ref"] in petty_refs))
        counts = sum(r["total"] for r in since if r["type"] == TYPE_COUNT)
        self.total.config(text=fmt_money(drawer_expected(records)))
        banked = records[last] if last >= 0 else None
        self.since.config(text=f"Since it was last emptied for banking on {banked['dt']:%d/%m/%Y}."
                          if banked else "Since the app started (never emptied for banking yet).")
        self.lines["takings"].config(text=fmt_money(takings))
        self.lines["petty"].config(text=fmt_money(-petty) if petty else "£0.00")
        self.lines["counts"].config(text=(("+" if counts > 0 else "") + fmt_money(counts)) if counts else "none",
                                    fg=AMBER if counts else TEXT)
        count = next((r for r in reversed(records) if r["type"] == TYPE_COUNT), None)
        self.lines["last_count"].config(text=f"{count['dt']:%d/%m %H:%M} · {count['staff']}" if count else "never")
        self.lines["last_banked"].config(text=f"{banked['dt']:%d/%m/%Y} · {fmt_money(-banked['total'])}"
                                         if banked else "never")
        self._fill(records if self.filter == "all" else since, petty_refs)

    def _fill(self, rows, petty_refs):
        self.tree.delete(*self.tree.get_children())
        # running balance: what the drawer held after each row (banking empties it)
        balance, shown, wanted = 0, [], {id(r) for r in rows}
        for r in self.app.records:
            if r["type"] in (TYPE_ACCEPTED,) + STOCK_TYPES:
                continue  # no cash involved
            balance = 0 if r["type"] == TYPE_BANKED else balance + r["total"]
            if id(r) in wanted:
                shown.append((r, balance))
        for n, (r, bal) in enumerate(reversed(shown)):
            kind = {TYPE_SALE: "Sale", TYPE_VOID: "Void", TYPE_PETTY: "Petty cash", TYPE_NOSALE: "No sale",
                    TYPE_COUNT: "Count", TYPE_BANKED: "Banked"}.get(r["type"], r["type"].title())
            if r["type"] == TYPE_SALE:
                details = f"Flat {r['flat']} · {r['name']} · {describe_items(r['wash'], r['dry'])}"
            else:
                details = r["notes"]
            amount = "" if r["type"] == TYPE_NOSALE else fmt_money(r["total"])
            tag = {TYPE_PETTY: "out", TYPE_COUNT: "count", TYPE_BANKED: "banked", TYPE_NOSALE: "open",
                   TYPE_VOID: "void"}.get(r["type"])
            tags = ["odd" if n % 2 else "even"] + ([tag] if tag else []) + (["tampered"] if r["tampered"] else [])
            self.tree.insert("", "end", values=(r["receipt"], r["dt"].strftime("%d/%m/%Y %H:%M"), kind, details,
                                                amount, fmt_money(bal), r["staff"]), tags=tags)
        self.tree.resort()
        if shown:
            self.empty.place_forget()
        else:
            self.empty.config(text="Nothing since the drawer was last emptied." if self.filter == "since"
                              else "Nothing yet.")
            self.empty.place(relx=0.5, rely=0.5, anchor="center")


# --------------------------------------------------------------------------
# Main application
# --------------------------------------------------------------------------

class App(tk.Tk):
    FLAT_COLUMNS = 6

    def __init__(self):
        super().__init__()
        self.report_callback_exception = self._on_error
        self.protocol("WM_DELETE_WINDOW", self.exit_app)  # the window's X closes it the same way as Exit
        prepare_data_dir()
        self.cfg, config_warning = load_config()
        apply_theme(self.cfg.get("theme", "dark"))  # before any widget (or rent/keys/post) uses a colour
        self.scale = self.winfo_fpixels("1i") / 96
        self.title(APP_TITLE)
        self.configure(bg=BG)
        self.geometry(f"{self.S(1360)}x{self.S(820)}")
        self.minsize(self.S(1120), self.S(700))
        try:
            self.state("zoomed")
        except tk.TclError:
            pass
        self._icons = {}
        self.apply_fullscreen()
        self.bind_all("<F11>", lambda e: self.attributes("-fullscreen", not self.attributes("-fullscreen")))
        self.tenants = load_tenants()
        ensure_sales_log()
        backup_log(SALES_PATH, "sales_log")
        upgraded = upgrade_sales_log()
        self.records, skipped, self.last_check = load_records()
        self.log_problems = self._check_log(upgraded)
        import rent  # here, not at the top: rent.py imports from this module
        self.rent_module = rent
        self.rent = rent.RentStore(self)
        import keys
        self.keys_module = keys
        self.keys = keys.KeyStore(self)
        import post
        self.post_module = post
        self.post = post.PostStore(self)
        import visitors
        self.visitors_module = visitors
        self.visitors = visitors.VisitorStore(self)
        self.selected_flat = None
        self.flat_buttons = {}
        self.view = None

        self.logo_image = self._load_logo()
        self._init_styles()
        self._build_ui()
        self.apply_prices()
        self.clear_sale()
        self.refresh_all()
        self.show_view("home")
        self.last_activity = datetime.now()
        for event in ("<KeyPress>", "<ButtonPress>", "<Motion>", "<MouseWheel>"):
            self.bind_all(event, self._note_activity, add="+")
        self.bind_all("<KeyPress>", self.wake, add="+")
        self.after(300, lambda: self._startup_checks(skipped, config_warning))
        self.after(500, self._poll)

    def S(self, px):
        return int(round(px * self.scale))

    ICON_FONT = r"C:\Windows\Fonts\segmdl2.ttf"  # Segoe MDL2 Assets: Windows' own icon font
    ICONS = {"home": "\ue80f", "tokens": "\ue719", "rent": "\uec07", "drawer": "\uec59", "keys": "\ue8d7",
             "post": "\ue715", "visitors": "\ue716", "lock": "\ue72e", "settings": "\ue713", "exit": "\ue7e8"}

    def icon(self, name, color, size=18):
        """A small icon drawn from Windows' icon font in the given colour, or None if it can't be
        (no Pillow, or no such font): buttons then just show their text."""
        key = (name, color, size)
        if key not in self._icons:
            try:
                from PIL import ImageDraw, ImageFont
                px = self.S(size)
                font = ImageFont.truetype(self.ICON_FONT, px)
                img = Image.new("RGBA", (px + 2, px + 2), (0, 0, 0, 0))
                ImageDraw.Draw(img).text((1, 1), self.ICONS[name], font=font, fill=color)
                self._icons[key] = ImageTk.PhotoImage(img)
            except Exception:
                self._icons[key] = None
        return self._icons[key]

    def _check_log(self, upgraded):
        """Returns a list of problems found in sales_log.csv (empty if it is intact).
        config.json remembers the last row the app wrote, which catches rows deleted
        from the end: the hash chain alone can't see those."""
        problems = []
        altered = [r["receipt"] for r in self.records if r["tampered"]]
        if altered:
            shown = ", ".join(f"#{n}" for n in altered[:12]) + (" …" if len(altered) > 12 else "")
            problems.append(f"{len(altered)} entr{'y was' if len(altered) == 1 else 'ies were'} changed, "
                            f"added or removed outside the app: {shown}")
        tip = self.cfg.get("log_tip")
        if tip and not any(r["receipt"] == tip["receipt"] and r["check"] == tip["check"] for r in self.records):
            problems.append(f"Entry #{tip['receipt']}, the last one the app wrote, is missing or changed. "
                            "Entries may have been deleted from the end of the log.")
        if not tip or upgraded:
            self._set_log_tip()
        return problems

    def accept_log(self, admin_name, reason, parent):
        """Accepts the log as it is now: recalculates every Check, then records who did it and why."""
        try:
            backup = reseal_sales_log("before_accept")
        except OSError as e:
            messagebox.showerror("Sales log", f"Could not update sales_log.csv (is it open in Excel?):\n\n{e}",
                                 parent=parent)
            return False
        self.records, _, self.last_check = load_records()
        self.log_problems = []
        self._set_log_tip()
        self.record_event(TYPE_ACCEPTED, "", "", 0, 0, 0, admin_name, parent=parent,
                          notes=f"{reason}. Flagged log kept as backups\\{os.path.basename(backup)}")
        return True

    def _set_log_tip(self):
        last = self.records[-1] if self.records else None
        self.cfg["log_tip"] = {"receipt": last["receipt"], "check": last["check"]} if last else None
        try:
            save_config(self.cfg)
        except OSError:
            pass  # the sale is already safely in the log; this is only a cross-check

    def _load_logo(self, height=36):
        """Sets the window icon and returns the sidebar logo image, or None. Never raises."""
        path = find_logo()
        if path is None:
            return None
        try:
            if Image is not None:
                img = Image.open(path).convert("RGBA")
                h = self.S(height)
                logo = ImageTk.PhotoImage(img.resize((round(img.width * h / img.height), h), Image.LANCZOS))
                mark = img.crop((0, 0, LOGO_MARK_WIDTH, img.height))
                self._icon = ImageTk.PhotoImage(mark.resize((64, round(64 * mark.height / mark.width)),
                                                            Image.LANCZOS))
                self.iconphoto(True, self._icon)
                return logo
            logo = tk.PhotoImage(file=path)
            return logo.subsample(max(1, round(logo.height() / self.S(height))))
        except Exception:
            return None  # a missing or broken logo must never stop the shop opening

    # ---- setup -----------------------------------------------------------
    def _init_styles(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure("Dark.Treeview", background=ROW_A, fieldbackground=ROW_A, foreground=TEXT,
                     rowheight=self.S(38), borderwidth=0, font=(FONT, 10))  # 38: big enough for a fingertip
        st.layout("Dark.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        st.map("Dark.Treeview", background=[("selected", TEAL)], foreground=[("selected", "#ffffff")])
        st.configure("Dark.Treeview.Heading", background=PANEL_2, foreground=MUTED, relief="flat",
                     borderwidth=0, font=(FONT, 10, "bold"), padding=(8, 8))
        st.map("Dark.Treeview.Heading", background=[("active", BORDER)])
        st.configure("Vertical.TScrollbar", background=PANEL_2, troughcolor=PANEL, bordercolor=PANEL,
                     arrowcolor=MUTED, lightcolor=PANEL_2, darkcolor=PANEL_2, gripcount=0)
        st.map("Vertical.TScrollbar", background=[("active", GREY_HOVER)])
        st.configure("TNotebook", background=PANEL, borderwidth=0, tabmargins=0,
                     bordercolor=BORDER, lightcolor=PANEL, darkcolor=PANEL)
        st.configure("TNotebook.Tab", background=PANEL_2, foreground=MUTED, padding=(18, 8),
                     font=(FONT, 10, "bold"), bordercolor=PANEL, lightcolor=PANEL_2, borderwidth=0)
        st.map("TNotebook.Tab", background=[("selected", TEAL)], foreground=[("selected", "#ffffff")],
               lightcolor=[("selected", TEAL)])
        st.configure("TCombobox", fieldbackground=PANEL_2, background=PANEL_2, foreground=TEXT,
                     arrowcolor=TEXT, bordercolor=BORDER, lightcolor=PANEL_2, darkcolor=PANEL_2,
                     selectbackground=PANEL_2, selectforeground=TEXT, insertcolor=TEXT)
        st.map("TCombobox", fieldbackground=[("readonly", PANEL_2)], foreground=[("readonly", TEXT)],
               background=[("active", GREY_HOVER)])
        self.option_add("*TCombobox*Listbox.background", PANEL_2)
        self.option_add("*TCombobox*Listbox.foreground", TEXT)
        self.option_add("*TCombobox*Listbox.selectBackground", TEAL)
        self.option_add("*TCombobox*Listbox.font", (FONT, 11))

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self._build_topbar()
        self._build_statusbar()
        self.token_view = tk.Frame(self, bg=BG)
        self.token_view.grid(row=1, column=0, sticky="nsew")
        self.token_view.columnconfigure(1, weight=1)
        self.token_view.rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_main()
        self.rent_view = self.rent_module.RentView(self, self)
        self.rent_view.grid(row=1, column=0, sticky="nsew")
        self.drawer_view = DrawerView(self, self)
        self.drawer_view.grid(row=1, column=0, sticky="nsew")
        self.key_view = self.keys_module.KeyView(self, self)
        self.key_view.grid(row=1, column=0, sticky="nsew")
        self.post_view = self.post_module.PostView(self, self)
        self.post_view.grid(row=1, column=0, sticky="nsew")
        self.visitor_view = self.visitors_module.VisitorView(self, self)
        self.visitor_view.grid(row=1, column=0, sticky="nsew")
        self.home_view = tk.Frame(self, bg=BG)
        self.home_view.grid(row=1, column=0, sticky="nsew")
        self._build_home()
        self._build_idle()

    def _build_topbar(self):
        bar = tk.Frame(self, bg=STATUS_BG)
        bar.grid(row=0, column=0, sticky="ew")
        if self.logo_image is not None:
            logo = tk.Label(bar, image=self.logo_image, bg=STATUS_BG, cursor="hand2")
            logo.pack(side="left", padx=(self.S(22), self.S(28)), pady=self.S(10))
            logo.bind("<Button-1>", lambda e: self.show_view("home"))
        self.view_buttons = {}
        for key, text, color, _ in self.sections():
            btn = HoverButton(bar, text=f" {text}", font=(FONT, 12, "bold"), padx=12, pady=8,
                              image=self.icon(key, TEXT), compound="left",
                              command=lambda k=key: self.show_view(k))
            btn.pack(side="left", padx=(0, 6))
            self.view_buttons[key] = btn
        right = {}
        for key, text, bg, hover, command, left in (
                ("exit", "Exit", RED, RED_HOVER, self.exit_app, 0),
                ("settings", "Settings", None, None, self.open_settings, 6),
                ("lock", "Lock", None, None, self.lock, 6)):
            fg = "#ffffff" if bg else TEXT
            right[key] = HoverButton(bar, text=f" {text}", bg=bg, hover=hover, fg=fg, font=(FONT, 10, "bold"),
                                     image=self.icon(key, fg, 15), compound="left", command=command)
            right[key].pack(side="right", padx=(0, self.S(22) if key == "exit" else 6))
        # A narrow screen (or Windows display scaling) would push Exit off the edge: tighten the tabs,
        # then show Lock and Settings as icons only, until the bar fits.
        screen = self.winfo_screenwidth()
        bar.update_idletasks()
        if bar.winfo_reqwidth() > screen:
            for btn in self.view_buttons.values():
                btn.config(font=(FONT, 11, "bold"), padx=8)
            bar.update_idletasks()
        if bar.winfo_reqwidth() > screen:
            for key in ("settings", "lock"):
                if right[key].cget("image"):
                    right[key].config(compound="none")  # with an image set, "none" shows only the image

    def sections(self):
        """(view, tab name, colour, hover): each section keeps one colour for its tab and Home tile."""
        return [("home", "Home", BLUE, BLUE_HOVER), ("tokens", "Laundry", TEAL, TEAL_HOVER),
                ("rent", "Rent", GREEN, GREEN_HOVER), ("drawer", "Drawer", AMBER, AMBER_HOVER),
                ("keys", "Keys", ORANGE, ORANGE_HOVER), ("post", "Post", PURPLE, PURPLE_HOVER),
                ("visitors", "Visitors", PINK, PINK_HOVER)]

    def show_view(self, key):
        views = {"home": self.home_view, "tokens": self.token_view, "rent": self.rent_view,
                 "drawer": self.drawer_view, "keys": self.key_view, "post": self.post_view,
                 "visitors": self.visitor_view}
        if key not in views:
            key = "home"
        self.view = key
        for k, _, color, hover in self.sections():
            btn = self.view_buttons[k]
            if k == key:
                btn.set_colors(color, hover, "#ffffff")
                btn.config(image=self.icon(k, "#ffffff"))
            else:
                btn.set_colors(GREY_BTN, GREY_HOVER, TEXT)
                btn.config(image=self.icon(k, TEXT))
        views[key].tkraise()
        if key == "rent":
            self.rent_view.refresh()
        elif key == "drawer":
            self.drawer_view.refresh()
        elif key == "keys":
            self.key_view.refresh()
        elif key == "post":
            self.post_view.refresh()
        elif key == "visitors":
            self.visitor_view.refresh()
        elif key == "home":
            self.refresh_home()

    # ---- home ------------------------------------------------------------
    def _build_home(self):
        S = self.S
        page = tk.Frame(self.home_view, bg=BG)
        page.place(relx=0.5, rely=0.45, relwidth=0.9, anchor="center")  # tiles share the screen's width
        self.home_greeting = tk.Label(page, text="", font=(FONT, 26, "bold"), fg=TEXT, bg=BG, anchor="w")
        self.home_greeting.pack(fill="x")
        self.home_date = tk.Label(page, text="", font=(FONT, 13), fg=MUTED, bg=BG, anchor="w")
        self.home_date.pack(fill="x", pady=(0, S(22)))

        tiles = tk.Frame(page, bg=BG)
        tiles.pack(fill="x")
        colors = {key: (color, hover) for key, _, color, hover in self.sections()}
        self.home_stats = {}
        for i, (key, title, blurb, rows) in enumerate([
                ("tokens", "Laundry", "Washing and dryer tokens",
                 ("Takings today", "Sales today", "Petty cash today", "In the drawer now", "Tokens in stock")),
                ("rent", "Rent", "Payments, receipts, checking and banking",
                 ("Taken today", "Awaiting check", "In safe · not banked", "Banked this month", "Last banked")),
                ("keys", "Keys", "Keys lent to contractors",
                 ("Keys out now", "Overdue", "Issued today", "Returned today")),
                ("post", "Post", "Letters for the CEO and DCEO, and post sent",
                 ("Waiting to hand over", "Received today", "Sent today", "Postage this month"))]):
            accent, hover = colors[key]
            tiles.columnconfigure(i, weight=1, uniform="tile")
            tile = tk.Frame(tiles, bg=PANEL, cursor="hand2")
            tile.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else S(18), 0))
            tk.Frame(tile, bg=accent, height=4).pack(fill="x")
            body = tk.Frame(tile, bg=PANEL)
            body.pack(fill="both", expand=True, padx=S(22), pady=S(18))
            head = tk.Frame(body, bg=PANEL)
            head.pack(fill="x")
            icon = self.icon(key, accent, 22)
            if icon is not None:
                tk.Label(head, image=icon, bg=PANEL).pack(side="left", padx=(0, S(10)))
            tk.Label(head, text=title, font=(FONT, 19, "bold"), fg=TEXT, bg=PANEL, anchor="w").pack(side="left")
            tk.Label(body, text=blurb, font=(FONT, 10), fg=MUTED, bg=PANEL, anchor="w").pack(fill="x")
            tk.Frame(body, bg=BORDER, height=1).pack(fill="x", pady=S(14))
            values = []
            for label in rows:
                row = tk.Frame(body, bg=PANEL)
                row.pack(fill="x", pady=2)
                tk.Label(row, text=label, font=(FONT, 11), fg=MUTED, bg=PANEL).pack(side="left")
                value = tk.Label(row, text="", font=(FONT, 12, "bold"), fg=TEXT, bg=PANEL)
                value.pack(side="right")
                values.append(value)
            self.home_stats[key] = values
            HoverButton(body, text=f"Open {title.lower()}  →", bg=accent, hover=hover, fg="#ffffff",
                        font=(FONT, 12, "bold"), pady=9,
                        command=lambda k=key: self.show_view(k)).pack(side="bottom", fill="x", pady=(S(16), 0))
            self._bind_click(tile, lambda e, k=key: self.show_view(k))

        tk.Label(page, text="NEEDS ATTENTION", font=(FONT, 10, "bold"), fg=MUTED, bg=BG,
                 anchor="w").pack(fill="x", pady=(S(28), S(8)))
        self.home_attention = tk.Frame(page, bg=BG)
        self.home_attention.pack(fill="x")

    def _bind_click(self, widget, handler):
        """Makes a whole tile clickable (buttons keep their own command)."""
        if not isinstance(widget, tk.Button):
            widget.bind("<Button-1>", handler)
            widget.config(cursor="hand2")
        for child in widget.winfo_children():
            self._bind_click(child, handler)

    def refresh_home(self):
        now = datetime.now()
        today = now.date()
        part = "morning" if now.hour < 12 else "afternoon" if now.hour < 17 else "evening"
        self.home_greeting.config(text=f"Good {part}")
        self.home_date.config(text=now.strftime("%A %d %B %Y"))

        s = summarise([r for r in self.records if r["dt"].date() == today], self.petty_receipts())
        takings, sales, petty, drawer, in_stock = self.home_stats["tokens"]
        takings.config(text=fmt_money(s["total"]))
        sales.config(text=str(s["sales"] - s["voids"]))
        petty.config(text=fmt_money(s["petty"]))
        in_drawer = drawer_expected(self.records)
        drawer_limit = self.bank_limit("drawer") * 100
        drawer_over = drawer_limit and in_drawer > drawer_limit
        drawer.config(text=fmt_money(in_drawer), fg=AMBER if drawer_over else TEXT)
        stock, low = token_stock(self.records), self.low_stock_levels()
        running_low = [name for key, name in (("wash", "washing"), ("dry", "dryer"))
                       if stock is not None and stock[key] < low[key]]
        in_stock.config(text="not tracked" if stock is None else f"{stock['wash']} washing · {stock['dry']} dryer",
                        fg=MUTED if stock is None else AMBER if running_low else TEXT)

        rent, store = self.rent_module, self.rent
        pays = list(store.payments().values())
        taken = [p for p in pays if p["dt"].date() == today and not p["cancelled_at"]]
        awaiting = [p for p in pays if store.status(p) == rent.S_AWAITING]
        in_safe = [p for p in pays if store.status(p) == rent.S_IN_SAFE]

        def money_count(ps):
            return f"{fmt_money(sum(p['total'] for p in ps))}  ({len(ps)})"

        unbanked = sum(p["total"] for p in awaiting + in_safe)
        rent_limit = self.bank_limit("rent") * 100
        rent_over = rent_limit and unbanked > rent_limit
        for i, (label, ps) in enumerate(zip(self.home_stats["rent"], (taken, awaiting, in_safe))):
            label.config(text=money_count(ps), fg=AMBER if rent_over and i and ps else TEXT)
        month = [d for d in store.deposits().values() if d["bank_date"] and d["bank_date"] >= today.replace(day=1)]
        self.home_stats["rent"][3].config(text=f"{fmt_money(sum(d['total'] for d in month))}  ({len(month)})")
        last = max((d["bank_date"] for d in store.deposits().values() if d["bank_date"]), default=None)
        self.home_stats["rent"][4].config(text=last.strftime("%d/%m/%Y") if last else "never")

        keys = self.keys_module
        out = self.keys.out_now()
        late = sorted((l for l in out if self.keys.status(l) == keys.S_OVERDUE), key=lambda l: l["return_by"])
        n_out = sum(keys.key_count(l["out"]) for l in out)
        loans = [l for l in self.keys.loans().values() if not l["voided_at"]]
        issued = [l for l in loans if l["dt"].date() == today]
        back = [e for l in loans for e in l["events"] if e["Type"] == keys.K_RETURNED and e["dt"].date() == today]
        out_now, overdue, issued_today, back_today = self.home_stats["keys"]
        out_now.config(text=f"{n_out}  ({keys.plural(len(out), 'slip')})" if out else "none")
        overdue.config(text=str(len(late)) if late else "none", fg=RED if late else TEXT)
        issued_today.config(text=str(sum(keys.key_count(l["keys"]) for l in issued)))
        back_today.config(text=str(sum(keys.key_count(keys.text_to_keys(e["Keys"])) for e in back)))

        waiting = sorted(self.post.waiting(), key=lambda it: it["dt"])
        items_ = [it for it in self.post.items().values() if not it["voided_at"]]
        w, received, sent, postage = self.home_stats["post"]
        w.config(text=str(len(waiting)) if waiting else "none", fg=YELLOW if waiting else TEXT)
        received.config(text=str(sum(1 for it in items_ if it["incoming"] and it["dt"].date() == today)))
        sent.config(text=str(sum(1 for it in items_ if not it["incoming"] and it["dt"].date() == today)))
        postage.config(text=fmt_money(sum(it["cost"] for it in items_ if not it["incoming"]
                                          and it["dt"].date() >= today.replace(day=1))))

        items = []
        if self.log_problems:
            items.append((RED, "The token sales log was changed outside the app. See Settings → Logs"))
        if self.rent.problems:
            items.append((RED, "The rent log was changed outside the app. See Settings → Logs"))
        if self.keys.problems:
            items.append((RED, "The key log was changed outside the app. See Settings → Logs"))
        if self.post.problems:
            items.append((RED, "The post log was changed outside the app. See Settings → Logs"))
        if self.visitors.problems:
            items.append((RED, "The visitor log was changed outside the app. See Settings → Logs"))
        forgotten = self.visitors.forgotten()
        if forgotten:
            items.append((AMBER, f"{keys.plural(len(forgotten), 'visitor')} from an earlier day never signed out. "
                                 "Sign them out on the Visitors tab"))
        if waiting:
            who = ", ".join(sorted({it["to"] for it in waiting}))
            items.append((YELLOW, f"{keys.plural(len(waiting), 'post item')} for {who} waiting to be handed over "
                                  f"(oldest arrived {waiting[0]['dt']:%d/%m})"))
        if late:
            first = late[0]
            items.append((RED, f"{keys.plural(len(late), 'key slip')} overdue. Oldest: "
                               f"{keys.keys_display(first['out'])} for {first['property']} with {first['contractor']}, "
                               f"due back {first['return_by']:%d/%m}"))
        if awaiting:
            n = len(awaiting)
            items.append((YELLOW, f"{n} rent payment{'s' if n != 1 else ''} ({fmt_money(sum(p['total'] for p in awaiting))}) "
                                  f"waiting to be checked and put in the safe"))
        if rent_over:
            items.append((AMBER, f"{fmt_money(unbanked)} of rent is waiting to be banked, over the "
                                 f"£{rent_limit // 100} limit. Let the person who banks know"))
        if drawer_over:
            items.append((AMBER, f"{fmt_money(in_drawer)} of token takings in the cash drawer, over the "
                                 f"£{drawer_limit // 100} limit. Due for banking (Drawer → Bank takings)"))
        if in_safe:
            oldest = min(p["dt"] for p in in_safe)
            days = (today - oldest.date()).days
            if days > store.unbanked_days():
                items.append((AMBER, f"Rent has been in the safe since {oldest:%d/%m} ({days} days). "
                                          "It is due for banking"))
        if running_low:
            items.append((AMBER, f"Running low on {' and '.join(running_low)} tokens "
                                 f"({stock['wash']} washing, {stock['dry']} dryer left). Order more"))
        if self.cfg["drawer_enabled"] and self.cfg["com_port"] not in {p.device for p in available_ports()}:
            items.append((YELLOW, "Cash drawer trigger not connected. Check the USB cable, or the port in Settings"))
        if any(pin_matches(DEFAULT_PIN, st) for st in self.cfg["staff"]):
            items.append((YELLOW, f"Default PIN {DEFAULT_PIN} is still in use. Change it in Settings → Staff & PINs"))
        if not items:
            items.append((GREEN, "Nothing needs attention"))

        for w in self.home_attention.winfo_children():
            w.destroy()
        for color, text in items:
            row = tk.Frame(self.home_attention, bg=PANEL)
            row.pack(fill="x", pady=2)
            tk.Frame(row, bg=color, width=4).pack(side="left", fill="y")
            tk.Label(row, text=text, font=(FONT, 11), fg=TEXT, bg=PANEL, anchor="w").pack(
                side="left", fill="x", padx=self.S(14), pady=self.S(9))

    def _section(self, parent, number, title, top=18):
        row = tk.Frame(parent, bg=PANEL)
        row.pack(fill="x", pady=(top, 8))
        tk.Label(row, text=number, font=(FONT, 9, "bold"), fg=BG, bg=MUTED, width=2).pack(side="left")
        tk.Label(row, text=title.upper(), font=(FONT, 10, "bold"), fg=MUTED, bg=PANEL).pack(side="left", padx=8)
        return row

    def _build_sidebar(self):
        side = tk.Frame(self.token_view, bg=PANEL, width=self.S(420))
        side.grid(row=0, column=0, sticky="ns")
        side.pack_propagate(False)
        p = tk.Frame(side, bg=PANEL)
        p.pack(fill="both", expand=True, padx=self.S(22), pady=self.S(14))

        self._section(p, "1", "Choose flat", top=0)
        self.flat_search = PlaceholderEntry(p, "Search flat number or name…")
        self.flat_search.pack(fill="x", ipady=5)
        self.flat_search.bind("<KeyRelease>", lambda e: self._paint_flat_buttons())
        self.flat_search.bind("<Return>", lambda e: self._select_single_match())
        self.flat_search.bind("<Escape>", lambda e: (self.flat_search.clear(), self._paint_flat_buttons()))
        self.flat_grid = tk.Frame(p, bg=PANEL)
        self.flat_grid.pack(fill="x", pady=(8, 0))
        tenant_row = tk.Frame(p, bg=PANEL)
        tenant_row.pack(fill="x", pady=(8, 0))
        self.history_link = tk.Label(tenant_row, text="History", font=(FONT, 10, "underline"), fg=TEAL,
                                     bg=PANEL, cursor="hand2")
        self.history_link.bind("<Button-1>", lambda e: self.open_history())
        self.tenant_label = tk.Label(tenant_row, text="", font=(FONT, 12, "bold"), fg=TEXT, bg=PANEL, anchor="w")
        self.tenant_label.pack(side="left", fill="x", expand=True)
        self._build_flat_grid()

        tokens = self._section(p, "2", "Add tokens")
        clear = tk.Label(tokens, text="Clear", font=(FONT, 10, "underline"), fg=MUTED, bg=PANEL, cursor="hand2")
        clear.pack(side="right")
        clear.bind("<Button-1>", lambda e: self.clear_sale())
        self.wash_stepper = TokenStepper(p, "Washing tokens", TEAL, TEAL_HOVER, self.update_total)
        self.wash_stepper.pack(fill="x")
        self.dry_stepper = TokenStepper(p, "Dryer tokens", ORANGE, ORANGE_HOVER, self.update_total)
        self.dry_stepper.pack(fill="x", pady=(8, 0))

        total_row = tk.Frame(p, bg=PANEL)
        total_row.pack(fill="x", pady=(16, 10))
        tk.Label(total_row, text="Total", font=(FONT, 13), fg=MUTED, bg=PANEL).pack(side="left", anchor="s")
        self.total_label = tk.Label(total_row, text="£0.00", font=(FONT, 28, "bold"), fg=TEXT, bg=PANEL)
        self.total_label.pack(side="right")

        self.sell_btn = HoverButton(p, text="Take payment  ·  Open drawer", bg=GREEN, hover=GREEN_HOVER,
                                    font=(FONT, 13, "bold"), pady=12, command=self.sell)
        self.sell_btn.pack(fill="x")

    def _build_flat_grid(self):
        for w in self.flat_grid.winfo_children():
            w.destroy()
        self.flat_buttons = {}
        for c in range(self.FLAT_COLUMNS):
            self.flat_grid.columnconfigure(c, weight=1, uniform="flat")
        for i, t in enumerate(self.tenants):
            btn = HoverButton(self.flat_grid, text=t["flat"], font=(FONT, 11, "bold"), padx=0, pady=7,
                              command=lambda f=t["flat"]: self.select_flat(f))
            btn.grid(row=i // self.FLAT_COLUMNS, column=i % self.FLAT_COLUMNS, sticky="nsew", padx=2, pady=2)
            btn.bind("<Enter>", lambda e, f=t["flat"]: self._show_tenant(f, hover=True), add="+")
            btn.bind("<Leave>", lambda e: self._show_tenant(self.selected_flat), add="+")
            self.flat_buttons[t["flat"]] = btn
        self._paint_flat_buttons()
        self._show_tenant(self.selected_flat)

    def _build_main(self):
        main = tk.Frame(self.token_view, bg=BG)
        main.grid(row=0, column=1, sticky="nsew", padx=self.S(24), pady=self.S(14))
        main.columnconfigure(0, weight=1)
        main.rowconfigure(3, weight=1)

        head = tk.Frame(main, bg=BG)
        head.grid(row=0, column=0, sticky="ew")
        tk.Label(head, text="Today", font=(FONT, 20, "bold"), fg=TEXT, bg=BG).pack(side="left")
        self.date_label = tk.Label(head, text="", font=(FONT, 12), fg=MUTED, bg=BG)
        self.date_label.pack(side="left", padx=12, pady=(8, 0))

        cards = tk.Frame(main, bg=BG)
        cards.grid(row=1, column=0, sticky="ew", pady=(12, 18))
        self.stat = {}
        for i, (key, title, accent) in enumerate([("wash", "Washing tokens", TEAL), ("dry", "Dryer tokens", ORANGE),
                                                   ("total", "Takings", GREEN), ("petty", "Petty cash", PURPLE),
                                                   ("sales", "Sales", MUTED)]):
            cards.columnconfigure(i, weight=1, uniform="card")
            card = tk.Frame(cards, bg=PANEL)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else self.S(12), 0))
            tk.Frame(card, bg=accent, height=3).pack(fill="x")
            tk.Label(card, text=title.upper(), font=(FONT, 9, "bold"), fg=MUTED, bg=PANEL,
                     anchor="w").pack(fill="x", padx=16, pady=(12, 0))
            value = tk.Label(card, text="0", font=(FONT, 24, "bold"), fg=TEXT, bg=PANEL, anchor="w")
            value.pack(fill="x", padx=16)
            sub = tk.Label(card, text="", font=(FONT, 9), fg=MUTED, bg=PANEL, anchor="w")
            sub.pack(fill="x", padx=16, pady=(0, 12))
            self.stat[key] = (value, sub)

        bar = tk.Frame(main, bg=BG)
        bar.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        tk.Label(bar, text="Transaction log", font=(FONT, 14, "bold"), fg=TEXT, bg=BG).pack(side="left")
        # the drawer's own actions (no sale, count, petty cash, banking) are on the Drawer screen
        for text, bg, hover, cmd in [("Reports", GREY_BTN, GREY_HOVER, self.open_reports),
                                     ("Stock", GREY_BTN, GREY_HOVER, self.open_stock),
                                     ("Void", RED, RED_HOVER, self.void_selected)]:
            HoverButton(bar, text=text, bg=bg, hover=hover, font=(FONT, 10, "bold"),
                        command=cmd).pack(side="right", padx=(8, 0))
        self.log_search = PlaceholderEntry(bar, "Search log…", width=18)
        self.log_search.pack(side="right", padx=(0, 4), ipady=4)
        self.log_search.bind("<KeyRelease>", lambda e: self.refresh_tree())

        table = tk.Frame(main, bg=PANEL)
        table.grid(row=3, column=0, sticky="nsew")
        columns = [("receipt", "#", 60, "w"), ("datetime", "Date & time", 140, "w"), ("type", "Type", 110, "w"),
                   ("flat", "Flat", 60, "w"), ("tenant", "Tenant", 180, "w"), ("wash", "Washing", 80, "e"),
                   ("dry", "Dryer", 70, "e"), ("total", "Total", 90, "e"), ("staff", "Staff", 110, "w"),
                   ("notes", "Notes", 220, "w")]
        self.tree = ttk.Treeview(table, columns=[c[0] for c in columns], show="headings",
                                 style="Dark.Treeview", selectmode="browse")
        for key, text, width, anchor in columns:
            self.tree.heading(key, text=text, anchor=anchor)
            self.tree.column(key, width=self.S(width), minwidth=self.S(40), anchor=anchor,
                             stretch=key in ("tenant", "notes"))
        sb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("tampered", background=TAMPERED_BG, foreground=TAMPERED_FG)  # first = wins
        self.tree.tag_configure("even", background=ROW_A)
        self.tree.tag_configure("odd", background=ROW_B)
        self.tree.tag_configure("void", foreground=RED)
        self.tree.tag_configure("nosale", foreground=YELLOW)
        self.tree.tag_configure("petty", foreground=PURPLE_HOVER)
        self.tree.tag_configure("accepted", foreground=TAMPERED_FG)
        self.tree.tag_configure("stock", foreground=TEAL)
        self.tree.tag_configure("banked", foreground=GREEN_HOVER)
        self.tree.tag_configure("count", foreground=AMBER)
        make_sortable(self.tree)
        self.tree.tag_configure("voided", foreground=DIM, font=(FONT, 10, "overstrike"))
        self.empty_label = tk.Label(table, text="No transactions yet. Sales will appear here.",
                                    font=(FONT, 11), fg=MUTED, bg=ROW_A)

    def _build_statusbar(self):
        bar = tk.Frame(self, bg=STATUS_BG)
        bar.grid(row=2, column=0, sticky="ew")
        self.drawer_dot = tk.Label(bar, text="●", font=(FONT, 11), fg=DIM, bg=STATUS_BG)
        self.drawer_dot.pack(side="left", padx=(14, 4), pady=6)
        self.drawer_status = tk.Label(bar, text="", font=(FONT, 10), fg=MUTED, bg=STATUS_BG)
        self.drawer_status.pack(side="left")
        self.message_label = tk.Label(bar, text="", font=(FONT, 10), fg=MUTED, bg=STATUS_BG)
        self.message_label.pack(side="left", padx=24)
        self.warning_label = tk.Label(bar, text="", font=(FONT, 10, "bold"), fg=YELLOW, bg=STATUS_BG)
        self.warning_label.pack(side="right", padx=14)

    # ---- state helpers ---------------------------------------------------
    def save_cfg(self):
        try:
            save_config(self.cfg)
        except OSError as e:
            messagebox.showerror("Settings", f"Could not save config.json:\n\n{e}", parent=self)
        self.refresh_warning()
        self.update_drawer_status()

    def apply_fullscreen(self):
        """Full screen hides the taskbar and the window's title bar; F11 switches it for now."""
        self.attributes("-fullscreen", bool(self.cfg.get("fullscreen", True)))

    def idle_minutes(self):
        return parse_int(self.cfg.get("idle_minutes", DEFAULT_IDLE_MINUTES), DEFAULT_IDLE_MINUTES)

    def bank_limit(self, which):
        """Whole pounds above which Home highlights money waiting to be banked (0 = never)."""
        default = DEFAULT_BANK_LIMITS[which]
        return parse_int(self.cfg.get(f"bank_limit_{which}", default), default)

    def low_stock_levels(self):
        return {"wash": parse_int(self.cfg.get("low_stock_washing", DEFAULT_LOW_STOCK), DEFAULT_LOW_STOCK),
                "dry": parse_int(self.cfg.get("low_stock_dryer", DEFAULT_LOW_STOCK), DEFAULT_LOW_STOCK)}

    def _note_activity(self, _event=None):
        self.last_activity = datetime.now()

    def _check_idle(self):
        """The idle screen, with half-entered forms cleared, once nobody has used the app for a while.
        Never while a dialog is open: someone may be in the middle of entering a PIN."""
        minutes = self.idle_minutes()
        if not minutes or self.view == "idle":
            return
        if datetime.now() - self.last_activity < timedelta(minutes=minutes):
            return
        if any(isinstance(w, tk.Toplevel) for w in self.winfo_children()):
            return
        self.lock()

    def lock(self):
        """Clears anything half-entered and shows the idle screen (the Lock button, or when idle)."""
        self.clear_sale()
        self.log_search.clear()
        self.refresh_tree()
        self.rent_view.reset()
        self.key_view.reset()
        self.post_view.reset()
        self.visitor_view.reset()
        self.set_message("")
        self.sleep()

    # ---- idle screen -----------------------------------------------------
    def _build_idle(self):
        """Covers the whole window while nobody is using the app: the time and what's waiting,
        but no money, since anyone passing can see the screen. A tap or key press wakes it."""
        S = self.S
        self.idle_view = tk.Frame(self, bg=BG, cursor="hand2")
        page = tk.Frame(self.idle_view, bg=BG)
        page.place(relx=0.5, rely=0.45, anchor="center")
        self.idle_logo = self._load_logo(height=96)
        if self.idle_logo is not None:
            tk.Label(page, image=self.idle_logo, bg=BG).pack(pady=(0, S(40)))
        self.idle_clock = tk.Label(page, text="", font=(FONT, 80, "bold"), fg=TEXT, bg=BG)
        self.idle_clock.pack()
        self.idle_date = tk.Label(page, text="", font=(FONT, 18), fg=MUTED, bg=BG)
        self.idle_date.pack(pady=(0, S(40)))
        self.idle_status = tk.Frame(page, bg=BG)
        self.idle_status.pack()
        tk.Label(page, text="Tap anywhere or press any key to start", font=(FONT, 12), fg=DIM,
                 bg=BG).pack(pady=(S(48), 0))
        self._idle_lines = None
        self._bind_click(self.idle_view, self.wake)

    def _idle_lines_now(self):
        """What's waiting, as counts only."""
        keys = self.keys_module
        out = self.keys.out_now()
        late = sum(1 for l in out if self.keys.status(l) == keys.S_OVERDUE)
        waiting = self.post.waiting()
        awaiting = sum(1 for p in self.rent.payments().values() if self.rent.status(p) == self.rent_module.S_AWAITING)
        visitors_in = len(self.visitors.in_now())
        lines = []
        if visitors_in:
            lines.append((PINK, f"{keys.plural(visitors_in, 'visitor')} signed in"))
        if late:
            lines.append((RED, f"{keys.plural(late, 'key slip')} overdue"))
        if out:
            lines.append((ORANGE, f"{keys.plural(sum(keys.key_count(l['out']) for l in out), 'key')} out"))
        if waiting:
            lines.append((YELLOW, f"{keys.plural(len(waiting), 'post item')} waiting to be handed over"))
        if awaiting:
            lines.append((YELLOW, f"{keys.plural(awaiting, 'rent payment')} waiting to be checked"))
        return lines

    def _tick_idle(self):
        if self.view != "idle":
            return
        now = datetime.now()
        self.idle_clock.config(text=now.strftime("%H:%M"))
        self.idle_date.config(text=now.strftime("%A %d %B %Y"))
        lines = self._idle_lines_now()
        if lines != self._idle_lines:  # rebuilt only when something changed, so it doesn't flicker
            self._idle_lines = lines
            for w in self.idle_status.winfo_children():
                w.destroy()
            for color, text in lines:
                row = tk.Frame(self.idle_status, bg=BG)
                row.pack(anchor="w", pady=3)
                tk.Label(row, text="●", font=(FONT, 12), fg=color, bg=BG).pack(side="left", padx=(0, 10))
                tk.Label(row, text=text, font=(FONT, 15), fg=TEXT, bg=BG).pack(side="left")
            self._bind_click(self.idle_status, self.wake)
        self.after(1000, self._tick_idle)

    def sleep(self):
        self.view = "idle"
        self._idle_lines = None
        self.idle_view.place(x=0, y=0, relwidth=1, relheight=1)
        self.idle_view.lift()
        self.idle_view.focus_set()  # so the key that wakes it isn't typed into a box underneath
        self._tick_idle()

    def wake(self, _event=None):
        if self.view != "idle":
            return
        self.idle_view.place_forget()
        self.last_activity = datetime.now()
        self.show_view("home")

    def find_staff(self, pin):
        if not pin:
            return None
        for s in self.cfg["staff"]:
            if pin_matches(pin, s):
                return s
        return None

    def tenant_name(self, flat):
        for t in self.tenants:
            if t["flat"] == flat:
                return t["name"]
        return ""

    def voided_receipts(self):
        return {r["ref"] for r in self.records if r["type"] == TYPE_VOID and r["ref"] is not None}

    def petty_receipts(self):
        return {r["receipt"] for r in self.records if r["type"] == TYPE_PETTY}

    def next_receipt(self):
        return (self.records[-1]["receipt"] + 1) if self.records else 1

    def set_message(self, text, color=None, action=None):
        """Shows a message in the status bar and, briefly, as a pop-up in the corner, which is much
        harder to miss. action: optional (button text, function), e.g. ("Give change", ...)."""
        color = color or MUTED
        self.message_label.config(text=text, fg=color)
        self._show_toast(text, color, action)

    def _show_toast(self, text, color, action=None):
        if getattr(self, "_toast", None) is not None:
            self._toast.destroy()
            self.after_cancel(self._toast_timer)
            self._toast = None
        if not text or self.view == "idle":  # never over the idle screen: it could show a payment
            return
        S = self.S
        toast = self._toast = tk.Frame(self, bg=PANEL_2, highlightthickness=1, highlightbackground=BORDER)
        tk.Frame(toast, bg=color, width=S(5)).pack(side="left", fill="y")
        tk.Label(toast, text=text, font=(FONT, 12), fg=TEXT, bg=PANEL_2, wraplength=S(520), justify="left",
                 anchor="w").pack(side="left", padx=S(16), pady=S(14))
        if action:
            label, command = action

            def run():
                self._show_toast("", color)
                command()

            HoverButton(toast, text=label, bg=color, hover=color, fg="#ffffff", font=(FONT, 11, "bold"),
                        command=run).pack(side="left", padx=(0, S(14)))
        toast.place(relx=1.0, rely=1.0, x=-S(24), y=-S(52), anchor="se")
        toast.lift()
        toast.bind("<Button-1>", lambda e: self._show_toast("", color))
        self._toast_timer = self.after(9000 if action else 5000, lambda: self._show_toast("", color))

    def apply_prices(self):
        self.wash_stepper.set_price(self.cfg["price_washing_pence"])
        self.dry_stepper.set_price(self.cfg["price_dryer_pence"])
        self.update_total()

    def on_tenants_changed(self):
        if self.selected_flat and self.selected_flat not in {t["flat"] for t in self.tenants}:
            self.selected_flat = None
        self._build_flat_grid()
        self.update_total()

    # ---- flat selection --------------------------------------------------
    def _flat_matches(self, tenant, query):
        if not query:
            return True
        q = query.lower()
        return tenant["flat"].lower().startswith(q) or q in tenant["name"].lower()

    def _paint_flat_buttons(self):
        query = self.flat_search.value()
        for t in self.tenants:
            btn = self.flat_buttons.get(t["flat"])
            if btn is None:
                continue
            if t["flat"] == self.selected_flat:
                btn.set_colors(TEAL, TEAL_HOVER, "#ffffff")
            elif self._flat_matches(t, query):
                btn.set_colors(GREY_BTN, GREY_HOVER, TEXT)
            else:
                btn.set_colors(PANEL, PANEL_2, DIM)

    def _select_single_match(self):
        query = self.flat_search.value()
        matches = [t for t in self.tenants if self._flat_matches(t, query)]
        exact = [t for t in matches if t["flat"].lower() == query.lower()]
        pick = exact or matches
        if query and len(pick) == 1:
            self.select_flat(pick[0]["flat"])

    def select_flat(self, flat):
        self.selected_flat = None if flat == self.selected_flat else flat
        self._paint_flat_buttons()
        self._show_tenant(self.selected_flat)
        self.update_total()

    def _show_tenant(self, flat, hover=False):
        if not hover:  # the link belongs to the selected flat, not the one under the mouse
            if flat:
                self.history_link.pack(side="right", before=self.tenant_label)
            else:
                self.history_link.pack_forget()
        if not flat:
            self.tenant_label.config(text="No flat selected", fg=DIM)
            return
        name = self.tenant_name(flat) or "(no name on file)"
        self.tenant_label.config(text=f"Flat {flat}  ·  {name}", fg=MUTED if hover else TEXT)

    # ---- sale form -------------------------------------------------------
    def current_total(self):
        return (self.wash_stepper.qty * self.cfg["price_washing_pence"]
                + self.dry_stepper.qty * self.cfg["price_dryer_pence"])

    def update_total(self):
        if not hasattr(self, "sell_btn"):
            return
        total = self.current_total()
        self.total_label.config(text=fmt_money(total), fg=TEXT if total else DIM)
        self.sell_btn.set_enabled(bool(self.selected_flat) and total > 0)

    def clear_sale(self):
        self.selected_flat = None
        self.flat_search.clear()
        self.wash_stepper.set(0)
        self.dry_stepper.set(0)
        self._paint_flat_buttons()
        self._show_tenant(None)

    # ---- transactions ----------------------------------------------------
    def record_event(self, rtype, flat, name, wash, dry, total, staff_name, ref=None, notes="", parent=None):
        """Append to the CSV first. Returns the record, or None if it could not be saved."""
        rec = {"receipt": self.next_receipt(), "dt": datetime.now().replace(microsecond=0), "type": rtype,
               "flat": flat, "name": name, "wash": wash, "dry": dry, "total": total,
               "staff": staff_name, "ref": ref, "notes": notes, "tampered": False}
        rec["check"] = row_check(self.last_check, rec)
        try:
            append_record(rec)
        except PermissionError:
            messagebox.showerror(
                "Could not save", "sales_log.csv is open in another program (probably Excel).\n\n"
                "Close it and try again. Nothing was recorded and the drawer was not opened.",
                parent=parent or self)
            return None
        except OSError as e:
            messagebox.showerror("Could not save", f"Could not write to sales_log.csv:\n\n{e}\n\n"
                                 "Nothing was recorded and the drawer was not opened.", parent=parent or self)
            return None
        self.last_check = rec["check"]
        self.records.append(rec)
        self._set_log_tip()
        self.refresh_all()
        return rec

    def _open_drawer_after(self, rec, what):
        self.config(cursor="watch")
        self.update_idletasks()
        ok, msg = fire_drawer(self.cfg["com_port"], self.cfg["baud_rate"], self.cfg["drawer_enabled"])
        self.config(cursor="")
        self.update_drawer_status()
        if not ok:
            messagebox.showwarning(
                "Cash drawer did not open",
                f"The {what} WAS recorded (receipt #{rec['receipt']}), but the drawer could not be opened:\n\n"
                f"{msg}\n\nOpen the drawer with its key. Check the USB cable and the COM port in Settings.",
                parent=self)
        return ok

    def _pin_locked_out(self, dlg):
        if dlg.locked_out:
            messagebox.showwarning("Cancelled", "Too many incorrect PIN attempts. Nothing was recorded.", parent=self)
            self.set_message("Cancelled: too many incorrect PIN attempts", RED)

    def sell(self):
        wash, dry = self.wash_stepper.qty, self.dry_stepper.qty
        if not self.selected_flat or wash + dry == 0:
            return
        flat, name = self.selected_flat, self.tenant_name(self.selected_flat)
        total = self.current_total()
        dlg = PinDialog(self, self, "Confirm sale", f"Take {fmt_money(total)}",
                        [f"Flat {flat}  ·  {name or '(no name on file)'}", describe_items(wash, dry)],
                        confirm_text="Confirm & open drawer", cash_due=total)
        result = dlg.show()
        if not result:
            self._pin_locked_out(dlg)
            if not dlg.locked_out:
                self.set_message("Sale cancelled")
            return
        staff, _ = result
        change = dlg.change_pence
        notes = f"Cash given {fmt_money(dlg.given_pence)}, change {fmt_money(change)}" if change is not None else ""
        rec = self.record_event(TYPE_SALE, flat, name, wash, dry, total, staff["name"], notes=notes)
        if rec is None:
            return
        ok = self._open_drawer_after(rec, "sale")
        self.clear_sale()
        self.set_message(f"✓ Sale #{rec['receipt']}: flat {flat}, {describe_items(wash, dry)}, {fmt_money(total)}"
                         + (f"  ·  GIVE {fmt_money(change)} CHANGE" if change else "")
                         + ("" if ok else "  (drawer did not open)"), GREEN if ok else YELLOW)

    def void_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Void", "Select the sale or petty cash entry you want to void in the log first.",
                                parent=self)
            return
        receipt = int(sel[0].split("-")[0])
        rec = next((r for r in self.records if r["receipt"] == receipt), None)
        if rec is None or rec["type"] not in (TYPE_SALE, TYPE_PETTY):
            messagebox.showinfo("Void", "Only sales and petty cash can be voided.", parent=self)
            return
        if receipt in self.voided_receipts():
            messagebox.showinfo("Void", f"#{receipt} has already been voided.", parent=self)
            return
        is_petty = rec["type"] == TYPE_PETTY
        if is_petty:
            amount = fmt_money(-rec["total"])
            heading, lines = f"Void petty cash #{receipt}", [
                f"Petty cash {amount}  ·  {rec['notes']}",
                f"Taken by {rec['staff']} on {rec['dt']:%d/%m/%Y %H:%M}",
                f"Put {amount} back in the drawer. The drawer will open."]
            confirm_text, notes = "Void & open drawer", f"Voids petty cash #{receipt}: "
        else:
            heading, lines = f"Void sale #{receipt}", [
                f"Flat {rec['flat']}  ·  {rec['name'] or '(no name)'}",
                f"{describe_items(rec['wash'], rec['dry'])} on {rec['dt']:%d/%m/%Y %H:%M}",
                f"Refund {fmt_money(rec['total'])} to the tenant. The drawer will open."]
            confirm_text, notes = "Void & refund", f"Voids #{receipt}: "
        dlg = PinDialog(self, self, "Void", heading, lines, admin_only=True, reason_label="Reason",
                        confirm_text=confirm_text, confirm_bg=RED, confirm_hover=RED_HOVER)
        result = dlg.show()
        if not result:
            self._pin_locked_out(dlg)
            return
        staff, reason = result
        void = self.record_event(TYPE_VOID, rec["flat"], rec["name"], -rec["wash"], -rec["dry"], -rec["total"],
                                 staff["name"], ref=receipt, notes=notes + reason)
        if void is None:
            return
        ok = self._open_drawer_after(void, "void")
        if is_petty:
            msg = f"Petty cash #{receipt} voided. Put {fmt_money(-rec['total'])} back in the drawer"
        else:
            msg = f"Sale #{receipt} voided. Refund {fmt_money(rec['total'])}"
        self.set_message(msg + ("" if ok else "  (drawer did not open)"), YELLOW)

    def petty_cash(self):
        dlg = PinDialog(self, self, "Petty cash", "Take petty cash",
                        ["Takes cash out of the drawer for an expense.",
                         "Recorded in the log with the amount, what it was for and your name. "
                         "Keep the shop receipt."],
                        amount_label="Amount (£)", reason_label="What is it for? (e.g. cleaning supplies)",
                        confirm_text="Take cash & open drawer", confirm_bg=PURPLE, confirm_hover=PURPLE_HOVER)
        result = dlg.show()
        if not result:
            self._pin_locked_out(dlg)
            return
        staff, reason = result
        amount = dlg.amount_pence
        rec = self.record_event(TYPE_PETTY, "", "", 0, 0, -amount, staff["name"], notes=reason)
        if rec is None:
            return
        ok = self._open_drawer_after(rec, "petty cash")
        self.set_message(f"Petty cash #{rec['receipt']}: {fmt_money(amount)} taken by {staff['name']} for {reason}"
                         + ("" if ok else "  (drawer did not open)"), PURPLE_HOVER if ok else YELLOW)

    def no_sale(self, reason=""):
        """Opens the drawer without a sale, e.g. to give change for a rent payment: a note goes in
        and the same amount in coins comes out, so the drawer total doesn't change. On the Drawer
        screen, and offered after a cash rent payment (reason filled in)."""
        dlg = PinDialog(self, self, "No sale", "Open drawer (no sale)",
                        ["Opens the cash drawer without a sale, e.g. to give change: put the note in and take "
                         "the same amount out in coins.", "This is recorded in the log with your name."],
                        reason_label="Reason (e.g. change for a £10 note, rent receipt 6201)", reason_value=reason,
                        confirm_text="Open drawer", confirm_bg=AMBER, confirm_hover=AMBER_HOVER)
        result = dlg.show()
        if not result:
            self._pin_locked_out(dlg)
            return
        staff, reason = result
        rec = self.record_event(TYPE_NOSALE, "", "", 0, 0, 0, staff["name"], notes=reason)
        if rec is None:
            return
        ok = self._open_drawer_after(rec, "no-sale")
        self.set_message(f"Drawer opened (no sale) by {staff['name']}" if ok else "No-sale recorded; drawer did not open",
                         MUTED if ok else YELLOW)

    def count_cash(self):
        """Count the drawer (e.g. at the end of the day). Any difference from what the app expects
        is recorded, so the drawer figure matches the real cash from then on."""
        expected = drawer_expected(self.records)
        dlg = PinDialog(self, self, "Count cash", "Count the cash in the drawer",
                        [f"Should be in the drawer: {fmt_money(expected)}",
                         "The drawer opens. Count all the cash, then put it back."],
                        confirm_text="Open drawer", confirm_bg=TEAL, confirm_hover=TEAL_HOVER)
        result = dlg.show()
        if not result:
            self._pin_locked_out(dlg)
            return
        staff = result[0]
        opened = self.record_event(TYPE_NOSALE, "", "", 0, 0, 0, staff["name"], notes="Opened to count the cash")
        if opened is None:
            return
        self._open_drawer_after(opened, "drawer opening")
        res = FieldsDialog(self, "Count cash", "How much cash is in the drawer?", [
            {"key": "counted", "label": "Cash counted (£)"},
            {"key": "note", "label": "Note (needed if it doesn't match), e.g. end of day"},
        ], validate=lambda v: "Enter the cash counted, e.g. 45.50." if not v["counted"] or
            parse_pence(v["counted"]) is None or parse_pence(v["counted"]) < 0 else
            "Please add a note saying why it doesn't match (or count again)." if
            parse_pence(v["counted"]) != expected and not v["note"] else None,
            confirm_text="Record count", sub=f"The app expects {fmt_money(expected)}.").show()
        if not res:
            self.set_message("Count cancelled. Nothing changed", MUTED)
            return
        counted = parse_pence(res["counted"])
        diff = counted - expected
        note = f"Counted {fmt_money(counted)}, expected {fmt_money(expected)}"
        note += f" ({'over' if diff > 0 else 'short'} {fmt_money(abs(diff))})" if diff else " (correct)"
        if self.record_event(TYPE_COUNT, "", "", 0, 0, diff, staff["name"],
                             notes=note + (f" · {res['note']}" if res["note"] else "")) is None:
            return
        self.set_message(f"✓ Drawer counted by {staff['name']}: {fmt_money(counted)}"
                         + (f"  ({'over' if diff > 0 else 'short'} {fmt_money(abs(diff))})" if diff else ", correct"),
                         YELLOW if diff else GREEN)

    def bank_takings(self):
        """Empty the drawer and take the cash to the bank (token takings account)."""
        expected = drawer_expected(self.records)
        last = next((r for r in reversed(self.records) if r["type"] == TYPE_BANKED), None)
        since = f"since it was last emptied on {last['dt']:%d/%m/%Y}" if last else "so far"
        dlg = PinDialog(self, self, "Bank takings", "Empty the drawer for banking",
                        [f"Should be in the drawer: {fmt_money(expected)}",
                         f"Token takings minus petty cash {since}.",
                         "The drawer opens. Take out ALL the cash and count it."],
                        confirm_text="Open drawer", confirm_bg=TEAL, confirm_hover=TEAL_HOVER)
        result = dlg.show()
        if not result:
            self._pin_locked_out(dlg)
            return
        staff = result[0]
        opened = self.record_event(TYPE_NOSALE, "", "", 0, 0, 0, staff["name"], notes="Emptied for banking")
        if opened is None:
            return
        self._open_drawer_after(opened, "drawer opening")

        def validate(v):
            paid = parse_pence(v["paid"])
            if paid is None or paid < 0 or not v["paid"]:
                return "Enter the cash counted, e.g. 45.50 (0 if the drawer was empty)."
            return None

        res = FieldsDialog(self, "Bank takings", "How much cash was in the drawer?", [
            {"key": "paid", "label": "Cash counted and taken out for banking (£)"},
            {"key": "slip", "label": "Paying-in slip reference (optional)"},
        ], validate=validate, confirm_text="Record banking",
            sub=f"The app expects {fmt_money(expected)}. Enter what you actually counted.").show()
        if not res:
            self.set_message("Banking cancelled. Put the cash back in the drawer", YELLOW)
            return
        paid = parse_pence(res["paid"])
        diff = paid - expected
        if diff and not messagebox.askyesno(
                "Bank takings", f"You counted {fmt_money(paid)} but the drawer should have had {fmt_money(expected)}: "
                f"{'over' if diff > 0 else 'short'} by {fmt_money(abs(diff))}.\n\nCount again if you're not sure. "
                f"Record {fmt_money(paid)} as banked, with the difference noted under your name?", parent=self):
            self.set_message("Banking not recorded. Count again, then press Bank takings", YELLOW)
            return
        rec = self.record_event(TYPE_BANKED, "", "", 0, 0, -paid, staff["name"],
                                notes=banking_note(res["slip"], expected, paid))
        if rec is None:
            return
        self.set_message(f"✓ {fmt_money(paid)} taken out for banking by {staff['name']}"
                         + (f"  ({'over' if diff > 0 else 'short'} {fmt_money(abs(diff))})" if diff else ""),
                         YELLOW if diff else GREEN)
        self.print_token_banking(rec["receipt"])

    def print_token_banking(self, receipt, keep=True, parent=None):
        """Opens the banking sheet for the banking in log row `receipt`. keep: save it in banking/."""
        rent = self.rent_module
        b = next(b for b in bankings(self.records) if b["rec"]["receipt"] == receipt)
        path = (rent.token_banking_path(b) if keep else
                os.path.join(tempfile.gettempdir(), f"Tokens_T{receipt}_copy.pdf"))
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            rent.make_token_banking_sheet_pdf(path, b, self.petty_receipts(), self.rent.org())
            rent.open_file(path)
        except Exception as e:
            messagebox.showwarning("Banking sheet", f"Could not create the banking sheet:\n\n{e}\n\n"
                                   "The banking IS recorded. Print it later from Reports → Banking.",
                                   parent=parent or self)

    def open_settings(self):
        dlg = PinDialog(self, self, "Settings", "Settings", ["Admin PIN needed to change settings."],
                        admin_only=True, confirm_text="Open settings", confirm_bg=TEAL, confirm_hover=TEAL_HOVER)
        result = dlg.show()
        if not result:
            self._pin_locked_out(dlg)
            return
        SettingsWindow(self, result[0]).show()
        self.refresh_all()
        self.rent_view.refresh()
        self.key_view.refresh()
        self.post_view.refresh()
        self.visitor_view.refresh()

    def restart_app(self):
        """Starts a fresh copy and closes this one (a theme change needs one)."""
        global _instance_mutex
        if _instance_mutex:  # let the new copy start: only one may run at a time
            ctypes.windll.kernel32.CloseHandle(ctypes.c_void_p(_instance_mutex))
            _instance_mutex = None
        env = dict(os.environ, PYINSTALLER_RESET_ENVIRONMENT="1")  # a new .exe, not a child of this one
        args = [sys.executable] if getattr(sys, "frozen", False) else [sys.executable, os.path.abspath(sys.argv[0])]
        subprocess.Popen(args, env=env, close_fds=True)
        self.destroy()
        os._exit(0)

    def exit_app(self):
        """Closes the app completely. Everything is already saved: each entry is written as it's made."""
        if not messagebox.askyesno("Exit", f"Close {APP_TITLE}?", parent=self):
            return
        self.destroy()
        # Ends this process outright (the .exe's launcher then exits too), so nothing is left running.
        os._exit(0)

    def open_reports(self):
        ReportsWindow(self).show()

    def open_stock(self):
        StockWindow(self).show()

    def open_history(self):
        if self.selected_flat:
            TenantHistoryWindow(self, self, self.selected_flat).show()

    # ---- refresh ---------------------------------------------------------
    def refresh_all(self):
        self.refresh_tree()
        self.refresh_stats()
        self.refresh_warning()
        if hasattr(self, "drawer_view"):
            self.drawer_view.refresh()

    def refresh_tree(self):
        query = self.log_search.value().lower()
        voided = self.voided_receipts()
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        shown = 0
        for rec in reversed(self.records):
            is_voided = rec["type"] in (TYPE_SALE, TYPE_PETTY) and rec["receipt"] in voided
            rtype = f"{rec['type']} (voided)" if is_voided else rec["type"]
            if rec["tampered"]:
                rtype = "⚠ " + rtype
            is_money = rec["type"] not in (TYPE_NOSALE, TYPE_ACCEPTED) + STOCK_TYPES
            has_tokens = rec["type"] in (TYPE_SALE,) + STOCK_TYPES or rec["wash"] or rec["dry"]
            values = (rec["receipt"], rec["dt"].strftime("%d/%m/%Y %H:%M"), rtype, rec["flat"], rec["name"],
                      rec["wash"] if has_tokens else "", rec["dry"] if has_tokens else "",
                      fmt_money(rec["total"]) if is_money else "", rec["staff"], rec["notes"])
            if query and not any(query in str(v).lower() for v in values):
                continue
            tags = ["odd" if shown % 2 else "even"]
            if rec["tampered"]:
                tags.append("tampered")
            elif rec["type"] == TYPE_VOID:
                tags.append("void")
            elif rec["type"] == TYPE_NOSALE:
                tags.append("nosale")
            elif rec["type"] == TYPE_COUNT:
                tags.append("count")
            elif rec["type"] == TYPE_ACCEPTED:
                tags.append("accepted")
            elif rec["type"] in STOCK_TYPES:
                tags.append("stock")
            elif rec["type"] == TYPE_BANKED:
                tags.append("banked")
            elif is_voided:
                tags.append("voided")
            elif rec["type"] == TYPE_PETTY:
                tags.append("petty")
            iid = str(rec["receipt"])
            if self.tree.exists(iid):  # duplicate receipt number in a hand-edited file
                iid = f"{rec['receipt']}-{shown}"
            self.tree.insert("", "end", iid=iid, values=values, tags=tags)
            shown += 1
        self.tree.resort()
        still_there = [iid for iid in selected if self.tree.exists(iid)]
        if still_there:
            self.tree.selection_set(still_there)
        if shown == 0:
            self.empty_label.config(text="No matching transactions." if query else
                                    "No transactions yet. Sales will appear here.")
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")
        else:
            self.empty_label.place_forget()

    def refresh_stats(self):
        today = date.today()
        self.date_label.config(text=today.strftime("%A %d %B %Y"))
        s = summarise([r for r in self.records if r["dt"].date() == today], self.petty_receipts())
        self.stat["wash"][0].config(text=str(s["wash"]))
        self.stat["wash"][1].config(text=f"{fmt_money(s['wash'] * self.cfg['price_washing_pence'])} at current price")
        self.stat["dry"][0].config(text=str(s["dry"]))
        self.stat["dry"][1].config(text=f"{fmt_money(s['dry'] * self.cfg['price_dryer_pence'])} at current price")
        self.stat["total"][0].config(text=fmt_money(s["total"]))
        self.stat["total"][1].config(text=f"In the drawer now: {fmt_money(drawer_expected(self.records))}")
        self.stat["petty"][0].config(text=fmt_money(s["petty"]))
        self.stat["petty"][1].config(text=f"Net cash today {fmt_money(s['total'] - s['petty'])}")
        self.stat["sales"][0].config(text=str(s["sales"] - s["voids"]))
        self.stat["sales"][1].config(text=f"{s['voids']} voided · {s['nosales']} no-sale opens")
        stock, low = token_stock(self.records), self.low_stock_levels()
        self.wash_stepper.set_stock(stock and stock["wash"], low["wash"])
        self.dry_stepper.set_stock(stock and stock["dry"], low["dry"])

    def refresh_warning(self):
        broken = [name for name, problems in (("token", self.log_problems), ("rent", self.rent.problems),
                                              ("key", self.keys.problems), ("post", self.post.problems),
                                              ("visitor", self.visitors.problems))
                  if problems]
        if broken:
            self.warning_label.config(text=f"⚠  The {' and '.join(broken)} log was changed outside the app. Entries "
                                           "highlighted in red can't be trusted. See Settings → Logs", fg=RED)
        elif any(pin_matches(DEFAULT_PIN, s) for s in self.cfg["staff"]):
            self.warning_label.config(text=f"⚠  Default PIN {DEFAULT_PIN} is still in use. Change it in Settings → Staff & PINs",
                                      fg=YELLOW)
        else:
            self.warning_label.config(text="")

    def update_drawer_status(self):
        port = self.cfg["com_port"]
        if not self.cfg["drawer_enabled"]:
            self.drawer_dot.config(fg=DIM)
            self.drawer_status.config(text="Cash drawer switched off in Settings")
        elif port in {p.device for p in available_ports()}:
            self.drawer_dot.config(fg=GREEN)
            self.drawer_status.config(text=f"Cash drawer ready on {port}")
        else:
            self.drawer_dot.config(fg=RED)
            self.drawer_status.config(text=f"Cash drawer trigger not found on {port}. Check USB, or choose the port in Settings")

    def _poll(self):
        self.update_drawer_status()
        if self.date_label.cget("text") != date.today().strftime("%A %d %B %Y"):
            self._new_day()
            self.refresh_all()  # new day
            self.rent_view.refresh()
            self.key_view.refresh()  # keys due yesterday are now overdue
            self.post_view.refresh()
            self.visitor_view.refresh()  # yesterday's visitors still signed in are now flagged
        self._check_idle()
        if self.view == "home":
            self.refresh_home()  # greeting, date and drawer notice stay current
        self.after(4000, self._poll)

    def _new_day(self):
        """The front-desk PC stays on for weeks, so what start-up does is also done every day:
        a dated backup of each log, and each log read back and checked against its codes, so a
        change made outside the app is flagged the same day."""
        backup_log(SALES_PATH, "sales_log")
        try:
            self.records, _, self.last_check = load_records()
        except OSError:
            pass  # e.g. open in Excel right now: checked again tomorrow
        else:
            self.log_problems = self._check_log(False)
        for store in (self.rent, self.keys, self.post, self.visitors):
            store.reverify()

    def _startup_checks(self, skipped, config_warning):
        if config_warning:
            messagebox.showwarning("Settings reset", config_warning, parent=self)
        if skipped:
            messagebox.showwarning(
                "Sales log", f"{skipped} row(s) in sales_log.csv could not be read and are not shown.\n\n"
                "They are still in the file. This usually happens if it was edited in Excel.", parent=self)
        for title, name, store in (("Rent log", "rent_log.csv", self.rent), ("Key log", "key_log.csv", self.keys),
                                   ("Post log", "post_log.csv", self.post),
                                   ("Visitor log", "visitor_log.csv", self.visitors)):
            if store.skipped:
                messagebox.showwarning(
                    title, f"{store.skipped} row(s) in {name} could not be read and are not shown.\n\n"
                    "They are still in the file. This usually happens if it was edited in Excel.", parent=self)
        for name, problems in (("sales_log.csv", self.log_problems), ("rent_log.csv", self.rent.problems),
                               ("key_log.csv", self.keys.problems), ("post_log.csv", self.post.problems),
                               ("visitor_log.csv", self.visitors.problems)):
            if problems:
                messagebox.showwarning(
                    "Log changed outside the app",
                    f"{name} has been changed by something other than this app:\n\n• " + "\n• ".join(problems)
                    + "\n\nThe affected entries are highlighted in red. Totals include them as they are now, so "
                    "they may be wrong.\n\nA copy of each log is saved every day in the backups folder. An admin "
                    "can compare, restore a good copy, or accept the log as it is in Settings → Logs.",
                    parent=self)
        # If the configured port is missing but exactly one USB serial device is plugged in, use it.
        # (Built-in ports such as COM1 are ignored; they are never the trigger box.)
        ports = [p for p in available_ports() if "USB" in (p.hwid or "").upper()]
        if self.cfg["com_port"] not in {p.device for p in available_ports()} and len(ports) == 1:
            self.cfg["com_port"] = ports[0].device
            self.save_cfg()
            self.set_message(f"Cash drawer found on {ports[0].device} ({ports[0].description})", TEAL)
        self.update_drawer_status()
        # Open on the idle screen (once any start-up warnings are dealt with): nothing on show until
        # someone taps to start.
        self.sleep()

    def _on_error(self, exc_type, exc, tb):
        text = "".join(traceback.format_exception(exc_type, exc, tb))
        try:
            with open(ERROR_LOG_PATH, "a", encoding="utf-8") as f:
                f.write(f"\n--- {datetime.now():%Y-%m-%d %H:%M:%S} ---\n{text}")
        except OSError:
            pass
        messagebox.showerror("Unexpected error", f"Something went wrong:\n\n{exc}\n\nDetails were saved to error.log.",
                             parent=self)


_instance_mutex = None


def already_running():
    """Two copies on the same data would clash (receipt numbers, the logs' check chain), so only
    one may run per folder. A named Windows mutex: Windows removes it when the process ends, even
    after a crash, so it can never be left 'stuck' the way a lock file can."""
    global _instance_mutex
    name = "Local\\Front-Office-" + hashlib.sha256(os.path.normcase(BASE_DIR).encode()).hexdigest()[:16]
    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.restype = ctypes.c_void_p
        _instance_mutex = kernel32.CreateMutexW(None, False, name)
        return ctypes.get_last_error() == 183  # ERROR_ALREADY_EXISTS
    except (AttributeError, OSError):
        return False


def main():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # crisp text on high-DPI screens
    except (AttributeError, OSError):
        pass
    if already_running():
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(APP_TITLE, f"{APP_TITLE} is already open.\n\nOnly one copy can run at a time. "
                                       "If you can't see it, it may be behind another window.")
        root.destroy()
        return
    try:
        app = App()
    except Exception as e:  # e.g. data folder not writable, sales_log.csv locked
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(APP_TITLE, f"The app could not start:\n\n{e}\n\nData folder:\n{BASE_DIR}")
        root.destroy()
        raise
    app.mainloop()


if __name__ == "__main__":
    # rent.py does "from front_office import ...": make that find this running module
    # rather than loading a second copy of it.
    sys.modules.setdefault("front_office", sys.modules["__main__"])
    main()
