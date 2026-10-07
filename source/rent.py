"""
Rent & deposits
===============

Cash and cheque payments taken at the office, from receipt to bank:

  1. Taken    Any staff PIN. A numbered receipt (PDF) is saved and opened for printing.
  2. Checked  A staff member marked as a checker (the office manager) counts the money and
              puts it in the safe. Nobody can check a payment they took themselves, so a
              payment the checker took can be checked by any other staff member.
  3. Banked   The payments paid in together become one deposit, with the bank date and the
              paying-in slip reference. The amount paid in must match their total.

A receipt can be cancelled (with a reason) until it is banked. Its number is never reused.

Data files, in data/ next to the .exe:
  rent_log.csv      every receipt, check, banking and cancellation (append-only, with the
                    same Check chain as sales_log.csv)
  rent_tenants.csv  Tenant Name, Address, Tenant Ref (the housing system's Ten Key; household
                    members share their tenancy's key)
and receipts/YYYY/ (next to data/) holds the PDF of every receipt exactly as issued.
banking/YYYY/ holds the banking sheet printed for each deposit, to keep with the paying-in slip.
"""

import csv
import os
import tempfile
import tkinter as tk
from datetime import date, datetime, timedelta
from tkinter import filedialog, messagebox, ttk

from front_office import (
    AMBER, BACKUP_DIR, BASE_DIR, DATA_DIR, BG, BORDER, CSV_DATE_FMT, DIM, FONT, GREEN, GREEN_HOVER, GREY_BTN, GREY_HOVER,
    MUTED, ORANGE, PANEL, PANEL_2, RED, RED_HOVER, ROW_A, ROW_B, TAMPERED_BG, TAMPERED_FG, TEAL, TYPE_COUNT,
    TEAL_HOVER, TEXT, UI_DATE_FMT, YELLOW, LOGO_MARK_WIDTH, Dialog, FieldsDialog, HoverButton, PinDialog, admin_save_path,
    make_sortable, ORANGE_HOVER,
    PlaceholderEntry, append_csv_rows, backup_log, chain_hash, csv_money, entry_opts, find_logo, fmt_money,
    natural_key, parse_csv_datetime, parse_int, parse_pence, parse_ui_date, read_csv_rows, save_config,
    summarise, write_csv_export,
)

try:
    from PIL import Image
    from reportlab.lib.colors import Color, HexColor
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader, simpleSplit
    from reportlab.pdfgen import canvas
except ImportError:  # payments are still recorded; only the PDF receipt is unavailable
    canvas = None


RENT_LOG_PATH = os.path.join(DATA_DIR, "rent_log.csv")
RENT_TENANTS_PATH = os.path.join(DATA_DIR, "rent_tenants.csv")
RECEIPTS_DIR = os.path.join(BASE_DIR, "receipts")
BANKING_DIR = os.path.join(BASE_DIR, "banking")

RENT_HEADERS = [
    "Entry No", "Date & Time", "Type", "Receipt No", "Tenant Name", "Tenant Ref", "Address", "For", "Details",
    "Cash (£)", "Cheque (£)", "Total (£)", "Staff", "Deposit No", "Bank Date", "Paying-in Ref", "Notes", "Check",
]
RENT_TENANT_HEADERS = ["Tenant Name", "Address", "Tenant Ref"]
RENT_SALT = b"front-office/rent-chain/v1"

R_RECEIPT = "RECEIPT"
R_CHECKED = "CHECKED"
R_BANKED = "BANKED"
R_CANCELLED = "CANCELLED"
R_ACCEPTED = "LOG ACCEPTED"
RENT_TYPES = (R_RECEIPT, R_CHECKED, R_BANKED, R_CANCELLED, R_ACCEPTED)

PAYMENT_KINDS = ["Rent", "Arrears", "Rent + arrears", "Key deposit", "Deposit", "Other"]
DEFAULT_FIRST_RECEIPT = 6200  # the paper receipt book ends at 6199
DEFAULT_UNBANKED_DAYS = 7

S_AWAITING, S_IN_SAFE, S_BANKED, S_CANCELLED = "awaiting", "in_safe", "banked", "cancelled"
STATUS_TEXT = {S_AWAITING: "Awaiting check", S_IN_SAFE: "In safe", S_BANKED: "Banked", S_CANCELLED: "Cancelled"}

DEFAULT_ORG = {
    "name1": "YOUR ORGANISATION",
    "name2": "NAME LINE 2",
    "address": "1 Example Street, Town, AB1 2CD",
    "contact": "Tel: 01234 567890    Email: office@example.org",
}


# --------------------------------------------------------------------------
# Amounts in words (for "The sum of" on the receipt)
# --------------------------------------------------------------------------

_ONES = ["", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven",
         "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"]
_TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]


def _words(n):
    """British English: 1,305 -> 'one thousand three hundred and five'."""
    if n < 20:
        return _ONES[n]
    if n < 100:
        return _TENS[n // 10] + ("-" + _ONES[n % 10] if n % 10 else "")
    if n < 1000:
        rest = n % 100
        return _ONES[n // 100] + " hundred" + (" and " + _words(rest) if rest else "")
    for size, name in ((1_000_000, "million"), (1000, "thousand")):
        if n >= size:
            head, rest = divmod(n, size)
            if not rest:
                tail = ""
            elif rest < 100:
                tail = " and " + _words(rest)
            else:
                tail = " " + _words(rest)
            return _words(head) + " " + name + tail
    return str(n)


def amount_in_words(pence):
    pounds, p = divmod(abs(pence), 100)
    pennies = f"{_words(p)} {'penny' if p == 1 else 'pence'}"
    if not pounds and p:
        text = pennies
    else:
        text = f"{_words(pounds) or 'zero'} pound{'' if pounds == 1 else 's'}"
        text += f" and {pennies}" if p else " only"
    return text[0].upper() + text[1:]


# --------------------------------------------------------------------------
# Log: rows <-> entries, integrity chain
# --------------------------------------------------------------------------

def _csv_date(d):
    return d.strftime("%Y-%m-%d") if d else ""


def row_to_entry(row):
    entry = parse_int(row.get("Entry No"), default=None)
    dt = parse_csv_datetime(row.get("Date & Time"))
    rtype = (row.get("Type") or "").strip().upper()
    cash, cheque, total = (parse_pence(row.get(k, "")) for k in ("Cash (£)", "Cheque (£)", "Total (£)"))
    if entry is None or dt is None or rtype not in RENT_TYPES or None in (cash, cheque, total):
        return None
    bank_text = (row.get("Bank Date") or "").strip()
    bank_date = parse_ui_date(bank_text) if bank_text else None
    if bank_text and bank_date is None:
        return None
    return {
        "entry": entry, "dt": dt, "type": rtype,
        "receipt": parse_int(row.get("Receipt No"), default=None),
        "name": (row.get("Tenant Name") or "").strip(),
        "ref": (row.get("Tenant Ref") or "").strip(),
        "address": (row.get("Address") or "").strip(),
        "kind": (row.get("For") or "").strip(),
        "details": (row.get("Details") or "").strip(),
        "cash": cash, "cheque": cheque, "total": total,
        "staff": (row.get("Staff") or "").strip(),
        "deposit": (row.get("Deposit No") or "").strip(),
        "bank_date": bank_date,
        "slip": (row.get("Paying-in Ref") or "").strip(),
        "notes": (row.get("Notes") or "").strip(),
        "check": (row.get("Check") or "").strip(),
        "tampered": False,
    }


def entry_to_row(e):
    return [
        e["entry"], e["dt"].strftime(CSV_DATE_FMT), e["type"], "" if e["receipt"] is None else e["receipt"],
        e["name"], e["ref"], e["address"], e["kind"], e["details"], csv_money(e["cash"]), csv_money(e["cheque"]),
        csv_money(e["total"]), e["staff"], e["deposit"], _csv_date(e["bank_date"]), e["slip"], e["notes"],
        e.get("check", ""),
    ]


def entry_check(prev_check, e):
    """Values, not text, and times to the minute, so an Excel re-save is not a change.
    Excel also drops leading zeros (Ten Key 00174 -> 174), so those are ignored too."""
    return chain_hash(RENT_SALT, prev_check, [
        e["entry"], e["dt"].strftime("%Y-%m-%d %H:%M"), e["type"], "" if e["receipt"] is None else e["receipt"],
        e["name"], e["ref"].lstrip("0"), e["address"], e["kind"], e["details"], e["cash"], e["cheque"], e["total"],
        e["staff"], e["deposit"], _csv_date(e["bank_date"]), e["slip"], e["notes"],
    ])


def ensure_rent_log():
    if not os.path.exists(RENT_LOG_PATH) or os.path.getsize(RENT_LOG_PATH) == 0:
        write_csv_export(RENT_LOG_PATH, RENT_HEADERS, [])


def load_rent_log():
    """Returns (entries oldest first, unreadable rows, Check of the last row)."""
    entries, skipped, prev = [], 0, ""
    if not os.path.exists(RENT_LOG_PATH):
        return entries, skipped, prev
    for row in read_csv_rows(RENT_LOG_PATH)[1]:
        stored = (row.get("Check") or "").strip()
        e = row_to_entry(row)
        if e is None:
            skipped += 1
        else:
            e["tampered"] = stored != entry_check(prev, e)
            entries.append(e)
        prev = stored or prev
    entries.sort(key=lambda e: e["entry"])
    return entries, skipped, prev


def reseal_rent_log():
    rows = read_csv_rows(RENT_LOG_PATH)[1]
    os.makedirs(BACKUP_DIR, exist_ok=True)
    backup = os.path.join(BACKUP_DIR, f"rent_log_before_accept_{datetime.now():%Y-%m-%d_%H%M%S}.csv")
    with open(RENT_LOG_PATH, "rb") as src, open(backup, "wb") as dst:
        dst.write(src.read())
    prev, out = "", []
    for row in rows:
        e = row_to_entry(row)
        if e is None:
            out.append([row.get(h) or "" for h in RENT_HEADERS])
            continue
        e["check"] = prev = entry_check(prev, e)
        out.append(entry_to_row(e))
    tmp = RENT_LOG_PATH + ".tmp"
    write_csv_export(tmp, RENT_HEADERS, out)
    os.replace(tmp, RENT_LOG_PATH)
    return backup


# --------------------------------------------------------------------------
# Tenants
# --------------------------------------------------------------------------

def load_rent_tenants():
    if not os.path.exists(RENT_TENANTS_PATH):
        return []
    tenants, seen = [], set()
    for row in read_csv_rows(RENT_TENANTS_PATH)[1]:
        t = {"name": (row.get("Tenant Name") or "").strip(), "address": (row.get("Address") or "").strip(),
             "ref": (row.get("Tenant Ref") or "").strip()}
        if t["name"] and tenant_key(t) not in seen:
            seen.add(tenant_key(t))
            tenants.append(t)
    tenants.sort(key=lambda t: natural_key(t["name"]))
    return tenants


def tenant_key(t):
    return t["name"].lower(), t["address"].lower()


def save_rent_tenants(tenants):
    tmp = RENT_TENANTS_PATH + ".tmp"
    write_csv_export(tmp, RENT_TENANT_HEADERS, [[t["name"], t["address"], t.get("ref", "")]
                                                for t in sorted(tenants, key=lambda t: natural_key(t["name"]))])
    os.replace(tmp, RENT_TENANTS_PATH)


ADDRESS_WORDS = ("room", "rm", "flat", "unit", "address", "property", "building", "house", "street", "road",
                 "block", "postcode", "scheme")
REF_HEADINGS = ("ten key", "tenant key", "tenant ref", "tenancy ref", "tenancy key", "tenant id", "tenant no",
                "tenant number")


def read_tenant_file(path):
    """Reads tenants from an Excel or CSV file, e.g. the housing system's tenant export.
    Finds the header row (the first with a 'name' column), joins every address-like column
    into one address, and takes the tenant reference from a column like 'Ten Key'.
    Returns [{"name", "address", "ref"}]."""
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        rows = [["" if v is None else str(v).strip() for v in r] for r in wb.active.iter_rows(values_only=True)]
        wb.close()
    else:
        with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
            rows = [[c.strip() for c in r] for r in csv.reader(f)]
    for i, row in enumerate(rows):
        heads = [h.lower() for h in row]
        if any("name" in h for h in heads):
            break
    else:
        raise ValueError("No column with 'name' in its heading was found.")

    def cols(pred):
        return [j for j, h in enumerate(heads) if pred(h)]

    full = cols(lambda h: "name" in h and not any(w in h for w in ("first", "fore", "sur", "last", "given")))
    first = cols(lambda h: any(w in h for w in ("first", "fore", "given")))
    last = cols(lambda h: any(w in h for w in ("surname", "last name", "family")))
    name_cols = full[:1] or (first[:1] + last[:1])
    ref_cols = cols(lambda h: h in REF_HEADINGS)
    # "Property Key" / "Prop Code" are system codes, not part of the address
    addr_cols = [j for j in cols(lambda h: any(w in h for w in ADDRESS_WORDS)
                                 and not any(w in h.split() for w in ("key", "ref", "id", "code", "no")))
                 if j not in name_cols + first + last + ref_cols]
    out = []
    for row in rows[i + 1:]:
        row = row + [""] * (len(heads) - len(row))
        name = " ".join(" ".join(row[j] for j in name_cols if row[j]).split())
        if name:
            out.append({"name": name, "address": ", ".join(row[j] for j in addr_cols if row[j]),
                        "ref": row[ref_cols[0]] if ref_cols else ""})
    return out


# --------------------------------------------------------------------------
# Receipt PDF
# --------------------------------------------------------------------------

INK = "#1d1d22"
RECEIPT_RED = "#b3261e"
BRAND = "#ef8541"


def receipt_path(p):
    return os.path.join(RECEIPTS_DIR, p["dt"].strftime("%Y"), f"Receipt_{p['receipt']}.pdf")


def make_receipt_pdf(path, p, org, stamp=None):
    """Draws a receipt laid out like the paper book, in the top half of an A4 page.
    stamp: None for the original, or e.g. 'COPY' / 'CANCELLED' across it."""
    if canvas is None:
        raise RuntimeError("The PDF library (reportlab) is not available.")
    W, H = A4
    c = canvas.Canvas(path, pagesize=A4)
    c.setTitle(f"Receipt {p['receipt']}")
    c.setAuthor(f"{org['name1']} {org['name2']}")
    left, right = 20 * mm, W - 20 * mm
    top = H - 18 * mm
    ink = HexColor(INK)

    c.setStrokeColor(HexColor("#c9c9cf"))
    c.setLineWidth(0.8)
    c.roundRect(left - 7 * mm, top - 124 * mm, (right - left) + 14 * mm, 132 * mm, 3 * mm)

    # Header: mark, name, address
    x = left
    logo = find_logo()
    if logo and Image is not None:
        try:
            img = Image.open(logo).convert("RGBA")
            mark = img.crop((0, 0, LOGO_MARK_WIDTH, img.height))
            mh = 21 * mm
            mw = mh * mark.width / mark.height
            c.drawImage(ImageReader(mark), left, top - mh, mw, mh, mask="auto")
            x = left + mw + 5 * mm
        except Exception:
            pass  # a receipt without the logo is still a receipt
    c.setFillColor(ink)
    c.setFont("Times-Bold", 23)
    c.drawString(x, top - 8 * mm, org["name1"])
    c.setFont("Times-Bold", 12.5)
    c.drawString(x, top - 13.5 * mm, org["name2"])
    c.setFont("Helvetica", 7.6)
    c.drawString(x, top - 18 * mm, org["address"])
    c.drawString(x, top - 21.5 * mm, org["contact"])

    c.setFont("Times-Bold", 14)
    c.drawRightString(right - 27 * mm, top - 8 * mm, "No.")
    c.setFillColor(HexColor(RECEIPT_RED))
    c.setFont("Helvetica-Bold", 20)
    c.drawRightString(right, top - 8.5 * mm, str(p["receipt"]))
    if p.get("ref"):
        c.setFillColor(ink)
        c.setFont("Helvetica-Bold", 11)
        c.drawRightString(right, top - 15.5 * mm, p["ref"])
        c.setFont("Times-Bold", 9)
        c.drawRightString(right - c.stringWidth(p["ref"], "Helvetica-Bold", 11) - 2.5 * mm, top - 15.5 * mm,
                          "TENANT REF")

    c.setStrokeColor(ink)
    c.setLineWidth(0.9)
    c.line(left, top - 26 * mm, right, top - 26 * mm)
    c.setFillColor(ink)
    c.setFont("Times-Bold", 19)
    c.drawCentredString((left + right) / 2 - 12 * mm, top - 33.5 * mm, "RECEIPT")
    c.setFont("Times-Bold", 11)
    c.drawRightString(right - 26 * mm, top - 33.5 * mm, "DATE:")
    c.setFont("Helvetica", 12)
    c.drawRightString(right, top - 33.5 * mm, p["dt"].strftime("%d/%m/%Y"))
    c.line(left, top - 37 * mm, right, top - 37 * mm)

    def field(y, label, value, value_x=None, until=right, size=12):
        c.setFillColor(ink)
        c.setFont("Times-Bold", 10.5)
        c.drawString(left, y, label)
        vx = value_x or left + c.stringWidth(label, "Times-Bold", 10.5) + 3 * mm
        c.setStrokeColor(HexColor("#8a8a92"))
        c.setLineWidth(0.6)
        c.setDash(0.8, 1.6)
        c.line(vx, y - 1.4 * mm, until, y - 1.4 * mm)
        c.setDash()
        room = until - vx - 2 * mm
        while size > 9 and c.stringWidth(value, "Helvetica", size) > room:
            size -= 0.5  # shrink a long name or address to fit rather than cut it off
        c.setFont("Helvetica", size)
        lines = simpleSplit(value, "Helvetica", size, room) or [""]
        c.drawString(vx + 1.5 * mm, y + 0.3 * mm, lines[0])
        return lines[1:]

    y = top - 46 * mm
    c.setFont("Times-Bold", 10.5)
    c.drawString(left, y, "RECEIVED WITH THANKS FROM")
    y -= 9 * mm
    more = field(y, "MR/MRS/MISS", p["name"])
    if more:  # still too long at the smallest size: finish it on the line below
        y -= 6 * mm
        field(y, "", " ".join(more), value_x=left + 27 * mm)
    y -= 9 * mm
    more = field(y, "OF", p["address"])
    if more:
        y -= 7 * mm
        field(y, "", " ".join(more), value_x=left + 9 * mm)
    y -= 9 * mm
    what = p["kind"] + (f" – {p['details']}" if p["details"] else "")
    field(y, "FOR", what)
    y -= 9 * mm
    field(y, "THE SUM OF", amount_in_words(p["total"]), size=11)
    y -= 10 * mm
    mid = (left + right) / 2 + 8 * mm
    field(y, "CASH £", f"{p['cash'] / 100:,.2f}" if p["cash"] else "—", until=mid - 10 * mm)
    c.setFont("Times-Bold", 10.5)
    c.drawString(mid, y, "CHEQUE £")
    c.setStrokeColor(HexColor("#8a8a92"))
    c.setDash(0.8, 1.6)
    c.line(mid + 20 * mm, y - 1.4 * mm, right, y - 1.4 * mm)
    c.setDash()
    c.setFont("Helvetica", 12)
    c.drawString(mid + 21.5 * mm, y + 0.3 * mm, f"{p['cheque'] / 100:,.2f}" if p["cheque"] else "—")
    y -= 11 * mm
    field(y, "RECEIVED BY", p["taken_by"], until=right)

    # No handwritten signature: the receipt says so, and who issued it and when.
    c.setFillColor(ink)
    c.setFont("Helvetica-Bold", 8)
    c.drawString(left, top - 115.5 * mm, "This is a computer-generated receipt and is valid without a signature.")
    c.setFillColor(HexColor("#6b6b73"))
    c.setFont("Helvetica", 7.5)
    c.drawString(left, top - 119.5 * mm, f"Issued by {p['taken_by']} on {p['dt']:%d/%m/%Y at %H:%M}. "
                                         "Please keep this receipt for your records.")

    if stamp:
        c.saveState()
        c.translate(W / 2, top - 62 * mm)
        c.rotate(18)
        red = stamp == "CANCELLED"
        c.setFillColor(Color(0.70, 0.15, 0.12) if red else Color(0.45, 0.45, 0.5))
        c.setFillAlpha(0.22)
        c.setFont("Helvetica-Bold", 64)
        c.drawCentredString(0, -10, stamp)
        c.restoreState()

    # cut line across the middle of the page
    c.setStrokeColor(HexColor("#b0b0b8"))
    c.setDash(4, 3)
    c.line(10 * mm, H / 2, W - 10 * mm, H / 2)
    c.setDash()
    c.setFillColor(HexColor("#9a9aa2"))
    c.setFont("Helvetica", 7)
    c.drawString(12 * mm, H / 2 + 1.5 * mm, "cut here")
    c.showPage()
    c.save()


# --------------------------------------------------------------------------
# A4 sheets: banking sheet and tenant statement
# --------------------------------------------------------------------------

SHEET_LEFT = 18  # mm margins; tables are 174 mm wide
SHEET_WIDTH = 174


def pdf_money(pence):
    return f"{pence / 100:,.2f}" if pence else "—"


def _letterhead(c, org, title, info):
    """Logo, organisation, and the title with label/value pairs on the right. Returns y below it."""
    W, H = A4
    left, right = SHEET_LEFT * mm, (SHEET_LEFT + SHEET_WIDTH) * mm
    top = H - 16 * mm
    ink = HexColor(INK)
    x = left
    logo = find_logo()
    if logo and Image is not None:
        try:
            img = Image.open(logo).convert("RGBA")
            mark = img.crop((0, 0, LOGO_MARK_WIDTH, img.height))
            mh = 16 * mm
            mw = mh * mark.width / mark.height
            c.drawImage(ImageReader(mark), left, top - mh, mw, mh, mask="auto")
            x = left + mw + 4 * mm
        except Exception:
            pass
    c.setFillColor(ink)
    c.setFont("Times-Bold", 18)
    c.drawString(x, top - 6 * mm, org["name1"])
    c.setFont("Times-Bold", 10)
    c.drawString(x, top - 10.5 * mm, org["name2"])
    c.setFont("Helvetica", 6.8)
    c.drawString(x, top - 14 * mm, org["address"])
    c.setFont("Helvetica-Bold", 15)
    c.drawRightString(right, top - 6 * mm, title)
    y = top - 11.5 * mm
    for label, value in info:
        c.setFont("Helvetica-Bold", 9.5)
        c.setFillColor(ink)
        c.drawRightString(right, y, value)
        c.setFont("Helvetica", 8.5)
        c.setFillColor(HexColor("#6b6b73"))
        c.drawRightString(right - c.stringWidth(value, "Helvetica-Bold", 9.5) - 2 * mm, y, label)
        y -= 4.6 * mm
    y = min(y + 1 * mm, top - 19 * mm) - 2 * mm
    c.setStrokeColor(ink)
    c.setLineWidth(0.9)
    c.line(left, y, right, y)
    return y - 8 * mm


def _table(c, y, columns, rows, total=None, new_page=None, struck=()):
    """columns: [(heading, width mm, 'l'|'r')]. Carries on over new pages: new_page() finishes
    the page and returns the y to continue from. struck: row indexes drawn crossed out.
    Returns the y below the table."""
    left = SHEET_LEFT * mm
    width = sum(w for _, w, _ in columns) * mm
    ink = HexColor(INK)

    def cells(y, values, font, size=8.5):
        c.setFont(font, size)
        x = left
        for (_, w, align), text in zip(columns, values):
            text = "" if text is None else str(text)
            room = w * mm - 3 * mm
            while len(text) > 1 and c.stringWidth(text, font, size) > room:
                text = text[:-2] + "…"
            if align == "r":
                c.drawRightString(x + w * mm - 1.5 * mm, y, text)
            else:
                c.drawString(x + 1.5 * mm, y, text)
            x += w * mm

    def heading(y):
        c.setFillColor(HexColor("#ececf0"))
        c.rect(left, y - 2.2 * mm, width, 6.6 * mm, stroke=0, fill=1)
        c.setFillColor(ink)
        cells(y, [h for h, _, _ in columns], "Helvetica-Bold", 8)
        return y - 7 * mm

    y = heading(y)
    for i, row in enumerate(rows):
        if y < 28 * mm and new_page:
            y = heading(new_page())
        if i % 2:
            c.setFillColor(HexColor("#f6f6f8"))
            c.rect(left, y - 2 * mm, width, 5.6 * mm, stroke=0, fill=1)
        gone = i in struck
        c.setFillColor(HexColor("#9a9aa2") if gone else ink)
        cells(y, row, "Helvetica")
        if gone:
            c.setStrokeColor(HexColor("#9a9aa2"))
            c.setLineWidth(0.5)
            c.line(left + 1.5 * mm, y + 1 * mm, left + width - 1.5 * mm, y + 1 * mm)
        y -= 5.6 * mm
    if total:
        c.setStrokeColor(ink)
        c.setLineWidth(0.8)
        c.line(left, y + 3.4 * mm, left + width, y + 3.4 * mm)
        c.setFillColor(ink)
        cells(y - 1 * mm, total, "Helvetica-Bold", 9)
        y -= 7 * mm
    return y


def _sheet(path, title, doc_title, info, footer_text, org):
    """Starts an A4 sheet. Returns (canvas, new_page, footer): new_page() ends the current page
    (with its footer) and starts the next, returning the y to continue from."""
    W, H = A4
    c = canvas.Canvas(path, pagesize=A4)
    c.setTitle(doc_title)
    c.setAuthor(f"{org['name1']} {org['name2']}")

    def footer():
        c.setFont("Helvetica", 7)
        c.setFillColor(HexColor("#6b6b73"))
        c.drawString(SHEET_LEFT * mm, 12 * mm, footer_text)
        c.drawRightString((SHEET_LEFT + SHEET_WIDTH) * mm, 12 * mm, f"Page {c.getPageNumber()}")

    def new_page():
        footer()
        c.showPage()
        return _letterhead(c, org, title, info)

    return c, new_page, footer


def banking_sheet_path(d):
    return os.path.join(BANKING_DIR, d["bank_date"].strftime("%Y"), f"Banking_{d['deposit']}.pdf")


def make_banking_sheet_pdf(path, d, pays, org):
    """One page (or more) listing every receipt in a deposit, to keep with the paying-in slip."""
    if canvas is None:
        raise RuntimeError("The PDF library (reportlab) is not available.")
    info = [("Deposit", d["deposit"]), ("Paid in", d["bank_date"].strftime("%d/%m/%Y")),
            ("Paying-in ref", d["slip"] or "—"), ("Recorded by", d["banked_by"])]
    c, new_page, footer = _sheet(path, "BANKING SHEET", f"Banking sheet {d['deposit']}", info,
                         f"Banking sheet {d['deposit']} · printed {datetime.now():%d/%m/%Y %H:%M} "
                         "from the Front Office rent log", org)
    y = _letterhead(c, org, "BANKING SHEET", info)
    columns = [("Receipt", 16, "l"), ("Taken", 17, "l"), ("Tenant", 41, "l"), ("Ref", 20, "l"), ("For", 29, "l"),
               ("Cash £", 17, "r"), ("Cheque £", 17, "r"), ("Total £", 17, "r")]
    rows = [[p["receipt"], p["dt"].strftime("%d/%m/%y"), p["name"], p["ref"] or "new", p["kind"],
             pdf_money(p["cash"]), pdf_money(p["cheque"]), pdf_money(p["total"])] for p in pays]
    y = _table(c, y, columns, rows, new_page=new_page,
               total=["", "", f"{len(pays)} receipt{'s' if len(pays) != 1 else ''}", "", "TOTAL",
                      pdf_money(d["cash"]), pdf_money(d["cheque"]), pdf_money(d["total"])])

    if y < 70 * mm:
        y = new_page()
    ink = HexColor(INK)
    left = SHEET_LEFT * mm
    cheques = sum(1 for p in pays if p["cheque"])
    y -= 6 * mm
    for label, value, bold in (("Cash", fmt_money(d["cash"]), False),
                               (f"Cheques ({cheques})", fmt_money(d["cheque"]), False),
                               ("Total paid in", fmt_money(d["total"]), True)):
        c.setFillColor(ink)
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 11 if bold else 10)
        c.drawString(left, y, label)
        c.drawRightString(left + 75 * mm, y, value)
        y -= 6 * mm
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(left, y, amount_in_words(d["total"]))
    y -= 16 * mm
    c.setFont("Helvetica", 9.5)
    c.setStrokeColor(HexColor("#8a8a92"))
    c.setLineWidth(0.6)
    for label in ("Paid in at the bank by", "Checked against the paying-in slip by"):
        c.drawString(left, y, label)
        start = left + c.stringWidth(label, "Helvetica", 9.5) + 3 * mm
        c.line(start, y - 1 * mm, left + 118 * mm, y - 1 * mm)
        c.drawString(left + 124 * mm, y, "Date")
        c.line(left + 133 * mm, y - 1 * mm, left + SHEET_WIDTH * mm, y - 1 * mm)
        y -= 13 * mm
    footer()
    c.showPage()
    c.save()


def token_banking_path(b):
    return os.path.join(BANKING_DIR, b["rec"]["dt"].strftime("%Y"), f"Tokens_T{b['rec']['receipt']}.pdf")


def make_token_banking_sheet_pdf(path, b, petty_receipts, org):
    """Laundry token takings paid in: day by day since the drawer was last emptied, what should
    have been in it, and what was paid in. b: one entry from front_office.bankings()."""
    if canvas is None:
        raise RuntimeError("The PDF library (reportlab) is not available.")
    rec = b["rec"]
    ref = f"T{rec['receipt']}"
    info = [("Banking ref", ref), ("Paid in", rec["dt"].strftime("%d/%m/%Y")), ("Paying-in ref", b["slip"] or "—"),
            ("Banked by", rec["staff"])]
    c, new_page, footer = _sheet(path, "TOKEN BANKING SHEET", f"Token banking {ref}", info,
                                 f"Token banking {ref} · printed {datetime.now():%d/%m/%Y %H:%M} "
                                 "from the Front Office token log", org)
    y = _letterhead(c, org, "TOKEN BANKING SHEET", info)
    ink = HexColor(INK)
    left = SHEET_LEFT * mm
    rows = b["rows"]
    c.setFillColor(ink)
    c.setFont("Helvetica-Bold", 10.5)
    c.drawString(left, y, "Laundry token takings")
    c.setFont("Helvetica", 9)
    c.setFillColor(HexColor("#6b6b73"))
    period = (f"{rows[0]['dt']:%d/%m/%Y} to {rows[-1]['dt']:%d/%m/%Y}" if rows else "no transactions")
    c.drawString(left, y - 5 * mm, f"Everything since the drawer was last emptied: {period}")
    y -= 14 * mm

    s = summarise(rows, petty_receipts)
    columns = [("Date", 34, "l"), ("Washing", 22, "r"), ("Dryer", 22, "r"), ("Sales", 22, "r"),
               ("Takings £", 26, "r"), ("Petty cash £", 24, "r"), ("Net £", 24, "r")]
    table = [[d.strftime("%a %d/%m/%Y"), w, dr, n, pdf_money(t), pdf_money(pc), pdf_money(t - pc)]
             for d, (w, dr, t, n, pc) in sorted(s["by_day"].items())]
    y = _table(c, y, columns, table, new_page=new_page,
               total=["TOTAL", s["wash"], s["dry"], s["sales"], pdf_money(s["total"]), pdf_money(s["petty"]),
                      pdf_money(s["total"] - s["petty"])])

    if y < 75 * mm:
        y = new_page()
    y -= 6 * mm
    diff = b["difference"]
    counts = sum(r["total"] for r in rows if r["type"] == TYPE_COUNT)
    lines = [("Drawer count corrections (see the log)", ("+" if counts > 0 else "") + fmt_money(counts), False)] \
        if counts else []
    lines += [("Should have been in the drawer", fmt_money(b["expected"]), False),
              ("Counted and paid in", fmt_money(b["paid"]), True)]
    if diff:
        lines.append(("Over" if diff > 0 else "Short", fmt_money(abs(diff)), True))
    for label, value, bold in lines:
        c.setFillColor(HexColor(RECEIPT_RED) if label in ("Over", "Short") else ink)
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 11 if bold else 10)
        c.drawString(left, y, label)
        c.drawRightString(left + 90 * mm, y, value)
        y -= 6 * mm
    c.setFillColor(ink)
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(left, y, amount_in_words(b["paid"]))
    y -= 16 * mm
    c.setFont("Helvetica", 9.5)
    c.setStrokeColor(HexColor("#8a8a92"))
    c.setLineWidth(0.6)
    for label in ("Paid in at the bank by", "Checked against the paying-in slip by"):
        c.drawString(left, y, label)
        start = left + c.stringWidth(label, "Helvetica", 9.5) + 3 * mm
        c.line(start, y - 1 * mm, left + 118 * mm, y - 1 * mm)
        c.drawString(left + 124 * mm, y, "Date")
        c.line(left + 133 * mm, y - 1 * mm, left + SHEET_WIDTH * mm, y - 1 * mm)
        y -= 13 * mm
    footer()
    c.showPage()
    c.save()


def make_statement_pdf(path, tenant, pays, period, org):
    """Payments one tenancy made at the office. tenant: {"name", "address", "ref"}.
    Cancelled receipts are listed crossed out and not counted."""
    if canvas is None:
        raise RuntimeError("The PDF library (reportlab) is not available.")
    info = [("Tenant ref", tenant["ref"] or "—"), ("Period", period), ("Date", date.today().strftime("%d/%m/%Y"))]
    c, new_page, footer = _sheet(path, "STATEMENT OF PAYMENTS", f"Statement {tenant['name']}", info,
                         f"Printed {datetime.now():%d/%m/%Y %H:%M} from the Front Office rent log", org)
    y = _letterhead(c, org, "STATEMENT OF PAYMENTS", info)
    ink = HexColor(INK)
    left = SHEET_LEFT * mm
    for label, value in (("Tenant", tenant["name"]), ("Address", tenant["address"])):
        c.setFillColor(HexColor("#6b6b73"))
        c.setFont("Helvetica", 9)
        c.drawString(left, y, label)
        c.setFillColor(ink)
        c.setFont("Helvetica-Bold", 10.5)
        c.drawString(left + 20 * mm, y, value)
        y -= 6 * mm
    y -= 5 * mm
    columns = [("Receipt", 18, "l"), ("Date", 21, "l"), ("For", 76, "l"), ("Paid by", 36, "l"), ("Amount £", 23, "r")]
    rows, struck = [], set()
    for i, p in enumerate(pays):
        what = p["kind"] + (f" – {p['details']}" if p["details"] else "")
        if p["cancelled_at"]:
            what += " (cancelled)"
            struck.add(i)
        rows.append([p["receipt"], p["dt"].strftime("%d/%m/%Y"), what,
                     " + ".join(m for m, v in (("Cash", p["cash"]), ("Cheque", p["cheque"])) if v),
                     pdf_money(p["total"])])
    live = [p for p in pays if not p["cancelled_at"]]
    y = _table(c, y, columns, rows, new_page=new_page, struck=struck,
               total=["", "", f"TOTAL · {len(live)} payment{'s' if len(live) != 1 else ''}", "",
                      pdf_money(sum(p["total"] for p in live))])
    if y < 40 * mm:
        y = new_page()
    y -= 6 * mm
    c.setFillColor(HexColor("#6b6b73"))
    c.setFont("Helvetica", 8)
    for line in simpleSplit("This statement lists payments received at our office by cash or cheque only. It does "
                            "not include payments made any other way (e.g. bank transfer, standing order or housing "
                            "benefit) and is not a statement of your rent account balance.",
                            "Helvetica", 8, SHEET_WIDTH * mm):
        c.drawString(left, y, line)
        y -= 4 * mm
    footer()
    c.showPage()
    c.save()


def write_xlsx(path, headers, rows, money_cols=()):
    """Excel export. Text stays text (so Ten Key 00174 keeps its zeros), amounts are numbers
    formatted as £, dates are real dates."""
    import openpyxl
    from openpyxl.styles import Font
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(row)
    for j, head in enumerate(headers, start=1):
        letter = openpyxl.utils.get_column_letter(j)
        for (cell,) in ws.iter_rows(min_row=2, min_col=j, max_col=j):
            if head in money_cols:
                cell.number_format = '"£"#,##0.00'
            elif isinstance(cell.value, date):
                cell.number_format = "dd/mm/yyyy"
        width = max([len(str(head)), 8] + [len(str(c.value)) for (c,) in
                                            ws.iter_rows(min_row=2, min_col=j, max_col=j) if c.value is not None])
        ws.column_dimensions[letter].width = min(width + 2, 50)
    ws.freeze_panes = "A2"
    wb.save(path)


def open_file(path):
    os.startfile(path)  # opens in the default PDF viewer, ready to print


# --------------------------------------------------------------------------
# Store: the log plus derived payments
# --------------------------------------------------------------------------

class RentStore:
    def __init__(self, app):
        self.app = app
        ensure_rent_log()
        backup_log(RENT_LOG_PATH, "rent_log")
        self.entries, self.skipped, self.last_check = load_rent_log()
        self.tenants = load_rent_tenants()
        self._payments = None
        self.problems = self._check()

    # ---- integrity -------------------------------------------------------
    def _check(self):
        problems = []
        altered = [e["entry"] for e in self.entries if e["tampered"]]
        if altered:
            shown = ", ".join(f"#{n}" for n in altered[:12]) + (" …" if len(altered) > 12 else "")
            problems.append(f"{len(altered)} rent log entr{'y was' if len(altered) == 1 else 'ies were'} changed, "
                            f"added or removed outside the app: {shown}")
        tip = self.app.cfg.get("rent_log_tip")
        if tip and not any(e["entry"] == tip["entry"] and e["check"] == tip["check"] for e in self.entries):
            problems.append(f"Rent log entry #{tip['entry']}, the last one the app wrote, is missing or changed. "
                            "Entries may have been deleted from the end of the rent log.")
        if not tip:
            self._set_tip()
        return problems

    def reverify(self):
        """Daily, on a PC left running: a dated backup, then the log read back and checked again."""
        backup_log(RENT_LOG_PATH, "rent_log")
        try:
            self.entries, _, self.last_check = load_rent_log()
        except OSError:
            return
        self._payments = None
        self.problems = self._check()

    def _set_tip(self):
        last = self.entries[-1] if self.entries else None
        self.app.cfg["rent_log_tip"] = {"entry": last["entry"], "check": last["check"]} if last else None
        try:
            save_config(self.app.cfg)
        except OSError:
            pass

    def accept(self, admin_name, reason, parent):
        try:
            backup = reseal_rent_log()
        except OSError as e:
            messagebox.showerror("Rent log", f"Could not update rent_log.csv (is it open in Excel?):\n\n{e}",
                                 parent=parent)
            return False
        self.entries, _, self.last_check = load_rent_log()
        self._payments = None
        self.problems = []
        self._set_tip()
        self.write([dict(type=R_ACCEPTED, staff=admin_name,
                         notes=f"{reason}. Flagged log kept as backups\\{os.path.basename(backup)}")], parent)
        return True

    # ---- writing ---------------------------------------------------------
    def write(self, items, parent):
        """Appends entries (dicts of fields). All or nothing. Returns the entries, or None."""
        now = datetime.now().replace(microsecond=0)
        next_no = (self.entries[-1]["entry"] + 1) if self.entries else 1
        prev, new = self.last_check, []
        for i, item in enumerate(items):
            e = {"entry": next_no + i, "dt": now, "receipt": None, "name": "", "ref": "", "address": "", "kind": "",
                 "details": "", "cash": 0, "cheque": 0, "total": 0, "staff": "", "deposit": "",
                 "bank_date": None, "slip": "", "notes": "", "tampered": False}
            e.update(item)
            e["check"] = prev = entry_check(prev, e)
            new.append(e)
        try:
            append_csv_rows(RENT_LOG_PATH, RENT_HEADERS, [entry_to_row(e) for e in new])
        except PermissionError:
            messagebox.showerror("Could not save", "rent_log.csv is open in another program (probably Excel).\n\n"
                                 "Close it and try again. Nothing was recorded.", parent=parent)
            return None
        except OSError as err:
            messagebox.showerror("Could not save", f"Could not write to rent_log.csv:\n\n{err}\n\n"
                                 "Nothing was recorded.", parent=parent)
            return None
        self.entries.extend(new)
        self.last_check = prev
        self._payments = None
        self._set_tip()
        return new

    # ---- reading ---------------------------------------------------------
    def payments(self):
        """receipt number -> payment, built from the log's entries."""
        if self._payments is not None:
            return self._payments
        pays = {}
        for e in self.entries:
            if e["type"] == R_RECEIPT and e["receipt"] is not None:
                pays[e["receipt"]] = {
                    "receipt": e["receipt"], "dt": e["dt"], "name": e["name"], "ref": e["ref"],
                    "address": e["address"],
                    "kind": e["kind"], "details": e["details"], "cash": e["cash"], "cheque": e["cheque"],
                    "total": e["total"], "taken_by": e["staff"], "notes": e["notes"],
                    "checked_by": "", "checked_at": None, "banked_by": "", "banked_at": None, "deposit": "",
                    "bank_date": None, "slip": "", "cancelled_by": "", "cancelled_at": None, "cancel_reason": "",
                    "tampered": e["tampered"],
                }
        for e in self.entries:
            p = pays.get(e["receipt"])
            if p is None or e["type"] == R_RECEIPT:
                continue
            p["tampered"] = p["tampered"] or e["tampered"]
            if e["type"] == R_CHECKED:
                p["checked_by"], p["checked_at"] = e["staff"], e["dt"]
            elif e["type"] == R_BANKED:
                p.update(banked_by=e["staff"], banked_at=e["dt"], deposit=e["deposit"], bank_date=e["bank_date"],
                         slip=e["slip"])
            elif e["type"] == R_CANCELLED:
                p.update(cancelled_by=e["staff"], cancelled_at=e["dt"], cancel_reason=e["notes"])
        self._payments = pays
        return pays

    @staticmethod
    def status(p):
        if p["cancelled_at"]:
            return S_CANCELLED
        if p["banked_at"]:
            return S_BANKED
        if p["checked_at"]:
            return S_IN_SAFE
        return S_AWAITING

    def next_receipt(self):
        first = int(self.app.cfg.get("rent_first_receipt", DEFAULT_FIRST_RECEIPT))
        used = [e["receipt"] for e in self.entries if e["receipt"] is not None]
        return max([first] + [n + 1 for n in used])

    def next_deposit(self):
        used = [parse_int(e["deposit"].lstrip("D"), 0) for e in self.entries if e["type"] == R_BANKED]
        return f"D{max(used + [0]) + 1:04d}"

    def deposits(self):
        """deposit no -> summary, newest first."""
        deps = {}
        for p in self.payments().values():
            if p["deposit"] and not p["cancelled_at"]:
                d = deps.setdefault(p["deposit"], {"deposit": p["deposit"], "bank_date": p["bank_date"],
                                                   "slip": p["slip"], "banked_by": p["banked_by"],
                                                   "banked_at": p["banked_at"], "receipts": [], "cash": 0,
                                                   "cheque": 0, "total": 0})
                d["receipts"].append(p["receipt"])
                for k in ("cash", "cheque", "total"):
                    d[k] += p[k]
        return dict(sorted(deps.items(), key=lambda kv: kv[0], reverse=True))

    def payments_for(self, tenant):
        """A tenancy's payments, oldest first: everything under its Tenant Ref (household members
        share it), plus anything taken under the same name and address before it had a ref."""
        ref = tenant["ref"].lstrip("0")
        key = tenant_key(tenant)
        return sorted((p for p in self.payments().values()
                       if (ref and p["ref"].lstrip("0") == ref) or tenant_key(p) == key),
                      key=lambda p: p["receipt"])

    def org(self):
        org = dict(DEFAULT_ORG)
        org.update(self.app.cfg.get("receipt_org") or {})
        return org

    def unbanked_days(self):
        return int(self.app.cfg.get("rent_unbanked_days", DEFAULT_UNBANKED_DAYS))


def is_checker(staff):
    return bool(staff.get("checker"))


def reprint_receipt(p, org, parent):
    """Opens a copy of a receipt, stamped COPY (or CANCELLED)."""
    stamp = "CANCELLED" if p["cancelled_at"] else "COPY"
    path = os.path.join(tempfile.gettempdir(), f"Receipt_{p['receipt']}_{stamp.lower()}.pdf")
    try:
        make_receipt_pdf(path, p, org, stamp=stamp)
        open_file(path)
    except Exception as e:
        messagebox.showerror("Reprint", f"Could not create the receipt:\n\n{e}", parent=parent)


def print_banking_sheet(store, deposit, parent, keep=True):
    """Opens the banking sheet for a deposit. keep: save it in banking/ (when first banked);
    otherwise it's a reprint made in the temp folder."""
    d = store.deposits()[deposit]
    pays = [store.payments()[n] for n in sorted(d["receipts"])]
    path = (banking_sheet_path(d) if keep else
            os.path.join(tempfile.gettempdir(), f"Banking_{d['deposit']}_copy.pdf"))
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        make_banking_sheet_pdf(path, d, pays, store.org())
        open_file(path)
    except Exception as e:
        messagebox.showwarning("Banking sheet", (f"Deposit {deposit} IS recorded, but the banking sheet could not "
                                                 f"be created or opened:\n\n{e}\n\nYou can print it later from "
                                                 "Reports → Deposits.") if keep else
                               f"Could not create the banking sheet:\n\n{e}", parent=parent)


# --------------------------------------------------------------------------
# Main view (the "Rent & deposits" tab)
# --------------------------------------------------------------------------

class RentView(tk.Frame):
    MATCH_ROWS = 5

    def __init__(self, master, app):
        super().__init__(master, bg=BG)
        self.app = app
        self.store = app.rent
        self.kind = None
        self.kind_buttons = {}
        self.filter = "open"
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self._build_form()
        self._build_main()
        self.clear_form()
        self.refresh()

    # ---- left: take a payment --------------------------------------------
    def _section(self, parent, number, title, top=16):
        row = tk.Frame(parent, bg=PANEL)
        row.pack(fill="x", pady=(top, 8))
        tk.Label(row, text=number, font=(FONT, 9, "bold"), fg=BG, bg=MUTED, width=2).pack(side="left")
        tk.Label(row, text=title.upper(), font=(FONT, 10, "bold"), fg=MUTED, bg=PANEL).pack(side="left", padx=8)
        return row

    def _build_form(self):
        S = self.app.S
        side = tk.Frame(self, bg=PANEL, width=S(420))
        side.grid(row=0, column=0, sticky="ns")
        side.pack_propagate(False)
        p = tk.Frame(side, bg=PANEL)
        p.pack(fill="both", expand=True, padx=S(22), pady=S(14))

        head = self._section(p, "1", "Tenant", top=0)
        clear = tk.Label(head, text="Clear", font=(FONT, 10, "underline"), fg=MUTED, bg=PANEL, cursor="hand2")
        clear.pack(side="right")
        clear.bind("<Button-1>", lambda e: self.clear_form())
        self.name_entry = PlaceholderEntry(p, "Name: type to search, or a new tenant's name")
        self.name_entry.pack(fill="x", ipady=5)
        self.name_entry.bind("<KeyRelease>", lambda e: self._update_matches())
        self.name_entry.bind("<Down>", lambda e: self._focus_matches())
        self.matches = tk.Listbox(p, height=self.MATCH_ROWS, bg=PANEL_2, fg=TEXT, font=(FONT, 10), relief="flat",
                                  highlightthickness=0, bd=0, selectbackground=TEAL, selectforeground="#ffffff",
                                  activestyle="none")
        self.matches.pack(fill="x", pady=(2, 0))
        self.matches.bind("<<ListboxSelect>>", lambda e: self._pick_match())
        self.matches.bind("<Return>", lambda e: self._pick_match(move_on=True))
        self.address_entry = PlaceholderEntry(p, "Address, e.g. Rm 3, 91 Wellesley Road")
        self.address_entry.pack(fill="x", ipady=5, pady=(8, 0))
        self.address_entry.bind("<KeyRelease>", lambda e: self._update_new_tenant())
        status = tk.Frame(p, bg=PANEL)
        status.pack(fill="x", pady=(4, 0))
        self.history_link = tk.Label(status, text="History", font=(FONT, 10, "underline"), fg=TEAL, bg=PANEL,
                                     cursor="hand2")
        self.history_link.bind("<Button-1>", lambda e: self.open_history())
        self.ref_label = tk.Label(status, text="", font=(FONT, 10, "bold"), fg=TEAL, bg=PANEL)
        self.ref_label.pack(side="right")
        self.add_tenant_var = tk.BooleanVar(value=True)
        self.add_tenant_check = tk.Checkbutton(
            status, text="New tenant: add to the tenant list", variable=self.add_tenant_var, font=(FONT, 10),
            fg=MUTED, bg=PANEL, selectcolor=PANEL_2, activebackground=PANEL, activeforeground=TEXT, anchor="w",
            bd=0, highlightthickness=0)
        self.add_tenant_check.pack(side="left")

        self._section(p, "2", "Payment for")
        grid = tk.Frame(p, bg=PANEL)
        grid.pack(fill="x")
        for i, kind in enumerate(PAYMENT_KINDS):
            grid.columnconfigure(i % 3, weight=1, uniform="kind")
            btn = HoverButton(grid, text=kind, font=(FONT, 10, "bold"), padx=0, pady=6,
                              command=lambda k=kind: self.set_kind(k))
            btn.grid(row=i // 3, column=i % 3, sticky="nsew", padx=2, pady=2)
            self.kind_buttons[kind] = btn
        self.details_entry = PlaceholderEntry(p, "Details (optional), e.g. October rent, cheque no.")
        self.details_entry.pack(fill="x", ipady=5, pady=(8, 0))

        self._section(p, "3", "Amount")
        amounts = tk.Frame(p, bg=PANEL)
        amounts.pack(fill="x")
        amounts.columnconfigure(1, weight=1)
        amounts.columnconfigure(3, weight=1)
        vcmd = (self.register(lambda s: all(ch in "0123456789.,£" for ch in s)), "%P")
        self.cash_entry = self._amount_box(amounts, "Cash £", 0, vcmd)
        self.cheque_entry = self._amount_box(amounts, "Cheque £", 2, vcmd)

        total_row = tk.Frame(p, bg=PANEL)
        total_row.pack(fill="x", pady=(12, 8))
        tk.Label(total_row, text="Total", font=(FONT, 13), fg=MUTED, bg=PANEL).pack(side="left", anchor="s")
        self.total_label = tk.Label(total_row, text="£0.00", font=(FONT, 26, "bold"), fg=DIM, bg=PANEL)
        self.total_label.pack(side="right")
        self.problem_label = tk.Label(p, text="", font=(FONT, 9), fg=MUTED, bg=PANEL, anchor="w")
        self.problem_label.pack(fill="x")
        self.take_btn = HoverButton(p, text="Take payment  ·  Print receipt", bg=GREEN, hover=GREEN_HOVER,
                                    font=(FONT, 13, "bold"), pady=11, command=self.take_payment)
        self.take_btn.pack(fill="x", pady=(4, 0))

    def _amount_box(self, parent, label, col, vcmd):
        tk.Label(parent, text=label, font=(FONT, 11, "bold"), fg=MUTED, bg=PANEL).grid(
            row=0, column=col, sticky="w", padx=(0 if col == 0 else 12, 6))
        e = tk.Entry(parent, width=9, validate="key", validatecommand=vcmd, justify="right",
                     **entry_opts(font=(FONT, 15)))
        e.grid(row=0, column=col + 1, sticky="ew", ipady=3)
        e.bind("<KeyRelease>", lambda ev: self.update_total())
        return e

    # ---- tenant search ---------------------------------------------------
    def _matching_tenants(self):
        q = self.name_entry.value().lower()
        tenants = self.store.tenants
        if not q:
            return tenants
        return [t for t in tenants if q in t["name"].lower() or q in t["address"].lower() or q == t["ref"].lower()
                or (q.isdigit() and t["ref"].lstrip("0") == q.lstrip("0"))]

    def _update_matches(self):
        self._shown = self._matching_tenants()
        self.matches.delete(0, "end")
        for t in self._shown:
            self.matches.insert("end", "  ·  ".join(x for x in (t["name"], t["address"], t["ref"]) if x))
        if not self.store.tenants:
            self.matches.insert("end", "No tenant list yet. Type the name and address,")
            self.matches.insert("end", "or import the list in Settings → Rent tenants.")
            self._shown = []
        elif not self._shown:
            self.matches.insert("end", "Not on the list: a new tenant. Type their address below.")
        self._update_new_tenant()

    def _focus_matches(self):
        if self._shown:
            self.matches.focus_set()
            self.matches.selection_clear(0, "end")
            self.matches.selection_set(0)
            self.matches.activate(0)

    def _pick_match(self, move_on=False):
        sel = self.matches.curselection()
        if not sel or sel[0] >= len(self._shown):
            return
        t = self._shown[sel[0]]
        self.name_entry.set(t["name"])
        self.address_entry.set(t["address"])
        self.details_entry.focus_set() if move_on else self.cash_entry.focus_set()
        self._update_new_tenant()

    def _form_tenant(self):
        """The tenant-list entry matching the name and address typed, or None (a new tenant)."""
        key = (self.name_entry.value().lower(), self.address_entry.value().lower())
        return next((t for t in self.store.tenants if tenant_key(t) == key), None)

    def _known_tenant(self):
        return self._form_tenant() is not None

    def _update_new_tenant(self):
        t = self._form_tenant()
        show = bool(self.name_entry.value()) and t is None
        self.add_tenant_check.config(state="normal" if show else "disabled", fg=TEXT if show else DIM)
        self.ref_label.config(text=f"Tenant ref {t['ref']}" if t and t["ref"] else "")
        if self._history_tenant():
            self.history_link.pack(side="right", padx=(10, 0), before=self.ref_label)
        else:
            self.history_link.pack_forget()
        self.update_total()

    def _history_tenant(self):
        """The tenant in the form, if they have any payments to show."""
        name, address = self.name_entry.value(), self.address_entry.value()
        t = self._form_tenant() or ({"name": name, "address": address, "ref": ""} if name and address else None)
        return t if t and self.store.payments_for(t) else None

    def open_history(self):
        t = self._history_tenant()
        if t:
            RentHistoryWindow(self.app, self.store, t).show()

    # ---- form state ------------------------------------------------------
    def set_kind(self, kind):
        self.kind = None if kind == self.kind else kind
        for k, btn in self.kind_buttons.items():
            if k == self.kind:
                btn.set_colors(TEAL, TEAL_HOVER, "#ffffff")
            else:
                btn.set_colors(GREY_BTN, GREY_HOVER, TEXT)
        self.update_total()

    def amounts(self):
        cash, cheque = parse_pence(self.cash_entry.get()), parse_pence(self.cheque_entry.get())
        return cash, cheque

    def form_problem(self):
        cash, cheque = self.amounts()
        if not self.name_entry.value():
            return "Enter the tenant's name"
        if not self.address_entry.value():
            return "Enter the tenant's address"
        if not self.kind:
            return "Choose what the payment is for"
        if cash is None or cheque is None:
            return "Amounts must be numbers, e.g. 300 or 300.50"
        if cash + cheque <= 0:
            return "Enter the cash and/or cheque amount"
        return None

    def update_total(self):
        if not hasattr(self, "take_btn"):
            return
        cash, cheque = self.amounts()
        total = (cash or 0) + (cheque or 0)
        self.total_label.config(text=fmt_money(total), fg=TEXT if total else DIM)
        problem = self.form_problem()
        self.problem_label.config(text=problem or f"Receipt no. {self.store.next_receipt()} will be issued")
        self.take_btn.set_enabled(problem is None)

    def reset(self):
        """Back to a fresh screen: empty form, no search, 'Not banked' filter."""
        self.clear_form()
        self.search.clear()
        self.set_filter("open")

    def clear_form(self):
        for e in (self.name_entry, self.address_entry, self.details_entry):
            e.clear()
        self.cash_entry.delete(0, "end")
        self.cheque_entry.delete(0, "end")
        self.add_tenant_var.set(True)
        self.kind = "Rent"
        self.set_kind(None)
        self._update_matches()

    # ---- right: status cards and payments table --------------------------
    def _build_main(self):
        S = self.app.S
        main = tk.Frame(self, bg=BG)
        main.grid(row=0, column=1, sticky="nsew", padx=S(24), pady=S(14))
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        cards = tk.Frame(main, bg=BG)
        cards.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        self.cards = {}
        for i, (key, title, accent) in enumerate([("awaiting", "Awaiting check", YELLOW),
                                                   ("safe", "In safe · not banked", ORANGE),
                                                   ("today", "Taken today", GREEN),
                                                   ("month", "Banked this month", TEAL)]):
            cards.columnconfigure(i, weight=1, uniform="card")
            card = tk.Frame(cards, bg=PANEL)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else S(12), 0))
            bar = tk.Frame(card, bg=accent, height=3)
            bar.pack(fill="x")
            tk.Label(card, text=title.upper(), font=(FONT, 9, "bold"), fg=MUTED, bg=PANEL,
                     anchor="w").pack(fill="x", padx=16, pady=(12, 0))
            value = tk.Label(card, text="£0.00", font=(FONT, 24, "bold"), fg=TEXT, bg=PANEL, anchor="w")
            value.pack(fill="x", padx=16)
            sub = tk.Label(card, text="", font=(FONT, 9), fg=MUTED, bg=PANEL, anchor="w")
            sub.pack(fill="x", padx=16, pady=(0, 12))
            self.cards[key] = (value, sub, bar, accent)

        bar = tk.Frame(main, bg=BG)
        bar.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.filter_buttons = {}
        for key, text in (("open", "Not banked"), ("awaiting", "Awaiting check"), ("in_safe", "In safe"),
                          ("banked", "Banked"), ("cancelled", "Cancelled"), ("all", "All")):
            btn = HoverButton(bar, text=text, font=(FONT, 9, "bold"), padx=10, pady=4,
                              command=lambda k=key: self.set_filter(k))
            btn.pack(side="left", padx=(0, 4))
            self.filter_buttons[key] = btn
        for text, bg, hover, cmd in [("Reports", GREY_BTN, GREY_HOVER, self.open_reports),
                                     ("Cancel receipt", RED, RED_HOVER, self.cancel_selected),
                                     ("Reprint", GREY_BTN, GREY_HOVER, self.reprint_selected),
                                     ("Bank…", TEAL, TEAL_HOVER, self.bank_selected),
                                     ("Check · in safe", ORANGE, ORANGE_HOVER, self.check_selected)]:
            HoverButton(bar, text=text, bg=bg, hover=hover, font=(FONT, 10, "bold"),
                        command=cmd).pack(side="right", padx=(8, 0))
        self.search = PlaceholderEntry(bar, "Search…", width=18)
        self.search.pack(side="right", padx=(8, 4), ipady=4)
        self.search.bind("<KeyRelease>", lambda e: self.refresh_table())

        table = tk.Frame(main, bg=PANEL)
        table.grid(row=2, column=0, sticky="nsew")
        columns = [("receipt", "Receipt", 70, "w"), ("date", "Date", 90, "w"), ("name", "Tenant", 170, "w"),
                   ("ref", "Ref", 60, "w"), ("address", "Address", 190, "w"), ("kind", "For", 170, "w"),
                   ("total", "Amount", 90, "e"),
                   ("method", "Paid by", 120, "w"), ("taken", "Taken by", 110, "w"), ("status", "Status", 220, "w")]
        self.tree = ttk.Treeview(table, columns=[c[0] for c in columns], show="headings", style="Dark.Treeview",
                                 selectmode="extended")
        for key, text, width, anchor in columns:
            self.tree.heading(key, text=text, anchor=anchor)
            self.tree.column(key, width=S(width), minwidth=S(40), anchor=anchor,
                             stretch=key in ("name", "address", "status"))
        sb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("tampered", background=TAMPERED_BG, foreground=TAMPERED_FG)  # first = wins
        self.tree.tag_configure("even", background=ROW_A)
        self.tree.tag_configure("odd", background=ROW_B)
        self.tree.tag_configure(S_AWAITING, foreground=YELLOW)
        self.tree.tag_configure(S_IN_SAFE, foreground=TEXT)
        self.tree.tag_configure(S_BANKED, foreground=MUTED)
        self.tree.tag_configure(S_CANCELLED, foreground=DIM, font=(FONT, 10, "overstrike"))
        self.tree.bind("<Double-1>", lambda e: self.reprint_selected())
        make_sortable(self.tree)
        self.empty_label = tk.Label(table, text="", font=(FONT, 11), fg=MUTED, bg=ROW_A)
        self.set_filter("open")

    def set_filter(self, key):
        self.filter = key
        for k, btn in self.filter_buttons.items():
            if k == key:
                btn.set_colors(TEAL, TEAL_HOVER, "#ffffff")
            else:
                btn.set_colors(GREY_BTN, GREY_HOVER, TEXT)
        self.refresh_table()

    def refresh(self):
        self.refresh_cards()
        self.refresh_table()
        self.update_total()

    def refresh_cards(self):
        pays = list(self.store.payments().values())
        today = date.today()
        by = {s: [p for p in pays if self.store.status(p) == s] for s in (S_AWAITING, S_IN_SAFE)}

        def summary(ps):
            if not ps:
                return "Nothing waiting"
            oldest = min(p["dt"] for p in ps)
            return f"{len(ps)} payment{'s' if len(ps) != 1 else ''} · oldest {oldest:%d/%m}"

        for key, status in (("awaiting", S_AWAITING), ("safe", S_IN_SAFE)):
            ps = by[status]
            value, sub, bar, accent = self.cards[key]
            value.config(text=fmt_money(sum(p["total"] for p in ps)))
            sub.config(text=summary(ps), fg=MUTED)
            if key == "safe" and ps:
                days = (today - min(p["dt"] for p in ps).date()).days
                late = days > self.store.unbanked_days()
                sub.config(text=f"{summary(ps)} ({days} day{'s' if days != 1 else ''})", fg=AMBER if late else MUTED)
                bar.config(bg=RED if late else accent)
            else:
                bar.config(bg=accent)
        taken = [p for p in pays if p["dt"].date() == today and not p["cancelled_at"]]
        value, sub, _, _ = self.cards["today"]
        value.config(text=fmt_money(sum(p["total"] for p in taken)))
        sub.config(text=f"{len(taken)} receipt{'s' if len(taken) != 1 else ''}")
        month = [d for d in self.store.deposits().values()
                 if d["bank_date"] and d["bank_date"] >= today.replace(day=1)]
        value, sub, _, _ = self.cards["month"]
        value.config(text=fmt_money(sum(d["total"] for d in month)))
        sub.config(text=f"{len(month)} deposit{'s' if len(month) != 1 else ''}")

    def status_text(self, p):
        s = self.store.status(p)
        if s == S_AWAITING:
            return "Awaiting check"
        if s == S_IN_SAFE:
            return f"In safe · checked by {p['checked_by']} {p['checked_at']:%d/%m}"
        if s == S_BANKED:
            return f"Banked {p['bank_date']:%d/%m/%Y} · {p['deposit']} · {p['banked_by']}"
        return f"Cancelled · {p['cancel_reason']}"

    def refresh_table(self):
        query = self.search.value().lower()
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        wanted = {"open": (S_AWAITING, S_IN_SAFE), "awaiting": (S_AWAITING,), "in_safe": (S_IN_SAFE,),
                  "banked": (S_BANKED,), "cancelled": (S_CANCELLED,),
                  "all": (S_AWAITING, S_IN_SAFE, S_BANKED, S_CANCELLED)}[self.filter]
        shown = 0
        for p in sorted(self.store.payments().values(), key=lambda p: p["receipt"], reverse=True):
            status = self.store.status(p)
            if status not in wanted:
                continue
            method = " + ".join(m for m, v in (("Cash", p["cash"]), ("Cheque", p["cheque"])) if v)
            kind = p["kind"] + (f" – {p['details']}" if p["details"] else "")
            values = (p["receipt"], p["dt"].strftime("%d/%m/%Y"), p["name"], p["ref"] or "new", p["address"], kind,
                      fmt_money(p["total"]), method, p["taken_by"], self.status_text(p))
            if query and not any(query in str(v).lower() for v in values):
                continue
            tags = ["odd" if shown % 2 else "even", status] + (["tampered"] if p["tampered"] else [])
            self.tree.insert("", "end", iid=str(p["receipt"]), values=values, tags=tags)
            shown += 1
        self.tree.resort()
        keep = [i for i in selected if self.tree.exists(i)]
        if keep:
            self.tree.selection_set(keep)
        if shown:
            self.empty_label.place_forget()
        else:
            self.empty_label.config(text="No matching payments." if query else {
                "open": "Nothing waiting to be checked or banked.", "awaiting": "Nothing waiting to be checked.",
                "in_safe": "Nothing in the safe waiting to be banked.", "banked": "Nothing banked yet.",
                "cancelled": "No cancelled receipts.", "all": "No payments yet."}[self.filter])
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")

    def selected_payments(self):
        pays = self.store.payments()
        return [pays[int(i)] for i in self.tree.selection() if int(i) in pays]

    # ---- actions ---------------------------------------------------------
    def take_payment(self):
        if self.form_problem():
            return
        cash, cheque = self.amounts()
        name, address = self.name_entry.value(), self.address_entry.value()
        tenant = self._form_tenant()
        ref = tenant["ref"] if tenant else ""
        details = self.details_entry.value()
        receipt = self.store.next_receipt()
        paid = " + ".join(f"{m} {fmt_money(v)}" for m, v in (("Cash", cash), ("Cheque", cheque)) if v)
        dlg = PinDialog(self.app, self.app, "Take payment", f"Receipt {receipt}  ·  {fmt_money(cash + cheque)}",
                        [name + (f"  ·  ref {ref}" if ref else "  ·  new tenant (no ref)"), address,
                         self.kind + (f" – {details}" if details else ""), paid],
                        confirm_text="Take payment & print receipt")
        result = dlg.show()
        if not result:
            self.app._pin_locked_out(dlg)
            return
        staff, _ = result
        saved = self.store.write([dict(type=R_RECEIPT, receipt=receipt, name=name, ref=ref, address=address,
                                       kind=self.kind,
                                       details=details, cash=cash, cheque=cheque, total=cash + cheque,
                                       staff=staff["name"])], self.app)
        if not saved:
            return
        if self.add_tenant_var.get() and not self._known_tenant():
            self.store.tenants.append({"name": name, "address": address, "ref": ""})
            try:
                save_rent_tenants(self.store.tenants)
            except OSError as e:
                messagebox.showwarning("Tenants", f"The payment is recorded, but the tenant list could not be "
                                       f"saved:\n\n{e}", parent=self.app)
        p = self.store.payments()[receipt]
        self.clear_form()
        self.set_filter("open")
        self.refresh()
        self.app.refresh_warning()
        self._issue_receipt(p)
        self.app.set_message(f"✓ Receipt {receipt}: {fmt_money(p['total'])} from {name}. Put the money with the "
                             "receipt copy for checking.", GREEN,
                             action=("Give change", lambda: self.app.no_sale(f"Change for rent receipt {receipt}"))
                             if cash else None)

    def _issue_receipt(self, p):
        path = receipt_path(p)
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            make_receipt_pdf(path, p, self.store.org())
            open_file(path)
        except Exception as e:
            messagebox.showwarning("Receipt", f"Payment {p['receipt']} IS recorded, but the receipt could not be "
                                   f"created or opened:\n\n{e}\n\nSelect it and press Reprint to try again.",
                                   parent=self.app)

    def reprint_selected(self):
        sel = self.selected_payments()
        if len(sel) != 1:
            messagebox.showinfo("Reprint", "Select one payment to reprint its receipt.", parent=self.app)
            return
        reprint_receipt(sel[0], self.store.org(), self.app)

    def check_selected(self):
        sel = [p for p in self.selected_payments() if self.store.status(p) == S_AWAITING]
        if not sel:
            messagebox.showinfo("Check", "Select the payments to check (Awaiting check). Hold Ctrl to select several.",
                                parent=self.app)
            return
        takers = {p["taken_by"] for p in sel}
        total = sum(p["total"] for p in sel)
        cash, cheque = sum(p["cash"] for p in sel), sum(p["cheque"] for p in sel)
        # A checker can't check their own payment, so when a checker took it anyone else may check it.
        checkers = {s["name"] for s in self.app.cfg["staff"] if is_checker(s)}
        taken_by_checkers = takers <= checkers

        def allow(staff):
            if staff["name"] in takers:
                return "You took one of these payments, so someone else must check it"
            if not is_checker(staff) and not taken_by_checkers:
                return "That PIN is not set up to check payments"
            return None

        receipts = ", ".join(str(p["receipt"]) for p in sel[:8]) + (" …" if len(sel) > 8 else "")
        dlg = PinDialog(self.app, self.app, "Check payments",
                        f"Check {len(sel)} payment{'s' if len(sel) != 1 else ''}  ·  {fmt_money(total)}",
                        [f"Receipts {receipts}", f"Cash {fmt_money(cash)}  ·  Cheques {fmt_money(cheque)}",
                         "Count the money against the receipts, then put it in the safe."],
                        allow=allow, confirm_text="Checked · put in safe", confirm_bg=ORANGE, confirm_hover=ORANGE_HOVER)
        result = dlg.show()
        if not result:
            self.app._pin_locked_out(dlg)
            return
        staff, _ = result
        if self.store.write([dict(type=R_CHECKED, receipt=p["receipt"], name=p["name"], total=p["total"],
                                  staff=staff["name"], notes="Checked and put in the safe") for p in sel], self.app):
            self.refresh()
            self.app.set_message(f"✓ {len(sel)} payment{'s' if len(sel) != 1 else ''} ({fmt_money(total)}) checked "
                                 f"by {staff['name']} and in the safe", GREEN)

    def bank_selected(self):
        sel = [p for p in self.selected_payments() if self.store.status(p) == S_IN_SAFE]
        skipped = [p for p in self.selected_payments() if self.store.status(p) == S_AWAITING]
        if not sel:
            messagebox.showinfo("Bank", "Select the payments being paid in (In safe). Payments must be checked "
                                "before they can be banked.", parent=self.app)
            return
        if skipped:
            messagebox.showinfo("Bank", f"{len(skipped)} selected payment(s) haven't been checked yet, so they are "
                                "left out.", parent=self.app)
        total = sum(p["total"] for p in sel)
        cash, cheque = sum(p["cash"] for p in sel), sum(p["cheque"] for p in sel)

        def validate(v):
            d = parse_ui_date(v["date"])
            if d is None:
                return "Enter the bank date as DD/MM/YYYY."
            if d > date.today():
                return "The bank date can't be in the future."
            if d < min(p["dt"].date() for p in sel):
                return "The bank date is before one of the payments was taken."
            amount = parse_pence(v["amount"])
            if amount != total:
                return (f"The paying-in slip must match the payments selected: {fmt_money(total)}. "
                        "Check the selection and the slip.")
            return None

        res = FieldsDialog(self.app, "Bank payments", f"Bank {len(sel)} payment{'s' if len(sel) != 1 else ''}", [
            {"key": "date", "label": "Date paid in at the bank", "value": date.today().strftime(UI_DATE_FMT)},
            {"key": "slip", "label": "Paying-in slip reference (optional)"},
            {"key": "amount", "label": f"Total on the paying-in slip (£)"},
        ], validate=validate, confirm_text="Next",
            sub=f"Selected: {fmt_money(total)}  (cash {fmt_money(cash)}, cheques {fmt_money(cheque)}). "
                "Enter the slip total to confirm it matches.").show()
        if not res:
            return
        deposit = self.store.next_deposit()
        bank_date = parse_ui_date(res["date"])
        dlg = PinDialog(self.app, self.app, "Bank payments", f"Deposit {deposit}  ·  {fmt_money(total)}",
                        [f"{len(sel)} payments paid in on {bank_date:%d/%m/%Y}",
                         f"Paying-in slip: {res['slip'] or '(no reference)'}"],
                        confirm_text="Record banking", confirm_bg=TEAL, confirm_hover=TEAL_HOVER)
        result = dlg.show()
        if not result:
            self.app._pin_locked_out(dlg)
            return
        staff, _ = result
        if self.store.write([dict(type=R_BANKED, receipt=p["receipt"], name=p["name"], total=p["total"],
                                  staff=staff["name"], deposit=deposit, bank_date=bank_date, slip=res["slip"])
                             for p in sel], self.app):
            self.refresh()
            self.app.set_message(f"✓ Deposit {deposit}: {len(sel)} payments, {fmt_money(total)} banked "
                                 f"{bank_date:%d/%m/%Y} by {staff['name']}. Keep the banking sheet with the slip",
                                 GREEN)
            print_banking_sheet(self.store, deposit, self.app)

    def cancel_selected(self):
        sel = self.selected_payments()
        if len(sel) != 1:
            messagebox.showinfo("Cancel receipt", "Select one receipt to cancel.", parent=self.app)
            return
        p = sel[0]
        status = self.store.status(p)
        if status == S_CANCELLED:
            messagebox.showinfo("Cancel receipt", f"Receipt {p['receipt']} is already cancelled.", parent=self.app)
            return
        if status == S_BANKED:
            messagebox.showinfo("Cancel receipt", f"Receipt {p['receipt']} has been banked, so it can't be cancelled "
                                "here. The accountants need to correct it.", parent=self.app)
            return
        where = ("It has been checked, so take the money back out of the safe."
                 if status == S_IN_SAFE else "Give the money back or keep it aside.")
        dlg = PinDialog(self.app, self.app, "Cancel receipt", f"Cancel receipt {p['receipt']}",
                        [f"{p['name']}  ·  {fmt_money(p['total'])}", f"{p['kind']} on {p['dt']:%d/%m/%Y}",
                         f"The number is not reused. {where}"],
                        allow=lambda s: None if (s["admin"] or is_checker(s)) else
                        "Only an admin or checker can cancel receipts",
                        reason_label="Reason (e.g. wrong amount, reissued as next receipt)",
                        confirm_text="Cancel receipt", confirm_bg=RED, confirm_hover=RED_HOVER)
        result = dlg.show()
        if not result:
            self.app._pin_locked_out(dlg)
            return
        staff, reason = result
        if self.store.write([dict(type=R_CANCELLED, receipt=p["receipt"], name=p["name"], total=p["total"],
                                  staff=staff["name"], notes=reason)], self.app):
            self.refresh()
            self.app.set_message(f"Receipt {p['receipt']} cancelled by {staff['name']}", YELLOW)

    def open_reports(self):
        RentReportsWindow(self.app, self.store).show()


# --------------------------------------------------------------------------
# Tenant history
# --------------------------------------------------------------------------

class RentHistoryWindow(Dialog):
    """Every payment for one tenancy, with reprints and a printable statement."""

    def __init__(self, app, store, tenant):
        super().__init__(app, "Tenant history", resizable=True)
        self.app = app
        self.store = store
        self.tenant = tenant
        self.all_pays = store.payments_for(tenant)
        S = app.S
        self.geometry(f"{S(1040)}x{S(660)}")
        self.minsize(S(880), S(520))
        self.heading(f"{tenant['name']}  ·  {tenant['address']}",
                     (f"Tenant ref {tenant['ref']}. Includes everyone on this tenancy. " if tenant["ref"] else
                      "New tenant (no ref yet). ") + "Cancelled receipts are struck through and not counted. "
                     "Double-click a payment to reprint its receipt.")

        bar = tk.Frame(self.body, bg=PANEL)
        bar.pack(fill="x", pady=(14, 0))
        tk.Label(bar, text="From", font=(FONT, 10, "bold"), fg=MUTED, bg=PANEL).pack(side="left")
        self.from_entry = tk.Entry(bar, width=11, **entry_opts())
        self.from_entry.pack(side="left", padx=(6, 12), ipady=4)
        tk.Label(bar, text="To", font=(FONT, 10, "bold"), fg=MUTED, bg=PANEL).pack(side="left")
        self.to_entry = tk.Entry(bar, width=11, **entry_opts())
        self.to_entry.pack(side="left", padx=(6, 8), ipady=4)
        HoverButton(bar, text="Show", bg=TEAL, hover=TEAL_HOVER, font=(FONT, 10, "bold"),
                    command=self.run).pack(side="left", padx=(0, 16))
        today = date.today()
        for text, start, end in (("All", None, None),
                                 ("Last 12 months", today - timedelta(days=364), today),
                                 ("This year", today.replace(month=1, day=1), today),
                                 ("Last year", date(today.year - 1, 1, 1), date(today.year - 1, 12, 31))):
            HoverButton(bar, text=text, font=(FONT, 9, "bold"), padx=10, pady=4,
                        command=lambda s=start, e=end: self.set_range(s, e)).pack(side="left", padx=2)
        for e in (self.from_entry, self.to_entry):
            e.bind("<Return>", lambda ev: self.run())

        cards = tk.Frame(self.body, bg=PANEL)
        cards.pack(fill="x", pady=(14, 0))
        self.card_values = {}
        for i, (key, title, accent) in enumerate([("paid", "Paid", GREEN), ("count", "Payments", MUTED),
                                                   ("last", "Last payment", MUTED), ("open", "Not banked yet", ORANGE)]):
            cards.columnconfigure(i, weight=1, uniform="card")
            card = tk.Frame(cards, bg=PANEL_2)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 6, 0))
            tk.Frame(card, bg=accent, height=3).pack(fill="x")
            tk.Label(card, text=title.upper(), font=(FONT, 8, "bold"), fg=MUTED, bg=PANEL_2,
                     anchor="w").pack(fill="x", padx=12, pady=(8, 0))
            val = tk.Label(card, text="", font=(FONT, 15, "bold"), fg=TEXT, bg=PANEL_2, anchor="w")
            val.pack(fill="x", padx=12, pady=(0, 8))
            self.card_values[key] = val

        frame = tk.Frame(self.body, bg=PANEL)
        frame.pack(fill="both", expand=True, pady=(14, 0))
        columns = [("receipt", "Receipt", 70, "w"), ("date", "Date", 90, "w"), ("name", "Paid by", 170, "w"),
                   ("kind", "For", 190, "w"), ("method", "Cash / cheque", 110, "w"), ("total", "Amount", 90, "e"),
                   ("taken", "Taken by", 100, "w"), ("status", "Status", 150, "w")]
        self.tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", style="Dark.Treeview",
                                 selectmode="browse")
        for key, text, width, anchor in columns:
            self.tree.heading(key, text=text, anchor=anchor)
            self.tree.column(key, width=S(width), anchor=anchor, stretch=key in ("name", "kind"))
        sb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("even", background=ROW_A)
        self.tree.tag_configure("odd", background=ROW_B)
        self.tree.tag_configure(S_AWAITING, foreground=YELLOW)
        self.tree.tag_configure(S_CANCELLED, foreground=DIM, font=(FONT, 10, "overstrike"))
        self.tree.bind("<Double-1>", lambda e: self.reprint())
        self.empty_label = tk.Label(frame, text="No payments in this period.", font=(FONT, 11), fg=MUTED, bg=ROW_A)

        foot = tk.Frame(self.body, bg=PANEL)
        foot.pack(fill="x", pady=(14, 0))
        HoverButton(foot, text="Close", command=self.cancel).pack(side="right")
        HoverButton(foot, text="Print statement", bg=TEAL, hover=TEAL_HOVER, font=(FONT, 10, "bold"),
                    command=self.statement).pack(side="left")
        HoverButton(foot, text="Reprint receipt", font=(FONT, 10, "bold"),
                    command=self.reprint).pack(side="left", padx=6)
        HoverButton(foot, text="Export to Excel…", font=(FONT, 10, "bold"),
                    command=self.export).pack(side="left")
        self.set_range(None, None)

    def set_range(self, start, end):
        if start is None:
            start = min((p["dt"].date() for p in self.all_pays), default=date.today())
            end = date.today()
        for entry, d in ((self.from_entry, start), (self.to_entry, end)):
            entry.delete(0, "end")
            entry.insert(0, d.strftime(UI_DATE_FMT))
        self.run()

    def run(self):
        start, end = parse_ui_date(self.from_entry.get()), parse_ui_date(self.to_entry.get())
        if start is None or end is None:
            messagebox.showwarning("Tenant history", "Please enter dates as DD/MM/YYYY.", parent=self)
            return
        if start > end:
            start, end = end, start
        self.range = (start, end)
        self.pays = [p for p in self.all_pays if start <= p["dt"].date() <= end]
        live = [p for p in self.pays if not p["cancelled_at"]]
        self.card_values["paid"].config(text=fmt_money(sum(p["total"] for p in live)))
        self.card_values["count"].config(text=str(len(live)))
        self.card_values["last"].config(text=live[-1]["dt"].strftime("%d/%m/%Y") if live else "—")
        unbanked = [p for p in live if not p["banked_at"]]
        self.card_values["open"].config(text=fmt_money(sum(p["total"] for p in unbanked)))

        self.tree.delete(*self.tree.get_children())
        for i, p in enumerate(reversed(self.pays)):
            status = RentStore.status(p)
            self.tree.insert("", "end", iid=str(p["receipt"]), values=(
                p["receipt"], p["dt"].strftime("%d/%m/%Y"), p["name"],
                p["kind"] + (f" – {p['details']}" if p["details"] else ""),
                " + ".join(m for m, v in (("Cash", p["cash"]), ("Cheque", p["cheque"])) if v),
                fmt_money(p["total"]), p["taken_by"],
                STATUS_TEXT[status] + (f" {p['deposit']}" if status == S_BANKED else "")),
                tags=("odd" if i % 2 else "even", status))
        if self.pays:
            self.empty_label.place_forget()
        else:
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")

    def reprint(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Reprint", "Select a payment to reprint its receipt.", parent=self)
            return
        reprint_receipt(self.store.payments()[int(sel[0])], self.store.org(), self)

    def _period(self):
        start, end = self.range
        return f"{start:%d/%m/%Y} – {end:%d/%m/%Y}"

    def statement(self):
        if not self.pays:
            messagebox.showinfo("Statement", "There are no payments in this period.", parent=self)
            return
        safe_name = "".join(ch for ch in self.tenant["name"] if ch.isalnum() or ch in " -")[:40].strip()
        path = os.path.join(tempfile.gettempdir(), f"Statement_{safe_name or 'tenant'}_{date.today():%Y-%m-%d}.pdf")
        try:
            make_statement_pdf(path, self.tenant, self.pays, self._period(), self.store.org())
            open_file(path)
        except Exception as e:
            messagebox.showerror("Statement", f"Could not create the statement:\n\n{e}", parent=self)

    def export(self):
        start, end = self.range
        path = admin_save_path(
            self.app, self, title="Export to Excel", defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")],
            initialfile=f"payments_{self.tenant['ref'] or 'new'}_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}.xlsx")
        if not path:
            return
        try:
            write_xlsx(path, RentReportsWindow.PAY_HEADERS, [RentReportsWindow.pay_row(p) for p in self.pays],
                       RentReportsWindow.PAY_MONEY)
        except Exception as e:
            messagebox.showerror("Export failed", f"Could not write the file (is it open in Excel?):\n\n{e}",
                                 parent=self)
            return
        messagebox.showinfo("Exported", f"Saved {len(self.pays)} rows to:\n{path}", parent=self)


# --------------------------------------------------------------------------
# Reports
# --------------------------------------------------------------------------

class RentReportsWindow(Dialog):
    def __init__(self, app, store):
        super().__init__(app, "Rent reports", resizable=True)
        self.app = app
        self.store = store
        S = app.S
        self.geometry(f"{S(1100)}x{S(760)}")
        self.minsize(S(900), S(600))
        self.heading("Rent reports", "Payments by the date they were taken; deposits by the date they were banked. "
                                     "Cancelled receipts are listed but not counted.")
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
        for text, start, end in (("Today", today, today),
                                 ("This week", today - timedelta(days=today.weekday()), today),
                                 ("This month", month_start, today),
                                 ("Last month", last_month_end.replace(day=1), last_month_end),
                                 ("All time", None, None)):
            HoverButton(bar, text=text, font=(FONT, 9, "bold"), padx=10, pady=4,
                        command=lambda s=start, e=end: self.set_range(s, e)).pack(side="left", padx=2)
        for e in (self.from_entry, self.to_entry):
            e.bind("<Return>", lambda ev: self.run())

        cards = tk.Frame(self.body, bg=PANEL)
        cards.pack(fill="x", pady=(16, 0))
        self.card_values = {}
        for i, (key, title, accent) in enumerate([("taken", "Received", GREEN), ("cash", "Cash", GREEN),
                                                   ("cheque", "Cheques", GREEN), ("banked", "Banked", TEAL),
                                                   ("count", "Receipts", MUTED), ("cancelled", "Cancelled", RED)]):
            cards.columnconfigure(i, weight=1, uniform="card")
            card = tk.Frame(cards, bg=PANEL_2)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 6, 0))
            tk.Frame(card, bg=accent, height=3).pack(fill="x")
            tk.Label(card, text=title.upper(), font=(FONT, 8, "bold"), fg=MUTED, bg=PANEL_2,
                     anchor="w").pack(fill="x", padx=12, pady=(8, 0))
            val = tk.Label(card, text="0", font=(FONT, 16, "bold"), fg=TEXT, bg=PANEL_2, anchor="w")
            val.pack(fill="x", padx=12, pady=(0, 8))
            self.card_values[key] = val

        nb = ttk.Notebook(self.body)
        nb.pack(fill="both", expand=True, pady=(16, 0))
        self.pay_tree = self._tree(nb, [("receipt", "Receipt", 70, "w"), ("date", "Taken", 90, "w"),
                                        ("name", "Tenant", 170, "w"), ("ref", "Ref", 60, "w"),
                                        ("address", "Address", 190, "w"),
                                        ("kind", "For", 160, "w"), ("cash", "Cash", 80, "e"),
                                        ("cheque", "Cheque", 80, "e"), ("total", "Total", 90, "e"),
                                        ("status", "Status", 150, "w")])
        self.dep_tree = self._tree(nb, [("deposit", "Deposit", 80, "w"), ("date", "Bank date", 100, "w"),
                                        ("slip", "Paying-in ref", 140, "w"), ("by", "Banked by", 120, "w"),
                                        ("count", "Receipts", 80, "e"), ("cash", "Cash", 90, "e"),
                                        ("cheque", "Cheques", 90, "e"), ("total", "Total", 100, "e")])
        self.kind_tree = self._tree(nb, [("kind", "For", 200, "w"), ("count", "Receipts", 100, "e"),
                                         ("total", "Total", 120, "e")])
        nb.add(self.pay_tree.master, text="Payments")
        nb.add(self.dep_tree.master, text="Deposits (select one to export it)")
        nb.add(self.kind_tree.master, text="By type")

        foot = tk.Frame(self.body, bg=PANEL)
        foot.pack(fill="x", pady=(14, 0))
        HoverButton(foot, text="Close", command=self.cancel).pack(side="right")
        HoverButton(foot, text="Export payments…", font=(FONT, 10, "bold"),
                    command=self.export_payments).pack(side="left")
        HoverButton(foot, text="Export deposits…", font=(FONT, 10, "bold"),
                    command=self.export_deposits).pack(side="left", padx=6)
        HoverButton(foot, text="Export selected deposit for accounts…", font=(FONT, 10, "bold"), bg=TEAL,
                    hover=TEAL_HOVER, command=self.export_one_deposit).pack(side="left")
        HoverButton(foot, text="Banking sheet…", font=(FONT, 10, "bold"),
                    command=self.banking_sheet).pack(side="left", padx=6)
        self.set_range(month_start, today)

    def _tree(self, parent, columns):
        frame = tk.Frame(parent, bg=PANEL)
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", style="Dark.Treeview",
                            selectmode="browse")
        for key, text, width, anchor in columns:
            tree.heading(key, text=text, anchor=anchor)
            tree.column(key, width=self.app.S(width), anchor=anchor, stretch=True)
        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        tree.tag_configure("odd", background=ROW_B)
        tree.tag_configure("even", background=ROW_A)
        tree.tag_configure("cancelled", foreground=DIM, font=(FONT, 10, "overstrike"))
        return tree

    def set_range(self, start, end):
        if start is None:
            pays = self.store.payments().values()
            start = min((p["dt"].date() for p in pays), default=date.today())
            end = date.today()
        for entry, d in ((self.from_entry, start), (self.to_entry, end)):
            entry.delete(0, "end")
            entry.insert(0, d.strftime(UI_DATE_FMT))
        self.run()

    def run(self):
        start, end = parse_ui_date(self.from_entry.get()), parse_ui_date(self.to_entry.get())
        if start is None or end is None:
            messagebox.showwarning("Reports", "Please enter dates as DD/MM/YYYY.", parent=self)
            return
        if start > end:
            start, end = end, start
        self.range = (start, end)
        self.pays = [p for p in sorted(self.store.payments().values(), key=lambda p: p["receipt"])
                     if start <= p["dt"].date() <= end]
        live = [p for p in self.pays if not p["cancelled_at"]]
        self.deps = [d for d in self.store.deposits().values() if d["bank_date"] and start <= d["bank_date"] <= end]
        self.card_values["taken"].config(text=fmt_money(sum(p["total"] for p in live)))
        self.card_values["cash"].config(text=fmt_money(sum(p["cash"] for p in live)))
        self.card_values["cheque"].config(text=fmt_money(sum(p["cheque"] for p in live)))
        self.card_values["banked"].config(text=fmt_money(sum(d["total"] for d in self.deps)))
        self.card_values["count"].config(text=str(len(live)))
        self.card_values["cancelled"].config(text=str(len(self.pays) - len(live)))

        self.pay_tree.delete(*self.pay_tree.get_children())
        for i, p in enumerate(reversed(self.pays)):
            status = RentStore.status(p)
            self.pay_tree.insert("", "end", values=(
                p["receipt"], p["dt"].strftime("%d/%m/%Y"), p["name"], p["ref"] or "new", p["address"],
                p["kind"] + (f" – {p['details']}" if p["details"] else ""), fmt_money(p["cash"]),
                fmt_money(p["cheque"]), fmt_money(p["total"]),
                STATUS_TEXT[status] + (f" {p['deposit']}" if status == S_BANKED else "")),
                tags=("odd" if i % 2 else "even",) + (("cancelled",) if status == S_CANCELLED else ()))

        self.dep_tree.delete(*self.dep_tree.get_children())
        for i, d in enumerate(self.deps):
            self.dep_tree.insert("", "end", iid=d["deposit"], values=(
                d["deposit"], d["bank_date"].strftime("%d/%m/%Y"), d["slip"], d["banked_by"], len(d["receipts"]),
                fmt_money(d["cash"]), fmt_money(d["cheque"]), fmt_money(d["total"])),
                tags=("odd" if i % 2 else "even",))

        kinds = {}
        for p in live:
            k = kinds.setdefault(p["kind"], [0, 0])
            k[0] += 1
            k[1] += p["total"]
        self.kind_tree.delete(*self.kind_tree.get_children())
        for i, (kind, (n, t)) in enumerate(sorted(kinds.items(), key=lambda kv: -kv[1][1])):
            self.kind_tree.insert("", "end", values=(kind, n, fmt_money(t)), tags=("odd" if i % 2 else "even",))

    def _save(self, name, headers, rows, money_cols=()):
        path = admin_save_path(self.app, self, title="Export to Excel", initialfile=name,
                               defaultextension=".xlsx", filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            write_xlsx(path, headers, rows, money_cols)
        except Exception as e:
            messagebox.showerror("Export failed", f"Could not write the file (is it open in Excel?):\n\n{e}",
                                 parent=self)
            return
        messagebox.showinfo("Exported", f"Saved {len(rows)} rows to:\n{path}", parent=self)

    PAY_HEADERS = ["Receipt No", "Date Taken", "Tenant Name", "Tenant Ref", "Address", "For", "Details", "Cash (£)",
                   "Cheque (£)", "Total (£)", "Taken By", "Checked By", "Status", "Deposit No", "Bank Date",
                   "Paying-in Ref", "Banked By", "Cancel Reason"]
    PAY_MONEY = ("Cash (£)", "Cheque (£)", "Total (£)")

    @staticmethod
    def pay_row(p):
        return [p["receipt"], p["dt"].date(), p["name"], p["ref"], p["address"], p["kind"], p["details"],
                p["cash"] / 100, p["cheque"] / 100, p["total"] / 100, p["taken_by"], p["checked_by"],
                STATUS_TEXT[RentStore.status(p)], p["deposit"], p["bank_date"], p["slip"], p["banked_by"],
                p["cancel_reason"]]

    def export_payments(self):
        start, end = self.range
        self._save(f"rent_payments_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}.xlsx", self.PAY_HEADERS,
                   [self.pay_row(p) for p in self.pays], self.PAY_MONEY)

    def export_deposits(self):
        start, end = self.range
        self._save(f"rent_deposits_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}.xlsx",
                   ["Deposit No", "Bank Date", "Paying-in Ref", "Banked By", "Receipts", "Cash (£)", "Cheques (£)",
                    "Total (£)"],
                   [[d["deposit"], d["bank_date"], d["slip"], d["banked_by"], " ".join(map(str, d["receipts"])),
                     d["cash"] / 100, d["cheque"] / 100, d["total"] / 100] for d in self.deps],
                   ("Cash (£)", "Cheques (£)", "Total (£)"))

    def banking_sheet(self):
        sel = self.dep_tree.selection()
        if not sel:
            messagebox.showinfo("Banking sheet", "Select a deposit on the Deposits tab first.", parent=self)
            return
        print_banking_sheet(self.store, sel[0], self, keep=False)

    def export_one_deposit(self):
        sel = self.dep_tree.selection()
        if not sel:
            messagebox.showinfo("Export deposit", "Select a deposit on the Deposits tab first.", parent=self)
            return
        d = self.store.deposits()[sel[0]]
        pays = self.store.payments()
        rows = [self.pay_row(pays[n]) for n in sorted(d["receipts"])]
        rows.append(["TOTAL", None, "", "", "", "", "", d["cash"] / 100, d["cheque"] / 100, d["total"] / 100,
                     "", "", "", d["deposit"], d["bank_date"], d["slip"], d["banked_by"], ""])
        self._save(f"deposit_{d['deposit']}_{d['bank_date']:%Y-%m-%d}.xlsx", self.PAY_HEADERS, rows, self.PAY_MONEY)


# --------------------------------------------------------------------------
# Settings tabs (added to the shared Settings window)
# --------------------------------------------------------------------------

def build_tenants_tab(win, p):
    """Settings → Rent tenants."""
    store = win.app.rent
    tree = win._tree(p, [("name", "Tenant name", 260), ("address", "Address", 380), ("ref", "Tenant ref", 110)])

    def fill(select=None):
        tree.delete(*tree.get_children())
        for i, t in enumerate(store.tenants):
            iid = tree.insert("", "end", values=(t["name"], t["address"] or "—", t["ref"] or "—"),
                              tags=("odd" if i % 2 else "even",))
            if t is select:
                tree.selection_set(iid)
                tree.see(iid)
        count.config(text=f"{len(store.tenants)} tenants")

    def selected():
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Tenants", "Select a tenant first.", parent=win)
            return None
        return store.tenants[tree.index(sel[0])]

    def save(select=None):
        store.tenants.sort(key=lambda t: natural_key(t["name"]))
        try:
            save_rent_tenants(store.tenants)
        except OSError as e:
            messagebox.showerror("Tenants", f"Could not save rent_tenants.csv:\n\n{e}", parent=win)
        fill(select)
        win.app.rent_view._update_matches()

    def dialog(heading, t=None):
        return FieldsDialog(win, "Tenant", heading, [
            {"key": "name", "label": "Tenant name", "value": t["name"] if t else ""},
            {"key": "address", "label": "Address (e.g. Rm 3, 91 Wellesley Road)", "value": t["address"] if t else ""},
            {"key": "ref", "label": "Tenant ref (Ten Key in the housing system; blank if not set up yet)",
             "value": t["ref"] if t else ""},
        ], validate=lambda v: None if v["name"] else "Please enter a name.").show()

    def edit():
        t = selected()
        if t:
            res = dialog("Edit tenant", t)
            if res:
                t.update(res)
                save(t)

    def add():
        res = dialog("Add tenant")
        if res:
            t = dict(res)
            store.tenants.append(t)
            save(t)

    def remove():
        t = selected()
        if t and messagebox.askyesno("Remove tenant", f"Remove {t['name']} from the list?\n"
                                     "Their past payments stay in the log.", parent=win):
            store.tenants.remove(t)
            save()

    def import_file():
        path = filedialog.askopenfilename(parent=win, title="Import tenants",
                                          filetypes=[("Excel or CSV", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            found = read_tenant_file(path)
        except Exception as e:
            messagebox.showerror("Import", f"Could not read the file:\n\n{e}", parent=win)
            return
        found = list({tenant_key(t): t for t in found}.values())
        known = {tenant_key(t): t for t in store.tenants}
        new = [t for t in found if tenant_key(t) not in known]
        refs = [(known[tenant_key(t)], t["ref"]) for t in found
                if tenant_key(t) in known and t["ref"] and known[tenant_key(t)]["ref"] != t["ref"]]
        dropped = len(known.keys() - {tenant_key(t) for t in found})
        sample = "\n".join("  " + "  ·  ".join(x for x in (t["name"], t["address"] or "(no address)", t["ref"]) if x)
                           for t in found[:4])
        choice = messagebox.askyesnocancel(
            "Import tenants",
            f"The file has {len(found)} tenants ({sum(1 for t in found if t['ref'])} with a tenant ref). "
            f"For example:\n\n{sample}\n\n"
            f"YES: replace the whole list with this file. Use this for a full export from the housing system. "
            f"{dropped} tenant(s) on the current list are not in the file and would be removed.\n\n"
            f"NO: keep the current list and only add the {len(new)} new tenant(s)"
            + (f" and update {len(refs)} tenant ref(s)" if refs else "") + ".\n\n"
            "CANCEL: change nothing.", parent=win)
        if choice is None:
            return
        if choice:
            store.tenants[:] = found
        else:
            store.tenants.extend(new)
            for old, ref in refs:
                old["ref"] = ref
        save()

    btns = tk.Frame(p, bg=PANEL)
    btns.pack(fill="x", pady=(12, 0))
    HoverButton(btns, text="Edit", bg=TEAL, hover=TEAL_HOVER, font=(FONT, 10, "bold"), command=edit).pack(side="left")
    for text, cmd in (("Add", add), ("Remove", remove), ("Import from Excel…", import_file)):
        HoverButton(btns, text=text, font=(FONT, 10, "bold"), command=cmd).pack(side="left", padx=(8, 0))
    count = tk.Label(btns, text="", font=(FONT, 10), fg=MUTED, bg=PANEL)
    count.pack(side="right")
    tree.bind("<Double-1>", lambda e: edit())
    win._note(p, "Used to fill in the name, address and tenant ref when taking a payment. New tenants can also be "
                 "typed in when taking a payment and added to this list then; add their ref here once they are on "
                 "the housing system. Import reads the housing system's tenant export (Forename, Surname, Address "
                 "Line 1, Ten Key) or any sheet with a 'name' column. Importing again adds new tenants and fills in "
                 "missing refs; it never removes anyone.")
    fill()


def build_receipts_tab(win, p):
    """Settings → Receipts: what is printed on the receipt, and numbering."""
    store = win.app.rent
    org = store.org()
    entries = {}
    for key, label in (("name1", "Organisation name (large)"), ("name2", "Second line"), ("address", "Address"),
                       ("contact", "Phone / email line")):
        win._label(p, label, top=0 if key == "name1" else 10)
        e = tk.Entry(p, width=70, **entry_opts())
        e.insert(0, org[key])
        e.pack(fill="x", ipady=4)
        entries[key] = e
    row = tk.Frame(p, bg=PANEL)
    row.pack(fill="x", pady=(10, 0))
    tk.Label(row, text="Warn when cash has been in the safe unbanked for more than", font=(FONT, 10), fg=MUTED,
             bg=PANEL).pack(side="left")
    days = tk.Entry(row, width=4, justify="center", **entry_opts())
    days.insert(0, str(store.unbanked_days()))
    days.pack(side="left", padx=6, ipady=3)
    tk.Label(row, text="days", font=(FONT, 10), fg=MUTED, bg=PANEL).pack(side="left")

    def save():
        new_org = {k: e.get().strip() for k, e in entries.items()}
        if not new_org["name1"]:
            messagebox.showwarning("Receipts", "The organisation name can't be empty.", parent=win)
            return
        n = parse_int(days.get(), 0)
        if n < 1:
            messagebox.showwarning("Receipts", "Enter a number of days, e.g. 7.", parent=win)
            return
        win.app.cfg["receipt_org"] = new_org
        win.app.cfg["rent_unbanked_days"] = n
        win.app.save_cfg()
        win.app.rent_view.refresh()
        messagebox.showinfo("Receipts", "Saved. New receipts will use these details.", parent=win)

    def preview():
        sample = {"receipt": store.next_receipt(), "dt": datetime.now(), "name": "Sample Tenant", "ref": "00000",
                  "address": "Rm 1, 12 High Street", "kind": "Rent", "details": "October", "cash": 30000,
                  "cheque": 0, "total": 30000, "taken_by": win.admin["name"]}
        path = os.path.join(tempfile.gettempdir(), "Receipt_preview.pdf")
        try:
            make_receipt_pdf(path, sample, {k: e.get().strip() for k, e in entries.items()}, stamp="SAMPLE")
            open_file(path)
        except Exception as e:
            messagebox.showerror("Preview", f"Could not create the preview:\n\n{e}", parent=win)

    btns = tk.Frame(p, bg=PANEL)
    btns.pack(fill="x", pady=(16, 0))
    HoverButton(btns, text="Save", bg=TEAL, hover=TEAL_HOVER, command=save).pack(side="left")
    HoverButton(btns, text="Preview receipt", command=preview).pack(side="left", padx=8)
    win._note(p, f"Receipts are numbered automatically. The next one is {store.next_receipt()}. Numbers are never "
                 "reused, even when a receipt is cancelled. Every receipt is saved as a PDF in the 'receipts' folder, "
                 "exactly as issued; reprints are marked COPY.")
