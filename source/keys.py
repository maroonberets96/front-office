"""
Key logbook
===========

Keys lent out (to contractors), from issue to return. Replaces the paper key logbook.

  Issue    Any staff PIN. The property (from the property list, or typed in), which keys and
           how many, the contractor, and the date they should be back. A slip can be printed
           for the contractor to sign.
  Return   Any staff PIN. Some of the keys can come back and the rest stay out.
  Extend   A new return-by date, when the contractor needs the keys for longer.
  Lost     Keys that won't come back, with what was done about it (e.g. lock changed).
  Void     A mistaken entry, by an admin or checker, with a reason. Nothing is ever deleted.

Data files, in data/:
  key_log.csv      every issue, return, extension, loss and void (append-only, with the same
                   Check chain as the other logs)
  properties.csv   Prop Key, Property, Postcode (imported from the housing system's property list)
  contractors.csv  Name, Company, Phone
and keys/YYYY/ (next to data/) holds each slip printed for signing.
"""

import os
import re
import tempfile
import tkinter as tk
from datetime import date, datetime, timedelta
from tkinter import filedialog, messagebox, ttk

from front_office import (
    AMBER, BASE_DIR, BG, BORDER, DATA_DIR, DIM, FONT, GREEN, GREEN_HOVER, GREY_BTN, GREY_HOVER, MUTED, ORANGE,
    PANEL, PANEL_2, RED, RED_HOVER, ROW_A, ROW_B, TAMPERED_BG, TAMPERED_FG, TEAL, TEAL_HOVER, TEXT, UI_DATE_FMT,
    Dialog, FieldsDialog, HoverButton, PinDialog, PlaceholderEntry, admin_save_path, entry_opts, make_sortable, natural_key, parse_int, parse_ui_date,
)
from logbook import ChainLog, load_list, read_sheet, save_list
from rent import INK, SHEET_LEFT, SHEET_WIDTH, _letterhead, _sheet, _table, canvas, open_file, write_xlsx

try:
    from reportlab.lib.colors import HexColor
    from reportlab.lib.units import mm
    from reportlab.lib.utils import simpleSplit
except ImportError:  # keys are still logged; only the printed slip is unavailable
    pass


KEY_LOG_PATH = os.path.join(DATA_DIR, "key_log.csv")
PROPERTIES_PATH = os.path.join(DATA_DIR, "properties.csv")
CONTRACTORS_PATH = os.path.join(DATA_DIR, "contractors.csv")
SLIPS_DIR = os.path.join(BASE_DIR, "keys")

KEY_HEADERS = [
    "Entry No", "Date & Time", "Type", "Slip No", "Property", "Prop Key", "Keys", "Contractor", "Company", "Phone",
    "Reason", "Return By", "Staff", "Returned By", "Notes", "Check",
]
KEY_SALT = b"front-office/key-chain/v1"

K_ISSUED, K_RETURNED, K_EXTENDED, K_LOST, K_VOIDED = "ISSUED", "RETURNED", "EXTENDED", "LOST", "VOIDED"
PROPERTY_COLUMNS = [("key", "Prop Key"), ("name", "Property"), ("postcode", "Postcode")]
CONTRACTOR_COLUMNS = [("name", "Name"), ("company", "Company"), ("phone", "Phone")]

# The property list only has homes and common parts; these say which key it is.
KEY_KINDS = ["Front / communal door", "Flat / room door", "Fob", "CCTV room", "Boiler room", "Bike store",
             "Bin store", "Meter cupboard"]

S_OUT, S_OVERDUE, S_RETURNED, S_LOST, S_VOIDED = "out", "overdue", "returned", "lost", "voided"


def slip_name(n):
    return f"K{n:04d}"


# --------------------------------------------------------------------------
# Keys as text: "Fob x1; Front / communal door x2" in the log
# --------------------------------------------------------------------------

def keys_to_text(keys):
    return "; ".join(f"{name} x{qty}" for name, qty in keys.items() if qty > 0)


def text_to_keys(text):
    keys = {}
    for part in (text or "").split(";"):
        m = re.fullmatch(r"\s*(.+?)\s+x\s*(\d+)\s*", part)
        if m:
            keys[m.group(1)] = keys.get(m.group(1), 0) + int(m.group(2))
        elif part.strip():
            keys[part.strip()] = keys.get(part.strip(), 0) + 1
    return keys


def keys_display(keys):
    return ", ".join(name if qty == 1 else f"{qty} × {name}" for name, qty in keys.items() if qty > 0) or "—"


def key_count(keys):
    return sum(keys.values())


def plural(n, word):
    return f"{n} {word}{'' if n == 1 else 's'}"


# --------------------------------------------------------------------------
# Properties and contractors
# --------------------------------------------------------------------------

UNIT_ONLY = re.compile(r"(flat|rm|room)\s*\w+(\s*,\s*(rm|room)\s*\w+)?(\s*\(.*\))?", re.I)


def property_label(lines):
    """A readable name from the housing system's address lines:
    'Common Parts' / '12 High Street' -> '12 High Street – Common parts',
    'Flat 1, Rm A' / 'Unity House' -> 'Flat 1, Rm A, Unity House'."""
    parts = [" ".join(p.split()) for p in lines if p and p.strip()]
    parts = [p for p in parts if p.lower() != "london"]
    if not parts:
        return ""
    first = parts[0]
    common = re.fullmatch(r"(?:common parts?|communal areas?)\b[\s,]*(.*)", first, re.I)
    if common:
        where = common.group(1) or (parts[1] if len(parts) > 1 else "")
        return f"{where} – Common parts" if where else "Common parts"
    common = re.fullmatch(r"(.+?)\s*[-–]\s*(?:common parts?|communal areas?|common)", first, re.I)
    if common:  # 'Example Court - Common'
        return f"{common.group(1)} – Common parts"
    if UNIT_ONLY.fullmatch(first) and len(parts) > 1:
        return f"{first}, {parts[1]}"
    return first


def read_property_file(path):
    """The housing system's property list: Prop Key, Address Line 1-4, Postcode."""
    rows = read_sheet(path)
    for i, row in enumerate(rows):
        heads = [h.lower() for h in row]
        if any("address" in h or "prop" in h for h in heads):
            break
    else:
        raise ValueError("No column with 'address' or 'prop' in its heading was found.")
    key_col = next((j for j, h in enumerate(heads) if any(w in h for w in ("key", "code", "ref"))), None)
    line_cols = [j for j, h in enumerate(heads) if ("address" in h or "line" in h) and j != key_col]
    post_col = next((j for j, h in enumerate(heads) if "post" in h), None)
    out, seen = [], set()
    for row in rows[i + 1:]:
        row = row + [""] * (len(heads) - len(row))
        name = property_label([row[j] for j in line_cols])
        key = row[key_col] if key_col is not None else ""
        # the same place twice under two codes (e.g. a block's common parts) is listed once
        if not name or name.lower() == "suspense" or {key.lower(), name.lower()} & seen:
            continue
        seen.update({key.lower(), name.lower()} - {""})
        out.append({"key": key, "name": name, "postcode": " ".join(row[post_col].split()) if post_col is not None
                    else ""})
    return out


COMPANY_WORDS = ("company", "firm", "business", "organisation", "organization")


def read_contractor_file(path):
    """A contractor or supplier list. Either a company column plus a person's name column, or (like
    the housing system's supplier list: Supplier, Name, Address, Postcode, Supplier Contact) a Name
    column holding the company and a Contact column holding the person. Phone if there is one."""
    rows = read_sheet(path)
    for i, row in enumerate(rows):
        heads = [h.lower() for h in row]
        if any("name" in h or "contractor" in h or "compan" in h for h in heads):
            break
    else:
        raise ValueError("No column with 'name', 'company' or 'contractor' in its heading was found.")

    def col(words, avoid=()):
        return next((j for j, h in enumerate(heads) if any(w in h for w in words)
                     and not any(w in h for w in avoid)), None)

    phone = col(("phone", "mobile", "tel"))
    company = col(COMPANY_WORDS)
    contact = col(("contact", "person"), avoid=("no", "number", "phone", "mobile", "tel", "email"))
    name = col(("name", "contractor"), avoid=COMPANY_WORDS)
    if company is None and contact is not None:
        company, name = name, contact  # 'Name' is the supplier; 'Supplier Contact' is the person
    out, seen = [], set()
    for row in rows[i + 1:]:
        row = row + [""] * (len(heads) - len(row))
        c = {k: " ".join(row[j].split()).rstrip(",") if j is not None else ""
             for k, j in (("name", name), ("company", company), ("phone", phone))}
        if (c["name"] or c["company"]) and contractor_key(c) not in seen:
            seen.add(contractor_key(c))
            out.append(c)
    return out


def contractor_key(c):
    return c["name"].lower(), c["company"].lower()


def sort_contractors(items):
    items.sort(key=lambda c: (natural_key(c["company"] or c["name"]), natural_key(c["name"])))


# --------------------------------------------------------------------------
# Slip PDF
# --------------------------------------------------------------------------

def slip_path(loan):
    return os.path.join(SLIPS_DIR, loan["dt"].strftime("%Y"), f"Key_slip_{slip_name(loan['slip'])}.pdf")


def make_slip_pdf(path, loan, org):
    """One A4 page: what was lent to whom, until when, and a place for the contractor to sign."""
    if canvas is None:
        raise RuntimeError("The PDF library (reportlab) is not available.")
    name = slip_name(loan["slip"])
    info = [("Slip", name), ("Issued", loan["dt"].strftime("%d/%m/%Y %H:%M")),
            ("Return by", loan["return_by"].strftime("%d/%m/%Y") if loan["return_by"] else "—"),
            ("Issued by", loan["issued_by"])]
    c, new_page, footer = _sheet(path, "KEY ISSUE SLIP", f"Key slip {name}", info,
                                 f"Key slip {name} · printed {datetime.now():%d/%m/%Y %H:%M} "
                                 "from the Front Office key log", org)
    y = _letterhead(c, org, "KEY ISSUE SLIP", info)
    ink = HexColor(INK)
    grey = HexColor("#6b6b73")
    left = SHEET_LEFT * mm
    width = SHEET_WIDTH * mm
    prop = loan["property"] + (f"  ({loan['prop_key']})" if loan["prop_key"] else "")
    for label, value in (("Property", prop), ("Issued to", loan["contractor"]), ("Company", loan["company"]),
                         ("Phone", loan["phone"]), ("Reason", loan["reason"])):
        if not value:
            continue
        c.setFillColor(grey)
        c.setFont("Helvetica", 9)
        c.drawString(left, y, label)
        c.setFillColor(ink)
        c.setFont("Helvetica-Bold", 11)
        for line in simpleSplit(value, "Helvetica-Bold", 11, width - 25 * mm):
            c.drawString(left + 25 * mm, y, line)
            y -= 5.5 * mm
        y -= 1.5 * mm
    y -= 5 * mm
    y = _table(c, y, [("Key", 144, "l"), ("Quantity", 30, "r")],
               [[k, q] for k, q in loan["keys"].items()],
               total=["Total keys", key_count(loan["keys"])], new_page=new_page)

    y -= 8 * mm
    due = loan["return_by"].strftime("%A %d %B %Y") if loan["return_by"] else "the date agreed"
    c.setFillColor(ink)
    c.setFont("Helvetica", 9.5)
    text = (f"I have received the keys listed above. I will return them to the office by {due}. "
            "I will not copy them or give them to anyone else, and I will tell the office straight away if "
            "any are lost.")
    for line in simpleSplit(text, "Helvetica", 9.5, width):
        c.drawString(left, y, line)
        y -= 5 * mm

    def sign_line(y, label, until=left + 118 * mm, date_field=True):
        c.setFont("Helvetica", 9.5)
        c.setFillColor(ink)
        c.drawString(left, y, label)
        start = left + c.stringWidth(label, "Helvetica", 9.5) + 3 * mm
        c.setStrokeColor(HexColor("#8a8a92"))
        c.setLineWidth(0.6)
        c.line(start, y - 1 * mm, until, y - 1 * mm)
        if date_field:
            c.drawString(left + 124 * mm, y, "Date")
            c.line(left + 133 * mm, y - 1 * mm, left + width, y - 1 * mm)

    y -= 12 * mm
    sign_line(y, "Signed (contractor)")
    y -= 12 * mm
    sign_line(y, "Name in capitals", date_field=False)

    y -= 18 * mm
    c.setStrokeColor(HexColor("#c9c9cf"))
    c.setLineWidth(0.8)
    c.roundRect(left - 3 * mm, y - 24 * mm, width + 6 * mm, 31 * mm, 2 * mm)
    c.setFillColor(grey)
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(left, y, "OFFICE USE · KEYS RETURNED")
    y -= 10 * mm
    sign_line(y, "All keys returned: received by")
    y -= 11 * mm
    sign_line(y, "Signed (contractor)")
    footer()
    c.showPage()
    c.save()


# --------------------------------------------------------------------------
# Store: the log plus derived loans
# --------------------------------------------------------------------------

class KeyStore(ChainLog):
    def __init__(self, app):
        self._loans = None
        super().__init__(app, KEY_LOG_PATH, KEY_HEADERS, KEY_SALT, "key_log_tip", "key log",
                         (K_ISSUED, K_RETURNED, K_EXTENDED, K_LOST, K_VOIDED))
        self.properties = load_list(PROPERTIES_PATH, PROPERTY_COLUMNS)
        self.contractors = load_list(CONTRACTORS_PATH, CONTRACTOR_COLUMNS)
        sort_contractors(self.contractors)

    def changed(self):
        self._loans = None

    def save_properties(self):
        save_list(PROPERTIES_PATH, PROPERTY_COLUMNS, self.properties)

    def save_contractors(self):
        sort_contractors(self.contractors)
        save_list(CONTRACTORS_PATH, CONTRACTOR_COLUMNS, self.contractors)

    def loans(self):
        """slip number -> loan, built from the log's entries."""
        if self._loans is not None:
            return self._loans
        loans = {}
        for e in self.entries:
            n = parse_int(e["Slip No"], default=None)
            if e["Type"] == K_ISSUED and n is not None:
                keys = text_to_keys(e["Keys"])
                loans[n] = {
                    "slip": n, "dt": e["dt"], "property": e["Property"], "prop_key": e["Prop Key"], "keys": keys,
                    "out": dict(keys), "returned": {}, "lost": {}, "contractor": e["Contractor"],
                    "company": e["Company"], "phone": e["Phone"], "reason": e["Reason"],
                    "return_by": parse_ui_date(e["Return By"]), "first_due": parse_ui_date(e["Return By"]),
                    "issued_by": e["Staff"], "events": [],
                    "voided_at": None, "voided_by": "", "void_reason": "", "tampered": e["tampered"],
                }
        for e in self.entries:
            loan = loans.get(parse_int(e["Slip No"], default=None))
            if loan is None or e["Type"] == K_ISSUED:
                continue
            loan["tampered"] = loan["tampered"] or e["tampered"]
            loan["events"].append(e)
            if e["Type"] in (K_RETURNED, K_LOST):
                into = loan["returned" if e["Type"] == K_RETURNED else "lost"]
                for k, q in text_to_keys(e["Keys"]).items():
                    q = min(q, loan["out"].get(k, 0))
                    if q:
                        into[k] = into.get(k, 0) + q
                        loan["out"][k] -= q
                        if not loan["out"][k]:
                            del loan["out"][k]
            elif e["Type"] == K_EXTENDED:
                loan["return_by"] = parse_ui_date(e["Return By"]) or loan["return_by"]
            elif e["Type"] == K_VOIDED:
                loan.update(voided_at=e["dt"], voided_by=e["Staff"], void_reason=e["Notes"])
        self._loans = loans
        return loans

    @staticmethod
    def status(loan, today=None):
        if loan["voided_at"]:
            return S_VOIDED
        if loan["out"]:
            due = loan["return_by"]
            return S_OVERDUE if due and due < (today or date.today()) else S_OUT
        return S_LOST if loan["lost"] else S_RETURNED

    def next_slip(self):
        return max([0] + list(self.loans())) + 1

    def out_now(self):
        return [l for l in self.loans().values() if self.status(l) in (S_OUT, S_OVERDUE)]

    def find_property(self, text):
        t = text.strip().lower()
        return next((p for p in self.properties if p["name"].lower() == t), None)

    def find_contractor(self, name, company=""):
        return next((c for c in self.contractors if contractor_key(c) == (name.lower(), company.lower())), None)


def last_event(loan, *types):
    return next((e for e in reversed(loan["events"]) if e["Type"] in types), None)


def status_text(loan):
    s = KeyStore.status(loan)
    if s == S_VOIDED:
        return f"Voided · {loan['void_reason']}"
    if s in (S_OUT, S_OVERDUE):
        due = loan["return_by"]
        if s == S_OVERDUE:
            days = (date.today() - due).days
            text = f"OVERDUE · was due {due:%d/%m/%Y} ({plural(days, 'day')})"
        else:
            text = f"Out · due {due:%d/%m/%Y}" if due else "Out"
        if loan["returned"] or loan["lost"]:
            text += f" · {key_count(loan['out'])} of {key_count(loan['keys'])} still out"
        return text
    back = last_event(loan, K_RETURNED)
    if s == S_LOST:
        lost = last_event(loan, K_LOST)
        return f"Lost: {keys_display(loan['lost'])}" + (f" · {lost['Notes']}" if lost and lost["Notes"] else "")
    return f"Returned {back['dt']:%d/%m/%Y} · to {back['Staff']}" if back else "Returned"


# --------------------------------------------------------------------------
# Dialogs
# --------------------------------------------------------------------------

class KeyCountsDialog(Dialog):
    """Which of a slip's keys, and how many: for a return or a loss.
    result = {"keys": {name: qty}, "person": str, "note": str}."""

    def __init__(self, parent, title, heading, sub, keys, person_label=None, person="", note_label=None,
                 confirm_text="Next"):
        super().__init__(parent, title)
        self.heading(heading, sub)
        self.label("Keys" if len(keys) == 1 and sum(keys.values()) == 1 else "Keys (set how many of each)")
        box = tk.Frame(self.body, bg=PANEL_2)
        box.pack(fill="x")
        self.spins = {}
        for i, (name, qty) in enumerate(keys.items()):
            row = tk.Frame(box, bg=PANEL_2)
            row.pack(fill="x", padx=14, pady=(10 if i == 0 else 4, 0))
            tk.Label(row, text=name, font=(FONT, 11), fg=TEXT, bg=PANEL_2, anchor="w").pack(side="left")
            tk.Label(row, text=f"of {qty}", font=(FONT, 10), fg=MUTED, bg=PANEL_2, width=5).pack(side="right")
            var = tk.StringVar(value=str(qty))
            tk.Spinbox(row, from_=0, to=qty, width=3, textvariable=var, justify="center", font=(FONT, 11),
                       state="readonly", readonlybackground=PANEL, fg=TEXT, buttonbackground=GREY_BTN,
                       relief="flat", highlightthickness=1, highlightbackground=BORDER).pack(side="right")
            self.spins[name] = (var, qty)
        tk.Frame(box, bg=PANEL_2, height=10).pack()
        self.person = self.note = None
        if person_label:
            self.label(person_label)
            self.person = tk.Entry(self.body, width=40, **entry_opts())
            self.person.insert(0, person)
            self.person.pack(fill="x", ipady=5)
            self._initial_focus = self.person
        if note_label:
            self.label(note_label)
            self.note = tk.Entry(self.body, width=40, **entry_opts())
            self.note.pack(fill="x", ipady=5)
            self._initial_focus = self._initial_focus or self.note
        self.error = tk.Label(self.body, text="", font=(FONT, 10), fg=RED, bg=PANEL, anchor="w", justify="left",
                              wraplength=380)
        self.error.pack(fill="x", pady=(8, 0))
        self.button_row(confirm_text)

    def confirm(self):
        counts = {n: max(0, min(q, parse_int(v.get()))) for n, (v, q) in self.spins.items()}
        counts = {n: c for n, c in counts.items() if c}
        person = self.person.get().strip() if self.person is not None else ""
        note = self.note.get().strip() if self.note is not None else ""
        if not counts:
            self.error.config(text="Set at least one key.")
        elif self.person is not None and not person:
            self.error.config(text="Enter who brought the keys back.")
        elif self.note is not None and not note:
            self.error.config(text="Please say what happened / what was done.")
        else:
            self.result = {"keys": counts, "person": person, "note": note}
            self.close()


# --------------------------------------------------------------------------
# Main view (the "Keys" tab)
# --------------------------------------------------------------------------

class KeyView(tk.Frame):
    PROPERTY_ROWS = 4
    CONTRACTOR_ROWS = 3

    def __init__(self, master, app):
        super().__init__(master, bg=BG)
        self.app = app
        self.store = app.keys
        self.chosen = {}         # key name -> how many, in the order picked
        self.other_key = None    # the "Other…" key's name, once typed
        self.kind_buttons = {}
        self.filter = "out"
        self._props_shown, self._contractors_shown = [], []
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self._build_form()
        self._build_main()
        self.clear_form()
        self.refresh()

    # ---- left: issue keys ------------------------------------------------
    def _section(self, parent, number, title, top=14):
        row = tk.Frame(parent, bg=PANEL)
        row.pack(fill="x", pady=(top, 6))
        tk.Label(row, text=number, font=(FONT, 9, "bold"), fg=BG, bg=MUTED, width=2).pack(side="left")
        tk.Label(row, text=title.upper(), font=(FONT, 10, "bold"), fg=MUTED, bg=PANEL).pack(side="left", padx=8)
        return row

    def _link(self, parent, text, command):
        link = tk.Label(parent, text=text, font=(FONT, 10, "underline"), fg=MUTED, bg=PANEL, cursor="hand2")
        link.pack(side="right")
        link.bind("<Button-1>", lambda e: command())
        return link

    def _listbox(self, parent, rows):
        return tk.Listbox(parent, height=rows, bg=PANEL_2, fg=TEXT, font=(FONT, 10), relief="flat",
                          highlightthickness=0, bd=0, selectbackground=TEAL, selectforeground="#ffffff",
                          activestyle="none", exportselection=False)

    def _build_form(self):
        S = self.app.S
        side = tk.Frame(self, bg=PANEL, width=S(420))
        side.grid(row=0, column=0, sticky="ns")
        side.pack_propagate(False)
        p = tk.Frame(side, bg=PANEL)
        p.pack(fill="both", expand=True, padx=S(22), pady=S(12))

        head = self._section(p, "1", "Property", top=0)
        self._link(head, "Clear all", self.clear_form)
        self.prop_entry = PlaceholderEntry(p, "Type to search: address, postcode or code")
        self.prop_entry.pack(fill="x", ipady=5)
        self.prop_entry.bind("<KeyRelease>", lambda e: self._update_props())
        self.prop_entry.bind("<Down>", lambda e: self._focus_list(self.prop_list, self._props_shown))
        self.prop_list = self._listbox(p, self.PROPERTY_ROWS)
        self.prop_list.pack(fill="x", pady=(2, 0))
        self.prop_list.bind("<<ListboxSelect>>", lambda e: self._pick_property())
        self.prop_list.bind("<Return>", lambda e: self._pick_property())
        self.prop_note = tk.Label(p, text="", font=(FONT, 9), fg=MUTED, bg=PANEL, anchor="w")
        self.prop_note.pack(fill="x")

        head = self._section(p, "2", "Keys")
        self._link(head, "Clear keys", self.clear_keys)
        grid = tk.Frame(p, bg=PANEL)
        grid.pack(fill="x")
        for i, kind in enumerate(KEY_KINDS + ["Other…"]):
            grid.columnconfigure(i % 3, weight=1, uniform="key")
            btn = HoverButton(grid, text=kind, font=(FONT, 9, "bold"), padx=0, pady=5, wraplength=S(120),
                              command=lambda k=kind: self.add_key(k))
            btn.grid(row=i // 3, column=i % 3, sticky="nsew", padx=2, pady=2)
            self.kind_buttons[kind] = btn
        self.keys_note = tk.Label(p, text="", font=(FONT, 9), fg=MUTED, bg=PANEL, anchor="w", justify="left",
                                  wraplength=S(370))
        self.keys_note.pack(fill="x", pady=(2, 0))

        self._section(p, "3", "Contractor")
        self.company_entry = PlaceholderEntry(p, "Company: type to search (or a person's name)")
        self.company_entry.pack(fill="x", ipady=5)
        self.company_entry.bind("<KeyRelease>", lambda e: self._update_contractors())
        self.company_entry.bind("<Down>", lambda e: self._focus_list(self.contractor_list, self._contractors_shown))
        self.contractor_list = self._listbox(p, self.CONTRACTOR_ROWS)
        self.contractor_list.pack(fill="x", pady=(2, 0))
        self.contractor_list.bind("<<ListboxSelect>>", lambda e: self._pick_contractor())
        self.contractor_list.bind("<Return>", lambda e: self._pick_contractor(move_on=True))
        row = tk.Frame(p, bg=PANEL)
        row.pack(fill="x", pady=(6, 0))
        row.columnconfigure(0, weight=3)
        row.columnconfigure(1, weight=2)
        self.name_entry = PlaceholderEntry(row, "Name of person collecting")
        self.name_entry.grid(row=0, column=0, sticky="ew", ipady=4, padx=(0, 6))
        self.phone_entry = PlaceholderEntry(row, "Phone")
        self.phone_entry.grid(row=0, column=1, sticky="ew", ipady=4)
        for e in (self.name_entry, self.phone_entry):
            e.bind("<KeyRelease>", lambda ev: self.update_form())
        self.add_contractor_var = tk.BooleanVar(value=True)
        self.add_contractor_check = tk.Checkbutton(
            p, text="New person: add to the contractor list", variable=self.add_contractor_var,
            font=(FONT, 9), fg=MUTED, bg=PANEL, selectcolor=PANEL_2, activebackground=PANEL,
            activeforeground=TEXT, anchor="w", bd=0, highlightthickness=0)
        self.add_contractor_check.pack(fill="x", pady=(2, 0))

        self._section(p, "4", "Return by")
        row = tk.Frame(p, bg=PANEL)
        row.pack(fill="x")
        self.due_entry = tk.Entry(row, width=11, justify="center", **entry_opts())
        self.due_entry.pack(side="left", ipady=4)
        self.due_entry.bind("<KeyRelease>", lambda e: self.update_form())
        for text, days in (("+1 week", 7), ("Tomorrow", 1), ("Today", 0)):
            HoverButton(row, text=text, font=(FONT, 9, "bold"), padx=8, pady=3,
                        command=lambda d=days: self.set_due(d)).pack(side="right", padx=(4, 0))
        self.reason_entry = PlaceholderEntry(p, "Reason (optional), e.g. boiler service")
        self.reason_entry.pack(fill="x", ipady=5, pady=(8, 0))
        self.print_var = tk.BooleanVar(value=True)
        tk.Checkbutton(p, text="Print a slip for the contractor to sign", variable=self.print_var, font=(FONT, 10),
                       fg=TEXT, bg=PANEL, selectcolor=PANEL_2, activebackground=PANEL, activeforeground=TEXT,
                       anchor="w", bd=0, highlightthickness=0, command=self.update_form).pack(fill="x", pady=(8, 0))

        self.problem_label = tk.Label(p, text="", font=(FONT, 9), fg=MUTED, bg=PANEL, anchor="w")
        self.problem_label.pack(fill="x", pady=(6, 0))
        self.issue_btn = HoverButton(p, text="Issue keys", bg=GREEN, hover=GREEN_HOVER, font=(FONT, 13, "bold"),
                                     pady=10, command=self.issue)
        self.issue_btn.pack(fill="x", pady=(4, 0))

    # ---- property search -------------------------------------------------
    def _update_props(self):
        q = self.prop_entry.value().lower()
        props = self.store.properties
        words = q.split()
        self._props_shown = [p for p in props if all(w in f"{p['name']} {p['postcode']} {p['key']}".lower()
                                                     for w in words)] if q else props
        self.prop_list.delete(0, "end")
        for p in self._props_shown:
            self.prop_list.insert("end", "  ·  ".join(x for x in (p["name"], p["postcode"]) if x))
        if not props:
            self.prop_list.insert("end", "No property list yet. Type the property, or import")
            self.prop_list.insert("end", "the list in Settings → Properties.")
            self._props_shown = []
        elif not self._props_shown:
            self.prop_list.insert("end", "Not on the list: it will be recorded as typed.")
        self.update_form()

    def _focus_list(self, box, shown):
        if shown:
            box.focus_set()
            box.selection_clear(0, "end")
            box.selection_set(0)
            box.activate(0)

    def _pick_property(self):
        sel = self.prop_list.curselection()
        if sel and sel[0] < len(self._props_shown):
            self.prop_entry.set(self._props_shown[sel[0]]["name"])
            self.update_form()

    def _form_property(self):
        """(name, prop key): from the list when it matches what's typed, otherwise as typed."""
        text = self.prop_entry.value()
        p = self.store.find_property(text)
        return (p["name"], p["key"]) if p else (text, "")

    # ---- keys ------------------------------------------------------------
    def add_key(self, kind):
        if kind == "Other…":
            res = FieldsDialog(self.app, "Other key", "Which key?", [
                {"key": "name", "label": "What it opens, e.g. plant room, roof hatch, window lock",
                 "value": self.other_key or ""},
                {"key": "qty", "label": "How many", "value": str(self.chosen.get(self.other_key, 1))},
            ], validate=lambda v: "Please say what the key opens." if not v["name"] else
                "Enter how many, e.g. 1" if parse_int(v["qty"], 0) < 1 else None, confirm_text="Add").show()
            if not res:
                return
            if self.other_key:
                self.chosen.pop(self.other_key, None)
            self.other_key = " ".join(res["name"].replace(";", ",").split())
            self.other_key = self.other_key[0].upper() + self.other_key[1:]
            self.chosen[self.other_key] = parse_int(res["qty"], 1)
        else:
            self.chosen[kind] = self.chosen.get(kind, 0) + 1
        self._paint_keys()

    def clear_keys(self):
        self.chosen.clear()
        self.other_key = None
        self._paint_keys()

    def _paint_keys(self):
        for kind, btn in self.kind_buttons.items():
            name = self.other_key if kind == "Other…" else kind
            qty = self.chosen.get(name, 0) if name else 0
            if qty:
                label = name if kind == "Other…" else kind
                btn.config(text=f"{label}  ×{qty}")
                btn.set_colors(TEAL, TEAL_HOVER, "#ffffff")
            else:
                btn.config(text=kind)
                btn.set_colors(GREY_BTN, GREY_HOVER, TEXT)
        n = key_count(self.chosen)
        self.keys_note.config(text=f"{plural(n, 'key')}: {keys_display(self.chosen)}" if n else
                              "Tap a key to add it. Tap again for another of the same.")
        self.update_form()

    # ---- contractor search -------------------------------------------------
    def _update_contractors(self):
        q = self.company_entry.value().lower()
        cs = self.store.contractors
        self._contractors_shown = [c for c in cs if q in c["name"].lower() or q in c["company"].lower()] if q else cs
        self.contractor_list.delete(0, "end")
        for c in self._contractors_shown:
            self.contractor_list.insert("end", "  ·  ".join(x for x in (c["company"], c["name"], c["phone"]) if x))
        if not cs:
            self.contractor_list.insert("end", "No contractor list yet. Type their details, or")
            self.contractor_list.insert("end", "import the list in Settings → Contractors.")
            self._contractors_shown = []
        elif not self._contractors_shown:
            self.contractor_list.insert("end", "Not on the list: recorded as typed.")
        self.update_form()

    def _pick_contractor(self, move_on=False):
        sel = self.contractor_list.curselection()
        if sel and sel[0] < len(self._contractors_shown):
            c = self._contractors_shown[sel[0]]
            self.company_entry.set(c["company"])  # replaces what was typed to search
            self.name_entry.set(c["name"])
            self.phone_entry.set(c["phone"])
            if move_on or not c["name"]:
                self.name_entry.focus_set()
            self.update_form()

    def _new_contractor(self):
        name = self.name_entry.value()
        return bool(name) and self.store.find_contractor(name, self.company_entry.value()) is None

    # ---- form state --------------------------------------------------------
    def set_due(self, days):
        self.due_entry.delete(0, "end")
        self.due_entry.insert(0, (date.today() + timedelta(days=days)).strftime(UI_DATE_FMT))
        self.update_form()

    def form_problem(self):
        due = parse_ui_date(self.due_entry.get())
        if not self.prop_entry.value():
            return "Choose or type the property"
        if not self.chosen:
            return "Choose which keys"
        if not self.name_entry.value():
            return "Enter the name of the person collecting the keys"
        if due is None:
            return "Enter the return-by date as DD/MM/YYYY"
        if due < date.today():
            return "The return-by date can't be in the past"
        return None

    def update_form(self):
        if not hasattr(self, "issue_btn"):
            return
        name, key = self._form_property()
        self.prop_note.config(text=(f"✓ {key}" if key else "Not on the property list: recorded as typed")
                              if name else "")
        new = self._new_contractor()
        self.add_contractor_check.config(state="normal" if new else "disabled", fg=TEXT if new else DIM)
        problem = self.form_problem()
        self.problem_label.config(text=problem or f"Slip {slip_name(self.store.next_slip())} will be issued")
        self.issue_btn.set_enabled(problem is None)
        self.issue_btn.config(text="Issue keys  ·  Print slip" if self.print_var.get() else "Issue keys")

    def clear_form(self):
        for e in (self.prop_entry, self.name_entry, self.company_entry, self.phone_entry, self.reason_entry):
            e.clear()
        self.add_contractor_var.set(True)
        self.print_var.set(True)
        self.chosen.clear()
        self.other_key = None
        self.set_due(0)
        self._update_props()
        self._update_contractors()
        self._paint_keys()

    def reset(self):
        """Back to a fresh screen: empty form, no search, 'Out now' filter."""
        self.clear_form()
        self.search.clear()
        self.set_filter("out")

    # ---- right: cards and table ---------------------------------------------
    def _build_main(self):
        S = self.app.S
        main = tk.Frame(self, bg=BG)
        main.grid(row=0, column=1, sticky="nsew", padx=S(24), pady=S(14))
        main.columnconfigure(0, weight=1)
        main.rowconfigure(2, weight=1)

        cards = tk.Frame(main, bg=BG)
        cards.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        self.cards = {}
        for i, (key, title, accent) in enumerate([("out", "Keys out now", TEAL), ("overdue", "Overdue", RED),
                                                   ("today", "Issued today", GREEN),
                                                   ("back", "Returned today", ORANGE)]):
            cards.columnconfigure(i, weight=1, uniform="card")
            card = tk.Frame(cards, bg=PANEL)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else S(12), 0))
            bar = tk.Frame(card, bg=accent, height=3)
            bar.pack(fill="x")
            tk.Label(card, text=title.upper(), font=(FONT, 9, "bold"), fg=MUTED, bg=PANEL,
                     anchor="w").pack(fill="x", padx=16, pady=(12, 0))
            value = tk.Label(card, text="0", font=(FONT, 24, "bold"), fg=TEXT, bg=PANEL, anchor="w")
            value.pack(fill="x", padx=16)
            sub = tk.Label(card, text="", font=(FONT, 9), fg=MUTED, bg=PANEL, anchor="w")
            sub.pack(fill="x", padx=16, pady=(0, 12))
            self.cards[key] = (value, sub, bar, accent)

        bar = tk.Frame(main, bg=BG)
        bar.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.filter_buttons = {}
        for key, text in (("out", "Out now"), ("overdue", "Overdue"), ("returned", "Returned"), ("lost", "Lost"),
                          ("voided", "Voided"), ("all", "All")):
            btn = HoverButton(bar, text=text, font=(FONT, 9, "bold"), padx=10, pady=4,
                              command=lambda k=key: self.set_filter(k))
            btn.pack(side="left", padx=(0, 4))
            self.filter_buttons[key] = btn
        for text, bg, hover, cmd in [("Export", GREY_BTN, GREY_HOVER, self.export),
                                     ("Void", RED, RED_HOVER, self.void_selected),
                                     ("Lost…", GREY_BTN, GREY_HOVER, self.lost_selected),
                                     ("Reprint slip", GREY_BTN, GREY_HOVER, self.reprint_selected),
                                     ("Extend…", GREY_BTN, GREY_HOVER, self.extend_selected),
                                     ("Return keys", GREEN, GREEN_HOVER, self.return_selected)]:
            HoverButton(bar, text=text, bg=bg, hover=hover, font=(FONT, 10, "bold"),
                        command=cmd).pack(side="right", padx=(8, 0))
        self.search = PlaceholderEntry(bar, "Search…", width=16)
        self.search.pack(side="right", padx=(8, 4), ipady=4)
        self.search.bind("<KeyRelease>", lambda e: self.refresh_table())

        table = tk.Frame(main, bg=PANEL)
        table.grid(row=2, column=0, sticky="nsew")
        columns = [("slip", "Slip", 70, "w"), ("date", "Issued", 90, "w"), ("property", "Property", 230, "w"),
                   ("keys", "Keys", 190, "w"), ("contractor", "Issued to", 170, "w"), ("due", "Return by", 90, "w"),
                   ("by", "Issued by", 100, "w"), ("status", "Status", 230, "w")]
        self.tree = ttk.Treeview(table, columns=[c[0] for c in columns], show="headings", style="Dark.Treeview",
                                 selectmode="browse")
        for key, text, width, anchor in columns:
            self.tree.heading(key, text=text, anchor=anchor)
            self.tree.column(key, width=S(width), minwidth=S(40), anchor=anchor,
                             stretch=key in ("property", "keys", "contractor", "status"))
        sb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("tampered", background=TAMPERED_BG, foreground=TAMPERED_FG)  # first = wins
        self.tree.tag_configure("even", background=ROW_A)
        self.tree.tag_configure("odd", background=ROW_B)
        self.tree.tag_configure(S_OUT, foreground=TEXT)
        self.tree.tag_configure(S_OVERDUE, foreground=RED)
        self.tree.tag_configure(S_RETURNED, foreground=MUTED)
        self.tree.tag_configure(S_LOST, foreground=AMBER)
        self.tree.tag_configure(S_VOIDED, foreground=DIM, font=(FONT, 10, "overstrike"))
        self.tree.bind("<Double-1>", lambda e: self.open_details())
        make_sortable(self.tree)
        self.empty_label = tk.Label(table, text="", font=(FONT, 11), fg=MUTED, bg=ROW_A)
        self.set_filter("out")

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
        self.update_form()

    def refresh_cards(self):
        loans = [l for l in self.store.loans().values() if not l["voided_at"]]
        today = date.today()
        out = self.store.out_now()
        late = [l for l in out if self.store.status(l) == S_OVERDUE]
        value, sub, _, _ = self.cards["out"]
        value.config(text=str(sum(key_count(l["out"]) for l in out)))
        sub.config(text=f"on {plural(len(out), 'slip')}" if out else "All keys are in")
        value, sub, bar, accent = self.cards["overdue"]
        value.config(text=str(len(late)), fg=RED if late else TEXT)
        oldest = min((l["return_by"] for l in late), default=None)
        sub.config(text=f"{plural(len(late), 'slip')} · due back {oldest:%d/%m}" if late else "Nothing overdue")
        issued = [l for l in loans if l["dt"].date() == today]
        value, sub, _, _ = self.cards["today"]
        value.config(text=str(sum(key_count(l["keys"]) for l in issued)))
        sub.config(text=plural(len(issued), "slip"))
        back = [e for l in loans for e in l["events"] if e["Type"] == K_RETURNED and e["dt"].date() == today]
        value, sub, _, _ = self.cards["back"]
        value.config(text=str(sum(key_count(text_to_keys(e["Keys"])) for e in back)))
        sub.config(text=plural(len({e["Slip No"] for e in back}), "slip"))

    def _row_values(self, l):
        keys = keys_display(l["keys"])
        if l["out"] and (l["returned"] or l["lost"]):
            keys = f"{keys_display(l['out'])} still out (of {key_count(l['keys'])})"
        who = l["contractor"] + (f" ({l['company']})" if l["company"] else "")
        return (slip_name(l["slip"]), l["dt"].strftime("%d/%m/%Y"), l["property"], keys, who,
                l["return_by"].strftime("%d/%m/%Y") if l["return_by"] else "", l["issued_by"], status_text(l))

    def refresh_table(self):
        query = self.search.value().lower()
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        wanted = {"out": (S_OUT, S_OVERDUE), "overdue": (S_OVERDUE,), "returned": (S_RETURNED,), "lost": (S_LOST,),
                  "voided": (S_VOIDED,), "all": (S_OUT, S_OVERDUE, S_RETURNED, S_LOST, S_VOIDED)}[self.filter]
        shown = 0
        for l in sorted(self.store.loans().values(), key=lambda l: l["slip"], reverse=True):
            status = self.store.status(l)
            if status not in wanted:
                continue
            values = self._row_values(l)
            if query and not any(query in str(v).lower() for v in values + (l["phone"], l["prop_key"], l["reason"])):
                continue
            tags = ["odd" if shown % 2 else "even", status] + (["tampered"] if l["tampered"] else [])
            self.tree.insert("", "end", iid=str(l["slip"]), values=values, tags=tags)
            shown += 1
        self.tree.resort()
        keep = [i for i in selected if self.tree.exists(i)]
        if keep:
            self.tree.selection_set(keep)
        if shown:
            self.empty_label.place_forget()
        else:
            self.empty_label.config(text="No matching keys." if query else {
                "out": "No keys are out. Everything is back.", "overdue": "Nothing overdue.",
                "returned": "Nothing returned yet.", "lost": "No keys reported lost.", "voided": "Nothing voided.",
                "all": "No keys issued yet."}[self.filter])
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")

    def selected_loan(self, what):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(what, "Select a slip in the list first.", parent=self.app)
            return None
        return self.store.loans().get(int(sel[0]))

    def _pin(self, title, heading, lines, **kw):
        dlg = PinDialog(self.app, self.app, title, heading, lines, **kw)
        result = dlg.show()
        if not result:
            self.app._pin_locked_out(dlg)
        return result

    def _done(self, text, color=GREEN):
        self.refresh()
        self.app.set_message(text, color)

    # ---- actions -----------------------------------------------------------
    def issue(self):
        if self.form_problem():
            return
        prop, prop_key = self._form_property()
        name, company, phone = self.name_entry.value(), self.company_entry.value(), self.phone_entry.value()
        due = parse_ui_date(self.due_entry.get())
        reason = self.reason_entry.value()
        keys = dict(self.chosen)
        slip = self.store.next_slip()
        printing = self.print_var.get()
        result = self._pin("Issue keys", f"Slip {slip_name(slip)}  ·  {plural(key_count(keys), 'key')}",
                           [keys_display(keys), prop + (f"  ({prop_key})" if prop_key else ""),
                            "To " + name + (f", {company}" if company else "") + (f"  ·  {phone}" if phone else ""),
                            f"Back by {due:%A %d/%m/%Y}" + (f"  ·  {reason}" if reason else "")],
                           confirm_text="Issue keys & print slip" if printing else "Issue keys")
        if not result:
            return
        staff = result[0]
        if not self.store.write([{"Type": K_ISSUED, "Slip No": slip, "Property": prop, "Prop Key": prop_key,
                                  "Keys": keys_to_text(keys), "Contractor": name, "Company": company,
                                  "Phone": phone, "Reason": reason, "Return By": due.strftime(UI_DATE_FMT),
                                  "Staff": staff["name"]}], self.app):
            return
        if self.add_contractor_var.get() and self._new_contractor():
            self.store.contractors.append({"name": name, "company": company, "phone": phone})
            try:
                self.store.save_contractors()
            except OSError as e:
                messagebox.showwarning("Contractors", f"The keys are logged, but the contractor list could not be "
                                       f"saved:\n\n{e}", parent=self.app)
        self.clear_form()
        self.set_filter("out")
        self._done(f"✓ Slip {slip_name(slip)}: {keys_display(keys)} issued to {name} by {staff['name']}, "
                   f"back by {due:%d/%m/%Y}")
        if printing:
            self._print_slip(self.store.loans()[slip])

    def _print_slip(self, loan, keep=True):
        path = slip_path(loan) if keep else os.path.join(
            tempfile.gettempdir(), f"Key_slip_{slip_name(loan['slip'])}_copy.pdf")
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            make_slip_pdf(path, loan, self.app.rent.org())
            open_file(path)
        except Exception as e:
            messagebox.showwarning("Key slip", (f"Slip {slip_name(loan['slip'])} IS logged, but the slip could not "
                                                f"be created or opened:\n\n{e}\n\nSelect it and press Reprint slip "
                                                "to try again.") if keep else
                                   f"Could not create the slip:\n\n{e}", parent=self.app)

    def reprint_selected(self):
        loan = self.selected_loan("Reprint slip")
        if loan:
            self._print_slip(loan, keep=False)

    def _open_loan(self, loan, what):
        """The loan if keys are still out on it, otherwise None (with a message)."""
        if loan and self.store.status(loan) not in (S_OUT, S_OVERDUE):
            messagebox.showinfo(what, f"Slip {slip_name(loan['slip'])} has no keys still out.", parent=self.app)
            return None
        return loan

    def return_selected(self):
        loan = self._open_loan(self.selected_loan("Return keys"), "Return keys")
        if not loan:
            return
        name = slip_name(loan["slip"])
        res = KeyCountsDialog(self.app, "Return keys", f"Return keys · slip {name}",
                              f"{loan['property']}. Issued {loan['dt']:%d/%m/%Y} to {loan['contractor']}. "
                              "Any keys not returned now stay out on this slip.",
                              loan["out"], person_label="Returned by", person=loan["contractor"]).show()
        if not res:
            return
        rest = {k: q - res["keys"].get(k, 0) for k, q in loan["out"].items() if q - res["keys"].get(k, 0)}
        result = self._pin("Return keys", f"Slip {name}  ·  {plural(key_count(res['keys']), 'key')} back",
                           [keys_display(res["keys"]), f"Returned by {res['person']}",
                            f"Still out: {keys_display(rest)}" if rest else "All keys on this slip are back"],
                           confirm_text="Record return")
        if result and self.store.write([{"Type": K_RETURNED, "Slip No": loan["slip"], "Property": loan["property"],
                                         "Keys": keys_to_text(res["keys"]), "Contractor": loan["contractor"],
                                         "Returned By": res["person"], "Staff": result[0]["name"]}], self.app):
            self._done(f"✓ Slip {name}: {keys_display(res['keys'])} returned by {res['person']}, received by "
                       f"{result[0]['name']}" + (f". Still out: {keys_display(rest)}" if rest else ""))

    def extend_selected(self):
        loan = self._open_loan(self.selected_loan("Extend"), "Extend")
        if not loan:
            return
        name = slip_name(loan["slip"])

        def validate(v):
            d = parse_ui_date(v["date"])
            if d is None:
                return "Enter the date as DD/MM/YYYY."
            if d < date.today():
                return "The new date can't be in the past."
            return None

        start = max(date.today(), loan["return_by"] or date.today()) + timedelta(days=7)
        res = FieldsDialog(self.app, "Extend", f"New return-by date · slip {name}", [
            {"key": "date", "label": "Keys to be back by (DD/MM/YYYY)", "value": start.strftime(UI_DATE_FMT)},
            {"key": "why", "label": "Reason (optional), e.g. job running over"},
        ], validate=validate, confirm_text="Next",
            sub=f"{keys_display(loan['out'])} with {loan['contractor']}. Was due "
                f"{loan['return_by']:%d/%m/%Y}." if loan["return_by"] else None).show()
        if not res:
            return
        d = parse_ui_date(res["date"])
        result = self._pin("Extend", f"Slip {name}  ·  back by {d:%d/%m/%Y}",
                           [keys_display(loan["out"]), f"{loan['contractor']}  ·  {loan['property']}"],
                           confirm_text="Change date", confirm_bg=TEAL, confirm_hover=TEAL_HOVER)
        if result and self.store.write([{"Type": K_EXTENDED, "Slip No": loan["slip"], "Property": loan["property"],
                                         "Contractor": loan["contractor"], "Return By": d.strftime(UI_DATE_FMT),
                                         "Staff": result[0]["name"], "Notes": res["why"]}], self.app):
            self._done(f"✓ Slip {name}: now due back {d:%d/%m/%Y}")

    def lost_selected(self):
        loan = self._open_loan(self.selected_loan("Lost keys"), "Lost keys")
        if not loan:
            return
        name = slip_name(loan["slip"])
        res = KeyCountsDialog(self.app, "Lost keys", f"Report lost keys · slip {name}",
                              f"{loan['property']}. With {loan['contractor']} since {loan['dt']:%d/%m/%Y}. "
                              "A lost key is a security risk: consider changing the lock.",
                              loan["out"], note_label="What happened / what was done (e.g. lock changed)",
                              confirm_text="Next").show()
        if not res:
            return
        result = self._pin("Lost keys", f"Slip {name}  ·  {plural(key_count(res['keys']), 'key')} lost",
                           [keys_display(res["keys"]), res["note"]],
                           confirm_text="Record as lost", confirm_bg=RED, confirm_hover=RED_HOVER)
        if result and self.store.write([{"Type": K_LOST, "Slip No": loan["slip"], "Property": loan["property"],
                                         "Keys": keys_to_text(res["keys"]), "Contractor": loan["contractor"],
                                         "Staff": result[0]["name"], "Notes": res["note"]}], self.app):
            self._done(f"Slip {name}: {keys_display(res['keys'])} recorded as lost", AMBER)

    def void_selected(self):
        loan = self.selected_loan("Void")
        if not loan:
            return
        name = slip_name(loan["slip"])
        if loan["voided_at"]:
            messagebox.showinfo("Void", f"Slip {name} is already voided.", parent=self.app)
            return
        if loan["returned"] or loan["lost"]:
            messagebox.showinfo("Void", f"Keys on slip {name} have already been returned or reported lost, so it "
                                "can't be voided.", parent=self.app)
            return
        result = self._pin("Void", f"Void slip {name}",
                           [f"{keys_display(loan['keys'])}  ·  {loan['property']}",
                            f"To {loan['contractor']} on {loan['dt']:%d/%m/%Y}",
                            "Only for a slip entered by mistake. It stays in the log, struck through."],
                           allow=lambda s: None if (s["admin"] or s.get("checker")) else
                           "Only an admin or checker can void a slip",
                           reason_label="Reason (e.g. wrong property, entered twice)",
                           confirm_text="Void slip", confirm_bg=RED, confirm_hover=RED_HOVER)
        if result and self.store.write([{"Type": K_VOIDED, "Slip No": loan["slip"], "Property": loan["property"],
                                         "Contractor": loan["contractor"], "Staff": result[0]["name"],
                                         "Notes": result[1]}], self.app):
            self._done(f"Slip {name} voided by {result[0]['name']}", AMBER)

    def open_details(self):
        loan = self.selected_loan("Slip")
        if loan:
            KeyDetailsWindow(self.app, self, loan).show()

    def export(self):
        path = admin_save_path(self.app, self.app, title="Export key log", defaultextension=".xlsx",
                               initialfile=f"Key log {date.today():%Y-%m-%d}.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        rows = []
        for l in sorted(self.store.loans().values(), key=lambda l: l["slip"]):
            back = last_event(l, K_RETURNED)
            rows.append([slip_name(l["slip"]), l["dt"].replace(second=0), l["property"], l["prop_key"],
                         keys_display(l["keys"]), keys_display(l["out"]) if l["out"] else "", l["contractor"],
                         l["company"], l["phone"], l["reason"], l["return_by"], l["issued_by"],
                         back["dt"].date() if back and not l["out"] else None, status_text(l)])
        try:
            write_xlsx(path, ["Slip", "Issued", "Property", "Prop Key", "Keys", "Still out", "Issued to", "Company",
                              "Phone", "Reason", "Return by", "Issued by", "All back on", "Status"], rows)
            open_file(path)
        except Exception as e:
            messagebox.showerror("Export", f"Could not export:\n\n{e}", parent=self.app)


class KeyDetailsWindow(Dialog):
    """Everything that happened to one slip."""

    def __init__(self, app, view, loan):
        super().__init__(app, "Key slip", resizable=True)
        S = app.S
        self.geometry(f"{S(760)}x{S(480)}")
        self.heading(f"Slip {slip_name(loan['slip'])}  ·  {loan['property']}",
                     f"{keys_display(loan['keys'])} to {loan['contractor']}"
                     + (f" ({loan['company']})" if loan["company"] else "")
                     + (f", {loan['phone']}" if loan["phone"] else "") + f". {status_text(loan)}.")
        frame = tk.Frame(self.body, bg=PANEL)
        frame.pack(fill="both", expand=True, pady=(14, 0))
        columns = [("when", "When", 130), ("what", "What", 110), ("keys", "Keys", 210), ("who", "Staff", 110),
                   ("notes", "Details", 220)]
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", style="Dark.Treeview",
                            selectmode="none")
        for key, text, width in columns:
            tree.heading(key, text=text, anchor="w")
            tree.column(key, width=S(width), anchor="w", stretch=key in ("keys", "notes"))
        tree.pack(fill="both", expand=True)
        tree.tag_configure("even", background=ROW_A)
        tree.tag_configure("odd", background=ROW_B)
        first = " · ".join(x for x in (f"Back by {loan['first_due']:%d/%m/%Y}" if loan["first_due"] else "",
                                       loan["reason"]) if x)
        rows = [(loan["dt"], "Issued", keys_display(loan["keys"]), loan["issued_by"], first)]
        for e in loan["events"]:
            detail = {K_RETURNED: f"Returned by {e['Returned By']}", K_EXTENDED: f"New date {e['Return By']}",
                      K_LOST: e["Notes"], K_VOIDED: e["Notes"]}.get(e["Type"], e["Notes"])
            if e["Type"] == K_EXTENDED and e["Notes"]:
                detail += f" · {e['Notes']}"
            rows.append((e["dt"], e["Type"].capitalize(), keys_display(text_to_keys(e["Keys"])) if e["Keys"] else "",
                         e["Staff"], detail))
        for i, (when, what, keys, who, notes) in enumerate(rows):
            tree.insert("", "end", values=(when.strftime("%d/%m/%Y %H:%M"), what, keys, who, notes),
                        tags=("odd" if i % 2 else "even",))
        foot = tk.Frame(self.body, bg=PANEL)
        foot.pack(fill="x", pady=(14, 0))
        HoverButton(foot, text="Close", command=self.cancel).pack(side="right")
        HoverButton(foot, text="Reprint slip", command=lambda: view._print_slip(loan, keep=False)).pack(
            side="right", padx=(0, 8))


# --------------------------------------------------------------------------
# Settings tabs
# --------------------------------------------------------------------------

def _list_tab(win, p, title, columns, items, save, dialog_fields, key_of, importer, note, after_save,
              validate=None):
    """A list with Edit / Add / Remove / Import from Excel. columns: [(field, heading, width)].
    validate(values) -> problem or None; by default the first field is required."""
    tree = win._tree(p, [(f, h, w) for f, h, w in columns])
    count = None

    def fill(select=None):
        tree.delete(*tree.get_children())
        for i, it in enumerate(items):
            iid = tree.insert("", "end", values=[it.get(f) or "—" for f, _, _ in columns],
                              tags=("odd" if i % 2 else "even",))
            if it is select:
                tree.selection_set(iid)
                tree.see(iid)
        count.config(text=f"{len(items)} {title.lower()}")

    def selected():
        sel = tree.selection()
        if not sel:
            messagebox.showinfo(title, "Select one in the list first.", parent=win)
            return None
        return items[tree.index(sel[0])]

    def do_save(select=None):
        try:
            save()
        except OSError as e:
            messagebox.showerror(title, f"Could not save the list:\n\n{e}", parent=win)
        fill(select)
        after_save()

    def dialog(heading, it=None):
        return FieldsDialog(win, title, heading, [dict(f, value=(it or {}).get(f["key"], "")) for f in dialog_fields],
                            validate=validate or (lambda v: None if v[dialog_fields[0]["key"]] else
                            f"Please enter the {dialog_fields[0]['label'].split(' (')[0].lower()}.")).show()

    def edit():
        it = selected()
        if it:
            res = dialog("Edit", it)
            if res:
                it.update(res)
                do_save(it)

    def add():
        res = dialog("Add")
        if res:
            it = dict(res)
            items.append(it)
            do_save(it)

    def remove():
        it = selected()
        if it and messagebox.askyesno(title, f"Remove {it[columns[0][0]] or it[columns[1][0]]} from the list?\n"
                                      "Entries already in the key log are not changed.", parent=win):
            items.remove(it)
            do_save()

    def import_file():
        path = filedialog.askopenfilename(parent=win, title=f"Import {title.lower()}",
                                          filetypes=[("Excel or CSV", "*.xlsx *.xlsm *.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            found = importer(path)
        except Exception as e:
            messagebox.showerror("Import", f"Could not read the file:\n\n{e}", parent=win)
            return
        known = {key_of(it) for it in items}
        new = [it for it in found if key_of(it) not in known]
        sample = "\n".join("  " + "  ·  ".join(it[f] for f, _, _ in columns if it.get(f)) for it in found[:4])
        choice = messagebox.askyesnocancel(
            f"Import {title.lower()}",
            f"The file has {len(found)} {title.lower()}. For example:\n\n{sample}\n\n"
            f"YES: replace the whole list with this file ({len(items)} on the list now).\n\n"
            f"NO: keep the current list and only add the {len(new)} new one(s).\n\n"
            "CANCEL: change nothing.", parent=win)
        if choice is None:
            return
        if choice:
            items[:] = found
        else:
            items.extend(new)
        do_save()

    btns = tk.Frame(p, bg=PANEL)
    btns.pack(fill="x", pady=(12, 0))
    HoverButton(btns, text="Edit", bg=TEAL, hover=TEAL_HOVER, font=(FONT, 10, "bold"), command=edit).pack(side="left")
    for text, cmd in (("Add", add), ("Remove", remove), ("Import from Excel…", import_file)):
        HoverButton(btns, text=text, font=(FONT, 10, "bold"), command=cmd).pack(side="left", padx=(8, 0))
    count = tk.Label(btns, text="", font=(FONT, 10), fg=MUTED, bg=PANEL)
    count.pack(side="right")
    tree.bind("<Double-1>", lambda e: edit())
    win._note(p, note)
    fill()


def build_properties_tab(win, p):
    """Settings → Properties."""
    store = win.app.keys
    _list_tab(win, p, "Properties", [("key", "Prop key", 110), ("name", "Property", 420), ("postcode", "Postcode", 110)],
              store.properties, store.save_properties,
              [{"key": "name", "label": "Property (e.g. 12 High Street – Common parts)"},
               {"key": "key", "label": "Prop key (from the housing system; optional)"},
               {"key": "postcode", "label": "Postcode (optional)"}],
              lambda it: (it["key"] or it["name"]).lower(), read_property_file,
              "Picked from when issuing keys; anything not on the list can still be typed in. Import reads the "
              "housing system's property list (Prop Key, Address Line 1-4, Postcode). Common parts are shown as "
              "'<building> – Common parts'. Rooms like the CCTV room or boiler room are not properties: choose the "
              "building, then the key.",
              lambda: win.app.key_view._update_props())


def build_contractors_tab(win, p):
    """Settings → Contractors."""
    store = win.app.keys
    _list_tab(win, p, "Contractors", [("company", "Company", 280), ("name", "Person", 200), ("phone", "Phone", 160)],
              store.contractors, store.save_contractors,
              [{"key": "company", "label": "Company"}, {"key": "name", "label": "Person's name"},
               {"key": "phone", "label": "Phone (optional)"}],
              contractor_key, read_contractor_file,
              "Picked from when issuing keys. A company can be listed once with no person, and again for each "
              "person who collects keys; new people are added here when they first collect keys. Import reads the "
              "housing system's supplier list (Name = company, Supplier Contact = person) or any sheet with "
              "Company / Name / Phone columns.",
              lambda: win.app.key_view._update_contractors(),
              validate=lambda v: None if v["company"] or v["name"] else "Enter a company or a person's name.")
