"""
Shared plumbing for the logbooks (keys, post): an append-only CSV log with the same Check
chain as sales_log.csv and rent_log.csv, and reading simple lists from Excel.

A log's columns are "Entry No", "Date & Time", then its own fields (always including "Type",
"Staff" and "Notes"), then "Check". Entries are dicts keyed by column heading, plus
"entry" (int), "dt" (datetime) and "tampered".
"""

import csv
import os
import re
from datetime import datetime
from decimal import Decimal
from tkinter import messagebox

from front_office import (
    BACKUP_DIR, CSV_DATE_FMT, append_csv_rows, backup_log, chain_hash, parse_csv_datetime, parse_int,
    parse_ui_date, read_csv_rows, save_config, write_csv_export,
)

T_ACCEPTED = "LOG ACCEPTED"


def _hash_value(value):
    """What the Check covers: values, not text, so Excel re-saving the file isn't a change.
    Excel turns 2026-09-28 into 28/09/2026 and drops leading zeros from numbers (phone numbers too)."""
    v = str(value).strip()
    d = parse_ui_date(v) if v else None
    if d:
        return d.isoformat()
    if v.isdigit():
        return v.lstrip("0")
    if re.fullmatch(r"-?\d+\.\d+", v):  # an amount: Excel writes 2.50 back as 2.5
        return str(Decimal(v).normalize())
    return v


class ChainLog:
    def __init__(self, app, path, headers, salt, tip_key, title, types):
        self.app = app
        self.path = path
        self.headers = headers
        self.fields = headers[2:-1]
        self.salt = salt
        self.tip_key = tip_key
        self.title = title
        self.types = tuple(types) + (T_ACCEPTED,)
        self.file = os.path.basename(path)
        if not os.path.exists(path) or os.path.getsize(path) == 0:
            write_csv_export(path, headers, [])
        backup_log(path, os.path.splitext(self.file)[0])
        self.entries, self.skipped, self.last_check = self._load()
        self.problems = self._check()

    # ---- rows <-> entries --------------------------------------------------
    def _to_entry(self, row):
        entry = parse_int(row.get("Entry No"), default=None)
        dt = parse_csv_datetime(row.get("Date & Time"))
        e = {f: (row.get(f) or "").strip() for f in self.fields}
        e["Type"] = e["Type"].upper()
        if entry is None or dt is None or e["Type"] not in self.types:
            return None
        e.update(entry=entry, dt=dt, check=(row.get("Check") or "").strip(), tampered=False)
        return e

    def _to_row(self, e):
        return [e["entry"], e["dt"].strftime(CSV_DATE_FMT)] + [e.get(f, "") for f in self.fields] + [e["check"]]

    def _check_of(self, prev, e):
        return chain_hash(self.salt, prev, [e["entry"], e["dt"].strftime("%Y-%m-%d %H:%M")]
                          + [_hash_value(e.get(f, "")) for f in self.fields])

    def _load(self):
        entries, skipped, prev = [], 0, ""
        for row in read_csv_rows(self.path)[1]:
            stored = (row.get("Check") or "").strip()
            e = self._to_entry(row)
            if e is None:
                skipped += 1
            else:
                e["tampered"] = stored != self._check_of(prev, e)
                entries.append(e)
            prev = stored or prev
        entries.sort(key=lambda e: e["entry"])
        return entries, skipped, prev

    # ---- integrity -----------------------------------------------------------
    def _check(self):
        problems = []
        altered = [e["entry"] for e in self.entries if e["tampered"]]
        if altered:
            shown = ", ".join(f"#{n}" for n in altered[:12]) + (" …" if len(altered) > 12 else "")
            problems.append(f"{len(altered)} {self.title} entr{'y was' if len(altered) == 1 else 'ies were'} "
                            f"changed, added or removed outside the app: {shown}")
        tip = self.app.cfg.get(self.tip_key)
        if tip and not any(e["entry"] == tip["entry"] and e["check"] == tip["check"] for e in self.entries):
            problems.append(f"{self.title.capitalize()} entry #{tip['entry']}, the last one the app wrote, is "
                            f"missing or changed. Entries may have been deleted from the end of {self.file}.")
        if not tip:
            self._set_tip()
        return problems

    def reverify(self):
        """Daily, on a PC left running: a dated backup, then the log read back and checked again."""
        backup_log(self.path, os.path.splitext(self.file)[0])
        try:
            self.entries, _, self.last_check = self._load()
        except OSError:
            return
        self.changed()
        self.problems = self._check()

    def _set_tip(self):
        last = self.entries[-1] if self.entries else None
        self.app.cfg[self.tip_key] = {"entry": last["entry"], "check": last["check"]} if last else None
        try:
            save_config(self.app.cfg)
        except OSError:
            pass  # the entry is already safely in the log; this is only a cross-check

    def accept(self, admin_name, reason, parent):
        """Recalculates every Check, accepting the log as it is now. A copy goes to backups first."""
        try:
            rows = read_csv_rows(self.path)[1]
            os.makedirs(BACKUP_DIR, exist_ok=True)
            backup = os.path.join(BACKUP_DIR, f"{os.path.splitext(self.file)[0]}_before_accept_"
                                              f"{datetime.now():%Y-%m-%d_%H%M%S}.csv")
            with open(self.path, "rb") as src, open(backup, "wb") as dst:
                dst.write(src.read())
            prev, out = "", []
            for row in rows:
                e = self._to_entry(row)
                if e is None:
                    out.append([row.get(h) or "" for h in self.headers])
                    continue
                e["check"] = prev = self._check_of(prev, e)
                out.append(self._to_row(e))
            tmp = self.path + ".tmp"
            write_csv_export(tmp, self.headers, out)
            os.replace(tmp, self.path)
        except OSError as err:
            messagebox.showerror(self.title.capitalize(), f"Could not update {self.file} (is it open in Excel?):"
                                                          f"\n\n{err}", parent=parent)
            return False
        self.entries, _, self.last_check = self._load()
        self.problems = []
        self._set_tip()
        self.write([{"Type": T_ACCEPTED, "Staff": admin_name,
                     "Notes": f"{reason}. Flagged log kept as backups\\{os.path.basename(backup)}"}], parent)
        return True

    # ---- writing ---------------------------------------------------------------
    def write(self, items, parent):
        """Appends entries (dicts keyed by column heading). All or nothing. Returns them, or None."""
        now = datetime.now().replace(microsecond=0)
        next_no = (self.entries[-1]["entry"] + 1) if self.entries else 1
        prev, new = self.last_check, []
        for i, item in enumerate(items):
            e = {f: "" for f in self.fields}
            e.update({k: ("" if v is None else str(v).strip()) for k, v in item.items()})
            e.update(entry=next_no + i, dt=now, tampered=False)
            e["check"] = prev = self._check_of(prev, e)
            new.append(e)
        try:
            append_csv_rows(self.path, self.headers, [self._to_row(e) for e in new])
        except PermissionError:
            messagebox.showerror("Could not save", f"{self.file} is open in another program (probably Excel).\n\n"
                                 "Close it and try again. Nothing was recorded.", parent=parent)
            return None
        except OSError as err:
            messagebox.showerror("Could not save", f"Could not write to {self.file}:\n\n{err}\n\n"
                                 "Nothing was recorded.", parent=parent)
            return None
        self.entries.extend(new)
        self.last_check = prev
        self._set_tip()
        self.changed()
        return new

    def changed(self):
        """Called after every write, so derived views can be rebuilt."""


# --------------------------------------------------------------------------
# Simple lists (properties, contractors) kept as CSV, imported from Excel
# --------------------------------------------------------------------------

def read_sheet(path):
    """Every row of the first sheet of an Excel file (or a CSV), as lists of stripped strings."""
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        rows = [["" if v is None else str(v).strip() for v in r] for r in wb.active.iter_rows(values_only=True)]
        wb.close()
        return rows
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        return [[c.strip() for c in r] for r in csv.reader(f)]


def load_list(path, columns):
    """Rows of a list CSV as dicts {key: value} for columns [(key, heading)]."""
    if not os.path.exists(path):
        return []
    return [{k: (row.get(h) or "").strip() for k, h in columns} for row in read_csv_rows(path)[1]]


def save_list(path, columns, items):
    tmp = path + ".tmp"
    write_csv_export(tmp, [h for _, h in columns], [[it.get(k, "") for k, _ in columns] for it in items])
    os.replace(tmp, path)
