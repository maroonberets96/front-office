"""
Visitor book
============

Replaces the paper visitor books. Staff fill it in (the PC is inside the office; visitors never
use it themselves).

  Sign in    Any staff PIN, when the visitor arrives: name, company, who they're visiting and
             car registration. The time in is when it's logged.
  Sign out   Any staff PIN, with the time they left (now, or earlier if someone forgot).
  Void       A mistaken entry, by an admin or checker, with a reason. Nothing is ever deleted.

Records are kept indefinitely (the user's choice).

Data file, in data/:
  visitor_log.csv   every sign-in, sign-out and void (append-only, with the same Check chain as
                    the other logs)
"""

import os
import re
import tkinter as tk
from datetime import date, datetime
from tkinter import messagebox, ttk

from front_office import (
    AMBER, BG, BLUE, DATA_DIR, DIM, FONT, GREEN, GREEN_HOVER, GREY_BTN, GREY_HOVER, MUTED, PANEL, PINK, RED,
    RED_HOVER, ROW_A, ROW_B, TAMPERED_BG, TAMPERED_FG, TEAL, TEAL_HOVER, TEXT, FieldsDialog, HoverButton, PinDialog,
    PlaceholderEntry, admin_save_path, make_sortable, parse_int,
)
from logbook import ChainLog
from rent import open_file, write_xlsx

VISITOR_LOG_PATH = os.path.join(DATA_DIR, "visitor_log.csv")
VISITOR_HEADERS = [
    "Entry No", "Date & Time", "Type", "Visit No", "Name", "Company", "Visiting", "Car Reg", "Time Out",
    "Staff", "Notes", "Check",
]
VISITOR_SALT = b"front-office/visitor-chain/v1"

V_IN, V_OUT, V_VOIDED = "SIGNED IN", "SIGNED OUT", "VOIDED"
S_IN, S_OUT, S_VOIDED = "in", "out", "voided"


def visit_name(n):
    return f"V{n:04d}"


def plural(n, word):
    return f"{n} {word}{'' if n == 1 else 's'}"


def parse_time(text):
    """'10:48', '1048' or '10.48' -> (10, 48), or None."""
    m = re.fullmatch(r"(\d{1,2})[:.]?(\d\d)", text.strip())
    if not m:
        return None
    h, mi = int(m.group(1)), int(m.group(2))
    return (h, mi) if h < 24 and mi < 60 else None


class VisitorStore(ChainLog):
    def __init__(self, app):
        self._visits = None
        super().__init__(app, VISITOR_LOG_PATH, VISITOR_HEADERS, VISITOR_SALT, "visitor_log_tip", "visitor log",
                         (V_IN, V_OUT, V_VOIDED))

    def changed(self):
        self._visits = None

    def visits(self):
        """visit number -> visit, built from the log's entries."""
        if self._visits is not None:
            return self._visits
        visits = {}
        for e in self.entries:
            n = parse_int(e["Visit No"], default=None)
            if e["Type"] == V_IN and n is not None:
                visits[n] = {
                    "no": n, "dt": e["dt"], "name": e["Name"], "company": e["Company"], "visiting": e["Visiting"],
                    "car": e["Car Reg"], "staff": e["Staff"], "notes": e["Notes"],
                    "out_at": None, "out_by": "", "voided_at": None, "voided_by": "", "void_reason": "",
                    "tampered": e["tampered"],
                }
        for e in self.entries:
            v = visits.get(parse_int(e["Visit No"], default=None))
            if v is None or e["Type"] == V_IN:
                continue
            v["tampered"] = v["tampered"] or e["tampered"]
            if e["Type"] == V_OUT:
                t = parse_time(e["Time Out"])
                out = v["dt"].replace(hour=t[0], minute=t[1], second=0) if t else e["dt"]
                v.update(out_at=out, out_by=e["Staff"])
            elif e["Type"] == V_VOIDED:
                v.update(voided_at=e["dt"], voided_by=e["Staff"], void_reason=e["Notes"])
        self._visits = visits
        return visits

    @staticmethod
    def status(v):
        if v["voided_at"]:
            return S_VOIDED
        return S_OUT if v["out_at"] else S_IN

    def next_visit(self):
        return max([0] + list(self.visits())) + 1

    def in_now(self):
        return [v for v in self.visits().values() if self.status(v) == S_IN]

    def forgotten(self):
        """Still signed in from an earlier day: almost certainly left without signing out."""
        today = date.today()
        return [v for v in self.in_now() if v["dt"].date() < today]


def used(visits, field, extra=()):
    """Values typed before (most used first) then any extras, for the drop-down suggestions."""
    counts = {}
    for v in visits:
        s = v[field].strip()
        if s:
            counts[s] = counts.get(s, 0) + 1
    names = sorted(counts, key=lambda s: -counts[s])
    return names + [s for s in extra if s.lower() not in {n.lower() for n in names}]


def status_text(v):
    s = VisitorStore.status(v)
    if s == S_VOIDED:
        return f"Voided · {v['void_reason']}"
    if s == S_OUT:
        return f"Left {v['out_at']:%H:%M}"
    days = (date.today() - v["dt"].date()).days
    return "In the building" if not days else f"Not signed out ({plural(days, 'day')} ago)"


# --------------------------------------------------------------------------
# Main view (the "Visitors" tab)
# --------------------------------------------------------------------------

class VisitorView(tk.Frame):
    def __init__(self, master, app):
        super().__init__(master, bg=BG)
        self.app = app
        self.store = app.visitors
        self.filter = "in"
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self._build_form()
        self._build_main()
        self.clear_form()
        self.refresh()

    # ---- left: sign a visitor in ---------------------------------------------
    def _section(self, parent, number, title, top=14):
        row = tk.Frame(parent, bg=PANEL)
        row.pack(fill="x", pady=(top, 6))
        tk.Label(row, text=number, font=(FONT, 9, "bold"), fg=BG, bg=MUTED, width=2).pack(side="left")
        tk.Label(row, text=title.upper(), font=(FONT, 10, "bold"), fg=MUTED, bg=PANEL).pack(side="left", padx=8)
        return row

    def _combo(self, parent):
        box = ttk.Combobox(parent, font=(FONT, 11))
        box.pack(fill="x", ipady=3)
        for ev in ("<KeyRelease>", "<<ComboboxSelected>>"):
            box.bind(ev, lambda e: self.update_form())
        return box

    @staticmethod
    def _paint(buttons, chosen):
        for opt, btn in buttons.items():
            if opt == chosen:
                btn.set_colors(TEAL, TEAL_HOVER, "#ffffff")
            else:
                btn.set_colors(GREY_BTN, GREY_HOVER, TEXT)

    def _build_form(self):
        S = self.app.S
        side = tk.Frame(self, bg=PANEL, width=S(420))
        side.grid(row=0, column=0, sticky="ns")
        side.pack_propagate(False)
        p = tk.Frame(side, bg=PANEL)
        p.pack(fill="both", expand=True, padx=S(22), pady=S(12))

        head = tk.Frame(p, bg=PANEL)
        head.pack(fill="x")
        tk.Label(head, text="Sign a visitor in", font=(FONT, 14, "bold"), fg=TEXT, bg=PANEL).pack(side="left")
        clear = tk.Label(head, text="Clear", font=(FONT, 10, "underline"), fg=MUTED, bg=PANEL, cursor="hand2")
        clear.pack(side="right")
        clear.bind("<Button-1>", lambda e: self.clear_form())

        self._section(p, "1", "Visitor's name")
        self.name_entry = PlaceholderEntry(p, "Full name")
        self.name_entry.pack(fill="x", ipady=5)
        self.name_entry.bind("<KeyRelease>", lambda e: self.update_form())
        self._section(p, "2", "Company")
        self.company_box = self._combo(p)
        tk.Label(p, text="Leave blank if they're not from a company", font=(FONT, 9), fg=DIM, bg=PANEL,
                 anchor="w").pack(fill="x")
        self._section(p, "3", "Visiting")
        self.visiting_box = self._combo(p)
        self._section(p, "4", "Car registration")
        self.car_entry = PlaceholderEntry(p, "Optional, e.g. AB12 CDE")
        self.car_entry.pack(fill="x", ipady=5)
        self.notes_entry = PlaceholderEntry(p, "Note (optional), e.g. boiler service")
        self.notes_entry.pack(fill="x", ipady=5, pady=(10, 0))

        bottom = tk.Frame(p, bg=PANEL)
        bottom.pack(fill="x", side="bottom")
        self.problem_label = tk.Label(bottom, text="", font=(FONT, 9), fg=MUTED, bg=PANEL, anchor="w")
        self.problem_label.pack(fill="x")
        self.log_btn = HoverButton(bottom, text="Sign in visitor", bg=GREEN, hover=GREEN_HOVER,
                                   font=(FONT, 13, "bold"), pady=10, command=self.sign_in)
        self.log_btn.pack(fill="x", pady=(4, 0))

    def form_problem(self):
        if not self.name_entry.value():
            return "Enter the visitor's name"
        if not self.visiting_box.get().strip():
            return "Enter who they're visiting"
        return None

    def update_form(self):
        if not hasattr(self, "log_btn"):
            return
        problem = self.form_problem()
        self.problem_label.config(text=problem or f"Visit {visit_name(self.store.next_visit())} will be logged")
        self.log_btn.set_enabled(problem is None)

    def clear_form(self):
        for e in (self.name_entry, self.car_entry, self.notes_entry):
            e.clear()
        for box in (self.company_box, self.visiting_box):
            box.set("")
        self._fill_suggestions()
        self.update_form()

    def _fill_suggestions(self):
        visits = list(self.store.visits().values())
        staff = sorted(s["name"] for s in self.app.cfg.get("staff", []))
        self.company_box["values"] = used(visits, "company")
        self.visiting_box["values"] = used(visits, "visiting", staff)

    def reset(self):
        """Back to a fresh screen: empty form, no search, 'In now' filter."""
        self.clear_form()
        self.search.clear()
        self.set_filter("in")

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
        for i, (key, title, accent) in enumerate([("in", "In the building now", PINK),
                                                   ("today", "Visitors today", TEAL),
                                                   ("left", "Signed out today", GREEN),
                                                   ("month", "Visitors this month", BLUE)]):
            cards.columnconfigure(i, weight=1, uniform="card")
            card = tk.Frame(cards, bg=PANEL)
            card.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else S(12), 0))
            tk.Frame(card, bg=accent, height=3).pack(fill="x")
            tk.Label(card, text=title.upper(), font=(FONT, 9, "bold"), fg=MUTED, bg=PANEL,
                     anchor="w").pack(fill="x", padx=16, pady=(12, 0))
            value = tk.Label(card, text="0", font=(FONT, 24, "bold"), fg=TEXT, bg=PANEL, anchor="w")
            value.pack(fill="x", padx=16)
            sub = tk.Label(card, text="", font=(FONT, 9), fg=MUTED, bg=PANEL, anchor="w")
            sub.pack(fill="x", padx=16, pady=(0, 12))
            self.cards[key] = (value, sub)

        bar = tk.Frame(main, bg=BG)
        bar.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self.filter_buttons = {}
        for key, text in (("in", "In now"), ("today", "Today"), ("voided", "Voided"), ("all", "All")):
            btn = HoverButton(bar, text=text, font=(FONT, 9, "bold"), padx=10, pady=4,
                              command=lambda k=key: self.set_filter(k))
            btn.pack(side="left", padx=(0, 4))
            self.filter_buttons[key] = btn
        for text, bg, hover, cmd in [("Export", GREY_BTN, GREY_HOVER, self.export),
                                     ("Void", RED, RED_HOVER, self.void_selected),
                                     ("Sign out…", GREEN, GREEN_HOVER, self.sign_out)]:
            HoverButton(bar, text=text, bg=bg, hover=hover, font=(FONT, 10, "bold"),
                        command=cmd).pack(side="right", padx=(8, 0))
        self.search = PlaceholderEntry(bar, "Search…", width=18)
        self.search.pack(side="right", padx=(8, 4), ipady=4)
        self.search.bind("<KeyRelease>", lambda e: self.refresh_table())

        table = tk.Frame(main, bg=PANEL)
        table.grid(row=2, column=0, sticky="nsew")
        columns = [("no", "Visit", 70, "w"), ("date", "Date", 100, "w"), ("in", "In", 60, "w"),
                   ("out", "Out", 60, "w"), ("name", "Name", 170, "w"), ("company", "Company", 150, "w"),
                   ("visiting", "Visiting", 150, "w"), ("car", "Car reg", 90, "w"), ("staff", "Logged by", 110, "w"),
                   ("status", "Status", 180, "w")]
        self.tree = ttk.Treeview(table, columns=[c[0] for c in columns], show="headings", style="Dark.Treeview",
                                 selectmode="browse")
        for key, text, width, anchor in columns:
            self.tree.heading(key, text=text, anchor=anchor)
            self.tree.column(key, width=S(width), minwidth=S(40), anchor=anchor,
                             stretch=key in ("name", "company", "visiting", "status"))
        sb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("tampered", background=TAMPERED_BG, foreground=TAMPERED_FG)  # first = wins
        self.tree.tag_configure("even", background=ROW_A)
        self.tree.tag_configure("odd", background=ROW_B)
        self.tree.tag_configure(S_IN, foreground=TEXT)
        self.tree.tag_configure("forgotten", foreground=AMBER)
        self.tree.tag_configure(S_OUT, foreground=MUTED)
        self.tree.tag_configure(S_VOIDED, foreground=DIM, font=(FONT, 10, "overstrike"))
        self.tree.bind("<Double-1>", lambda e: self.sign_out())
        make_sortable(self.tree)
        self.empty_label = tk.Label(table, text="", font=(FONT, 11), fg=MUTED, bg=ROW_A)
        self.set_filter("in")

    def set_filter(self, key):
        self.filter = key
        self._paint(self.filter_buttons, key)
        self.refresh_table()

    def refresh(self):
        self.refresh_cards()
        self.refresh_table()
        self._fill_suggestions()
        self.update_form()

    def refresh_cards(self):
        visits = [v for v in self.store.visits().values() if not v["voided_at"]]
        today = date.today()
        in_now = self.store.in_now()
        value, sub = self.cards["in"]
        value.config(text=str(len(in_now)), fg=PINK if in_now else TEXT)
        forgotten = self.store.forgotten()
        sub.config(text=f"{plural(len(forgotten), 'from an earlier day')}" if forgotten else
                   f"since {min(v['dt'] for v in in_now):%H:%M}" if in_now else "Nobody signed in",
                   fg=AMBER if forgotten else MUTED)
        todays = [v for v in visits if v["dt"].date() == today]
        value, sub = self.cards["today"]
        value.config(text=str(len(todays)))
        sub.config(text=", ".join(sorted({v["company"] for v in todays if v["company"]})) or "—")
        left = [v for v in visits if v["out_at"] and v["out_at"].date() == today]
        value, sub = self.cards["left"]
        value.config(text=str(len(left)))
        sub.config(text=f"last at {max(v['out_at'] for v in left):%H:%M}" if left else "—")
        month = [v for v in visits if v["dt"].date() >= today.replace(day=1)]
        value, sub = self.cards["month"]
        value.config(text=str(len(month)))
        sub.config(text=plural(len({v["company"] for v in month if v["company"]}), "company"))

    def refresh_table(self):
        query = self.search.value().lower()
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        today = date.today()
        shown = 0
        for v in sorted(self.store.visits().values(), key=lambda x: x["no"], reverse=True):
            status = self.store.status(v)
            if not {"in": status == S_IN, "today": v["dt"].date() == today and status != S_VOIDED,
                    "voided": status == S_VOIDED, "all": True}[self.filter]:
                continue
            values = (visit_name(v["no"]), v["dt"].strftime("%d/%m/%Y"), v["dt"].strftime("%H:%M"),
                      v["out_at"].strftime("%H:%M") if v["out_at"] else "", v["name"], v["company"], v["visiting"],
                      v["car"], v["staff"], status_text(v))
            if query and not any(query in str(x).lower() for x in values + (v["notes"],)):
                continue
            forgotten = status == S_IN and v["dt"].date() < today
            tags = ["odd" if shown % 2 else "even", "forgotten" if forgotten else status]
            if v["tampered"]:
                tags.append("tampered")
            self.tree.insert("", "end", iid=str(v["no"]), values=values, tags=tags)
            shown += 1
        self.tree.resort()
        keep = [i for i in selected if self.tree.exists(i)]
        if keep:
            self.tree.selection_set(keep)
        if shown:
            self.empty_label.place_forget()
        else:
            self.empty_label.config(text="No matching visitors." if query else {
                "in": "Nobody is signed in.", "today": "No visitors today yet.", "voided": "Nothing voided.",
                "all": "No visitors logged yet."}[self.filter])
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")

    def selected_visit(self, what):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(what, "Select a visitor in the list first.", parent=self.app)
            return None
        return self.store.visits().get(int(sel[0]))

    def _pin(self, title, heading, lines, **kw):
        dlg = PinDialog(self.app, self.app, title, heading, lines, **kw)
        result = dlg.show()
        if not result:
            self.app._pin_locked_out(dlg)
        return result

    def _done(self, text, color=GREEN):
        self.refresh()
        self.app.set_message(text, color)

    # ---- actions -------------------------------------------------------------
    def sign_in(self):
        if self.form_problem():
            return
        n = self.store.next_visit()
        name, company = self.name_entry.value(), self.company_box.get().strip()
        visiting, car, notes = self.visiting_box.get().strip(), self.car_entry.value().upper(), self.notes_entry.value()
        result = self._pin("Visitor sign in", f"Visit {visit_name(n)}  ·  {name}",
                           [x for x in (company and f"From {company}", f"Visiting {visiting}",
                                        car and f"Car {car}", notes) if x],
                           confirm_text="Sign in visitor")
        if not result:
            return
        row = {"Type": V_IN, "Visit No": n, "Name": name, "Company": company, "Visiting": visiting,
               "Car Reg": car, "Staff": result[0]["name"], "Notes": notes}
        if self.store.write([row], self.app):
            self.clear_form()
            self.set_filter("in")
            self._done(f"✓ {name} signed in at {datetime.now():%H:%M} to see {visiting}")

    def sign_out(self):
        v = self.selected_visit("Sign out")
        if not v:
            return
        if self.store.status(v) != S_IN:
            messagebox.showinfo("Sign out", f"{v['name']} is not signed in.", parent=self.app)
            return
        now = datetime.now()
        same_day = v["dt"].date() == now.date()

        def check(values):
            t = parse_time(values["time"])
            if not t:
                return "Enter the time they left, e.g. 14:30"
            out = v["dt"].replace(hour=t[0], minute=t[1], second=0)
            if out < v["dt"].replace(second=0):
                return f"They arrived at {v['dt']:%H:%M}, so they can't have left before then."
            if out > now:
                return "That time hasn't happened yet."
            return None

        res = FieldsDialog(self.app, "Sign out", f"Sign out {v['name']}", [
            {"key": "time", "label": "Time they left (HH:MM)", "value": now.strftime("%H:%M") if same_day else ""},
        ], validate=check, confirm_text="Next",
            sub=f"Arrived {v['dt']:%d/%m/%Y %H:%M} to see {v['visiting']}."
                + ("" if same_day else " They weren't signed out that day: enter the time they left, "
                                       "or your best guess.")).show()
        if not res:
            return
        h, mi = parse_time(res["time"])
        time_out = f"{h:02d}:{mi:02d}"
        result = self._pin("Sign out", f"Visit {visit_name(v['no'])}  ·  {v['name']} left at {time_out}",
                           [x for x in (v["company"] and f"From {v['company']}", f"Visiting {v['visiting']}",
                                        f"Arrived {v['dt']:%d/%m/%Y %H:%M}") if x],
                           confirm_text="Sign out")
        if result and self.store.write([{"Type": V_OUT, "Visit No": v["no"], "Name": v["name"],
                                         "Time Out": time_out, "Staff": result[0]["name"]}], self.app):
            self._done(f"✓ {v['name']} signed out at {time_out}")

    def void_selected(self):
        v = self.selected_visit("Void")
        if not v:
            return
        name = visit_name(v["no"])
        if v["voided_at"]:
            messagebox.showinfo("Void", f"Visit {name} is already voided.", parent=self.app)
            return
        result = self._pin("Void", f"Void visit {name}",
                           [f"{v['name']}" + (f" from {v['company']}" if v["company"] else ""),
                            f"Signed in {v['dt']:%d/%m/%Y %H:%M} by {v['staff']}",
                            "Only for an entry made by mistake. It stays in the log, struck through."],
                           allow=lambda s: None if (s["admin"] or s.get("checker")) else
                           "Only an admin or checker can void an entry",
                           reason_label="Reason (e.g. entered twice)",
                           confirm_text="Void visit", confirm_bg=RED, confirm_hover=RED_HOVER)
        if result and self.store.write([{"Type": V_VOIDED, "Visit No": v["no"], "Name": v["name"],
                                         "Staff": result[0]["name"], "Notes": result[1]}], self.app):
            self._done(f"Visit {name} voided by {result[0]['name']}", AMBER)

    def export(self):
        path = admin_save_path(self.app, self.app, title="Export visitor book", defaultextension=".xlsx",
                               initialfile=f"Visitor book {date.today():%Y-%m-%d}.xlsx",
                               filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        rows = [[visit_name(v["no"]), v["dt"].replace(second=0), v["out_at"].replace(second=0) if v["out_at"] else None,
                 v["name"], v["company"], v["visiting"], v["car"], v["notes"], v["staff"], v["out_by"], status_text(v)]
                for v in sorted(self.store.visits().values(), key=lambda x: x["no"])]
        try:
            write_xlsx(path, ["Visit", "Time in", "Time out", "Name", "Company", "Visiting", "Car reg", "Note",
                              "Signed in by", "Signed out by", "Status"], rows)
            open_file(path)
        except Exception as e:
            messagebox.showerror("Export", f"Could not export:\n\n{e}", parent=self.app)
