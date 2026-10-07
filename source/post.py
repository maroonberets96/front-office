"""
Post logbook
============

Only post worth a record is logged, not everything that arrives: letters for the CEO or DCEO,
and anything else special. Replaces the handwritten post book.

  Incoming   Any staff PIN, on arrival: who it's for, who it's from, what it is (letter,
             signed-for, parcel), tracking number and a short note. Then "Hand over" records
             who it was given to and when (again any staff PIN).
  Outgoing   Any staff PIN: who it went to, and optionally the address, how it was sent,
             tracking number, postage, who it was sent for and what it was about.
  Void       A mistaken entry, by an admin or checker, with a reason. Nothing is ever deleted.

Data file, in data/:
  post_log.csv   every item received, handed over, sent or voided (append-only, with the same
                 Check chain as the other logs)
"""

import os
import tkinter as tk
from datetime import date
from tkinter import messagebox, ttk

from front_office import (
    AMBER, BG, DATA_DIR, DIM, FONT, GREEN, GREEN_HOVER, GREY_BTN, GREY_HOVER, MUTED, ORANGE, PANEL, PURPLE, RED,
    RED_HOVER, ROW_A, ROW_B, TAMPERED_BG, TAMPERED_FG, TEAL, TEAL_HOVER, TEXT, YELLOW, FieldsDialog, HoverButton,
    PinDialog, PlaceholderEntry, admin_save_path, csv_money, make_sortable, entry_opts, fmt_money, parse_int, parse_pence,
)
from logbook import ChainLog
from rent import open_file, write_xlsx

POST_LOG_PATH = os.path.join(DATA_DIR, "post_log.csv")
POST_HEADERS = [
    "Entry No", "Date & Time", "Type", "Item No", "To", "From", "Address", "Item", "Tracking No", "Cost (£)",
    "Subject", "Handed To", "Staff", "Notes", "Check",
]
POST_SALT = b"front-office/post-chain/v1"

P_RECEIVED, P_HANDED, P_SENT, P_VOIDED = "RECEIVED", "HANDED OVER", "SENT", "VOIDED"

RECIPIENTS = ["CEO", "DCEO"]                      # quick buttons for incoming post
ITEM_KINDS = ["Letter", "Signed for", "Parcel", "Other"]
SEND_METHODS = ["2nd class", "1st class", "Signed For", "Special Delivery", "Hand delivered", "Other"]
COMMON_SENDERS = ["HMRC", "Council", "Court", "Solicitor", "Bank", "Regulator of Social Housing"]

S_WAITING, S_HANDED, S_SENT, S_VOIDED = "waiting", "handed", "sent", "voided"


def item_name(n):
    return f"P{n:04d}"


def plural(n, word):
    return f"{n} {word}{'' if n == 1 else 's'}"


class PostStore(ChainLog):
    def __init__(self, app):
        self._items = None
        super().__init__(app, POST_LOG_PATH, POST_HEADERS, POST_SALT, "post_log_tip", "post log",
                         (P_RECEIVED, P_HANDED, P_SENT, P_VOIDED))

    def changed(self):
        self._items = None

    def items(self):
        """item number -> item, built from the log's entries."""
        if self._items is not None:
            return self._items
        items = {}
        for e in self.entries:
            n = parse_int(e["Item No"], default=None)
            if e["Type"] in (P_RECEIVED, P_SENT) and n is not None:
                items[n] = {
                    "no": n, "dt": e["dt"], "incoming": e["Type"] == P_RECEIVED, "to": e["To"], "from": e["From"],
                    "address": e["Address"], "item": e["Item"], "tracking": e["Tracking No"],
                    "cost": parse_pence(e["Cost (£)"]) or 0, "subject": e["Subject"], "staff": e["Staff"],
                    "handed_to": "", "handed_at": None, "handed_by": "",
                    "voided_at": None, "voided_by": "", "void_reason": "", "tampered": e["tampered"],
                }
        for e in self.entries:
            it = items.get(parse_int(e["Item No"], default=None))
            if it is None or e["Type"] in (P_RECEIVED, P_SENT):
                continue
            it["tampered"] = it["tampered"] or e["tampered"]
            if e["Type"] == P_HANDED:
                it.update(handed_to=e["Handed To"], handed_at=e["dt"], handed_by=e["Staff"])
            elif e["Type"] == P_VOIDED:
                it.update(voided_at=e["dt"], voided_by=e["Staff"], void_reason=e["Notes"])
        self._items = items
        return items

    @staticmethod
    def status(it):
        if it["voided_at"]:
            return S_VOIDED
        if not it["incoming"]:
            return S_SENT
        return S_HANDED if it["handed_at"] else S_WAITING

    def next_item(self):
        return max([0] + list(self.items())) + 1

    def waiting(self):
        return [it for it in self.items().values() if self.status(it) == S_WAITING]


def used(items, field, extra=()):
    """Values typed before (most used first) then the usual ones, for the drop-down suggestions."""
    counts = {}
    for it in items:
        v = it[field].strip()
        if v:
            counts[v] = counts.get(v, 0) + 1
    names = sorted(counts, key=lambda v: -counts[v])
    return names + [v for v in extra if v.lower() not in {n.lower() for n in names}]


def status_text(it):
    s = PostStore.status(it)
    if s == S_VOIDED:
        return f"Voided · {it['void_reason']}"
    if s == S_WAITING:
        days = (date.today() - it["dt"].date()).days
        return "Waiting to be handed over" + (f" ({plural(days, 'day')})" if days else "")
    if s == S_HANDED:
        return f"Handed to {it['handed_to']} {it['handed_at']:%d/%m/%Y}"
    return "Sent" + (f" · {it['item']}" if it["item"] else "")


# --------------------------------------------------------------------------
# Main view (the "Post" tab)
# --------------------------------------------------------------------------

class PostView(tk.Frame):
    def __init__(self, master, app):
        super().__init__(master, bg=BG)
        self.app = app
        self.store = app.post
        self.mode = "in"
        self.kind = self.method = None
        self.filter = "waiting"
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self._build_form()
        self._build_main()
        self.set_mode("in")
        self.clear_form()
        self.refresh()

    # ---- left: log an item -------------------------------------------------
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

    def _choices(self, parent, options, cols, command):
        grid = tk.Frame(parent, bg=PANEL)
        grid.pack(fill="x")
        buttons = {}
        for i, opt in enumerate(options):
            grid.columnconfigure(i % cols, weight=1, uniform="choice")
            btn = HoverButton(grid, text=opt, font=(FONT, 9, "bold"), padx=0, pady=5,
                              command=lambda o=opt: command(o))
            btn.grid(row=i // cols, column=i % cols, sticky="nsew", padx=2, pady=2)
            buttons[opt] = btn
        return buttons

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
        head.columnconfigure(0, weight=1, uniform="mode")
        head.columnconfigure(1, weight=1, uniform="mode")
        self.mode_buttons = {}
        for i, (key, text) in enumerate((("in", "Incoming post"), ("out", "Outgoing post"))):
            btn = HoverButton(head, text=text, font=(FONT, 11, "bold"), pady=7, command=lambda k=key: self.set_mode(k))
            btn.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 4, 0))
            self.mode_buttons[key] = btn
        clear = tk.Label(p, text="Clear", font=(FONT, 10, "underline"), fg=MUTED, bg=PANEL, cursor="hand2")
        clear.pack(anchor="e", pady=(6, 0))
        clear.bind("<Button-1>", lambda e: self.clear_form())

        # incoming
        f = self.in_form = tk.Frame(p, bg=PANEL)
        self._section(f, "1", "Addressed to", top=0)
        self.to_box = self._combo(f)
        quick = tk.Frame(f, bg=PANEL)
        quick.pack(fill="x", pady=(4, 0))
        for name in RECIPIENTS:
            HoverButton(quick, text=name, font=(FONT, 10, "bold"), padx=16, pady=3,
                        command=lambda n=name: self._set(self.to_box, n)).pack(side="left", padx=(0, 4))
        self._section(f, "2", "From")
        self.from_box = self._combo(f)
        tk.Label(f, text="e.g. HMRC, the council, a solicitor", font=(FONT, 9), fg=DIM, bg=PANEL,
                 anchor="w").pack(fill="x")
        self._section(f, "3", "What it is")
        self.kind_buttons = self._choices(f, ITEM_KINDS, 4, self.set_kind)
        self.in_tracking = PlaceholderEntry(f, "Tracking number (optional)")
        self.in_tracking.pack(fill="x", ipady=5, pady=(8, 0))
        self.in_subject = PlaceholderEntry(f, "Note (optional), e.g. tax exemption letter")
        self.in_subject.pack(fill="x", ipady=5, pady=(8, 0))

        # outgoing
        f = self.out_form = tk.Frame(p, bg=PANEL)
        self._section(f, "1", "Sent to", top=0)
        self.recipient_box = self._combo(f)
        self.address_entry = PlaceholderEntry(f, "Address (optional)")
        self.address_entry.pack(fill="x", ipady=5, pady=(8, 0))
        self._section(f, "2", "How it was sent")
        self.method_buttons = self._choices(f, SEND_METHODS, 3, self.set_method)
        row = tk.Frame(f, bg=PANEL)
        row.pack(fill="x", pady=(8, 0))
        row.columnconfigure(0, weight=1)
        self.out_tracking = PlaceholderEntry(row, "Tracking number (optional)")
        self.out_tracking.grid(row=0, column=0, sticky="ew", ipady=4)
        tk.Label(row, text="Postage £", font=(FONT, 10, "bold"), fg=MUTED, bg=PANEL).grid(row=0, column=1,
                                                                                         padx=(10, 4))
        vcmd = (self.register(lambda s: all(ch in "0123456789.,£" for ch in s)), "%P")
        self.cost_entry = tk.Entry(row, width=6, justify="right", validate="key", validatecommand=vcmd,
                                   **entry_opts())
        self.cost_entry.grid(row=0, column=2, ipady=4)
        self.cost_entry.bind("<KeyRelease>", lambda e: self.update_form())
        self._section(f, "3", "About")
        self.sender_entry = PlaceholderEntry(f, "Sent by / for (optional), e.g. Housing team")
        self.sender_entry.pack(fill="x", ipady=5)
        self.out_subject = PlaceholderEntry(f, "Subject (optional), e.g. notice to tenant")
        self.out_subject.pack(fill="x", ipady=5, pady=(8, 0))

        self.bottom = tk.Frame(p, bg=PANEL)
        self.bottom.pack(fill="x", side="bottom")
        self.problem_label = tk.Label(self.bottom, text="", font=(FONT, 9), fg=MUTED, bg=PANEL, anchor="w")
        self.problem_label.pack(fill="x")
        self.log_btn = HoverButton(self.bottom, text="Log", bg=GREEN, hover=GREEN_HOVER, font=(FONT, 13, "bold"),
                                   pady=10, command=self.log_item)
        self.log_btn.pack(fill="x", pady=(4, 0))

    def _set(self, box, value):
        box.set(value)
        self.update_form()

    def set_mode(self, mode):
        self.mode = mode
        self._paint(self.mode_buttons, mode)
        (self.out_form if mode == "in" else self.in_form).pack_forget()
        (self.in_form if mode == "in" else self.out_form).pack(fill="x", pady=(4, 0))
        self.log_btn.config(text="Log incoming post" if mode == "in" else "Log outgoing post")
        self.update_form()

    def set_kind(self, kind):
        self.kind = kind
        self._paint(self.kind_buttons, kind)
        self.update_form()

    def set_method(self, method):
        self.method = None if method == self.method else method
        self._paint(self.method_buttons, self.method)
        self.update_form()

    def form_problem(self):
        if self.mode == "in":
            if not self.to_box.get().strip():
                return "Enter who it's addressed to"
            if not self.from_box.get().strip():
                return "Enter who it's from (or 'unknown')"
            return None
        if not self.recipient_box.get().strip():
            return "Enter who it was sent to"
        if parse_pence(self.cost_entry.get()) is None:
            return "Postage must be an amount, e.g. 2.70"
        return None

    def update_form(self):
        if not hasattr(self, "log_btn"):
            return
        problem = self.form_problem()
        self.problem_label.config(text=problem or f"Item {item_name(self.store.next_item())} will be logged")
        self.log_btn.set_enabled(problem is None)

    def clear_form(self):
        for box in (self.to_box, self.from_box, self.recipient_box):
            box.set("")
        for e in (self.in_tracking, self.in_subject, self.address_entry, self.out_tracking, self.sender_entry,
                  self.out_subject):
            e.clear()
        self.cost_entry.delete(0, "end")
        self.set_kind("Letter")
        self.method = None
        self._paint(self.method_buttons, None)
        self._fill_suggestions()
        self.update_form()

    def _fill_suggestions(self):
        incoming = [it for it in self.store.items().values() if it["incoming"]]
        outgoing = [it for it in self.store.items().values() if not it["incoming"]]
        self.to_box["values"] = used(incoming, "to", RECIPIENTS)
        self.from_box["values"] = used(incoming, "from", COMMON_SENDERS)
        self.recipient_box["values"] = used(outgoing, "to")

    def reset(self):
        """Back to a fresh screen: incoming, empty form, no search, 'Waiting' filter."""
        self.set_mode("in")
        self.clear_form()
        self.search.clear()
        self.set_filter("waiting")

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
        for i, (key, title, accent) in enumerate([("waiting", "Waiting to be handed over", YELLOW),
                                                   ("in", "Received today", TEAL), ("out", "Sent today", PURPLE),
                                                   ("cost", "Postage this month", ORANGE)]):
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
        for key, text in (("waiting", "Waiting"), ("in", "Incoming"), ("out", "Outgoing"), ("voided", "Voided"),
                          ("all", "All")):
            btn = HoverButton(bar, text=text, font=(FONT, 9, "bold"), padx=10, pady=4,
                              command=lambda k=key: self.set_filter(k))
            btn.pack(side="left", padx=(0, 4))
            self.filter_buttons[key] = btn
        for text, bg, hover, cmd in [("Export", GREY_BTN, GREY_HOVER, self.export),
                                     ("Void", RED, RED_HOVER, self.void_selected),
                                     ("Hand over…", GREEN, GREEN_HOVER, self.hand_over)]:
            HoverButton(bar, text=text, bg=bg, hover=hover, font=(FONT, 10, "bold"),
                        command=cmd).pack(side="right", padx=(8, 0))
        self.search = PlaceholderEntry(bar, "Search…", width=18)
        self.search.pack(side="right", padx=(8, 4), ipady=4)
        self.search.bind("<KeyRelease>", lambda e: self.refresh_table())

        table = tk.Frame(main, bg=PANEL)
        table.grid(row=2, column=0, sticky="nsew")
        columns = [("no", "Item", 70, "w"), ("date", "Date", 120, "w"), ("dir", "In / out", 70, "w"),
                   ("to", "To", 170, "w"), ("from", "From", 170, "w"), ("item", "What / how", 120, "w"),
                   ("tracking", "Tracking", 120, "w"), ("subject", "Note", 180, "w"), ("staff", "Logged by", 100, "w"),
                   ("status", "Status", 200, "w")]
        self.tree = ttk.Treeview(table, columns=[c[0] for c in columns], show="headings", style="Dark.Treeview",
                                 selectmode="browse")
        for key, text, width, anchor in columns:
            self.tree.heading(key, text=text, anchor=anchor)
            self.tree.column(key, width=S(width), minwidth=S(40), anchor=anchor,
                             stretch=key in ("to", "from", "subject", "status"))
        sb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("tampered", background=TAMPERED_BG, foreground=TAMPERED_FG)  # first = wins
        self.tree.tag_configure("even", background=ROW_A)
        self.tree.tag_configure("odd", background=ROW_B)
        self.tree.tag_configure(S_WAITING, foreground=YELLOW)
        self.tree.tag_configure(S_HANDED, foreground=MUTED)
        self.tree.tag_configure(S_SENT, foreground=TEXT)
        self.tree.tag_configure(S_VOIDED, foreground=DIM, font=(FONT, 10, "overstrike"))
        self.tree.bind("<Double-1>", lambda e: self.hand_over())
        make_sortable(self.tree)
        self.empty_label = tk.Label(table, text="", font=(FONT, 11), fg=MUTED, bg=ROW_A)
        self.set_filter("waiting")

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
        items = [it for it in self.store.items().values() if not it["voided_at"]]
        today = date.today()
        waiting = self.store.waiting()
        value, sub = self.cards["waiting"]
        value.config(text=str(len(waiting)), fg=YELLOW if waiting else TEXT)
        oldest = min((it["dt"] for it in waiting), default=None)
        sub.config(text=f"oldest arrived {oldest:%d/%m}" if oldest else "Nothing waiting")
        received = [it for it in items if it["incoming"] and it["dt"].date() == today]
        value, sub = self.cards["in"]
        value.config(text=str(len(received)))
        sub.config(text=", ".join(sorted({it["to"] for it in received})) or "—")
        sent = [it for it in items if not it["incoming"] and it["dt"].date() == today]
        value, sub = self.cards["out"]
        value.config(text=str(len(sent)))
        sub.config(text=f"postage {fmt_money(sum(it['cost'] for it in sent))}" if sent else "—")
        month = [it for it in items if not it["incoming"] and it["dt"].date() >= today.replace(day=1)]
        value, sub = self.cards["cost"]
        value.config(text=fmt_money(sum(it["cost"] for it in month)))
        sub.config(text=plural(len(month), "item") + " sent")

    def refresh_table(self):
        query = self.search.value().lower()
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        shown = 0
        for it in sorted(self.store.items().values(), key=lambda i: i["no"], reverse=True):
            status = self.store.status(it)
            if not {"waiting": status == S_WAITING, "in": it["incoming"] and status != S_VOIDED,
                    "out": not it["incoming"] and status != S_VOIDED, "voided": status == S_VOIDED,
                    "all": True}[self.filter]:
                continue
            what = it["item"] + (f" · {fmt_money(it['cost'])}" if it["cost"] else "")
            values = (item_name(it["no"]), it["dt"].strftime("%d/%m/%Y %H:%M"), "In" if it["incoming"] else "Out",
                      it["to"], it["from"], what, it["tracking"], it["subject"], it["staff"], status_text(it))
            if query and not any(query in str(v).lower() for v in values + (it["address"],)):
                continue
            tags = ["odd" if shown % 2 else "even", status] + (["tampered"] if it["tampered"] else [])
            self.tree.insert("", "end", iid=str(it["no"]), values=values, tags=tags)
            shown += 1
        self.tree.resort()
        keep = [i for i in selected if self.tree.exists(i)]
        if keep:
            self.tree.selection_set(keep)
        if shown:
            self.empty_label.place_forget()
        else:
            self.empty_label.config(text="No matching post." if query else {
                "waiting": "Nothing waiting to be handed over.", "in": "No incoming post logged yet.",
                "out": "No outgoing post logged yet.", "voided": "Nothing voided.",
                "all": "No post logged yet."}[self.filter])
            self.empty_label.place(relx=0.5, rely=0.5, anchor="center")

    def selected_item(self, what):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(what, "Select an item in the list first.", parent=self.app)
            return None
        return self.store.items().get(int(sel[0]))

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
    def log_item(self):
        if self.form_problem():
            return
        n = self.store.next_item()
        if self.mode == "in":
            to, frm = self.to_box.get().strip(), self.from_box.get().strip()
            tracking, subject = self.in_tracking.value(), self.in_subject.value()
            result = self._pin("Incoming post", f"Item {item_name(n)}  ·  for {to}",
                               [f"{self.kind} from {frm}"] + [x for x in (tracking and f"Tracking {tracking}",
                                                                          subject) if x],
                               confirm_text="Log incoming post")
            if not result:
                return
            row = {"Type": P_RECEIVED, "Item No": n, "To": to, "From": frm, "Item": self.kind,
                   "Tracking No": tracking, "Subject": subject, "Staff": result[0]["name"]}
            text = f"✓ Item {item_name(n)}: {self.kind.lower()} for {to} from {frm} logged. Hand it over when collected"
        else:
            to, cost = self.recipient_box.get().strip(), parse_pence(self.cost_entry.get()) or 0
            address, tracking = self.address_entry.value(), self.out_tracking.value()
            sender, subject = self.sender_entry.value(), self.out_subject.value()
            how = " · ".join(x for x in (self.method, tracking and f"tracking {tracking}",
                                         cost and f"postage {fmt_money(cost)}") if x)
            result = self._pin("Outgoing post", f"Item {item_name(n)}  ·  to {to}",
                               [x for x in (address, how, sender and f"For {sender}", subject) if x],
                               confirm_text="Log outgoing post")
            if not result:
                return
            row = {"Type": P_SENT, "Item No": n, "To": to, "From": sender, "Address": address,
                   "Item": self.method or "", "Tracking No": tracking, "Cost (£)": csv_money(cost) if cost else "",
                   "Subject": subject, "Staff": result[0]["name"]}
            text = f"✓ Item {item_name(n)}: post to {to} logged"
        if self.store.write([row], self.app):
            mode = self.mode
            self.clear_form()
            self.set_filter("waiting" if mode == "in" else "out")
            self._done(text)

    def hand_over(self):
        it = self.selected_item("Hand over")
        if not it:
            return
        if self.store.status(it) != S_WAITING:
            messagebox.showinfo("Hand over", f"Item {item_name(it['no'])} is not waiting to be handed over.",
                                parent=self.app)
            return
        res = FieldsDialog(self.app, "Hand over", f"Hand over item {item_name(it['no'])}", [
            {"key": "to", "label": "Given to (e.g. the CEO, or their PA)", "value": it["to"]},
        ], validate=lambda v: None if v["to"] else "Enter who it was given to.", confirm_text="Next",
            sub=f"{it['item']} for {it['to']} from {it['from']}, arrived {it['dt']:%d/%m/%Y}.").show()
        if not res:
            return
        result = self._pin("Hand over", f"Item {item_name(it['no'])}  ·  handed to {res['to']}",
                           [f"{it['item']} from {it['from']}", f"Arrived {it['dt']:%d/%m/%Y %H:%M}"],
                           confirm_text="Record hand-over")
        if result and self.store.write([{"Type": P_HANDED, "Item No": it["no"], "To": it["to"], "From": it["from"],
                                         "Handed To": res["to"], "Staff": result[0]["name"]}], self.app):
            self._done(f"✓ Item {item_name(it['no'])} handed to {res['to']} by {result[0]['name']}")

    def void_selected(self):
        it = self.selected_item("Void")
        if not it:
            return
        name = item_name(it["no"])
        if it["voided_at"]:
            messagebox.showinfo("Void", f"Item {name} is already voided.", parent=self.app)
            return
        if it["handed_at"]:
            messagebox.showinfo("Void", f"Item {name} has already been handed over, so it can't be voided.",
                                parent=self.app)
            return
        result = self._pin("Void", f"Void item {name}",
                           [f"{'For' if it['incoming'] else 'To'} {it['to']}"
                            + (f" from {it['from']}" if it["incoming"] else ""),
                            f"Logged {it['dt']:%d/%m/%Y} by {it['staff']}",
                            "Only for an entry made by mistake. It stays in the log, struck through."],
                           allow=lambda s: None if (s["admin"] or s.get("checker")) else
                           "Only an admin or checker can void an entry",
                           reason_label="Reason (e.g. entered twice)",
                           confirm_text="Void item", confirm_bg=RED, confirm_hover=RED_HOVER)
        if result and self.store.write([{"Type": P_VOIDED, "Item No": it["no"], "To": it["to"],
                                         "Staff": result[0]["name"], "Notes": result[1]}], self.app):
            self._done(f"Item {name} voided by {result[0]['name']}", AMBER)

    def export(self):
        path = admin_save_path(self.app, self.app, title="Export post log", defaultextension=".xlsx",
                               initialfile=f"Post log {date.today():%Y-%m-%d}.xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        rows = [[item_name(it["no"]), it["dt"].replace(second=0), "In" if it["incoming"] else "Out", it["to"],
                 it["from"], it["address"], it["item"], it["tracking"], it["cost"] / 100 if it["cost"] else None,
                 it["subject"], it["staff"], it["handed_to"], it["handed_at"].replace(second=0) if it["handed_at"]
                 else None, status_text(it)]
                for it in sorted(self.store.items().values(), key=lambda i: i["no"])]
        try:
            write_xlsx(path, ["Item", "Date", "In/Out", "To", "From / sent for", "Address", "What / how", "Tracking",
                              "Postage", "Note", "Logged by", "Handed to", "Handed over", "Status"], rows,
                       money_cols=("Postage",))
            open_file(path)
        except Exception as e:
            messagebox.showerror("Export", f"Could not export:\n\n{e}", parent=self.app)
