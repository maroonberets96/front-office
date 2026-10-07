# Front Office

A Windows desktop app (Python + Tkinter) for a small office front desk. It has these tabs, which share staff, PINs, settings and backups but keep their money and logs separate:

- **Laundry tokens**: point of sale for washing and dryer tokens, with a USB-triggered cash drawer
- **Drawer**: cash in the drawer, no sale / change, counts, petty cash and banking
- **Rent & deposits**: rent, arrears and deposits from receipt (numbered PDF) to bank, with checker sign-off
- **Keys**: keys lent to contractors, with printed slips to sign, through to their return
- **Post**: recorded post in (handed over) and out (method, tracking, cost)
- **Visitors**: the visitor book, with sign-in and sign-out by staff

Every action needs a staff PIN (PINs are stored hashed). Logs are append-only CSV files. Each row carries a hash chain, so edits made outside the app (for example in Excel) are flagged. A copy of each log is saved every day.

## Running

```
pip install pyserial pillow reportlab openpyxl
python source/front_office.py
```

The first run creates `data/`, `receipts/`, `banking/`, `keys/` and `backups/` next to `source/`. These folders are git-ignored. The default admin PIN is `1234`; the app warns you until you change it.

Set your organisation's name, address and contact line for receipts in **Settings → Receipts**.

## Building the .exe

```
pip install pyinstaller
source\build_exe.bat
```

This builds `Front Office.exe` as a single file in the repository root. To brand it, put a `logo.png` (and optionally an `app.ico`) in `source/` before you build.

## Updating a live install

Close the app, rename the old `.exe` to `.exe.old`, and copy in only the new `.exe`. Never copy the data folders from a test copy over the live ones.

Requires Windows 10/11, a PDF viewer and a default printer.
