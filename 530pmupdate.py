# ===== CHUNK 1 / 4 — START =====
# deathstar.py — The Death Star (GUI)
# v2025-10-01 — Email + Dialer + Warm Leads + Plain Paste RC + Resizable Columns
# Requires: PySimpleGUI, tksheet; (optional) pywin32 for Outlook

import os, sys, csv, html, time, re, hashlib, configparser, io, base64
from datetime import datetime, timedelta
from pathlib import Path

# --- Force vendored PySimpleGUI 4.60.x ---
_VENDOR_PSG = Path(__file__).parent / "vendor_psg"
if str(_VENDOR_PSG) not in sys.path:
    sys.path.insert(0, str(_VENDOR_PSG))
import PySimpleGUI as sg
# -----------------------------------------

set_theme = getattr(sg, "theme", getattr(sg, "ChangeLookAndFeel", None))


def _resolve_popup_error():
    for attr_name in ("popup_error", "PopupError", "popup", "Popup"):
        if hasattr(sg, attr_name):
            return getattr(sg, attr_name)
    return lambda *a, **k: None


popup_error = _resolve_popup_error()
if not hasattr(sg, "popup_error"):
    setattr(sg, "popup_error", popup_error)  # legacy compatibility

APP_VERSION = "2025-10-01"

# -------------------- App data locations --------------------
APP_DIR      = Path(os.environ.get("APPDATA", str(Path.home()))) / "DeathStarApp"  # Roaming
CSV_PATH     = APP_DIR / "kybercrystals.csv"
STATE_PATH   = APP_DIR / "annihilated_planets.txt"
TPL_PATH     = APP_DIR / "templates.ini"
RESULTS_PATH = APP_DIR / "results.csv"

# Dialer / warm / no-interest / customers / orders files
DIALER_RESULTS_PATH = APP_DIR / "dialer_results.csv"
WARM_LEADS_PATH     = APP_DIR / "warm_leads.csv"
NO_INTEREST_PATH    = APP_DIR / "no_interest.csv"
CUSTOMERS_PATH      = APP_DIR / "customers.csv"
ORDERS_PATH         = APP_DIR / "orders.csv"

# -------------------- Outlook settings ----------------------
DEATHSTAR_SUBFOLDER = "Order 66"   # Drafts subfolder
TARGET_MAILBOX_HINT = ""           # Optional: part of SMTP/display name to target a specific account

# -------------------- Columns / headers ---------------------
HEADER_FIELDS = [
    "Email","First Name","Last Name","Company","Industry","Phone",
    "Address","City","State","Reviews","Website"
]

# Warm Leads v1 headers are kept for compatibility; Part 2 upgrades to v2 at runtime.
WARM_FIELDS = [
    "Company","Prospect Name","Phone #","Email",
    "Location","Industry","Google Reviews","Rep","Samples?","Timestamp",
    "Call 1 Notes","Call 2 Date","Call 2 Notes","Call 3 Date","Call 3 Notes",
    "Call 4 Date","Call 4 Notes","Call 5 Date","Call 5 Notes","Call 6 Date","Call 6 Notes",
    "Call 7 Dates","Call 7 Notes","Call 8 Date","Call 8 Notes","Call 9 Date","Call 9 Notes",
    "Call 10 Date","Call 10 Notes","Call 11 Date","Call 11 Notes","Call 12 Date","Call 12 Notes",
    "Call 13 Date","Call 13 Notes","Call 14 Date","Call 14 Notes"
]

# Customers sheet headers (new schema)
CUSTOMER_FIELDS = [
    "Company","Name","Phone","Industry",
    "First Order Date","Last Order Date","First Contact","Days To Close",
    "CLTV","Sku's Ordered","Notes","Reorder?","Days","Sales/Day","Notes 2"
]

# -------------------- UI tuning -----------------------------
START_ROWS = 200
DEFAULT_COL_WIDTH = 140

# -------------------- Generic, brand-safe defaults ----------
DEFAULT_SUBJECT = "Quick intro from YOUR COMPANY"

DEFAULT_TEMPLATES = {
    "default": (
        "Hey {First Name},\n\n"
        "My name is YOUR NAME with YOUR COMPANY. We help {Industry} MAIN GOAL. "
        "If it’s useful, I can share examples or send over a couple of samples.\n\n"
        "Thanks,\n"
        "YOUR NAME\n"
        "YOUR COMPANY\n"
        "PHONE\n"
        "WEBSITE"
    ),
    "butcher_shop": (
        "Hey {First Name},\n\n"
        "My name is YOUR NAME with YOUR COMPANY. We help butcher shops MAIN GOAL. "
        "If it’s useful, I can share examples or send over a couple of samples.\n\n"
        "Thanks,\n"
        "YOUR NAME\n"
        "YOUR COMPANY\n"
        "PHONE\n"
        "WEBSITE"
    ),
    "farm_orchard": (
        "Hey {First Name},\n\n"
        "My name is YOUR NAME with YOUR COMPANY. We help farms & orchards MAIN GOAL. "
        "If it’s useful, I can share examples or send over a couple of samples.\n\n"
        "Thanks,\n"
        "YOUR NAME\n"
        "YOUR COMPANY\n"
        "PHONE\n"
        "WEBSITE"
    ),
}

DEFAULT_SUBJECTS = {
    "default":      DEFAULT_SUBJECT,
    "butcher_shop": DEFAULT_SUBJECT,
    "farm_orchard": DEFAULT_SUBJECT,
}

DEFAULT_MAP = {}

# -------------------- Regex helpers -------------------------
EMAIL_RE       = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
PLACEHOLDER_RE = re.compile(r"\{([^}]+)\}")
REF_RE         = re.compile(r"\[ref:([0-9a-f]{6,12})\]", re.IGNORECASE)

# ============================================================
# Bootstrapping & Persistence
# ============================================================

def ensure_app_files():
    APP_DIR.mkdir(parents=True, exist_ok=True)
    if not CSV_PATH.exists():
        with CSV_PATH.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(HEADER_FIELDS)
    if not STATE_PATH.exists():
        STATE_PATH.touch()
    if not TPL_PATH.exists():
        save_templates_ini(DEFAULT_TEMPLATES, DEFAULT_SUBJECTS, DEFAULT_MAP)
    if not RESULTS_PATH.exists():
        with RESULTS_PATH.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(["Ref","Email","Company","Industry","DateSent","DateReplied","Status","Subject"])
    ensure_dialer_files()
    ensure_warm_file()
    ensure_no_interest_file()
    ensure_orders_file()
    # customers.csv ensured when mounting grid (and in Part 2 before saving)

def save_templates_ini(templates_dict, subjects_dict, map_dict):
    cfg = configparser.ConfigParser()
    cfg["templates"] = templates_dict
    cfg["subjects"]  = subjects_dict
    cfg["map"]       = map_dict
    with TPL_PATH.open("w", encoding="utf-8") as f:
        cfg.write(f)

def load_templates_ini():
    cfg = configparser.ConfigParser()
    cfg.read(TPL_PATH, encoding="utf-8")
    tpls = dict(cfg["templates"]) if "templates" in cfg else dict(DEFAULT_TEMPLATES)
    subs = dict(cfg["subjects"])  if "subjects"  in cfg else dict(DEFAULT_SUBJECTS)
    mp   = dict(cfg["map"])       if "map"       in cfg else dict(DEFAULT_MAP)
    for k,v in DEFAULT_TEMPLATES.items(): tpls.setdefault(k, v)
    for k,v in DEFAULT_SUBJECTS.items():  subs.setdefault(k, v)
    return tpls, subs, mp

# ============================================================
# Results persistence
# ============================================================

def upsert_result(ref_short, email, company, industry, subject, sent_dt=None, replied_dt=None):
    rows = []
    if RESULTS_PATH.exists():
        with RESULTS_PATH.open("r", encoding="utf-8", newline="") as f:
            rows = list(csv.DictReader(f))
    idx = next((i for i,x in enumerate(rows) if x.get("Ref")==ref_short), None)
    rec = {
        "Ref": ref_short,
        "Email": email or "",
        "Company": company or "",
        "Industry": industry or "",
        "DateSent": sent_dt or (rows[idx]["DateSent"] if idx is not None else ""),
        "DateReplied": replied_dt or (rows[idx]["DateReplied"] if idx is not None else ""),
        "Status": rows[idx]["Status"] if idx is not None else "",
        "Subject": subject or (rows[idx]["Subject"] if idx is not None else ""),
    }
    if idx is None: rows.append(rec)
    else: rows[idx] = rec
    with RESULTS_PATH.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["Ref","Email","Company","Industry","DateSent","DateReplied","Status","Subject"])
        w.writeheader(); w.writerows(rows)

def set_status(ref_short, status):
    rows = []
    if RESULTS_PATH.exists():
        with RESULTS_PATH.open("r", encoding="utf-8", newline="") as f:
            rows = list(csv.DictReader(f))
    for r in rows:
        if r.get("Ref")==ref_short:
            r["Status"] = status; break
    with RESULTS_PATH.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["Ref","Email","Company","Industry","DateSent","DateReplied","Status","Subject"])
        w.writeheader(); w.writerows(rows)

def load_results_rows_sorted():
    rows = []
    if RESULTS_PATH.exists():
        with RESULTS_PATH.open("r", encoding="utf-8", newline="") as f:
            rows = list(csv.DictReader(f))
    def sk(r): return (r.get("DateReplied",""), r.get("DateSent",""))
    rows.sort(key=sk, reverse=True)
    return rows

# ============================================================
# Dialer / Warm / No Interest helpers
# ============================================================

def ensure_dialer_files():
    if not DIALER_RESULTS_PATH.exists():
        with DIALER_RESULTS_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow([
                "Timestamp","Outcome","Email","First Name","Last Name","Company","Industry",
                "Phone","Address","City","State","Reviews","Website","Note"
            ])

def ensure_warm_file():
    WARM_LEADS_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not WARM_LEADS_PATH.exists():
        with WARM_LEADS_PATH.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(WARM_FIELDS)
        return

    needs_rewrite = False
    rows = []

    with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        existing_fields = rdr.fieldnames or []
        if existing_fields != WARM_FIELDS:
            needs_rewrite = True
        rows = list(rdr) if rdr.fieldnames else []

    if needs_rewrite:
        with WARM_LEADS_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=WARM_FIELDS)
            w.writeheader()
            for row in rows:
                w.writerow({h: row.get(h, "") for h in WARM_FIELDS})

def ensure_no_interest_file():
    if not NO_INTEREST_PATH.exists():
        with NO_INTEREST_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow([
                "Timestamp","Email","First Name","Last Name","Company","Industry",
                "Phone","City","State","Website","Note","Source","NoContact"
            ])

def load_dialer_matrix_from_email_csv():
    """(Legacy) Start from the same CSV backing Email Leads (kybercrystals.csv)."""
    rows = []
    if CSV_PATH.exists():
        with CSV_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                rows.append([ (r.get(h,"") or "") for h in HEADER_FIELDS ])
    return rows

def dialer_save_call(row_dict, outcome, note):
    """Persist a single call. If outcome is green, also add to warm_leads.csv."""
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ensure_dialer_files()
    with DIALER_RESULTS_PATH.open("a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            ts, outcome,
            row_dict.get("Email",""), row_dict.get("First Name",""), row_dict.get("Last Name",""),
            row_dict.get("Company",""), row_dict.get("Industry",""), row_dict.get("Phone",""),
            row_dict.get("Address",""), row_dict.get("City",""), row_dict.get("State",""),
            row_dict.get("Reviews",""), row_dict.get("Website",""), note
        ])
    if (outcome or "").lower() == "green":
        with WARM_LEADS_PATH.open("a", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            # Warm v1 row; Part 2 will upgrade schema on load/save
            row = {h:"" for h in WARM_FIELDS}
            row["Company"]   = row_dict.get("Company","")
            row["Prospect Name"] = f"{row_dict.get('First Name','')} {row_dict.get('Last Name','')}".strip()
            row["Phone #"]   = row_dict.get("Phone","")
            row["Email"]     = row_dict.get("Email","")
            row["Location"]  = f"{row_dict.get('City','')}, {row_dict.get('State','')}".strip(", ")
            row["Industry"]  = row_dict.get("Industry","")
            row["Google Reviews"] = row_dict.get("Reviews","")
            row["Timestamp"] = ts
            w.writerow([row.get(h,"") for h in WARM_FIELDS])

def add_no_interest(row_dict, note, no_contact_flag: int, source: str):
    ensure_no_interest_file()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with NO_INTEREST_PATH.open("a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            ts,
            row_dict.get("Email",""), row_dict.get("First Name",""), row_dict.get("Last Name",""),
            row_dict.get("Company",""), row_dict.get("Industry",""), row_dict.get("Phone",""),
            row_dict.get("City",""), row_dict.get("State",""), row_dict.get("Website",""),
            note, source, no_contact_flag
        ])

# Email Results → Warm / No-interest helpers
def lead_lookup(email, company):
    rec = {}
    if CSV_PATH.exists():
        with CSV_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            rows = list(rdr)
        email_l = (email or "").strip().lower()
        for r in rows:
            if email_l and (r.get("Email","").strip().lower() == email_l):
                rec = r; break
        if not rec and company:
            comp_l = company.strip().lower()
            for r in rows:
                if (r.get("Company","").strip().lower() == comp_l):
                    rec = r; break
    return rec

def add_warm_from_result(result_row, note="Email Results"):
    ensure_warm_file()
    enr = lead_lookup(result_row.get("Email",""), result_row.get("Company",""))
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with WARM_LEADS_PATH.open("a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        row = {h:"" for h in WARM_FIELDS}
        row["Company"] = result_row.get("Company","")
        row["Prospect Name"] = f"{enr.get('First Name','')} {enr.get('Last Name','')}".strip()
        row["Phone #"] = enr.get("Phone","")
        row["Email"] = result_row.get("Email","")
        row["Location"] = f"{enr.get('City','')}, {enr.get('State','')}".strip(", ")
        row["Industry"] = result_row.get("Industry","")
        row["Google Reviews"] = enr.get("Reviews","")
        row["Timestamp"] = ts
        row["Call 1 Notes"] = note
        w.writerow([row.get(h,"") for h in WARM_FIELDS])

def add_no_interest_from_result(result_row, note="Email Results", no_contact_flag=0):
    ensure_no_interest_file()
    enr = lead_lookup(result_row.get("Email",""), result_row.get("Company",""))
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with NO_INTEREST_PATH.open("a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            ts,
            result_row.get("Email",""),
            enr.get("First Name",""), enr.get("Last Name",""),
            result_row.get("Company",""), result_row.get("Industry",""),
            enr.get("Phone",""), enr.get("City",""), enr.get("State",""),
            enr.get("Website",""),
            note, "EmailResults", no_contact_flag
        ])

# ============================================================
# Clipboard / Right-click / Column-resize helpers for tksheet
# ============================================================

def _bind_plaintext_paste_for_tksheet(sheet_obj, tk_root):
    """Force paste to use raw text (tabs/newlines). Bind Ctrl+V and Ctrl+Shift+V."""
    def _paste_plain(_evt=None):
        try:
            clip = tk_root.clipboard_get()
        except Exception:
            return "break"
        if not clip:
            return "break"
        rows = [line.split("\t") for line in clip.splitlines()]
        try:
            # tksheet >= 6
            sheet_obj.paste_data(rows)
        except Exception:
            # Fallback: manual cell set
            try:
                r0, c0 = sheet_obj.get_currently_selected()
            except Exception:
                r0, c0 = 0, 0
            for r_off, row in enumerate(rows):
                for c_off, val in enumerate(row):
                    sheet_obj.set_cell_data(r0 + r_off, c0 + c_off, val)
            sheet_obj.refresh()
        return "break"

    for w in (sheet_obj, getattr(sheet_obj, "MT", None), getattr(sheet_obj, "Toplevel", None)):
        if not w:
            continue
        try:
            w.bind("<Control-v>", _paste_plain)
            w.bind("<Control-V>", _paste_plain)
            w.bind("<Control-Shift-v>", _paste_plain)
            w.bind("<Control-Shift-V>", _paste_plain)
        except Exception:
            pass


def _ensure_rc_menu_plain_paste(sheet_obj, tk_root):
    """Replace/create right-click menu with only 'Paste (Ctrl+Shift+V)'."""
    def _paste_plain_from_menu():
        try:
            clip = tk_root.clipboard_get()
        except Exception:
            return
        if not clip:
            return
        rows = [line.split("\t") for line in clip.splitlines()]
        try:
            sheet_obj.paste_data(rows)
        except Exception:
            try:
                r0, c0 = sheet_obj.get_currently_selected()
            except Exception:
                r0, c0 = 0, 0
            for r_off, row in enumerate(rows):
                for c_off, val in enumerate(row):
                    sheet_obj.set_cell_data(r0 + r_off, c0 + c_off, val)
            sheet_obj.refresh()

    try:
        import tkinter as tk
        m = tk.Menu(sheet_obj.MT, tearoff=0)
        m.add_command(label="Paste (Ctrl+Shift+V)", command=_paste_plain_from_menu)

        def _popup(evt):
            try:
                m.tk_popup(evt.x_root, evt.y_root)
            finally:
                try:
                    m.grab_release()
                except Exception:
                    pass

        try:
            sheet_obj.MT.bind("<Button-3>", _popup)  # right click
            sheet_obj.MT.bind("<Button-2>", _popup)  # middle click (macs)
        except Exception:
            pass
    except Exception:
        pass


def _enable_column_resizing(sheet_obj):
    """
    Turn on column resizing regardless of tksheet version.
    Enables a bunch of possible flags; ignores ones that don't exist.
    """
    flags = [
        "column_width_resize",   # tksheet 6.x
        "column_resize",         # older alias
        "resize_columns",        # older alias
        "drag_select",
        "column_drag_and_drop",
    ]
    try:
        # tksheet >= 6.x
        sheet_obj.enable_bindings(tuple(set(sheet_obj.get_bindings() + tuple(flags))))
    except Exception:
        try:
            for fl in flags:
                try:
                    sheet_obj.enable_bindings((fl,))
                except Exception:
                    pass
        except Exception:
            pass
# ===== CHUNK 1 / 4 — END =====
# ===== CHUNK 2 / 4 — START =====
# ============================================================
# GUI
# ============================================================

def main():
    print(">>> ENTERING main()")
    ensure_app_files()
    templates, subjects, mapping = load_templates_ini()

    theme_applied = False
    if callable(set_theme):
        try:
            set_theme("DarkGrey13")
            theme_applied = True
        except Exception:
            theme_applied = False
    if not theme_applied:
        fallback_theme = getattr(sg, "SetOptions", None)
        if callable(fallback_theme):
            try:
                fallback_theme(background_color="#1B1B1B", text_color="#FFFFFF")
            except Exception:
                pass

    # ---------------- Toolbar with Update button ----------------
    top_bar = [
        sg.Text(f"Death Star v{APP_VERSION}", text_color="#9EE493"),
        sg.Push(),
        sg.Button("Update", key="-UPDATE-", button_color=("white","#444444"))
    ]

    # -------- Email Leads tab (host frame for tksheet) --------
    leads_host = sg.Frame(
        "KYBER CHAMBER (Spreadsheet — paste directly from Google Sheets / Excel)",
        [[sg.Text("Loading grid…", key="-LOADING-", text_color="#9EE493")]],
        relief=sg.RELIEF_GROOVE, border_width=2, background_color="#1B1B1B",
        title_color="#9EE493", expand_x=True, expand_y=True, key="-LEADS_HOST-",
    )

    # Controls UNDER the sheet, left-aligned
    leads_buttons_row1 = [
        sg.Button("Open Folder", key="-OPENFOLDER-"),
        sg.Button("Add 10 Rows", key="-ADDROWS-"),
        sg.Button("Delete Selected Rows", key="-DELROWS-"),
        sg.Button("Save Now", key="-SAVECSV-"),
        sg.Text("Status:", text_color="#A0A0A0"),
        sg.Text("Idle", key="-STATUS-", text_color="#FFFFFF"),
    ]
    leads_buttons_row2 = [
        sg.Button("Fire the Death Star", key="-FIRE-", size=(25,2), disabled=True, button_color=("white","#700000")),
        sg.Text(" (disabled: add valid NEW leads)", key="-FIRE_HINT-", text_color="#BBBBBB")
    ]

    leads_tab = [
        [leads_host],
        [sg.Text("Columns / placeholders:", text_color="#CCCCCC")],
        [sg.Text(", ".join(HEADER_FIELDS), text_color="#9EE493", font=("Consolas", 9))],
        [sg.Column([leads_buttons_row1], pad=(0,0))],
        [sg.Column([leads_buttons_row2], pad=(0,0))],
    ]

    # -------- Email Templates tab --------
    def tpl_val(k): return templates.get(k,"")
    def sub_val(k): return subjects.get(k, DEFAULT_SUBJECTS.get(k, DEFAULT_SUBJECT))
    known_keys = list(templates.keys())

    tpl_rows = []
    order = ["default","butcher_shop","farm_orchard"] + [k for k in known_keys if k not in ("default","butcher_shop","farm_orchard")]
    for key in order:
        body_height = 8 if key in ("butcher_shop","farm_orchard") else 6
        tpl_rows += [
            [sg.Text(key, size=(18,1), text_color="#CCCCCC")],
            [sg.Column([ [sg.Text("Subject", text_color="#9EE493")],
                         [sg.Input(default_text=sub_val(key), key=f"-SUBJ_{key}-", size=(48,1), enable_events=True)] ], pad=(0,0)),
             sg.Text("   "),
             sg.Column([ [sg.Text("Body", text_color="#9EE493")],
                         [sg.Multiline(default_text=tpl_val(key), key=f"-TPL_{key}-", size=(90, body_height),
                                       font=("Consolas",10), text_color="#EEE", background_color="#111", enable_events=True)] ],
                       pad=(0,0), expand_x=True)]
        ]

    tpl_tab = [
        [sg.Text("Templates support ANY header as {Placeholder}. {First Name} falls back to 'there' if blank.", text_color="#CCCCCC")],
        *tpl_rows,
        [sg.HorizontalSeparator(color="#4CAF50")],
        [sg.Text("Industry → Template mapping (needle -> template_key). Case-insensitive substring match.", text_color="#CCCCCC")],
        [sg.Multiline(default_text="\n".join([f"{k} -> {v}" for k,v in mapping.items()]) if mapping else "",
                      key="-MAP-", size=(120,6), font=("Consolas",10),
                      text_color="#EEE", background_color="#111", enable_events=True)],
        [sg.Button("Save Templates & Mapping", key="-SAVETPL-"),
         sg.Button("Reset to Defaults", key="-RESETTPL-"),
         sg.Push(),
         sg.Button("＋ Add New Template", key="-ADDNEW-", button_color=("white","#2E7D32")),
         sg.Button("Reload Templates", key="-RELOADTPL-"),
         sg.Text("", key="-TPL_STATUS-", text_color="#A0FFA0")]
    ]

    # -------- Email Results tab --------
    def results_table_data():
        rows = load_results_rows_sorted()
        data = [[r.get("Ref",""), r.get("Email",""), r.get("Company",""), r.get("Industry",""),
                 r.get("DateSent",""), r.get("DateReplied",""), r.get("Status",""), r.get("Subject","")] for r in rows]
        return rows, data
    rs_rows, rs_data = results_table_data()
    results_tab = [
        [sg.Text("Sync replies from Outlook; tag Green (good), Gray (neutral), Red (negative).", text_color="#CCCCCC")],
        [sg.Text("Lookback days:", text_color="#CCCCCC"), sg.Input("60", key="-LOOKBACK-", size=(6,1)),
         sg.Button("Sync from Outlook", key="-SYNC-"), sg.Text("", key="-RS_STATUS-", text_color="#A0FFA0")],
        [sg.Table(values=rs_data, headings=["Ref","Email","Company","Industry","DateSent","DateReplied","Status","Subject"],
                  auto_size_columns=False, col_widths=[10,26,26,14,18,18,8,40], justification="left", num_rows=15,
                  key="-RSTABLE-", enable_events=True, alternating_row_color="#2a2a2a",
                  text_color="#EEE", background_color="#111", header_text_color="#FFF", header_background_color="#333")],
        [sg.Button("Mark Green", key="-MARK_GREEN-", button_color=("white","#2E7D32")),
         sg.Button("Mark Gray",  key="-MARK_GRAY-",  button_color=("black","#DDDDDD")),
         sg.Button("Mark Red",   key="-MARK_RED-",   button_color=("white","#C62828")),
         sg.Text("   Warm Leads:", text_color="#A0A0A0"), sg.Text("0", key="-WARM-", text_color="#9EE493"),
         sg.Text("   Replies:", text_color="#A0A0A0"), sg.Text("0 / 0", key="-REPLRATE-", text_color="#FFFFFF")]
    ]

    # -------- Dialer tab (grid + outcome dots + notes) --------
    ensure_dialer_files()
    ensure_no_interest_file()

    DIALER_EXTRA_COLS = ["🙂","😐","☹️","Note1","Note2","Note3","Note4","Note5","Note6","Note7","Note8"]
    DIALER_HEADERS = HEADER_FIELDS + DIALER_EXTRA_COLS

    try:
        dialer_matrix = load_dialer_leads_matrix()
    except Exception:
        dialer_matrix = load_dialer_matrix_from_email_csv()
        if not dialer_matrix:
            dialer_matrix = [[""] * len(HEADER_FIELDS) for _ in range(50)]
        dialer_matrix = [row + ["○","○","○"] + ([""]*8) for row in dialer_matrix]

    if len(dialer_matrix) < 100:
        dialer_matrix += [[""] * len(HEADER_FIELDS) + ["○","○","○"] + ([""]*8)
                          for _ in range(100 - len(dialer_matrix))]

    dialer_host = sg.Frame(
        "DIALER GRID",
        [[sg.Text("Loading dialer grid…", key="-DIAL_LOAD-", text_color="#9EE493")]],
        relief=sg.RELIEF_GROOVE, border_width=2, background_color="#1B1B1B",
        title_color="#9EE493", expand_x=True, expand_y=True, key="-DIAL_HOST-",
    )

    dialer_controls_right = [
        [sg.Text("Outcome:", text_color="#CCCCCC")],
        [sg.Button("🟢 Green", key="-DIAL_SET_GREEN-", button_color=("white","#2E7D32"), size=(14,1))],
        [sg.Button("⚪ Gray",  key="-DIAL_SET_GRAY-",  button_color=("black","#DDDDDD"), size=(14,1))],
        [sg.Button("🔴 Red",   key="-DIAL_SET_RED-",   button_color=("white","#C62828"), size=(14,1))],
        [sg.Text("Note:", text_color="#CCCCCC", pad=((0,0),(10,0)))],
        [sg.Multiline(key="-DIAL_NOTE-", size=(28,6), font=("Consolas",10), background_color="#111", text_color="#EEE")],
        [sg.Button("Confirm Call", key="-DIAL_CONFIRM-", size=(16,2), disabled=True, button_color=("white","#444444"))],
        [sg.Text("", key="-DIAL_MSG-", text_color="#A0FFA0", size=(28,2))]
    ]

    dialer_buttons_under = [
        sg.Button("Add 100 Rows", key="-DIAL_ADD100-"),
    ]

    dialer_tab = [
        [sg.Column([[dialer_host],
                    [sg.Column([dialer_buttons_under], pad=(0,0))]],
                   expand_x=True, expand_y=True),
         sg.Column(dialer_controls_right, vertical_alignment="top", pad=((10,0),(0,0)))]
    ]

    # -------- Warm Leads tab --------
    warm_host = sg.Frame(
        "WARM LEADS GRID",
        [[sg.Text("Loading warm grid…", key="-WARM_LOAD-", text_color="#9EE493")]],
        relief=sg.RELIEF_GROOVE, border_width=2, background_color="#1B1B1B",
        title_color="#9EE493", expand_x=True, expand_y=True, key="-WARM_HOST-",
    )

    warm_controls_right = [
        [sg.Text("Outcome:", text_color="#CCCCCC")],
        [sg.Button("🟢 Green", key="-WARM_SET_GREEN-", button_color=("white","#2E7D32"), size=(14,1))],
        [sg.Button("⚪ Gray",  key="-WARM_SET_GRAY-",  button_color=("black","#DDDDDD"), size=(14,1))],
        [sg.Button("🔴 Red",   key="-WARM_SET_RED-",   button_color=("white","#C62828"), size=(14,1))],
        [sg.Text("Note:", text_color="#CCCCCC", pad=((0,0),(10,0)))],
        [sg.Multiline(key="-WARM_NOTE-", size=(28,6), font=("Consolas",10), background_color="#111", text_color="#EEE")],
        [sg.Button("Confirm", key="-WARM_CONFIRM-", size=(16,2), disabled=True, button_color=("white","#444444"))],
        [sg.Text("", key="-WARM_STATUS_SIDE-", text_color="#A0FFA0", size=(28,2))],
    ]

    warm_buttons_under = [
        sg.Button("Save Warm Leads", key="-WARM_SAVE-"),
        sg.Button("Export Warm Leads CSV", key="-WARM_EXPORT-"),
        sg.Button("Reload Warm", key="-WARM_RELOAD-"),
        sg.Button("Add 100 Rows", key="-WARM_ADD100-"),
        sg.Button("→ Confirm New Customer", key="-WARM_MARK_CUSTOMER-", button_color=("white","#2E7D32")),
        sg.Text("", key="-WARM_STATUS-", text_color="#A0FFA0"),
    ]

    warm_tab = [
        [sg.Column([[warm_host],
                    [sg.Column([warm_buttons_under], pad=(0,0))]],
                   expand_x=True, expand_y=True),
         sg.Column(warm_controls_right, vertical_alignment="top", pad=((10,0),(0,0)))]
    ]

    # -------- Customers tab (mirror Warm layout: grid left, slim panel right) --------
    if not CUSTOMERS_PATH.exists():
        with CUSTOMERS_PATH.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(CUSTOMER_FIELDS)

    customers_host = sg.Frame(
        "CUSTOMERS GRID",
        [[sg.Text("Loading customers grid…", key="-CUST_LOAD-", text_color="#9EE493")]],
        relief=sg.RELIEF_GROOVE, border_width=2, background_color="#1B1B1B",
        title_color="#9EE493", expand_x=True, expand_y=True, key="-CUST_HOST-",
    )

    customers_buttons_under = [
        sg.Button("Save Customers", key="-CUST_SAVE-"),
        sg.Button("Export Customers CSV", key="-CUST_EXPORT-"),
        sg.Button("Reload Customers", key="-CUST_RELOAD-"),
        sg.Button("Add 50 Rows", key="-CUST_ADD50-"),
        sg.Button("Add Order", key="-CUST_ADD_ORDER-", button_color=("white", "#2E7D32")),
        sg.Text("", key="-CUST_STATUS-", text_color="#A0FFA0")
    ]

    # Analytics panel (fixed width, compact like Warm's right panel)
    an_lines_top = [
        [sg.Text("CUSTOMER ANALYTICS:", text_color="#9EE493")],
        [sg.Text("Pipeline — Total Warm Leads:  "), sg.Text("0",    key="-AN_WARMS-",      text_color="#A0FFA0")],
        [sg.Text("Pipeline — Total Samples Sent ($):  "), sg.Text("0.00", key="-AN_SAMPLES-",    text_color="#A0FFA0")],
        [sg.Text("Pipeline — New Customers:  "), sg.Text("0",       key="-AN_NEWCUS-",     text_color="#A0FFA0")],
        [sg.Text("Pipeline — Close Rate:  "), sg.Text("0%",         key="-AN_CLOSERATE-",  text_color="#A0FFA0")],
        [sg.HorizontalSeparator(color="#4CAF50")],
        [sg.Text("CAC (Samples ÷ New Customers):  "), sg.Text("0.00", key="-AN_CAC-",        text_color="#A0FFA0")],
        [sg.Text("Average LTV (All Customers):  "),   sg.Text("0.00", key="-AN_AVGLTV-",     text_color="#A0FFA0")],
        [sg.Text("Reorder Rate (All Customers):  "),  sg.Text("0%",   key="-AN_REORDER-",    text_color="#A0FFA0")],
    ]
    analytics_panel = sg.Frame(
        "", [[sg.Column(an_lines_top, pad=(6,6), expand_x=True, expand_y=False)]],
        relief=sg.RELIEF_GROOVE, border_width=2,
        background_color="#1B1B1B", title_color="#9EE493",
        expand_x=False, expand_y=False
    )

    customers_tab = [
        [sg.Column([[customers_host],
                    [sg.Column([customers_buttons_under], pad=(0,0))]],
                   expand_x=True, expand_y=True),
         sg.Column([[analytics_panel]],
                   vertical_alignment="top",
                   pad=((10,0),(0,0)),
                   size=(320, 300))]  # ~Warm panel width
    ]

    # -------- Compose full layout --------
    layout = [
        top_bar,
        [sg.TabGroup([[sg.Tab("Email Leads",    leads_tab,   expand_x=True, expand_y=True),
                       sg.Tab("Email Templates",tpl_tab,     expand_x=True, expand_y=True),
                       sg.Tab("Email Results",  results_tab, expand_x=True, expand_y=True),
                       sg.Tab("Dialer",         dialer_tab,  expand_x=True, expand_y=True),
                       sg.Tab("Warm Leads",     warm_tab,    expand_x=True, expand_y=True),
                       sg.Tab("Customers",      customers_tab, expand_x=True, expand_y=True)]],
                     expand_x=True, expand_y=True)]
    ]

    # Build the window
    window = sg.Window(f"The Death Star — {APP_VERSION}", layout, finalize=True, resizable=True,
                       background_color="#202020", size=(1300, 900))

    # ---------- Mount tksheet grids AFTER window is finalized ----------
    try:
        from tksheet import Sheet as DialerSheet
        from tksheet import Sheet
    except Exception:
        popup_error("tksheet not installed. Run: pip install tksheet")
        return

    # Email Leads sheet in leads_host
    host_frame_tk = leads_host.Widget
    for child in host_frame_tk.winfo_children():
        try: child.destroy()
        except Exception: pass
    sheet_holder = sg.tk.Frame(host_frame_tk, bg="#111111")
    sheet_holder.pack(side="top", fill="both", expand=True)

    existing = []
    if CSV_PATH.exists():
        with CSV_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                existing.append([r.get(h,"") for h in HEADER_FIELDS])

    if existing:
        data = existing + [[""] * len(HEADER_FIELDS) for _ in range(max(0, START_ROWS - len(existing)))]
    else:
        data = [[""] * len(HEADER_FIELDS) for _ in range(START_ROWS)]

    sheet = Sheet(
        sheet_holder,
        data=data,
        headers=HEADER_FIELDS,
        show_x_scrollbar=True,
        show_y_scrollbar=True
    )
    sheet.enable_bindings((
        "single_select","row_select","column_select",
        "drag_select","column_drag_and_drop","row_drag_and_drop",
        "copy","cut","delete","undo","edit_cell","return_edit_cell",
        "select_all","right_click_popup_menu",
        "column_width_resize","column_resize","resize_columns"
    ))
    try:
        sheet.set_options(
            expand_sheet_if_paste_too_big=True,
            data_change_detected=True,
            show_vertical_grid=True,
            show_horizontal_grid=True,
        )
    except Exception:
        pass
    sheet.pack(fill="both", expand=True)
    for c in range(len(HEADER_FIELDS)):
        try: sheet.column_width(c, width=DEFAULT_COL_WIDTH)
        except Exception: pass

    _bind_plaintext_paste_for_tksheet(sheet, window.TKroot)
    _ensure_rc_menu_plain_paste(sheet, window.TKroot)
    _enable_column_resizing(sheet)

    # ---------- Dialer grid ----------
    dial_host_tk = dialer_host.Widget
    for child in dial_host_tk.winfo_children():
        try: child.destroy()
        except Exception: pass
    dial_sheet_holder = sg.tk.Frame(dial_host_tk, bg="#111111")
    dial_sheet_holder.pack(side="top", fill="both", expand=True)

    dial_sheet = DialerSheet(
        dial_sheet_holder,
        data=dialer_matrix,
        headers=DIALER_HEADERS,
        show_x_scrollbar=True,
        show_y_scrollbar=True
    )
    dial_sheet.enable_bindings((
        "single_select","row_select","column_select",
        "drag_select","copy","cut","delete","undo",
        "edit_cell","return_edit_cell","select_all","right_click_popup_menu",
        "column_width_resize","column_resize","resize_columns"
    ))
    try:
        dial_sheet.set_options(
            expand_sheet_if_paste_too_big=True,
            data_change_detected=True,
            show_vertical_grid=True,
            show_horizontal_grid=True,
            row_selected_background="#FFF8B3",
            row_selected_foreground="#000000",
        )
    except Exception:
        pass
    dial_sheet.pack(fill="both", expand=True)

    # Set dialer column widths
    def _idx(colname, default=None):
        try: return DIALER_HEADERS.index(colname)
        except Exception: return default
    idx_address = _idx("Address", 6)
    idx_city    = _idx("City", 7)
    idx_state   = _idx("State", 8)
    idx_reviews = _idx("Reviews", 9)
    idx_website = _idx("Website", 10)
    first_dot = len(HEADER_FIELDS)
    first_note = len(HEADER_FIELDS) + 3
    last_note  = first_note + 7

    for c in range(len(DIALER_HEADERS)):
        width = DEFAULT_COL_WIDTH
        if c == idx_address: width = 120
        if c == idx_city:    width = 90
        if c == idx_state:   width = 42
        if c == idx_reviews: width = 60
        if c == idx_website: width = 160
        if first_dot <= c < first_note:
            width = 36
        if first_note <= c <= last_note:
            width = 120
        try: dial_sheet.column_width(c, width=width)
        except Exception: pass
    _bind_plaintext_paste_for_tksheet(dial_sheet, window.TKroot)
    _ensure_rc_menu_plain_paste(dial_sheet, window.TKroot)
    _enable_column_resizing(dial_sheet)

    # ---------- Warm grid (tksheet) ----------
    warm_host_tk = warm_host.Widget
    for child in warm_host_tk.winfo_children():
        try: child.destroy()
        except Exception: pass
    warm_holder = sg.tk.Frame(warm_host_tk, bg="#111111")
    warm_holder.pack(side="top", fill="both", expand=True)

    warm_matrix = []
    headers_for_warm = None
    try:
        warm_matrix = load_warm_leads_matrix_v2()
        headers_for_warm = globals().get("WARM_V2_FIELDS", None)
    except Exception:
        pass

    if not warm_matrix:
        try:
            with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
                rdr = csv.DictReader(f)
                for r in rdr:
                    warm_matrix.append([r.get(h, "") for h in WARM_FIELDS])
            headers_for_warm = WARM_FIELDS
        except Exception:
            warm_matrix = []
            headers_for_warm = WARM_FIELDS

    if len(warm_matrix) < 100:
        warm_matrix += [[""] * len(headers_for_warm) for _ in range(100 - len(warm_matrix))]

    try:
        from tksheet import Sheet as WarmSheet
    except Exception:
        popup_error("tksheet not installed. Run: pip install tksheet")
        return

    warm_sheet = WarmSheet(
        warm_holder,
        data=warm_matrix,
        headers=headers_for_warm,
        show_x_scrollbar=True,
        show_y_scrollbar=True
    )
    warm_sheet.enable_bindings((
        "single_select","row_select","column_select",
        "drag_select","column_drag_and_drop","row_drag_and_drop",
        "copy","cut","delete","undo","edit_cell","return_edit_cell",
        "select_all","right_click_popup_menu",
        "column_width_resize","column_resize","resize_columns"
    ))
    try:
        warm_sheet.set_options(
            expand_sheet_if_paste_too_big=True,
            data_change_detected=True,
            show_vertical_grid=True,
            show_horizontal_grid=True,
        )
    except Exception:
        pass
    warm_sheet.pack(fill="both", expand=True)

    for c, name in enumerate(headers_for_warm):
        width = 120
        if name in ("Company","Prospect Name"): width = 180
        if name in ("Phone #","Rep","Samples?"): width = 90
        if name in ("Email","Google Reviews","Industry","Location"): width = 160
        if name.endswith("Date"): width = 110
        if name in ("Timestamp","Cost ($)"): width = 120
        try: warm_sheet.column_width(c, width=width)
        except Exception: pass

    _bind_plaintext_paste_for_tksheet(warm_sheet, window.TKroot)
    _ensure_rc_menu_plain_paste(warm_sheet, window.TKroot)
    _enable_column_resizing(warm_sheet)

    # ---------- Customers grid (tksheet) ----------
    cust_host_tk = customers_host.Widget
    for child in cust_host_tk.winfo_children():
        try: child.destroy()
        except Exception: pass
    cust_holder = sg.tk.Frame(cust_host_tk, bg="#111111")
    cust_holder.pack(side="top", fill="both", expand=True)

    customers_matrix = []
    try:
        with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                customers_matrix.append([r.get(h, "") for h in CUSTOMER_FIELDS])
    except Exception:
        customers_matrix = []
    if len(customers_matrix) < 50:
        customers_matrix += [[""] * len(CUSTOMER_FIELDS) for _ in range(50 - len(customers_matrix))]

    try:
        from tksheet import Sheet as CustomerSheet
    except Exception:
        popup_error("tksheet not installed. Run: pip install tksheet")
        return

    customer_sheet = CustomerSheet(
        cust_holder,
        data=customers_matrix,
        headers=CUSTOMER_FIELDS,
        show_x_scrollbar=True,
        show_y_scrollbar=True
    )
    customer_sheet.enable_bindings((
        "single_select","row_select","column_select",
        "drag_select","column_drag_and_drop","row_drag_and_drop",
        "copy","cut","delete","undo","edit_cell","return_edit_cell",
        "select_all","right_click_popup_menu",
        "column_width_resize","column_resize","resize_columns"
    ))
    try:
        customer_sheet.set_options(
            expand_sheet_if_paste_too_big=True,
            data_change_detected=True,
            show_vertical_grid=True,
            show_horizontal_grid=True,
        )
    except Exception:
        pass
    customer_sheet.pack(fill="both", expand=True)

    for c, name in enumerate(CUSTOMER_FIELDS):
        width = 120
        if name in ("Company","Prospect Name"): width = 180
        if name in ("Phone #","Rep","Samples?"): width = 90
        if name in ("Email","Google Reviews","Industry","Location"): width = 160
        if name in ("Opening Order $","Customer Since","Timestamp"): width = 150
        try: customer_sheet.column_width(c, width=width)
        except Exception: pass

    _bind_plaintext_paste_for_tksheet(customer_sheet, window.TKroot)
    _ensure_rc_menu_plain_paste(customer_sheet, window.TKroot)
    _enable_column_resizing(customer_sheet)

    # ---------- Handoff to Part 2 event loop ----------
    try:
        main_after_mount(
            window=window,
            sheet=sheet,
            dial_sheet=dial_sheet,
            leads_host=leads_host,
            dialer_host=dialer_host,
            templates=templates,
            subjects=subjects,
            mapping=mapping,
            warm_sheet=warm_sheet,
            customer_sheet=customer_sheet,
        )
    finally:
        try:
            window.close()
        except Exception:
            pass
    print(">>> EXITING main()")
# ===== CHUNK 2 / 4 — END =====
# ===== CHUNK 3 / 4 — START =====
# ============================================================
# CSV I/O
# ============================================================

def load_csv_to_matrix():
    rows = []
    if not CSV_PATH.exists():
        return rows
    with CSV_PATH.open("r", encoding="utf-8", newline="") as f:
        data = list(csv.reader(f))
    if not data:
        return rows
    for row in data[1:]:
        rows.append((row + [""] * len(HEADER_FIELDS))[:len(HEADER_FIELDS)])
    return rows


def save_matrix_to_csv(matrix):
    with CSV_PATH.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(HEADER_FIELDS)
        for row in matrix:
            w.writerow((row + [""] * len(HEADER_FIELDS))[:len(HEADER_FIELDS)])


# ---- Safe write helpers for Warm Leads / Customers / Dialer leads ----
BACKUP_DIR = APP_DIR / "_backups"

def _ts():
    from datetime import datetime as _dt
    return _dt.now().strftime("%Y%m%d-%H%M%S")

def _backup(path: Path):
    """Best-effort backup (binary copy) to APP_DIR/_backups with timestamp."""
    try:
        if path.exists() and path.is_file():
            BACKUP_DIR.mkdir(parents=True, exist_ok=True)
            bak = BACKUP_DIR / f"{path.stem}.{_ts()}{path.suffix}.bak"
            with path.open("rb") as s, bak.open("wb") as d:
                d.write(s.read())
    except Exception:
        pass

def _atomic_write_csv(path: Path, headers: list, rows: list):
    """Write CSV to a temporary file, then replace target atomically."""
    tmp = path.with_suffix(path.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        for row in rows:
            out = (list(row) + [""] * len(headers))[:len(headers)]
            w.writerow(out)
    tmp.replace(path)


# ============================================================
# Warm Leads v2 schema helpers
# ============================================================

def _build_warm_v2_fields():
    # Part 1 ensures file migration; we mirror the header list here
    base = [
        "Company","Prospect Name","Phone #","Email",
        "Location","Industry","Google Reviews","Rep","Samples?","Timestamp"
    ]
    try:
        ts_idx = base.index("Timestamp")
    except ValueError:
        ts_idx = len(base)
    new_fields = base[:ts_idx] + ["Cost ($)", "Timestamp"] + [f"Call {i}" for i in range(1, 16)] + base[ts_idx+1:]
    return new_fields

WARM_V2_FIELDS = _build_warm_v2_fields()

def load_warm_leads_matrix_v2():
    """Load rows under WARM_V2_FIELDS (auto-filling missing cols with '')."""
    rows = []
    if not WARM_LEADS_PATH.exists():
        return rows
    with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            rows.append([r.get(h, "") for h in WARM_V2_FIELDS])
    return rows

def save_warm_leads_matrix_v2(matrix):
    """Save rows using the WARM_V2_FIELDS header."""
    _backup(WARM_LEADS_PATH)
    _atomic_write_csv(WARM_LEADS_PATH, WARM_V2_FIELDS, matrix)


# ============================================================
# Dialer leads own CSV helpers
# ============================================================

DIALER_LEADS_PATH = APP_DIR / "dialer_leads.csv"

def ensure_dialer_leads_file():
    """Ensure the dialer grid CSV exists with the expected headers."""
    if not DIALER_LEADS_PATH.exists():
        with DIALER_LEADS_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            # Header: left EMAIL HEADER_FIELDS + three outcome columns + 8 notes
            hdr = HEADER_FIELDS + ["🙂","😐","☹️"] + [f"Note{i}" for i in range(1,9)]
            w.writerow(hdr)

def load_dialer_leads_matrix():
    """Load dialer leads rows. If legacy header is detected, adapt best-effort."""
    ensure_dialer_leads_file()
    rows = []
    with DIALER_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.reader(f)
        raw = list(rdr)
    if not raw:
        return rows
    hdr = raw[0]
    expected = HEADER_FIELDS + ["🙂","😐","☹️"] + [f"Note{i}" for i in range(1,9)]
    # If headers match, simple load; else adapt columns where possible
    idx_map = [hdr.index(h) if h in hdr else None for h in expected]
    for row in raw[1:]:
        out = []
        for i, idx in enumerate(idx_map):
            if idx is None:
                # outcome cols default to hollow dots, notes blank
                if i >= len(HEADER_FIELDS) and i < len(HEADER_FIELDS)+3:
                    out.append("○")
                else:
                    out.append("")
            else:
                try:
                    out.append(row[idx])
                except Exception:
                    out.append("")
        rows.append(out)
    return rows

def save_dialer_leads_matrix(matrix):
    """Save dialer grid with its own headers."""
    _backup(DIALER_LEADS_PATH)
    headers = HEADER_FIELDS + ["🙂","😐","☹️"] + [f"Note{i}" for i in range(1,9)]
    _atomic_write_csv(DIALER_LEADS_PATH, headers, matrix)


# ============================================================
# Customers CSV helpers
# ============================================================

def ensure_customers_file():
    """Ensure customers.csv exists with the new schema."""
    if not CUSTOMERS_PATH.exists():
        _backup(CUSTOMERS_PATH)
        with CUSTOMERS_PATH.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(CUSTOMER_FIELDS)

def load_customers_matrix():
    """Load customers.csv into a matrix matching CUSTOMER_FIELDS."""
    ensure_customers_file()
    rows = []
    with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        if not rdr.fieldnames:
            return rows
        for r in rdr:
            rows.append([r.get(h, "") for h in CUSTOMER_FIELDS])
    return rows

def save_customers_matrix(matrix):
    """Save matrix to customers.csv with backup + atomic replace."""
    ensure_customers_file()
    _backup(CUSTOMERS_PATH)
    _atomic_write_csv(CUSTOMERS_PATH, CUSTOMER_FIELDS, matrix)

def update_customer_row_fields_by_company(company, updates: dict):
    """
    Update the first matching customer row (by Company exact match, case-insensitive)
    with fields from `updates`. Only keys present in CUSTOMER_FIELDS are applied.
    """
    ensure_customers_file()
    try:
        with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            rows = list(rdr)
            flds = rdr.fieldnames or CUSTOMER_FIELDS
    except Exception:
        rows, flds = [], CUSTOMER_FIELDS

    comp_l = (company or "").strip().lower()
    found = False
    for r in rows:
        if (r.get("Company","") or "").strip().lower() == comp_l:
            for k,v in updates.items():
                if k in CUSTOMER_FIELDS:
                    r[k] = v
            found = True
            break

    if not found:
        # If company not found, append a new row with only provided fields.
        new_row = {h:"" for h in CUSTOMER_FIELDS}
        new_row["Company"] = company or ""
        for k,v in updates.items():
            if k in CUSTOMER_FIELDS:
                new_row[k] = v
        rows.append(new_row)

    _backup(CUSTOMERS_PATH)
    with CUSTOMERS_PATH.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CUSTOMER_FIELDS)
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h,"") for h in CUSTOMER_FIELDS})


# ============================================================
# Orders CSV helpers (for analytics / CLTV updates)
# ============================================================

# Guard in case not declared in Chunk 1 for any reason
if "ORDERS_PATH" not in globals():
    ORDERS_PATH = APP_DIR / "orders.csv"

def ensure_orders_file():
    """Create orders.csv if missing."""
    if not ORDERS_PATH.exists():
        with ORDERS_PATH.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(["Company","Order Date","Amount"])

def _parse_date_mmddyyyy(s):
    s = (s or "").strip()
    if not s:
        return None
    # Try common formats: YYYY-MM-DD, MM/DD/YYYY, MM-DD-YYYY
    for fmt in ("%Y-%m-%d","%m/%d/%Y","%m-%d-%Y","%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    # Also accept MM/DD and assume current year
    for fmt in ("%m/%d","%m-%d"):
        try:
            d = datetime.strptime(s, fmt)
            return d.replace(year=datetime.now().year).date()
        except Exception:
            pass
    return None

def _money_to_float(val):
    s = (val or "").strip().replace(",", "").replace("$","")
    if not s:
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0

def _float_to_money(x):
    try:
        return f"{float(x):.2f}"
    except Exception:
        return ""

def append_order_row(company: str, order_date: str, amount: str):
    """
    Append an order to orders.csv. order_date can be YYYY-MM-DD or MM/DD/YYYY.
    amount can include $ and commas. After writing, recompute and update the
    customer's First/Last Order, CLTV, Days, Sales/Day.
    """
    ensure_orders_file()
    company = (company or "").strip()
    if not company:
        raise ValueError("Company is required for orders.")

    # Normalize fields
    dt = _parse_date_mmddyyyy(order_date)
    if dt is None:
        # If no/invalid date, default to today
        dt = datetime.now().date()
    amt_f = _money_to_float(amount)

    # Write the order
    with ORDERS_PATH.open("a", encoding="utf-8", newline="") as f:
        csv.writer(f).writerow([company, dt.strftime("%Y-%m-%d"), _float_to_money(amt_f)])

    # Recompute stats and update customers.csv
    stats = compute_customer_order_stats(company)
    # Prepare updates
    updates = {}
    if stats["first_order_date"]:
        updates["First Order Date"] = stats["first_order_date"].strftime("%Y-%m-%d")
    if stats["last_order_date"]:
        updates["Last Order Date"] = stats["last_order_date"].strftime("%Y-%m-%d")
    updates["CLTV"] = _float_to_money(stats["cltv"])
    updates["Days"] = str(stats["days_since_first"]) if stats["days_since_first"] is not None else ""
    updates["Sales/Day"] = _float_to_money(stats["sales_per_day"]) if stats["sales_per_day"] is not None else ""

    update_customer_row_fields_by_company(company, updates)

def compute_customer_order_stats(company: str):
    """
    Scan orders.csv for this company and compute:
      - first_order_date (date or None)
      - last_order_date  (date or None)
      - cltv (float)
      - days_since_first (int or None)
      - sales_per_day (float or None)
    """
    ensure_orders_file()
    comp_l = (company or "").strip().lower()
    first_dt = None
    last_dt  = None
    total = 0.0

    with ORDERS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            if (r.get("Company","") or "").strip().lower() != comp_l:
                continue
            d = _parse_date_mmddyyyy(r.get("Order Date",""))
            a = _money_to_float(r.get("Amount",""))
            if d:
                if first_dt is None or d < first_dt:
                    first_dt = d
                if last_dt is None or d > last_dt:
                    last_dt = d
            total += a

    days_since_first = None
    sales_per_day = None
    if first_dt:
        days_since_first = max(1, (datetime.now().date() - first_dt).days)  # avoid div by zero
        sales_per_day = total / float(days_since_first) if days_since_first else None

    return {
        "first_order_date": first_dt,
        "last_order_date": last_dt,
        "cltv": total,
        "days_since_first": days_since_first,
        "sales_per_day": sales_per_day,
    }


# ============================================================
# Utilities (placeholders, keys, fingerprints)
# ============================================================

def normalize_header_map(row_dict):
    norm = {}
    for k,v in row_dict.items():
        if not k: continue
        k2 = k.strip()
        if not k2: continue
        norm[k2] = v
        norm[k2.lower()] = v
        norm[k2.replace(" ","_").lower()] = v
    return norm

PLACEHOLDER_RE = re.compile(r"\{([^}]+)\}")

def apply_placeholders(text, row_dict, profile=None):
    if not text:
        return ""
    values = normalize_header_map(row_dict)
    def repl(m):
        token = m.group(1).strip()
        cands = [token, token.lower(), token.replace(" ","_").lower()]
        token_norm = token.replace("_"," ").strip().lower()
        for c in cands:
            if c in values and values[c] != "":
                return values[c]
        if token_norm in ("first name","firstname","first"):
            return "there"
        return "{"+token+"}"
    return PLACEHOLDER_RE.sub(repl, text)

def dict_from_row(row):
    return {HEADER_FIELDS[i]: (row[i] if i < len(HEADER_FIELDS) else "") for i in range(len(HEADER_FIELDS))}

def get_val(d, name):
    lower = { (k or "").lower(): v for k,v in d.items() }
    return lower.get((name or "").lower(), "") or ""

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
def valid_email(addr): return bool(addr and EMAIL_RE.match(addr))

def row_fingerprint_from_dict(d):
    key = "|".join([
        (get_val(d,"Email") or "").lower(),
        (get_val(d,"First Name") or "").lower(),
        (get_val(d,"Company") or "").lower(),
        (get_val(d,"Industry") or "").lower(),
    ])
    return hashlib.sha1(key.encode("utf-8")).hexdigest()

def choose_template_key(industry_value, mapping):
    ind = (industry_value or "").lower()
    for needle, key in mapping.items():
        if needle.lower() in ind:
            return key
    return "default"

def blocks_to_html(text):
    text = text.replace("\r\n","\n")
    parts = [p for p in text.split("\n\n") if p.strip()!=""]
    out = []
    for p in parts:
        esc = html.escape(p).replace("\n","<br>")
        out.append(f'<p style="margin:0 0 12px 0;">{esc}</p>')
    return "\n".join(out) if out else "<p></p>"


# ============================================================
# Outlook helpers
# ============================================================

REF_RE = re.compile(r"\[ref:([0-9a-f]{6,12})\]", re.IGNORECASE)

def require_pywin32():
    try:
        import win32com.client as win32  # noqa
        return True
    except Exception:
        return False

def pick_store(session):
    store = session.DefaultStore
    if TARGET_MAILBOX_HINT:
        hint = TARGET_MAILBOX_HINT.lower()
        for i in range(1, session.Accounts.Count + 1):
            acc = session.Accounts.Item(i)
            smtp = (getattr(acc, "SmtpAddress", "") or "").lower()
            disp = (acc.DisplayName or "").lower()
            if hint in smtp or hint in disp:
                try:
                    store = acc.DeliveryStore
                    return store
                except Exception:
                    pass
        for i in range(1, session.Stores.Count + 1):
            st = session.Stores.Item(i)
            if hint in (st.DisplayName or "").lower():
                store = st
                break
    return store

def outlook_draft_many(rows_matrix, seen_set, templates, subjects, mapping):
    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    session = outlook.GetNamespace("MAPI")
    store = pick_store(session)
    drafts_root = store.GetDefaultFolder(16)  # olFolderDrafts
    target_folder = None
    for i in range(1, drafts_root.Folders.Count + 1):
        f = drafts_root.Folders.Item(i)
        if (f.Name or "").lower() == DEATHSTAR_SUBFOLDER.lower():
            target_folder = f; break
    if target_folder is None:
        target_folder = drafts_root.Folders.Add(DEATHSTAR_SUBFOLDER)
    made = 0
    new_fps = []
    for row in rows_matrix:
        d = dict_from_row(row)
        if not valid_email(d.get("Email","")):
            continue
        fp = row_fingerprint_from_dict(d)
        if fp in seen_set:
            continue
        ref_short = fp[:8]
        tpl_key = choose_template_key(d.get("Industry",""), mapping)
        body_tpl = templates.get(tpl_key, templates.get("default",""))
        subj_tpl = subjects.get(tpl_key) or subjects.get("default") or DEFAULT_SUBJECT
        subj_text = apply_placeholders(subj_tpl, d)
        body_text = apply_placeholders(body_tpl, d)
        body_html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"></head>
        <body style="margin:0;padding:0;"><div style="font-family:Segoe UI, Arial, sans-serif; font-size:14px; line-height:1.5; color:#111;">
        {blocks_to_html(body_text)}<!-- ref:{ref_short} --></div></body></html>"""
        msg = drafts_root.Items.Add("IPM.Note")
        msg.To = d.get("Email","")
        msg.Subject = f"{subj_text} [ref:{ref_short}]"
        msg.BodyFormat = 2
        msg.HTMLBody = body_html
        msg.Save()
        msg.Move(target_folder)
        upsert_result(ref_short, d.get("Email",""), d.get("Company",""), d.get("Industry",""), subj_text)
        made += 1
        new_fps.append(fp)
        time.sleep(0.02)
    with STATE_PATH.open("a", encoding="utf-8") as f:
        for fp in new_fps:
            f.write(fp+"\n")
    return made

def load_state_set():
    if not STATE_PATH.exists():
        return set()
    return {line.strip() for line in STATE_PATH.read_text(encoding="utf-8").splitlines() if line.strip()}

def load_results_rows_sorted():
    rows = []
    if RESULTS_PATH.exists():
        with RESULTS_PATH.open("r", encoding="utf-8", newline="") as f:
            rows = list(csv.DictReader(f))
    def sk(r): return (r.get("DateReplied",""), r.get("DateSent",""))
    rows.sort(key=sk, reverse=True)
    return rows

def outlook_sync_results(lookback_days=60):
    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    session = outlook.GetNamespace("MAPI")
    store = pick_store(session)
    sent = store.GetDefaultFolder(5)  # Sent Items
    inbox = store.GetDefaultFolder(6) # Inbox
    since = (datetime.now() - timedelta(days=lookback_days)).strftime("%m/%d/%Y %I:%M %p")
    sent_items = sent.Items; sent_items.IncludeRecurrences = True; sent_items.Sort("[SentOn]", True)
    inbox_items = inbox.Items; inbox_items.IncludeRecurrences = True; inbox_items.Sort("[ReceivedTime]", True)
    sent_recent = sent_items.Restrict(f"[SentOn] >= '{since}'")
    inbox_recent = inbox_items.Restrict(f"[ReceivedTime] >= '{since}'")
    sent_map, reply_map = {}, {}
    for i in range(1, min(2000, sent_recent.Count)+1):
        try:
            m = sent_recent.Item(i)
            subj = str(getattr(m,"Subject","") or "")
            rm = REF_RE.search(subj)
            if rm:
                sent_map[rm.group(1).lower()] = str(getattr(m,"SentOn","") or "")
        except Exception:
            continue
    for i in range(1, min(2000, inbox_recent.Count)+1):
        try:
            m = inbox_recent.Item(i)
            subj = str(getattr(m,"Subject","") or "")
            rm = REF_RE.search(subj)
            if rm:
                reply_map[rm.group(1).lower()] = str(getattr(m,"ReceivedTime","") or "")
        except Exception:
            continue
    rows = load_results_rows_sorted()
    byref = {r["Ref"].lower(): r for r in rows}
    for ref, dt in sent_map.items():
        if ref in byref:
            byref[ref]["DateSent"] = dt
        else:
            byref[ref] = {"Ref":ref,"Email":"","Company":"","Industry":"","DateSent":dt,"DateReplied":"","Status":"","Subject":""}
    for ref, dt in reply_map.items():
        if ref in byref:
            byref[ref]["DateReplied"] = dt
        else:
            byref[ref] = {"Ref":ref,"Email":"","Company":"","Industry":"","DateSent":"","DateReplied":dt,"Status":"","Subject":""}
    out = list(byref.values())
    with RESULTS_PATH.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["Ref","Email","Company","Industry","DateSent","DateReplied","Status","Subject"])
        w.writeheader(); w.writerows(out)
    return len(sent_map), len(reply_map)


# ============================================================
# In-window helpers, event loop, entry point
# ============================================================

# -------- Dialer helper shims (used by main_after_mount) --------
def dialer_cols_info(_headers=None):
    first_dot  = len(HEADER_FIELDS)          # 🙂/😐/☹️
    last_dot   = first_dot + 2
    first_note = len(HEADER_FIELDS) + 3      # Note1
    last_note  = first_note + 7              # Note8
    return {"first_dot": first_dot, "last_dot": last_dot,
            "first_note": first_note, "last_note": last_note}

_DOT_BG = {"green": "#2E7D32", "gray": "#9E9E9E", "red": "#C62828"}
_DOT_FG = {"green": "#FFFFFF", "gray": "#000000", "red": "#FFFFFF"}

def dialer_clear_dot_highlights(sheet, row, cols):
    try:
        for c in range(cols["first_dot"], cols["last_dot"]+1):
            sheet.highlight_cells(row=row, column=c, bg=None, fg=None)
    except Exception:
        pass

def dialer_colorize_outcome(sheet, row, outcome, cols=None):
    if cols is None: cols = dialer_cols_info(None)
    base = cols["first_dot"]
    try:
        for i in range(3):
            sheet.set_cell_data(row, base+i, "○")
        idx = {"green":0, "gray":1, "red":2}[outcome]
        c = base + idx
        sheet.set_cell_data(row, c, "●")
        try:
            sheet.highlight_cells(row=row, column=c, bg=_DOT_BG[outcome], fg=_DOT_FG[outcome])
        except Exception:
            pass
        sheet.refresh()
    except Exception:
        pass

def dialer_next_empty_note_col(sheet, row, cols=None):
    if cols is None: cols = dialer_cols_info(None)
    try:
        r = sheet.get_row_data(row) or []
    except Exception:
        return None
    for c in range(cols["first_note"], cols["last_note"]+1):
        if c >= len(r) or not (r[c] or "").strip():
            return c
    return None

def dialer_move_to_next_row(sheet, current_row):
    try:
        total = sheet.get_total_rows()
    except Exception:
        total = 0
    nxt = current_row + 1 if total == 0 else min(current_row + 1, max(0, total - 1))
    try:
        sheet.set_currently_selected(nxt, 0)
        sheet.see(nxt, 0)
    except Exception:
        pass
    return nxt


# -------------------- Warm helpers --------------------
def warm_get_col_index_map():
    """Return quick indices for WARM_V2_FIELDS."""
    idx = {name: i for i, name in enumerate(WARM_V2_FIELDS)}
    return {
        "cost": idx.get("Cost ($)"),
        "timestamp": idx.get("Timestamp"),
        "first_call": idx.get("Call 1"),
        "last_call": idx.get("Call 15"),
    }

def warm_next_empty_call_col(row_values, col_map):
    """Find next empty Call N cell (1..15)."""
    if not row_values: return None
    c1, cN = col_map["first_call"], col_map["last_call"]
    if c1 is None or cN is None: return None
    for c in range(c1, cN+1):
        cell = row_values[c] if c < len(row_values) else ""
        if not (cell or "").strip():
            return c
    return None

def warm_format_cost(val):
    """Normalize to dollars string '123.45'. Keep empty if blank."""
    s = (val or "").strip().replace(",", "")
    if not s:
        return ""
    try:
        return f"{float(s):.2f}"
    except Exception:
        return s  # leave as-is if non-numeric
# ===== CHUNK 3 / 4 — END =====
# ===== CHUNK 4 / 4 — START =====
# ============================================================
# main_after_mount
# ============================================================

def main_after_mount(window, sheet, dial_sheet, leads_host, dialer_host, templates, subjects, mapping, warm_sheet, customer_sheet):
    # ---------- extractors ----------
    def matrix_from_sheet():
        raw = sheet.get_sheet_data() or []
        trimmed = []
        for row in raw:
            row = (list(row) + [""] * len(HEADER_FIELDS))[:len(HEADER_FIELDS)]
            trimmed.append([str(c) for c in row])
        while trimmed and not any((cell or "").strip() for cell in trimmed[-1]):
            trimmed.pop()
        return trimmed

    def warm_matrix_from_sheet_v2():
        raw = warm_sheet.get_sheet_data() or []
        trimmed = []
        for row in raw:
            row = (list(row) + [""] * len(WARM_V2_FIELDS))[:len(WARM_V2_FIELDS)]
            trimmed.append([str(c) for c in row])
        while trimmed and not any((cell or "").strip() for cell in trimmed[-1]):
            trimmed.pop()
        return trimmed

    def customers_matrix_from_sheet():
        raw = customer_sheet.get_sheet_data() or []
        trimmed = []
        for row in raw:
            row = (list(row) + [""] * len(CUSTOMER_FIELDS))[:len(CUSTOMER_FIELDS)]
            trimmed.append([str(c) for c in row])
        while trimmed and not any((cell or "").strip() for cell in trimmed[-1]):
            trimmed.pop()
        return trimmed

    # ----- dialer helpers / state -----
    cols = dialer_cols_info(dial_sheet.headers() if hasattr(dial_sheet, "headers") else None)
    state = {
        "row": None,
        "outcome": None,          # "green" | "gray" | "red" | None
        "note_col_by_row": {},    # sticky preview slot: {row_idx: col_idx}
        "colored_row": None,      # which row currently has the colored outcome dot
    }

    def _row_selected(sheet_obj):
        try:
            sel = sheet_obj.get_selected_rows() or []
            if sel:
                return sel[0]
        except Exception:
            pass
        try:
            r, _ = sheet_obj.get_currently_selected()
            if isinstance(r, int) and r >= 0:
                return r
        except Exception:
            pass
        return None

    def _set_working_row(r):
        state["row"] = r
        try:
            dial_sheet.set_currently_selected(r, 0)
            dial_sheet.see(r, 0)
        except Exception:
            pass

    def _current_note_text():
        return (window["-DIAL_NOTE-"].get() or "").strip()

    def _confirm_enabled():
        r = state["row"]
        if r is None:
            return False
        have_outcome = state["outcome"] in ("green","gray","red")
        have_text    = bool(_current_note_text())
        sticky = state["note_col_by_row"].get(r)
        have_slot = (sticky is not None) or (dialer_next_empty_note_col(dial_sheet, r, cols) is not None)
        return have_outcome and have_text and have_slot

    def _update_confirm_button():
        ok = _confirm_enabled()
        window["-DIAL_CONFIRM-"].update(disabled=not ok, button_color=("white", "#2E7D32" if ok else "#444444"))

    def _apply_outcome(r, which):
        state["outcome"] = which
        try:
            xv = dial_sheet.MT.xview(); yv = dial_sheet.MT.yview()
        except Exception:
            xv = yv = None
        try:
            if state["colored_row"] is not None and state["colored_row"] != r:
                dialer_clear_dot_highlights(dial_sheet, state["colored_row"], cols)
        except Exception:
            pass
        f = cols["first_dot"]
        try:
            for i in range(3):
                dial_sheet.set_cell_data(r, f+i, "○")
            dial_sheet.set_cell_data(r, f + {"green":0,"gray":1,"red":2}[which], "●")
            dialer_colorize_outcome(dial_sheet, r, which, cols)
            dial_sheet.refresh()
        finally:
            try:
                if xv: dial_sheet.MT.xview_moveto(xv[0])
                if yv: dial_sheet.MT.yview_moveto(yv[0])
            except Exception:
                pass
        state["colored_row"] = r
        _update_confirm_button()

    def _apply_note_preview(r):
        txt = _current_note_text()
        c = state["note_col_by_row"].get(r)
        if c is None:
            c = dialer_next_empty_note_col(dial_sheet, r, cols)
            state["note_col_by_row"][r] = c
        if c is None:
            _update_confirm_button()
            return
        try:
            xv = dial_sheet.MT.xview(); yv = dial_sheet.MT.yview()
        except Exception:
            xv = yv = None
        try:
            dial_sheet.set_cell_data(r, c, txt)
            dial_sheet.refresh()
        finally:
            try:
                if xv: dial_sheet.MT.xview_moveto(xv[0])
                if yv: dial_sheet.MT.yview_moveto(yv[0])
            except Exception:
                pass
        _update_confirm_button()

    # --- persist dialer grid to CSV (full grid) ---
    def _save_dialer_grid_to_csv():
        try:
            data = dial_sheet.get_sheet_data() or []
        except Exception:
            data = []
        # Clip/pad to expected header length: HEADER_FIELDS + 3 + 8
        expected_len = len(HEADER_FIELDS) + 3 + 8
        matrix = []
        for row in data:
            r = (list(row) + [""] * expected_len)[:expected_len]
            # normalize empty outcome cells to "○"
            for i in range(len(HEADER_FIELDS), len(HEADER_FIELDS)+3):
                r[i] = r[i] if (r[i] or "").strip() else "○"
            matrix.append(r)
        save_dialer_leads_matrix(matrix)

    # prime the dialer selected row (first non-empty, else 0)
    try:
        if state["row"] is None:
            r0 = _row_selected(dial_sheet)
            if r0 is None:
                r0 = 0
            _set_working_row(r0)
    except Exception:
        pass

    # ============================================================
    # Warm tab state & helpers
    # ============================================================
    warm_state = {
        "row": None,
        "outcome": None,  # "green"|"gray"|"red"|None
    }
    warm_cols = warm_get_col_index_map()

    def _warm_selected_row():
        return _row_selected(warm_sheet)

    def _warm_note_text():
        return (window["-WARM_NOTE-"].get() or "").strip()

    def _warm_confirm_enabled():
        r = warm_state["row"]
        if r is None: return False
        if warm_state["outcome"] not in ("green","gray","red"): return False
        if not _warm_note_text(): return False
        # must have an empty Call slot available
        try:
            row_vals = warm_sheet.get_row_data(r) or []
        except Exception:
            row_vals = []
        return warm_next_empty_call_col(row_vals, warm_cols) is not None

    def _warm_update_confirm_button():
        ok = _warm_confirm_enabled()
        window["-WARM_CONFIRM-"].update(disabled=not ok, button_color=("white", "#2E7D32" if ok else "#444444"))

    def _warm_set_row(r):
        warm_state["row"] = r
        try:
            warm_sheet.set_currently_selected(r, 0)
            warm_sheet.see(r, 0)
        except Exception:
            pass
        _warm_update_confirm_button()

    def _warm_apply_outcome(which):
        warm_state["outcome"] = which
        _warm_update_confirm_button()

    def _warm_cost_normalize_in_row(r):
        """Format Cost ($) in-place for a given row."""
        ci = warm_cols["cost"]
        if ci is None: return
        try:
            row_vals = warm_sheet.get_row_data(r) or []
        except Exception:
            return
        val = row_vals[ci] if ci < len(row_vals) else ""
        newv = warm_format_cost(val)
        try:
            warm_sheet.set_cell_data(r, ci, newv)
            warm_sheet.refresh()
        except Exception:
            pass

    def _save_warm_grid_to_csv_v2():
        # Normalize all Cost ($) before save
        try:
            total = warm_sheet.get_total_rows()
        except Exception:
            total = 0
        for r in range(total):
            _warm_cost_normalize_in_row(r)
        matrix = warm_matrix_from_sheet_v2()
        save_warm_leads_matrix_v2(matrix)

    # ============================================================
    # Customers helpers (save / selection / add order / analytics)
    # ============================================================
    def _save_customers_grid_to_csv():
        try:
            matrix = customers_matrix_from_sheet()
            save_customers_matrix(matrix)
            window["-CUST_STATUS-"].update("Saved ✓")
        except Exception as e:
            window["-CUST_STATUS-"].update(f"Save error: {e}")

    def _customer_selected_row():
        return _row_selected(customer_sheet)

    def _cust_idx(name, default=None):
        try:
            return CUSTOMER_FIELDS.index(name)
        except Exception:
            return default

    def _popup_add_order(company):
        """Modal dialog to collect Amount + Date. Returns (amount_str, date_str) or None."""
        comp_disp = company or "(unknown)"
        layout = [
            [sg.Text(f"Add Order for: {comp_disp}", text_color="#9EE493")],
            [sg.Text("Amount ($):", size=(12,1)), sg.Input(key="-AO_AMOUNT-", size=(20,1))],
            [sg.Text("Order Date:", size=(12,1)), sg.Input(datetime.now().strftime("%Y-%m-%d"), key="-AO_DATE-", size=(20,1)),
             sg.Text(" (YYYY-MM-DD or MM/DD/YYYY)", text_color="#AAAAAA")],
            [sg.Push(), sg.Button("Cancel"), sg.Button("Add", button_color=("white","#2E7D32"))]
        ]
        win = sg.Window("Add Order", layout, modal=True, finalize=True)
        amount, date_s = None, None
        while True:
            ev, vals = win.read()
            if ev in (sg.WINDOW_CLOSE_ATTEMPTED_EVENT, sg.WIN_CLOSED, "Cancel"):
                win.close()
                return None
            if ev == "Add":
                amount = (vals.get("-AO_AMOUNT-","") or "").strip()
                date_s = (vals.get("-AO_DATE-","") or "").strip()
                if not amount:
                    sg.popup_error("Amount is required.")
                    continue
                win.close()
                return (amount, date_s)

    # ---- analytics helpers ----
    def _safe_update(key, text):
        try:
            if key in window.AllKeysDict:
                window[key].update(text)
        except Exception:
            pass

    def _orders_count_by_company():
        """Return {company: (count, total_amount)} from orders.csv."""
        counts = {}
        if ORDERS_PATH.exists():
            with ORDERS_PATH.open("r", encoding="utf-8", newline="") as f:
                rdr = csv.DictReader(f)
                for r in rdr:
                    comp = (r.get("Company","") or "").strip()
                    amt = r.get("Amount","") or ""
                    try:
                        val = float(str(amt).replace(",","").strip() or "0")
                    except Exception:
                        val = 0.0
                    if comp:
                        c, s = counts.get(comp, (0, 0.0))
                        counts[comp] = (c+1, s+val)
        return counts

    def refresh_customer_analytics():
        """
        Pipeline metrics + CAC + LTV + Reorder rate.
        - Warm leads: count rows in warm_leads.csv (v2) that look non-empty (Company or Email)
        - Samples sum: sum of 'Cost ($)' in warm v2
        - New customers: rows in customers.csv where 'Customer Since' is non-empty
        - Close rate: new_customers / warm_leads
        - CAC: samples_sum / max(1, new_customers)
        - Avg LTV: average 'CLTV' across all customers with a numeric value
        - Reorder rate: percent customers with Reorder? == 'Yes' (case-insensitive)
        - Also auto-mark Reorder? to 'Yes' if orders count for company >= 2
        """
        # Warm counts / samples
        warm_leads = 0
        samples_sum = 0.0
        try:
            if WARM_LEADS_PATH.exists():
                with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
                    rdr = csv.DictReader(f)
                    for r in rdr:
                        non_empty = (r.get("Company","") or r.get("Email","") or "").strip()
                        if non_empty:
                            warm_leads += 1
                        try:
                            samples_sum += float((r.get("Cost ($)","") or "0").replace(",","").strip() or "0")
                        except Exception:
                            pass
        except Exception:
            pass

        # Customers stats
        new_customers = 0
        total_customers = 0
        ltv_vals = []
        reorder_yes = 0

        orders_counts = _orders_count_by_company()

        # optionally mutate grid for Reorder? auto-update (>=2 orders)
        idx_reorder = _cust_idx("Reorder?")
        idx_company = _cust_idx("Company")

        try:
            rows = customer_sheet.get_sheet_data() or []
        except Exception:
            rows = []

        # mirror on-disk too for correctness:
        try:
            if CUSTOMERS_PATH.exists():
                with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
                    rdr = csv.DictReader(f)
                    disk_rows = list(rdr)
            else:
                disk_rows = []
        except Exception:
            disk_rows = []

        # Build a map to update disk if needed
        disk_changed = False

        for r_idx, row in enumerate(rows):
            # Compose row dict by headers
            rec = {CUSTOMER_FIELDS[i]: (row[i] if i < len(CUSTOMER_FIELDS) else "") for i in range(len(CUSTOMER_FIELDS))}
            comp = (rec.get("Company","") or "").strip()
            if not any((rec.get(h,"") or "").strip() for h in CUSTOMER_FIELDS):
                continue
            total_customers += 1
            if (rec.get("Customer Since","") or "").strip():
                new_customers += 1
            # LTV
            try:
                v = float((rec.get("CLTV","") or "").replace(",","").strip())
                if v > 0:
                    ltv_vals.append(v)
            except Exception:
                pass
            # Reorder? (explicit)
            if (rec.get("Reorder?","") or "").strip().lower() == "yes":
                is_yes_now = True
            else:
                is_yes_now = False

            # Auto YES if orders >= 2
            oc = orders_counts.get(comp, (0, 0.0))[0] if comp else 0
            if oc >= 2 and not is_yes_now:
                # mark in grid
                if idx_reorder is not None:
                    try:
                        customer_sheet.set_cell_data(r_idx, idx_reorder, "Yes")
                        is_yes_now = True
                    except Exception:
                        pass
                # also mark in disk mirror
                for drow in disk_rows:
                    if (drow.get("Company","") or "").strip() == comp:
                        if (drow.get("Reorder?","") or "").strip().lower() != "yes":
                            drow["Reorder?"] = "Yes"
                            disk_changed = True

            if is_yes_now:
                reorder_yes += 1

        # If we auto-updated "Reorder?" on-disk, write it
        if disk_changed:
            try:
                _backup(CUSTOMERS_PATH)
                _atomic_write_csv(CUSTOMERS_PATH, CUSTOMER_FIELDS, [[r.get(h,"") for h in CUSTOMER_FIELDS] for r in disk_rows])
            except Exception:
                pass

        # Compute metrics
        close_rate = (new_customers / warm_leads * 100.0) if warm_leads else 0.0
        cac = (samples_sum / new_customers) if new_customers else 0.0
        avg_ltv = (sum(ltv_vals) / len(ltv_vals)) if ltv_vals else 0.0
        reorder_rate = (reorder_yes / total_customers * 100.0) if total_customers else 0.0

        # Update labels safely (these keys must match what we placed in CHUNK 2)
        _safe_update("-AN_WARMS-", str(warm_leads))
        _safe_update("-AN_SAMPLES-", f"{samples_sum:.2f}")
        _safe_update("-AN_NEWCUS-", str(new_customers))
        _safe_update("-AN_CLOSERATE-", f"{close_rate:.1f}%")
        _safe_update("-AN_CAC-", f"{cac:.2f}")
        _safe_update("-AN_AVGLTV-", f"{avg_ltv:.2f}")
        _safe_update("-AN_REORDER-", f"{reorder_rate:.1f}%")

    # ============================================================
    # prime UI (Email Results stats + analytics)
    # ============================================================
    def refresh_fire_state():
        matrix = matrix_from_sheet()
        seen = load_state_set()
        new_count = 0
        for row in matrix:
            d = dict_from_row(row)
            if not valid_email(d.get("Email","")):
                continue
            fp = row_fingerprint_from_dict(d)
            if fp not in seen:
                new_count += 1
        if new_count > 0:
            window["-FIRE-"].update(disabled=False, button_color=("white","#C00000"))
            window["-FIRE_HINT-"].update(f" Ready: {new_count} new lead(s).")
        else:
            window["-FIRE-"].update(disabled=True, button_color=("white","#700000"))
            window["-FIRE_HINT-"].update(" (no NEW leads; already drafted or no valid emails)")

    def refresh_results_metrics():
        rows = load_results_rows_sorted()
        total_sent = sum(1 for r in rows if r.get("DateSent"))
        total_replied = sum(1 for r in rows if r.get("DateReplied"))
        warm = sum(1 for r in rows if (r.get("Status","").lower()=="green"))
        window["-WARM-"].update(str(warm))
        window["-REPLRATE-"].update(f"{total_replied} / {total_sent}")

    refresh_fire_state()
    refresh_results_metrics()
    _update_confirm_button()
    _warm_update_confirm_button()
    # first analytics draw
    try:
        refresh_customer_analytics()
    except Exception:
        pass

    # ============================================================
    # Event loop
    # ============================================================
    while True:
        event, values = window.read(timeout=300)
        if event in (sg.WINDOW_CLOSE_ATTEMPTED_EVENT, sg.WIN_CLOSED):
            break

        # Track grid selection changes
        r_now = _row_selected(dial_sheet)
        if r_now is not None and r_now != state["row"]:
            _set_working_row(r_now)
            state["note_col_by_row"].setdefault(r_now, None)
            _update_confirm_button()

        r_warm_now = _warm_selected_row()
        if r_warm_now is not None and r_warm_now != warm_state["row"]:
            _warm_set_row(r_warm_now)

        # ---------------- Email Leads tab buttons ----------------
        if event == "-OPENFOLDER-":
            try:
                os.startfile(str(APP_DIR))
            except Exception as e:
                popup_error(f"Open folder error: {e}")

        elif event == "-ADDROWS-":
            try:
                sheet.insert_rows(sheet.get_total_rows(), number_of_rows=10); sheet.refresh()
            except Exception:
                try: sheet.insert_rows(sheet.get_total_rows(), amount=10); sheet.refresh()
                except Exception as e: popup_error(f"Could not add rows: {e}")
            refresh_fire_state()

        elif event == "-DELROWS-":
            try:
                sels = sheet.get_selected_rows() or []
                if sels:
                    for r in sorted(sels, reverse=True):
                        try: sheet.delete_rows(r, 1)
                        except Exception:
                            try: sheet.delete_rows(r)
                            except Exception:
                                try: sheet.del_rows(r, 1)
                                except Exception:
                                    pass
                    sheet.refresh()
            except Exception as e:
                popup_error(f"Could not delete rows: {e}")
            refresh_fire_state()

        elif event == "-SAVECSV-":
            try:
                save_matrix_to_csv(matrix_from_sheet())
                window["-STATUS-"].update("Saved CSV")
            except Exception as e:
                window["-STATUS-"].update(f"Save error: {e}")

        # ---------------- Templates tab ----------------
        elif event == "-ADDNEW-":
            info = add_new_template_dialog()
            if info:
                try:
                    tpls, subs, mp = load_templates_ini()
                    name = info["name"]
                    tpls[name] = info["body"]
                    subs[name] = info["subject"]
                    for k in info["keys"]:
                        mp[k] = name
                    save_templates_ini(tpls, subs, mp)
                    window["-TPL_STATUS-"].update(f"Template '{name}' saved. Click Reload Templates to show it.")
                except Exception as e:
                    window["-TPL_STATUS-"].update(f"Save error: {e}")

        elif event == "-RELOADTPL-":
            window.close(); main(); return

        elif event == "-SAVETPL-":
            try:
                tpls_out, subs_out = {}, {}
                current_keys = set()
                for k in values:
                    if k.startswith("-TPL_"):
                        current_keys.add(k[5:-1] if k.endswith("-") else k[5:])
                for tkey in current_keys:
                    tpls_out[tkey] = values.get(f"-TPL_{tkey}-","")
                    subs_out[tkey] = values.get(f"-SUBJ_{tkey}-","") or DEFAULT_SUBJECT
                new_map = {}
                for line in (values.get("-MAP-","") or "").splitlines():
                    if "->" in line:
                        left,right = line.split("->",1); left,right = left.strip(), right.strip()
                        if left and right: new_map[left] = right
                save_templates_ini(tpls_out, subs_out, new_map)
                window["-TPL_STATUS-"].update("Templates & mapping saved ✓")
            except Exception as e:
                window["-TPL_STATUS-"].update(f"Save error: {e}")

        elif event == "-RESETTPL-":
            save_templates_ini(DEFAULT_TEMPLATES, DEFAULT_SUBJECTS, DEFAULT_MAP)
            window.close(); main(); return

        # ---------------- Email Results tab ----------------
        elif event == "-SYNC-":
            if not require_pywin32():
                window["-RS_STATUS-"].update("pywin32 missing (Outlook COM). Install pywin32.")
                continue
            try:
                days = int((values.get("-LOOKBACK-","") or "60").strip())
            except Exception:
                days = 60
            window["-RS_STATUS-"].update("Syncing…")
            try:
                s_count, r_count = outlook_sync_results(days)
                rows = load_results_rows_sorted()
                data = [[r.get("Ref",""), r.get("Email",""), r.get("Company",""), r.get("Industry",""),
                         r.get("DateSent",""), r.get("DateReplied",""), r.get("Status",""), r.get("Subject","")] for r in rows]
                window["-RSTABLE-"].update(values=data)
                refresh_results_metrics()
                window["-RS_STATUS-"].update(f"Synced: {s_count} sent refs; {r_count} replies.")
            except Exception as e:
                window["-RS_STATUS-"].update(f"Sync error: {e}")

        elif event == "-MARK_GREEN-":
            sels = values.get("-RSTABLE-", [])
            if sels:
                idx = sels[0]
                rows = load_results_rows_sorted()
                if 0 <= idx < len(rows):
                    set_status(rows[idx]["Ref"], "green")
                    try: add_warm_from_result(rows[idx], note="Marked Green on Email Results")
                    except Exception as e: print("Warm add error:", e)
                rows = load_results_rows_sorted()
                data = [[r.get("Ref",""), r.get("Email",""), r.get("Company",""), r.get("Industry",""),
                         r.get("DateSent",""), r.get("DateReplied",""), r.get("Status",""), r.get("Subject","")] for r in rows]
                window["-RSTABLE-"].update(values=data)
                refresh_results_metrics()
                try: refresh_customer_analytics()
                except Exception: pass

        elif event == "-MARK_GRAY-":
            sels = values.get("-RSTABLE-", [])
            if sels:
                idx = sels[0]
                rows = load_results_rows_sorted()
                if 0 <= idx < len(rows):
                    set_status(rows[idx]["Ref"], "gray")
                rows = load_results_rows_sorted()
                data = [[r.get("Ref",""), r.get("Email",""), r.get("Company",""), r.get("Industry",""),
                         r.get("DateSent",""), r.get("DateReplied",""), r.get("Status",""), r.get("Subject","")] for r in rows]
                window["-RSTABLE-"].update(values=data)
                refresh_results_metrics()

        elif event == "-MARK_RED-":
            sels = values.get("-RSTABLE-", [])
            if sels:
                idx = sels[0]
                rows = load_results_rows_sorted()
                if 0 <= idx < len(rows):
                    set_status(rows[idx]["Ref"], "red")
                    try: add_no_interest_from_result(rows[idx], note="Marked Red on Email Results", no_contact_flag=0)
                    except Exception as e: print("No-interest add error:", e)
                rows = load_results_rows_sorted()
                data = [[r.get("Ref",""), r.get("Email",""), r.get("Company",""), r.get("Industry",""),
                         r.get("DateSent",""), r.get("DateReplied",""), r.get("Status",""), r.get("Subject","")] for r in rows]
                window["-RSTABLE-"].update(values=data)
                refresh_results_metrics()

        # ---------------- Dialer tab: buttons ----------------
        elif event == "-DIAL_SET_GREEN-":
            if state["row"] is None:
                r = _row_selected(dial_sheet)
                if r is None:
                    window["-DIAL_MSG-"].update("Pick a row first.")
                    continue
                _set_working_row(r)
            _apply_outcome(state["row"], "green")

        elif event == "-DIAL_SET_GRAY-":
            if state["row"] is None:
                r = _row_selected(dial_sheet)
                if r is None:
                    window["-DIAL_MSG-"].update("Pick a row first.")
                    continue
                _set_working_row(r)
            _apply_outcome(state["row"], "gray")

        elif event == "-DIAL_SET_RED-":
            if state["row"] is None:
                r = _row_selected(dial_sheet)
                if r is None:
                    window["-DIAL_MSG-"].update("Pick a row first.")
                    continue
                _set_working_row(r)
            _apply_outcome(state["row"], "red")

        elif event == "-DIAL_NOTE-":
            if state["row"] is None:
                r = _row_selected(dial_sheet)
                if r is None:
                    continue
                _set_working_row(r)
            _apply_note_preview(state["row"])

        elif event == "-DIAL_CONFIRM-":
            if state["row"] is None:
                window["-DIAL_MSG-"].update("Pick a row first.")
            else:
                r = state["row"]
                note_text = _current_note_text()
                if not note_text:
                    window["-DIAL_MSG-"].update("Type a note.")
                else:
                    outcome = state["outcome"] or "gray"
                    try:
                        row_vals = dial_sheet.get_row_data(r) or []
                        base = dict_from_row([row_vals[i] if i < len(row_vals) else "" for i in range(len(HEADER_FIELDS))])

                        c = state["note_col_by_row"].get(r)
                        if c is None:
                            c = dialer_next_empty_note_col(dial_sheet, r, cols)
                        if c is not None:
                            try:
                                xv = dial_sheet.MT.xview(); yv = dial_sheet.MT.yview()
                            except Exception:
                                xv = yv = None
                            try:
                                dial_sheet.set_cell_data(r, c, note_text)
                                dial_sheet.refresh()
                            finally:
                                try:
                                    if xv: dial_sheet.MT.xview_moveto(xv[0])
                                    if yv: dial_sheet.MT.yview_moveto(yv[0])
                                except Exception:
                                    pass

                        # Persist CSV logs
                        dialer_save_call(base, outcome, note_text)
                        if outcome == "red":
                            add_no_interest(base, note_text, no_contact_flag=0, source="Dialer")
                        elif outcome == "gray":
                            # Count filled notes for no-contact rule
                            filled = 0
                            row_vals2 = dial_sheet.get_row_data(r) or []
                            for k in range(cols["first_note"], cols["last_note"]+1):
                                if k < len(row_vals2) and (row_vals2[k] or "").strip():
                                    filled += 1
                            if filled >= 8:
                                add_no_interest(base, "No Contact after 8 calls. " + note_text, no_contact_flag=1, source="Dialer")

                        window["-DIAL_MSG-"].update("Saved ✓")
                        window["-DIAL_NOTE-"].update("")
                        state["outcome"] = None
                        state["note_col_by_row"].pop(r, None)

                        # If green or red → remove row and persist dialer grid to CSV
                        if outcome in ("green", "red"):
                            # delete row in-grid
                            try:
                                dial_sheet.delete_rows(r, 1)
                            except Exception:
                                try:
                                    dial_sheet.delete_rows(r)
                                except Exception:
                                    try:
                                        dial_sheet.del_rows(r, 1)
                                    except Exception:
                                        pass
                            try:
                                dial_sheet.refresh()
                            except Exception:
                                pass
                            # Save updated grid to CSV
                            _save_dialer_grid_to_csv()

                            # move selection to the next row (same index now points to next)
                            try:
                                total = dial_sheet.get_total_rows()
                            except Exception:
                                total = 0
                            if total <= 0:
                                state["row"] = None
                            else:
                                new_idx = min(r, max(0, total - 1))
                                _set_working_row(new_idx)
                        else:
                            # gray: stay and move to next row
                            _save_dialer_grid_to_csv()
                            new_row = dialer_move_to_next_row(dial_sheet, r)
                            _set_working_row(new_row)

                        _update_confirm_button()

                    except Exception as e:
                        window["-DIAL_MSG-"].update(f"Save error: {e}")

        elif event == "-DIAL_ADD100-":
            try:
                # Append 100 blank dialer rows: left headers empty + three "○" + 8 blanks
                add = [[""] * len(HEADER_FIELDS) + ["○","○","○"] + ([""]*8) for _ in range(100)]
                try:
                    cur = dial_sheet.get_sheet_data() or []
                except Exception:
                    cur = []
                dial_sheet.set_sheet_data((cur or []) + add)
                dial_sheet.refresh()
                _save_dialer_grid_to_csv()
            except Exception as e:
                window["-DIAL_MSG-"].update(f"Add rows error: {e}")

        # ---------------- Warm tab: buttons ----------------
        elif event == "-WARM_SET_GREEN-":
            if warm_state["row"] is None:
                r = _warm_selected_row()
                if r is None:
                    window["-WARM_STATUS-"].update("Pick a warm row first.")
                    continue
                _warm_set_row(r)
            _warm_apply_outcome("green")

        elif event == "-WARM_SET_GRAY-":
            if warm_state["row"] is None:
                r = _warm_selected_row()
                if r is None:
                    window["-WARM_STATUS-"].update("Pick a warm row first.")
                    continue
                _warm_set_row(r)
            _warm_apply_outcome("gray")

        elif event == "-WARM_SET_RED-":
            if warm_state["row"] is None:
                r = _warm_selected_row()
                if r is None:
                    window["-WARM_STATUS-"].update("Pick a warm row first.")
                    continue
                _warm_set_row(r)
            _warm_apply_outcome("red")

        elif event == "-WARM_NOTE-":
            _warm_update_confirm_button()

        elif event == "-WARM_CONFIRM-":
            r = warm_state["row"]
            if r is None:
                window["-WARM_STATUS-"].update("Pick a warm row first.")
                continue
            if warm_state["outcome"] not in ("green","gray","red"):
                window["-WARM_STATUS-"].update("Choose an outcome first.")
                continue
            note_text = _warm_note_text()
            if not note_text:
                window["-WARM_STATUS-"].update("Type a note.")
                continue

            # Prepare row values
            try:
                row_vals = warm_sheet.get_row_data(r) or []
            except Exception:
                row_vals = []
            # Normalize cost
            ci = warm_cols["cost"]
            if ci is not None:
                try:
                    cur_cost = row_vals[ci] if ci < len(row_vals) else ""
                except Exception:
                    cur_cost = ""
                new_cost = warm_format_cost(cur_cost)
                try:
                    warm_sheet.set_cell_data(r, ci, new_cost)
                except Exception:
                    pass

            # Stamp note into next empty Call N column
            col_call = warm_next_empty_call_col(row_vals, warm_cols)
            if col_call is None:
                window["-WARM_STATUS-"].update("All 15 call slots are filled.")
                continue
            stamp = datetime.now().strftime("%m-%d")
            stamped_note = f"{stamp}: {note_text}"
            try:
                warm_sheet.set_cell_data(r, col_call, stamped_note)
            except Exception:
                pass

            # Update the Timestamp column to today (MM-DD)
            ti = warm_cols["timestamp"]
            if ti is not None:
                try:
                    warm_sheet.set_cell_data(r, ti, stamp)
                except Exception:
                    pass

            # Persist to CSV
            try:
                warm_sheet.refresh()
            except Exception:
                pass
            _save_warm_grid_to_csv_v2()

            # Additional routing based on outcome
            outcome = warm_state["outcome"]
            # Build a minimal dict for no-interest if needed
            # We'll map by known WARM_V2_FIELDS names where possible
            wmap = {WARM_V2_FIELDS[i]: (row_vals[i] if i < len(WARM_V2_FIELDS) else "") for i in range(len(WARM_V2_FIELDS))}
            base = {
                "Email": wmap.get("Email",""),
                "First Name": (wmap.get("Prospect Name","") or "").split(" ")[0] if wmap.get("Prospect Name") else "",
                "Last Name": " ".join((wmap.get("Prospect Name","") or "").split(" ")[1:]) if wmap.get("Prospect Name") else "",
                "Company": wmap.get("Company",""),
                "Industry": wmap.get("Industry",""),
                "Phone": wmap.get("Phone #",""),
                "City": (wmap.get("Location","") or "").split(",")[0] if wmap.get("Location") else "",
                "State": (wmap.get("Location","") or "").split(",")[-1].strip() if wmap.get("Location") and "," in wmap.get("Location") else "",
                "Website": "",
            }

            if outcome == "red":
                try:
                    add_no_interest(base, stamped_note, no_contact_flag=0, source="Warm")
                except Exception as e:
                    print("Warm->NoInterest error:", e)

            window["-WARM_STATUS-"].update("Saved ✓")
            window["-WARM_NOTE-"].update("")
            warm_state["outcome"] = None
            _warm_update_confirm_button()

        elif event == "-WARM_ADD100-":
            try:
                add = [[""] * len(WARM_V2_FIELDS) for _ in range(100)]
                cur = warm_sheet.get_sheet_data() or []
                warm_sheet.set_sheet_data((cur or []) + add)
                warm_sheet.refresh()
                _save_warm_grid_to_csv_v2()
            except Exception as e:
                window["-WARM_STATUS-"].update(f"Add rows error: {e}")

        elif event == "-WARM_SAVE-":
            try:
                _save_warm_grid_to_csv_v2()
                window["-WARM_STATUS-"].update("Saved ✓")
                try: refresh_customer_analytics()
                except Exception: pass
            except Exception as e:
                window["-WARM_STATUS-"].update(f"Save error: {e}")

        elif event == "-WARM_EXPORT-":
            path = sg.popup_get_file("Save warm_leads.csv", save_as=True, default_extension=".csv",
                                     file_types=(("CSV","*.csv"),), no_window=True)
            if path:
                try:
                    _save_warm_grid_to_csv_v2()
                    # Copy file to path chosen
                    with WARM_LEADS_PATH.open("rb") as s, open(path, "wb") as d:
                        d.write(s.read())
                    window["-WARM_STATUS-"].update("Exported ✓")
                except Exception as e:
                    window["-WARM_STATUS-"].update(f"Export error: {e}")

        elif event == "-WARM_RELOAD-":
            try:
                rows = load_warm_leads_matrix_v2()
                if len(rows) < 100:
                    rows += [[""] * len(WARM_V2_FIELDS) for _ in range(100 - len(rows))]
                warm_sheet.set_sheet_data(rows)
                warm_sheet.refresh()
                window["-WARM_STATUS-"].update("Reloaded ✓")
                try: refresh_customer_analytics()
                except Exception: pass
            except Exception as e:
                window["-WARM_STATUS-"].update(f"Reload error: {e}")

        elif event == "-WARM_MARK_CUSTOMER-":
            try:
                sel_rows = warm_sheet.get_selected_rows() or []
            except Exception:
                sel_rows = []
            if not sel_rows:
                window["-WARM_STATUS-"].update("Pick a row in the Warm grid first.")
                continue
            r_idx = sel_rows[0]
            row = warm_sheet.get_row_data(r_idx) or []
            warm_row = { WARM_V2_FIELDS[i]: (row[i] if i < len(WARM_V2_FIELDS) else "") for i in range(len(WARM_V2_FIELDS)) }
            yn = sg.popup_yes_no("Mark this Warm Lead as a NEW CUSTOMER?\n\nYou’ll be asked for the Opening Order $ next.")
            if yn != "Yes":
                window["-WARM_STATUS-"].update("Canceled.")
                continue
            amt = sg.popup_get_text("Opening Order $ (numbers only, e.g. 1500 or 1500.00):", default_text="")
            if amt is None:
                window["-WARM_STATUS-"].update("Canceled.")
                continue
            amt = (amt or "").strip().replace(",", "")
            try:
                float_amt = float(amt)
                amt = f"{float_amt:.2f}"
            except Exception:
                window["-WARM_STATUS-"].update("Invalid amount. Use numbers only, e.g. 1500 or 1500.00")
                continue

            cust = {h:"" for h in CUSTOMER_FIELDS}
            cust["Company"] = warm_row.get("Company","")
            cust["Prospect Name"] = warm_row.get("Prospect Name","")
            cust["Phone #"] = warm_row.get("Phone #","")
            cust["Email"] = warm_row.get("Email","")
            cust["Location"] = warm_row.get("Location","")
            cust["Industry"] = warm_row.get("Industry","")
            cust["Google Reviews"] = warm_row.get("Google Reviews","")
            cust["Rep"] = warm_row.get("Rep","")
            cust["Samples?"] = warm_row.get("Samples?","")
            cust["Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if "Opening Order $" in CUSTOMER_FIELDS:
                cust["Opening Order $"] = amt
            if "Customer Since" in CUSTOMER_FIELDS:
                cust["Customer Since"] = datetime.now().strftime("%Y-%m-%d")
            if "Notes" in CUSTOMER_FIELDS:
                # Grab latest filled Call note as default notes
                last_note = ""
                for i in range(15, 0, -1):
                    v = warm_row.get(f"Call {i}", "")
                    if (v or "").strip():
                        last_note = v
                        break
                cust["Notes"] = last_note

            try:
                existing = []
                if CUSTOMERS_PATH.exists():
                    with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
                        rdr = csv.DictReader(f)
                        for r in rdr:
                            existing.append([r.get(h,"") for h in CUSTOMER_FIELDS])
                existing.append([cust.get(h,"") for h in CUSTOMER_FIELDS])
                _backup(CUSTOMERS_PATH)
                _atomic_write_csv(CUSTOMERS_PATH, CUSTOMER_FIELDS, existing)
                # Update grid & persist warm
                try:
                    warm_sheet.delete_rows(r_idx, 1)
                except Exception:
                    try:
                        warm_sheet.delete_rows(r_idx)
                    except Exception:
                        try:
                            warm_sheet.del_rows(r_idx, 1)
                        except Exception:
                            pass
                warm_sheet.refresh()
                _save_warm_grid_to_csv_v2()
                window["-WARM_STATUS-"].update("Promoted to Customer ✓")
                try: refresh_customer_analytics()
                except Exception: pass
            except Exception as e:
                window["-WARM_STATUS-"].update(f"Move error (customers): {e}")

        # ---------------- Customers tab: buttons ----------------
        elif event == "-CUST_ADD50-":
            try:
                add = [[""] * len(CUSTOMER_FIELDS) for _ in range(50)]
                cur = customer_sheet.get_sheet_data() or []
                customer_sheet.set_sheet_data((cur or []) + add)
                customer_sheet.refresh()
                _save_customers_grid_to_csv()
                try: refresh_customer_analytics()
                except Exception: pass
            except Exception as e:
                window["-CUST_STATUS-"].update(f"Add rows error: {e}")

        elif event == "-CUST_ADD_ORDER-":
            try:
                r_sel = _customer_selected_row()
            except Exception:
                r_sel = None
            if r_sel is None:
                window["-CUST_STATUS-"].update("Pick a customer row first.")
                continue

            # Pull company from selected row
            try:
                row_vals = customer_sheet.get_row_data(r_sel) or []
            except Exception:
                row_vals = []
            idx_company = _cust_idx("Company", 0)
            company = row_vals[idx_company] if idx_company is not None and idx_company < len(row_vals) else ""
            if not (company or "").strip():
                window["-CUST_STATUS-"].update("Company is required on the selected row.")
                continue

            # Pop dialog to collect order details
            res = _popup_add_order(company)
            if not res:
                window["-CUST_STATUS-"].update("Canceled.")
                continue
            amount_s, date_s = res

            try:
                append_order_row(company, date_s, amount_s)
            except Exception as e:
                window["-CUST_STATUS-"].update(f"Add order error: {e}")
                continue

            # Compute stats and reflect them in-grid
            stats = compute_customer_order_stats(company)
            # Column indices
            idx_fod  = _cust_idx("First Order Date")
            idx_lod  = _cust_idx("Last Order Date")
            idx_cltv = _cust_idx("CLTV")
            idx_days = _cust_idx("Days")
            idx_spd  = _cust_idx("Sales/Day")

            def _set_if(idx, val):
                if idx is None: return
                try:
                    customer_sheet.set_cell_data(r_sel, idx, val)
                except Exception:
                    pass

            _set_if(idx_fod, stats["first_order_date"].strftime("%Y-%m-%d") if stats["first_order_date"] else "")
            _set_if(idx_lod, stats["last_order_date"].strftime("%Y-%m-%d") if stats["last_order_date"] else "")
            _set_if(idx_cltv, f"{float(stats['cltv']):.2f}")
            _set_if(idx_days, str(stats["days_since_first"]) if stats["days_since_first"] is not None else "")
            _set_if(idx_spd,  f"{float(stats['sales_per_day']):.2f}" if stats["sales_per_day"] is not None else "")

            try:
                customer_sheet.refresh()
            except Exception:
                pass

            # Persist the whole customers grid (the append_order_row already updated customers.csv too)
            _save_customers_grid_to_csv()
            window["-CUST_STATUS-"].update("Order added ✓")

            # Refresh analytics (CAC/LTV/Reorder may change)
            try:
                refresh_customer_analytics()
            except Exception:
                pass

        elif event == "-CUST_SAVE-":
            _save_customers_grid_to_csv()
            try: refresh_customer_analytics()
            except Exception: pass

        elif event == "-CUST_EXPORT-":
            path = sg.popup_get_file("Save customers.csv", save_as=True, default_extension=".csv",
                                     file_types=(("CSV","*.csv"),), no_window=True)
            if path:
                try:
                    # Ensure on-disk content matches the grid first
                    _save_customers_grid_to_csv()
                    with CUSTOMERS_PATH.open("rb") as s, open(path, "wb") as d:
                        d.write(s.read())
                    window["-CUST_STATUS-"].update("Exported ✓")
                except Exception as e:
                    window["-CUST_STATUS-"].update(f"Export error: {e}")

        elif event == "-CUST_RELOAD-":
            try:
                rows = load_customers_matrix()
                if len(rows) < 50:
                    rows += [[""] * len(CUSTOMER_FIELDS) for _ in range(50 - len(rows))]
                customer_sheet.set_sheet_data(rows)
                customer_sheet.refresh()
                window["-CUST_STATUS-"].update("Reloaded ✓")
                try: refresh_customer_analytics()
                except Exception: pass
            except Exception as e:
                window["-CUST_STATUS-"].update(f"Reload error: {e}")

        # Keep analytics and fire button fresh
        refresh_fire_state()
        try:
            refresh_customer_analytics()
        except Exception:
            pass

    window.close()


# ============================================================
# Entry
# ============================================================
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        try:
            popup_error(f"Fatal error starting app: {e}")
        except Exception:
            import traceback
            traceback.print_exc()
        input("\n\n[ERROR] Press Enter to exit...")
# ===== CHUNK 4 / 4 — END =====
