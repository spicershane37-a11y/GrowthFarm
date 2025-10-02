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

# Dialer / warm / no-interest / customers files
DIALER_RESULTS_PATH = APP_DIR / "dialer_results.csv"
WARM_LEADS_PATH     = APP_DIR / "warm_leads.csv"
NO_INTEREST_PATH    = APP_DIR / "no_interest.csv"
CUSTOMERS_PATH      = APP_DIR / "customers.csv"

# NEW: dedicated Dialer leads store (separate from kybercrystals.csv)
DIALER_LEADS_PATH   = APP_DIR / "dialer_leads.csv"

# -------------------- Outlook settings ----------------------
DEATHSTAR_SUBFOLDER = "Order 66"   # Drafts subfolder
TARGET_MAILBOX_HINT = ""           # Optional: part of SMTP/display name to target a specific account

# -------------------- Columns / headers ---------------------
HEADER_FIELDS = [
    "Email","First Name","Last Name","Company","Industry","Phone",
    "Address","City","State","Reviews","Website"
]

# Warm Leads sheet headers (grid) — WITHOUT Engagement/Potential/Timing/Total
WARM_FIELDS = [
    "Company","Prospect Name","Phone #","Email",
    "Location","Industry","Google Reviews","Rep","Samples?","Timestamp",
    "Call 1 Notes","Call 2 Date","Call 2 Notes","Call 3 Date","Call 3 Notes",
    "Call 4 Date","Call 4 Notes","Call 5 Date","Call 5 Notes","Call 6 Date","Call 6 Notes",
    "Call 7 Dates","Call 7 Notes","Call 8 Date","Call 8 Notes","Call 9 Date","Call 9 Notes",
    "Call 10 Date","Call 10 Notes","Call 11 Date","Call 11 Notes","Call 12 Date","Call 12 Notes",
    "Call 13 Date","Call 13 Notes","Call 14 Date","Call 14 Notes"
]

# Customers sheet headers
CUSTOMER_FIELDS = [
    "Company","Prospect Name","Phone #","Email",
    "Location","Industry","Google Reviews","Rep","Samples?","Timestamp",
    "Opening Order $","Customer Since","Notes"
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
    """(legacy) Start from the same CSV backing Email Leads; not used by the new Dialer."""
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
    leads_buttons = [
        sg.Button("Open Folder", key="-OPENFOLDER-"),
        sg.Button("Add 10 Rows", key="-ADDROWS-"),
        sg.Button("Delete Selected Rows", key="-DELROWS-"),
        sg.Button("Save Now", key="-SAVECSV-"),
        sg.Text("Status:", text_color="#A0A0A0"), sg.Text("Idle", key="-STATUS-", text_color="#FFFFFF"),
    ]
    fire_row = [
        sg.Button("Fire the Death Star", key="-FIRE-", size=(25,2), disabled=True, button_color=("white","#700000")),
        sg.Text(" (disabled: add valid NEW leads)", key="-FIRE_HINT-", text_color="#BBBBBB")
    ]
    leads_tab = [
        [leads_host],
        [sg.Text("Columns / placeholders:", text_color="#CCCCCC")],
        [sg.Text(", ".join(HEADER_FIELDS), text_color="#9EE493", font=("Consolas", 9))],
        leads_buttons,
        [sg.HorizontalSeparator(color="#4CAF50")],
        fire_row
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

    DIALER_EXTRA_COLS = ["🙂","😐","☹️",
                         "Note1","Note2","Note3","Note4","Note5","Note6","Note7","Note8"]  # emoji headers
    DIALER_HEADERS = HEADER_FIELDS + DIALER_EXTRA_COLS

    # NEW: load from the Dialer’s own CSV (created if missing) — function defined in Part 2
    dialer_matrix = []
    try:
        dialer_matrix = load_dialer_leads_matrix()
    except Exception:
        # if Part 2 isn't loaded yet, fall back to an empty grid; Part 2 will manage persistence
        dialer_matrix = []

    if len(dialer_matrix) < 100:
        dialer_matrix += [[""] * len(HEADER_FIELDS) + ["○","○","○"] + ([""] * 8)
                          for _ in range(100 - len(dialer_matrix))]

    dialer_host = sg.Frame(
        "DIALER GRID",
        [[sg.Text("Loading dialer grid…", key="-DIAL_LOAD-", text_color="#9EE493")]],
        relief=sg.RELIEF_GROOVE, border_width=2, background_color="#1B1B1B",
        title_color="#9EE493", expand_x=True, expand_y=True, key="-DIAL_HOST-",
    )

    dialer_controls = [
        [sg.Text("Outcome:", text_color="#CCCCCC")],
        [sg.Button("🟢 Green", key="-DIAL_SET_GREEN-", button_color=("white","#2E7D32"), size=(12,1))],
        [sg.Button("⚪ Gray",  key="-DIAL_SET_GRAY-",  button_color=("black","#DDDDDD"), size=(12,1))],
        [sg.Button("🔴 Red",   key="-DIAL_SET_RED-",   button_color=("white","#C62828"), size=(12,1))],
        [sg.Text("Note (goes into next empty Note1–Note8):", text_color="#CCCCCC", pad=((0,0),(10,0)))],
        [sg.Multiline(key="-DIAL_NOTE-", size=(28,6), font=("Consolas",10), background_color="#111", text_color="#EEE", enable_events=True)],
        [sg.Button("Confirm Call", key="-DIAL_CONFIRM-", size=(14,2), disabled=True, button_color=("white","#444444"))],
        [sg.Button("＋ Add 100 Rows", key="-DIAL_ADD100-", size=(14,1))],
        [sg.Text("", key="-DIAL_MSG-", text_color="#A0FFA0", size=(28,2))]
    ]

    dialer_tab = [
        [sg.Column([[dialer_host]], expand_x=True, expand_y=True),
         sg.Column(dialer_controls, vertical_alignment="top", pad=((10,0),(0,0)))]
    ]

    # -------- Warm Leads tab (tksheet grid) --------
    ensure_warm_file()

    warm_host = sg.Frame(
        "WARM LEADS GRID (paste from Google Sheets / Excel)",
        [[sg.Text("Loading warm grid…", key="-WARM_LOAD-", text_color="#9EE493")]],
        relief=sg.RELIEF_GROOVE, border_width=2, background_color="#1B1B1B",
        title_color="#9EE493", expand_x=True, expand_y=True, key="-WARM_HOST-",
    )

    warm_controls = [
        sg.Button("Save Warm Leads", key="-WARM_SAVE-"),
        sg.Button("Export Warm Leads CSV", key="-WARM_EXPORT-"),
        sg.Button("Reload Warm", key="-WARM_RELOAD-"),
        sg.Button("→ Confirm New Customer", key="-WARM_MARK_CUSTOMER-", button_color=("white","#2E7D32")),
        sg.Text("", key="-WARM_STATUS-", text_color="#A0FFA0")
    ]

    warm_tab = [
        [warm_host],
        warm_controls
    ]

    # -------- Customers tab (tksheet grid) --------
    if not CUSTOMERS_PATH.exists():
        with CUSTOMERS_PATH.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(CUSTOMER_FIELDS)

    customers_host = sg.Frame(
        "CUSTOMERS GRID",
        [[sg.Text("Loading customers grid…", key="-CUST_LOAD-", text_color="#9EE493")]],
        relief=sg.RELIEF_GROOVE, border_width=2, background_color="#1B1B1B",
        title_color="#9EE493", expand_x=True, expand_y=True, key="-CUST_HOST-",
    )

    customers_controls = [
        sg.Button("Save Customers", key="-CUST_SAVE-"),
        sg.Button("Export Customers CSV", key="-CUST_EXPORT-"),
        sg.Button("Reload Customers", key="-CUST_RELOAD-"),
        sg.Text("", key="-CUST_STATUS-", text_color="#A0FFA0")
    ]

    customers_tab = [
        [customers_host],
        customers_controls
    ]

    # -------- Compose full layout (Analytics removed) --------
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
            # selection colors (row selection)
            row_selected_background="#FFF59D",  # light highlighter yellow
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
    first_dot = len(HEADER_FIELDS)         # 🙂 column index
    first_note = len(HEADER_FIELDS) + 3    # Note1
    last_note  = first_note + 7            # Note8

    for c in range(len(DIALER_HEADERS)):
        width = DEFAULT_COL_WIDTH
        if c == idx_address: width = 120
        if c == idx_city:    width = 90
        if c == idx_state:   width = 42
        if c == idx_reviews: width = 60
        if c == idx_website: width = 160
        if first_dot <= c < first_note:  # three dot columns
            width = 40
        if first_note <= c <= last_note: # notes
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
    if WARM_LEADS_PATH.exists():
        try:
            with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
                rdr = csv.DictReader(f)
                for r in rdr:
                    warm_matrix.append([r.get(h, "") for h in WARM_FIELDS])
        except Exception:
            warm_matrix = []
    if len(warm_matrix) < 100:
        warm_matrix += [[""] * len(WARM_FIELDS) for _ in range(100 - len(warm_matrix))]

    try:
        from tksheet import Sheet as WarmSheet
    except Exception:
        popup_error("tksheet not installed. Run: pip install tksheet")
        return

    warm_sheet = WarmSheet(
        warm_holder,
        data=warm_matrix,
        headers=WARM_FIELDS,
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

    for c, name in enumerate(WARM_FIELDS):
        width = 120
        if name in ("Company","Prospect Name"): width = 180
        if name in ("Phone #","Rep","Samples?"): width = 90
        if name in ("Email","Google Reviews","Industry","Location"): width = 160
        if name.endswith("Date"): width = 110
        if name == "Timestamp": width = 150
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


# ---- Safe write helpers for Warm Leads / Customers ----
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
    since = (datetime.now() - timedelta(days=lookback_days)).strftime("%m/%d/%Y %H:%M %p")
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
# ===================== Dialer helper shims (rescue) =====================
# These helpers are required by the Dialer logic used in Part 2.

if "dialer_cols_info" not in globals():
    def dialer_cols_info(_headers=None):
        first_dot  = len(HEADER_FIELDS)          # outcome dots start
        last_dot   = first_dot + 2
        first_note = len(HEADER_FIELDS) + 3      # Note1
        last_note  = first_note + 7              # Note8
        return {"first_dot": first_dot, "last_dot": last_dot,
                "first_note": first_note, "last_note": last_note}

if "dialer_clear_dot_highlights" not in globals():
    def dialer_clear_dot_highlights(sheet, row, cols):
        try:
            for c in range(cols["first_dot"], cols["last_dot"]+1):
                sheet.dehighlight_cells(row, c, 1, 1)
        except Exception:
            try:
                for c in range(cols["first_dot"], cols["last_dot"]+1):
                    sheet.highlight_cells(row=row, column=c, bg=None, fg=None)
            except Exception:
                pass

if "dialer_colorize_outcome" not in globals():
    _DOT_BG = {"green": "#2E7D32", "gray": "#9E9E9E", "red": "#C62828"}
    _DOT_FG = {"green": "#FFFFFF", "gray": "#000000", "red": "#FFFFFF"}
    def dialer_colorize_outcome(sheet, row, outcome, cols=None):
        if cols is None: cols = dialer_cols_info(None)
        base = cols["first_dot"]
        try:
            for i in range(3):
                sheet.set_cell_data(row, base+i, "○")
        except Exception:
            pass
        which_idx = {"green":0, "gray":1, "red":2}.get(outcome)
        if which_idx is None: return
        c = base + which_idx
        try:
            sheet.set_cell_data(row, c, "●")
            try:
                sheet.highlight_cells(row=row, column=c, bg=_DOT_BG[outcome], fg=_DOT_FG[outcome])
            except Exception:
                pass
            sheet.refresh()
        except Exception:
            pass

if "dialer_next_empty_note_col" not in globals():
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

if "dialer_preview_note" not in globals():
    def dialer_preview_note(sheet, row, text, cols=None):
        if cols is None: cols = dialer_cols_info(None)
        c = dialer_next_empty_note_col(sheet, row, cols)
        if c is None: return None
        try:
            sheet.set_cell_data(row, c, text)
            sheet.refresh()
        except Exception:
            pass
        return c

if "dialer_move_to_next_row" not in globals():
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

if "dialer_mark_working_row" not in globals():
    def dialer_mark_working_row(sheet, row):
        # default; not used after we added custom painter
        try:
            sheet.dehighlight_rows()
        except Exception:
            pass
        try:
            sheet.highlight_rows(rows=[row], bg="#FFF59D", fg="#000000")  # highlighter yellow
            sheet.refresh()
        except Exception:
            pass
# =================== end Dialer helper shims (rescue) ===================

# Persist in-progress dialer grid here
DIALER_STATE_PATH = APP_DIR / "dialer_grid.csv"

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

    def warm_matrix_from_sheet():
        raw = warm_sheet.get_sheet_data() or []
        trimmed = []
        for row in raw:
            row = (list(row) + [""] * len(WARM_FIELDS))[:len(WARM_FIELDS)]
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

    # Save full dialer grid to disk (including dots/notes)
    def _save_dialer_grid():
        try:
            data = dial_sheet.get_sheet_data() or []
            with DIALER_STATE_PATH.open("w", encoding="utf-8", newline="") as f:
                w = csv.writer(f)
                hdrs = HEADER_FIELDS + ["🙂","😐","🙁","Note1","Note2","Note3","Note4","Note5","Note6","Note7","Note8"]
                w.writerow(hdrs)
                for row in data:
                    w.writerow((list(row) + [""] * len(hdrs))[:len(hdrs)])
        except Exception as e:
            print("Persist dialer_grid.csv failed:", e)

    # Add 100 rows and seed ○ in dot cols
    def _append_100_rows():
        try:
            total = dial_sheet.get_total_rows()
        except Exception:
            total = 0
        ok=False
        try: dial_sheet.insert_rows(total, 100); ok=True
        except Exception: pass
        if not ok:
            try: dial_sheet.insert_rows(total, number_of_rows=100); ok=True
            except Exception: pass
        if not ok:
            try: dial_sheet.insert_rows(total, amount=100); ok=True
            except Exception: pass
        try:
            new_total = dial_sheet.get_total_rows()
        except Exception:
            new_total = total + 100
        first_dot = len(HEADER_FIELDS)
        for r in range(total, new_total):
            for i in range(3):
                try: dial_sheet.set_cell_data(r, first_dot+i, "○")
                except Exception: pass
        try: dial_sheet.refresh()
        except Exception: pass
        _save_dialer_grid()

    # ----- dialer helpers / state -----
    cols = dialer_cols_info(dial_sheet.headers() if hasattr(dial_sheet, "headers") else None)
    state = {
        "row": None,
        "outcome": None,          # "green" | "gray" | "red" | None
        "note_col_by_row": {},    # sticky preview slot: {row_idx: col_idx}
        "done_rows": set(),       # rows confirmed as gray (still working)
        "colored_row": None,      # row that currently has colored outcome dot
    }

    def _paint_highlights():
        """Re-apply row highlights: dark for 'done', yellow for working row."""
        try:
            dial_sheet.dehighlight_rows()
        except Exception:
            pass
        # Dark gray for confirmed-but-still-working rows
        try:
            if state["done_rows"]:
                dial_sheet.highlight_rows(rows=sorted(state["done_rows"]), bg="#424242", fg="#FFFFFF")
        except Exception:
            pass
        # Yellow for working row
        if state["row"] is not None:
            try:
                dial_sheet.highlight_rows(rows=[state["row"]], bg="#FFF59D", fg="#000000")
            except Exception:
                pass
        try:
            dial_sheet.refresh()
        except Exception:
            pass

    # Try to set 🙂 😐 🙁 headers at runtime (safety net) and center them
    try:
        hdrs = None
        try:
            hdrs = list(dial_sheet.headers())
        except Exception:
            pass
        if not hdrs:
            try:
                hdrs = list(dial_sheet.get_headers())
            except Exception:
                pass
        if hdrs:
            base = len(HEADER_FIELDS)
            if len(hdrs) >= base + 3:
                hdrs[base + 0] = "🙂"
                hdrs[base + 1] = "😐"
                hdrs[base + 2] = "🙁"
                ok = False
                try:
                    dial_sheet.headers(hdrs); ok = True
                except Exception:
                    pass
                if not ok:
                    try:
                        dial_sheet.set_headers(hdrs); ok = True
                    except Exception:
                        pass
    except Exception:
        pass

    # Center-align the three outcome dot columns (compatible attempts)
    try:
        first_dot = len(HEADER_FIELDS)
        for col in range(first_dot, first_dot + 3):
            done = False
            try:
                dial_sheet.align_columns(columns=[col], align="center"); done = True
            except Exception:
                pass
            if not done:
                try:
                    dial_sheet.column_align(col, align="center"); done = True
                except Exception:
                    pass
            if not done:
                try:
                    total_rows = dial_sheet.get_total_rows()
                    for r in range(total_rows):
                        dial_sheet.set_cell_alignments(row=r, column=col, align="center")
                except Exception:
                    pass
    except Exception:
        pass

    def _current_note_text():
        return (window["-DIAL_NOTE-"].get() or "").strip()

    def _row_selected():
        try:
            sel = dial_sheet.get_selected_rows() or []
            if sel:
                return sel[0]
        except Exception:
            pass
        try:
            r, _ = dial_sheet.get_currently_selected()
            if isinstance(r, int) and r >= 0:
                return r
        except Exception:
            pass
        return None

    def _ensure_row_selected():
        r = _row_selected()
        if r is None:
            try:
                dial_sheet.set_currently_selected(0, 0)
                dial_sheet.see(0, 0)
                state["row"] = 0
                _paint_highlights()
                return 0
            except Exception:
                return None
        return r

    def _set_working_row(r):
        state["row"] = r
        try:
            dial_sheet.set_currently_selected(r, 0)
            dial_sheet.see(r, 0)
        except Exception:
            pass
        _paint_highlights()

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
        # Preserve viewport
        try:
            xv = dial_sheet.MT.xview(); yv = dial_sheet.MT.yview()
        except Exception:
            xv = yv = None
        # Clear colored dots on previously colored row
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

    # --- tksheet row delete (compat) ---
    def _delete_dialer_row(idx: int):
        try:
            dial_sheet.delete_rows(idx, 1); return True
        except Exception:
            pass
        try:
            dial_sheet.delete_rows(idx); return True
        except Exception:
            pass
        try:
            dial_sheet.del_rows(idx, 1); return True
        except Exception:
            pass
        try:
            dial_sheet.del_rows(idx); return True
        except Exception:
            pass
        return False

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

    # bind: clicks select working row
    def _on_grid_click(_evt=None):
        r = _row_selected()
        if r is not None:
            _set_working_row(r)
            state["note_col_by_row"].setdefault(r, None)
            _update_confirm_button()

    try:
        dial_sheet.MT.bind("<ButtonRelease-1>", _on_grid_click)
        dial_sheet.MT.bind("<ButtonRelease-3>", _on_grid_click)
    except Exception:
        pass

    # prime UI metrics
    refresh_fire_state()
    refresh_results_metrics()
    _update_confirm_button()

    # ============================================================
    # Event loop
    # ============================================================
    while True:
        event, values = window.read(timeout=300)
        if event in (sg.WINDOW_CLOSE_ATTEMPTED_EVENT, sg.WIN_CLOSED):
            break

        # track selection
        r_now = _row_selected()
        if r_now is not None and r_now != state["row"]:
            _set_working_row(r_now)
            state["note_col_by_row"].setdefault(r_now, None)
            _update_confirm_button()

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
                            except Exception: sheet.del_rows(r, 1)
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

        elif event == "-FIRE-":
            window["-STATUS-"].update("Preparing to fire…")
            if not require_pywin32():
                window["-STATUS-"].update("pywin32 missing (Outlook COM). Install pywin32.")
                continue
            try:
                matrix = matrix_from_sheet()
                save_matrix_to_csv(matrix)
                seen = load_state_set()
                new_rows = []
                for row in matrix:
                    d = dict_from_row(row)
                    if not valid_email(d.get("Email","")): continue
                    fp = row_fingerprint_from_dict(d)
                    if fp not in seen: new_rows.append(row)
                if not new_rows:
                    window["-STATUS-"].update("No new leads to draft.")
                else:
                    made = outlook_draft_many(new_rows, seen, templates, subjects, mapping)
                    window["-STATUS-"].update(f"Done. Created {made} drafts → Outlook/Drafts/{DEATHSTAR_SUBFOLDER}")
                refresh_fire_state()
            except Exception as e:
                window["-STATUS-"].update(f"Error: {e}")

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

        # ----- Email Results tagging -----
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

        # ----- Dialer outcome buttons -----
        elif event in ("-DIAL_SET_GREEN-","-DIAL_SET_GRAY-","-DIAL_SET_RED-"):
            r = state["row"]
            if r is None:
                r = _ensure_row_selected()
                if r is None:
                    window["-DIAL_MSG-"].update("Pick a row first.")
                    continue
                if state["row"] != r:
                    _set_working_row(r)
                    state["note_col_by_row"].setdefault(r, None)
            which = {"-DIAL_SET_GREEN-":"green","-DIAL_SET_GRAY-":"gray","-DIAL_SET_RED-":"red"}[event]
            _apply_outcome(state["row"], which)
            _save_dialer_grid()

        # Note typing → live preview in sticky column
        elif event == "-DIAL_NOTE-":
            r = state["row"]
            if r is None:
                r = _ensure_row_selected()
                if r is None:
                    continue
                _set_working_row(r)
                state["note_col_by_row"].setdefault(r, None)
            _apply_note_preview(r)
            _save_dialer_grid()

        elif event == "-DIAL_ADD100-":
            _append_100_rows()

        # Confirm Call → persist; green/red remove row; gray stays (dark gray)
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
                        # Build lead dict
                        row_vals = dial_sheet.get_row_data(r) or []
                        base = dict_from_row([row_vals[i] if i < len(row_vals) else "" for i in range(len(HEADER_FIELDS))])

                        # Ensure note saved
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

                        # Persist CSVs
                        dialer_save_call(base, outcome, note_text)
                        if outcome == "red":
                            add_no_interest(base, note_text, no_contact_flag=0, source="Dialer")
                        elif outcome == "gray":
                            filled = 0
                            row_vals2 = dial_sheet.get_row_data(r) or []
                            for k in range(cols["first_note"], cols["last_note"]+1):
                                if k < len(row_vals2) and (row_vals2[k] or "").strip():
                                    filled += 1
                            if filled >= 8:
                                add_no_interest(base, "No Contact after 8 calls. " + note_text, no_contact_flag=1, source="Dialer")

                        # UI feedback
                        window["-DIAL_MSG-"].update("Saved ✓")
                        window["-DIAL_NOTE-"].update("")
                        state["outcome"] = None
                        state["note_col_by_row"].pop(r, None)

                        # If green or red → delete row & advance
                        if outcome in ("green", "red"):
                            try:
                                if state["colored_row"] is not None:
                                    dialer_clear_dot_highlights(dial_sheet, state["colored_row"], cols)
                            except Exception:
                                pass
                            state["colored_row"] = None
                            state["done_rows"].discard(r)

                            ok_del = _delete_dialer_row(r)
                            if not ok_del:
                                window["-DIAL_MSG-"].update("Save ok, but could not delete grid row (tksheet).")
                            else:
                                try: dial_sheet.refresh()
                                except Exception: pass
                                state["done_rows"] = {(x-1 if x > r else x) for x in state["done_rows"]}
                                state["note_col_by_row"] = {(ri-1 if ri > r else ri): ci
                                    for ri, ci in state["note_col_by_row"].items() if ri != r}
                                try:
                                    total = dial_sheet.get_total_rows()
                                except Exception:
                                    total = 0
                                if total <= 0:
                                    state["row"] = None
                                else:
                                    new_idx = min(r, max(0, total - 1))
                                    _set_working_row(new_idx)
                                    state["note_col_by_row"].setdefault(new_idx, None)

                            _paint_highlights()
                            _update_confirm_button()
                            _save_dialer_grid()
                            continue

                        # gray → keep row, mark dark, advance
                        state["done_rows"].add(r)
                        try:
                            if state["colored_row"] is not None:
                                dialer_clear_dot_highlights(dial_sheet, state["colored_row"], cols)
                        except Exception:
                            pass
                        state["colored_row"] = None

                        new_row = dialer_move_to_next_row(dial_sheet, r)
                        _set_working_row(new_row)
                        state["note_col_by_row"].setdefault(new_row, None)
                        _paint_highlights()
                        _update_confirm_button()
                        _save_dialer_grid()

                    except Exception as e:
                        window["-DIAL_MSG-"].update(f"Save error: {e}")

        # Keep the Fire button state fresh
        refresh_fire_state()

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

