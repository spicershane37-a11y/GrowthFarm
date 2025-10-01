# deathstar.py — The Death Star (GUI)
# v2025-09-29 — Email + Dialer + Warm Leads + Analytics (charts) + Plain Paste RC + Resizable Columns
# Requires: PySimpleGUI, tksheet, matplotlib; (optional) pywin32 for Outlook

import os, sys, csv, html, time, re, hashlib, configparser, io, base64
from datetime import datetime, timedelta
from pathlib import Path
import PySimpleGUI as sg

# Matplotlib headless for PySimpleGUI Image
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

APP_VERSION = "2025-09-29"

# -------------------- App data locations --------------------
APP_DIR      = Path(os.environ.get("APPDATA", str(Path.home()))) / "DeathStarApp"  # Roaming
CSV_PATH     = APP_DIR / "kybercrystals.csv"
STATE_PATH   = APP_DIR / "annihilated_planets.txt"
TPL_PATH     = APP_DIR / "templates.ini"
RESULTS_PATH = APP_DIR / "results.csv"

# Dialer / warm / no-interest files
DIALER_RESULTS_PATH = APP_DIR / "dialer_results.csv"
WARM_LEADS_PATH     = APP_DIR / "warm_leads.csv"
NO_INTEREST_PATH    = APP_DIR / "no_interest.csv"

# -------------------- Outlook settings ----------------------
DEATHSTAR_SUBFOLDER = "Order 66"   # Drafts subfolder
TARGET_MAILBOX_HINT = ""           # Optional: part of SMTP/display name to target a specific account

# -------------------- Columns / headers ---------------------
HEADER_FIELDS = [
    "Email","First Name","Last Name","Company","Industry","Phone",
    "Address","City","State","Reviews","Website"
]

WARM_FIELDS = [
    "Timestamp","Email","First Name","Last Name","Company","Industry",
    "Phone","City","State","Website","Note","Source"
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
    """Start from the same CSV backing Email Leads (kybercrystals.csv)."""
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
            w.writerow([
                ts,
                row_dict.get("Email",""), row_dict.get("First Name",""), row_dict.get("Last Name",""),
                row_dict.get("Company",""), row_dict.get("Industry",""), row_dict.get("Phone",""),
                row_dict.get("City",""), row_dict.get("State",""), row_dict.get("Website",""),
                note, "Dialer"
            ])

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
        w.writerow([
            ts,
            result_row.get("Email",""),
            enr.get("First Name",""), enr.get("Last Name",""),
            result_row.get("Company",""), result_row.get("Industry",""),
            enr.get("Phone",""), enr.get("City",""), enr.get("State",""),
            enr.get("Website",""),
            note, "EmailResults"
        ])

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
# Analytics helpers & charts
# ============================================================

def _parse_dt(s: str):
    if not s: return None
    s = s.strip()
    fmts = [
        "%Y-%m-%d %H:%M:%S",     # our CSVs
        "%Y-%m-%d",
        "%m/%d/%Y %I:%M %p",     # Outlook
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
    ]
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None

def _in_range(dt_obj, start_dt, end_dt):
    if dt_obj is None: return False
    if start_dt and dt_obj < start_dt: return False
    if end_dt and dt_obj > end_dt: return False
    return True

def count_dials_between(start_dt, end_dt):
    total = 0
    if DIALER_RESULTS_PATH.exists():
        with DIALER_RESULTS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                ts = _parse_dt(r.get("Timestamp",""))
                if _in_range(ts, start_dt, end_dt):
                    total += 1
    return total

def count_emails_sent_between(start_dt, end_dt):
    total = 0
    if RESULTS_PATH.exists():
        with RESULTS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                ts = _parse_dt(r.get("DateSent",""))
                if _in_range(ts, start_dt, end_dt):
                    total += 1
    return total

def warm_leads_breakdown_between(start_dt, end_dt):
    warm_total = 0
    warm_email = 0
    warm_dialer = 0
    if WARM_LEADS_PATH.exists():
        with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                ts = _parse_dt(r.get("Timestamp",""))
                if not _in_range(ts, start_dt, end_dt):
                    continue
                warm_total += 1
                src = (r.get("Source","") or "").strip().lower()
                if src.startswith("email"):
                    warm_email += 1
                elif src.startswith("dialer"):
                    warm_dialer += 1
    return warm_total, warm_email, warm_dialer

def _date_key(dt_obj, groupby):
    if dt_obj is None: return None
    if groupby == "Day":
        return dt_obj.strftime("%Y-%m-%d")
    if groupby == "Week":
        monday = dt_obj - timedelta(days=dt_obj.weekday())
        return monday.strftime("%Y-%m-%d")
    if groupby == "Month":
        return dt_obj.strftime("%Y-%m")
    return dt_obj.strftime("%Y-%m-%d")

def _series_from_rows(rows, ts_field, start_dt, end_dt, groupby, filt=lambda r: True):
    buckets = {}
    for r in rows:
        if not filt(r):
            continue
        ts = _parse_dt(r.get(ts_field,""))
        if not _in_range(ts, start_dt, end_dt):
            continue
        key = _date_key(ts, groupby)
        if key is None:
            continue
        buckets[key] = buckets.get(key, 0) + 1
    labels = sorted(buckets.keys())
    values = [buckets[k] for k in labels]
    return labels, values

def _png_from_bar(labels, values, title):
    fig = plt.figure(figsize=(6,2.4), dpi=110)
    ax = fig.add_subplot(111)
    ax.bar(labels, values)
    ax.set_title(title)
    ax.set_ylabel("Count")
    ax.tick_params(axis='x', labelrotation=45)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("ascii")
# ============================================================
# “Add New Template” dialog
# ============================================================

def add_new_template_dialog():
    layout = [
        [sg.Text("Industry keyword(s) (comma-separated):")],
        [sg.Input(key="-NEW_KEYS-", size=(60,1))],
        [sg.Text("Template name (optional; defaults to first keyword):")],
        [sg.Input(key="-NEW_NAME-", size=(60,1))],
        [sg.Text("Subject:")],
        [sg.Input(key="-NEW_SUBJ-", size=(60,1))],
        [sg.Text("Body:")],
        [sg.Multiline(key="-NEW_BODY-", size=(60,10), font=("Consolas",10))],
        [sg.Push(), sg.Button("Save", key="-SAVE-"), sg.Button("Cancel")]
    ]
    win = sg.Window("Add New Template", layout, modal=True)
    evt, vals = win.read()
    res = None
    if evt == "-SAVE-":
        keys_raw = (vals.get("-NEW_KEYS-","") or "").strip()
        subj     = (vals.get("-NEW_SUBJ-","") or "").strip()
        body     = (vals.get("-NEW_BODY-","") or "").strip()
        name     = (vals.get("-NEW_NAME-","") or "").strip()
        if keys_raw and subj and body:
            keys = [k.strip() for k in keys_raw.split(",") if k.strip()]
            if not name:
                name = keys[0].replace(" ","_").lower()
            res = {"keys": keys, "name": name, "subject": subj, "body": body}
        else:
            sg.popup_error("Please provide industry keyword(s), subject, and body.")
    win.close()
    return res

# ============================================================
# Clipboard → plain-text paste helper for tksheet (both grids)
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
            sheet_obj.paste_data(rows)  # tksheet >=6
        except Exception:
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
        if not w: continue
        try:
            w.bind("<Control-v>", _paste_plain)
            w.bind("<Control-V>", _paste_plain)
            w.bind("<Control-Shift-v>", _paste_plain)
            w.bind("<Control-Shift-V>", _paste_plain)
        except Exception:
            pass

# ============================================================
# Right-click menu: only "Paste (Ctrl+Shift+V)" as plain text
# ============================================================

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
                try: m.grab_release()
                except Exception: pass

        # Bind right-click (and Button-2 for macs)
        try:
            sheet_obj.MT.bind("<Button-3>", _popup)
            sheet_obj.MT.bind("<Button-2>", _popup)
        except Exception:
            pass
    except Exception:
        pass

# ============================================================
# tksheet: enable column resizing across versions
# ============================================================

def _enable_column_resizing(sheet_obj):
    """
    Turn on column resizing regardless of tksheet version.
    Enables a bunch of possible flags; ignores ones that don't exist.
    """
    flags = [
        "column_width_resize",   # tksheet 6.x
        "column_resize",         # older alias
        "resize_columns",        # older alias
        "drag_select",           # keep drag select on
        "column_drag_and_drop",  # retain column reordering
    ]
    try:
        # tksheet >= 6.x: you can pass a tuple of bindings to enable_bindings
        sheet_obj.enable_bindings(tuple(set(sheet_obj.get_bindings() + tuple(flags))))  # if get_bindings exists
    except Exception:
        try:
            # Try enabling each individually
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
    ensure_app_files()
    templates, subjects, mapping = load_templates_ini()
    sg.theme("DarkGrey13")

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

    # -------- Dialer tab (grid like Email Leads + dots + notes) --------
    ensure_dialer_files()
    ensure_no_interest_file()

    DIALER_EXTRA_COLS = ["🟢","⚪","🔴","Note1","Note2","Note3","Note4","Note5","Note6","Note7","Note8"]
    DIALER_HEADERS = HEADER_FIELDS + DIALER_EXTRA_COLS

    dialer_matrix = load_dialer_matrix_from_email_csv()
    if not dialer_matrix:
        dialer_matrix = [[""] * len(HEADER_FIELDS) for _ in range(50)]
    dialer_matrix = [row + ["","",""] + ([""]*8) for row in dialer_matrix]

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
        [sg.Multiline(key="-DIAL_NOTE-", size=(28,6), font=("Consolas",10), background_color="#111", text_color="#EEE")],
        [sg.Button("Confirm Call", key="-DIAL_CONFIRM-", size=(14,2), disabled=True, button_color=("white","#444444"))],
        [sg.Text("", key="-DIAL_MSG-", text_color="#A0FFA0", size=(28,2))]
    ]

    dialer_tab = [
        [sg.Column([[dialer_host]], expand_x=True, expand_y=True),
         sg.Column(dialer_controls, vertical_alignment="top", pad=((10,0),(0,0)))]
    ]

    # -------- Warm Leads tab --------
    ensure_warm_file()
    warm_rows = []
    if WARM_LEADS_PATH.exists():
        with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                warm_rows.append(r)

    warm_table_data = [[r.get(h, "") for h in WARM_FIELDS] for r in warm_rows]

    warm_tab = [
        [sg.Text("Live from warm_leads.csv", text_color="#CCCCCC")],
        [sg.Table(values=warm_table_data, headings=WARM_FIELDS,
                  auto_size_columns=False, col_widths=[18,26,12,12,22,14,14,14,8,26,28,12],
                  key="-WARM_TABLE-", num_rows=14, alternating_row_color="#2a2a2a",
                  text_color="#EEE", background_color="#111",
                  header_text_color="#FFF", header_background_color="#333")],
        [sg.Button("Export Warm Leads CSV", key="-WARM_EXPORT-"),
         sg.Button("Reload Warm", key="-WARM_RELOAD-"),
         sg.Text("", key="-WARM_STATUS-", text_color="#A0FFA0")]
    ]

    # -------- Analytics tab --------
    analytics_tab = [
        [sg.Text("Analytics — filter by date, choose grouping, then Apply.", text_color="#CCCCCC")],
        [sg.Text("From:"), sg.Input(key="-AN_FROM-", size=(12,1)),
         sg.CalendarButton("📅", target="-AN_FROM-", format="%Y-%m-%d"),
         sg.Text("   To:"), sg.Input(key="-AN_TO-", size=(12,1)),
         sg.CalendarButton("📅", target="-AN_TO-", format="%Y-%m-%d"),
         sg.Text("   Group by:"), sg.Combo(["Day","Week","Month"], default_value="Week", key="-AN_GROUP-", readonly=True, size=(8,1)),
         sg.Button("Apply", key="-AN_APPLY-"),
         sg.Push(),
         sg.Button("This Week", key="-AN_PRESET_WEEK-"),
         sg.Button("This Month", key="-AN_PRESET_MONTH-"),
         sg.Button("This Year", key="-AN_PRESET_YEAR-"),
         sg.Button("All Time", key="-AN_PRESET_ALL-")],
        [sg.HorizontalSeparator(color="#4CAF50")],
        [sg.Text("Totals in range:", text_color="#9EE493"),
         sg.Text("  Dials: 0", key="-AN_DIALS-", text_color="#9EE493"),
         sg.Text("  Emails: 0", key="-AN_EMAILS-", text_color="#9EE493"),
         sg.Text("  Warm(all): 0", key="-AN_WARM_ALL-", text_color="#FFFFFF"),
         sg.Text("  •Email: 0", key="-AN_WARM_EMAIL-", text_color="#CCCCCC"),
         sg.Text("  •Dialer: 0", key="-AN_WARM_DIALER-", text_color="#CCCCCC")],
        [sg.Text("", key="-AN_STATUS-", text_color="#A0FFA0")],
        [sg.HorizontalSeparator(color="#4CAF50")],
        [sg.Text("Dials by period")],
        [sg.Image(key="-CH_DIALS-")],
        [sg.Text("Emails Sent by period")],
        [sg.Image(key="-CH_EMAILS-")],
        [sg.Text("Warm Leads (All) by period")],
        [sg.Image(key="-CH_WARM_ALL-")],
        [sg.Text("Warm from Email by period")],
        [sg.Image(key="-CH_WARM_EMAIL-")],
        [sg.Text("Warm from Dialer by period")],
        [sg.Image(key="-CH_WARM_DIALER-")],
    ]

    # -------- Compose full layout --------
    layout = [
        top_bar,
        [sg.TabGroup([[sg.Tab("Email Leads",    leads_tab,   expand_x=True, expand_y=True),
                       sg.Tab("Email Templates",tpl_tab,     expand_x=True, expand_y=True),
                       sg.Tab("Email Results",  results_tab, expand_x=True, expand_y=True),
                       sg.Tab("Dialer",         dialer_tab,  expand_x=True, expand_y=True),
                       sg.Tab("Warm Leads",     warm_tab,    expand_x=True, expand_y=True),
                       sg.Tab("Analytics",      analytics_tab, expand_x=True, expand_y=True)]],
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
        sg.popup_error("tksheet not installed. Run: pip install tksheet")
        return

    # Email Leads sheet in leads_host
    host_frame_tk = leads_host.Widget
    for child in host_frame_tk.winfo_children():
        try: child.destroy()
        except Exception: pass
    sheet_holder = sg.tk.Frame(host_frame_tk, bg="#111111")
    sheet_holder.pack(side="top", fill="both", expand=True)

    existing = load_csv_to_matrix()
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
        # attempt common resize flags while we're here (some versions accept it here)
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

    # Force plain-text paste + RC menu
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
    first_dot = len(HEADER_FIELDS)         # 🟢
    first_note = len(HEADER_FIELDS) + 3    # Note1
    last_note  = first_note + 7            # Note8

    for c in range(len(DIALER_HEADERS)):
        width = DEFAULT_COL_WIDTH
        if c == idx_address: width = 120
        if c == idx_city:    width = 90
        if c == idx_state:   width = 42
        if c == idx_reviews: width = 60
        if c == idx_website: width = 160
        if first_dot <= c < first_note:  # 3 dots
            width = 36
        if first_note <= c <= last_note: # notes
            width = 120
        try: dial_sheet.column_width(c, width=width)
        except Exception: pass

    _bind_plaintext_paste_for_tksheet(dial_sheet, window.TKroot)
    _ensure_rc_menu_plain_paste(dial_sheet, window.TKroot)
    _enable_column_resizing(dial_sheet)
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
    if not text: return ""
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
        # Prefer Accounts (primary stores)
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
        # Fallback: match by store display
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

    # Drafts root and target subfolder
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
        if not valid_email(d.get("Email","")): continue
        fp = row_fingerprint_from_dict(d)
        if fp in seen_set: continue
        ref_short = fp[:8]

        tpl_key   = choose_template_key(d.get("Industry",""), mapping)
        body_tpl  = templates.get(tpl_key, templates.get("default",""))
        subj_tpl  = subjects.get(tpl_key) or subjects.get("default") or DEFAULT_SUBJECT

        subj_text = apply_placeholders(subj_tpl, d)
        body_text = apply_placeholders(body_tpl, d)

        body_html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;">
<div style="font-family:Segoe UI, Arial, sans-serif; font-size:14px; line-height:1.5; color:#111;">
{blocks_to_html(body_text)}
<!-- ref:{ref_short} -->
</div></body></html>"""

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
        for fp in new_fps: f.write(fp+"\n")
    return made

def load_state_set():
    if not STATE_PATH.exists(): return set()
    return {line.strip() for line in STATE_PATH.read_text(encoding="utf-8").splitlines() if line.strip()}

def load_results_rows_sorted():
    # (Already defined in Part 2, but ensure available if file concatenation changes)
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
    sent_items = sent.Items;   sent_items.IncludeRecurrences = True;   sent_items.Sort("[SentOn]", True)
    inbox_items = inbox.Items; inbox_items.IncludeRecurrences = True; inbox_items.Sort("[ReceivedTime]", True)
    sent_recent  = sent_items.Restrict(f"[SentOn] >= '{since}'")
    inbox_recent = inbox_items.Restrict(f"[ReceivedTime] >= '{since}'")

    sent_map, reply_map = {}, {}
    for i in range(1, min(2000, sent_recent.Count)+1):
        try:
            m = sent_recent.Item(i)
            subj = str(getattr(m,"Subject","") or "")
            rm = REF_RE.search(subj)
            if rm: sent_map[rm.group(1).lower()] = str(getattr(m,"SentOn","") or "")
        except Exception:
            continue
    for i in range(1, min(2000, inbox_recent.Count)+1):
        try:
            m = inbox_recent.Item(i)
            subj = str(getattr(m,"Subject","") or "")
            rm = REF_RE.search(subj)
            if rm: reply_map[rm.group(1).lower()] = str(getattr(m,"ReceivedTime","") or "")
        except Exception:
            continue

    rows = load_results_rows_sorted()
    byref = {r["Ref"].lower(): r for r in rows}
    for ref, dt in sent_map.items():
        if ref in byref: byref[ref]["DateSent"] = dt
        else:
            byref[ref] = {"Ref":ref,"Email":"","Company":"","Industry":"","DateSent":dt,"DateReplied":"","Status":"","Subject":""}
    for ref, dt in reply_map.items():
        if ref in byref: byref[ref]["DateReplied"] = dt
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

def _series_from_rows(rows, ts_field, start_dt, end_dt, groupby, filt=lambda r: True):
    # (Already defined in Part 2 — redefine defensively for completeness)
    buckets = {}
    for r in rows:
        if not filt(r):
            continue
        ts = _parse_dt(r.get(ts_field,""))
        if not _in_range(ts, start_dt, end_dt):
            continue
        key = _date_key(ts, groupby)
        if key is None:
            continue
        buckets[key] = buckets.get(key, 0) + 1
    labels = sorted(buckets.keys())
    values = [buckets[k] for k in labels]
    return labels, values

def main_after_mount(window, sheet, dial_sheet, leads_host, dialer_host, templates, subjects, mapping):
    # ---------- helpers inside main() ----------

    def matrix_from_sheet():
        raw = sheet.get_sheet_data() or []
        trimmed = []
        for row in raw:
            row = (list(row) + [""] * len(HEADER_FIELDS))[:len(HEADER_FIELDS)]
            trimmed.append([str(c) for c in row])
        while trimmed and not any(cell.strip() for cell in trimmed[-1]):
            trimmed.pop()
        return trimmed

    def refresh_fire_state():
        matrix = matrix_from_sheet()
        seen = load_state_set()
        new_count = 0
        for row in matrix:
            d = dict_from_row(row)
            if not valid_email(d.get("Email","")): continue
            fp = row_fingerprint_from_dict(d)
            if fp not in seen: new_count += 1
        if new_count > 0:
            window["-FIRE-"].update(disabled=False, button_color=("white","#C00000"))
            window["-FIRE_HINT-"].update(f"  Ready: {new_count} new lead(s).")
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

    # Prime initial UI states
    refresh_fire_state()
    refresh_results_metrics()

    # --------- Analytics initial paint (All Time) ----------
    def analytics_refresh_ui(window, from_str, to_str, groupby):
        start_dt = datetime.strptime(from_str, "%Y-%m-%d") if from_str else None
        end_dt   = (datetime.strptime(to_str, "%Y-%m-%d") + timedelta(days=1) - timedelta(seconds=1)) if to_str else None

        # Load files once
        dial_rows = []
        if DIALER_RESULTS_PATH.exists():
            with DIALER_RESULTS_PATH.open("r", encoding="utf-8", newline="") as f:
                dial_rows = list(csv.DictReader(f))

        res_rows = []
        if RESULTS_PATH.exists():
            with RESULTS_PATH.open("r", encoding="utf-8", newline="") as f:
                res_rows = list(csv.DictReader(f))

        warm_rows = []
        if WARM_LEADS_PATH.exists():
            with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
                warm_rows = list(csv.DictReader(f))

        # Totals
        dials_total = count_dials_between(start_dt, end_dt)
        emails_total = count_emails_sent_between(start_dt, end_dt)
        warm_total, warm_email, warm_dialer = warm_leads_breakdown_between(start_dt, end_dt)

        window["-AN_DIALS-"].update(f"  Dials: {dials_total}")
        window["-AN_EMAILS-"].update(f"  Emails: {emails_total}")
        window["-AN_WARM_ALL-"].update(f"  Warm(all): {warm_total}")
        window["-AN_WARM_EMAIL-"].update(f"  •Email: {warm_email}")
        window["-AN_WARM_DIALER-"].update(f"  •Dialer: {warm_dialer}")

        # Charts
        def _update_chart(img_key, labels, values, title):
            png_b64 = _png_from_bar(labels, values, title)
            window[img_key].update(data=base64.b64decode(png_b64))

        lbl, val = _series_from_rows(dial_rows, "Timestamp", start_dt, end_dt, groupby)
        _update_chart("-CH_DIALS-", lbl, val, f"Dials per {groupby}")

        lbl, val = _series_from_rows(res_rows, "DateSent", start_dt, end_dt, groupby)
        _update_chart("-CH_EMAILS-", lbl, val, f"Emails Sent per {groupby}")

        lbl, val = _series_from_rows(warm_rows, "Timestamp", start_dt, end_dt, groupby)
        _update_chart("-CH_WARM_ALL-", lbl, val, f"Warm Leads (All) per {groupby}")

        lbl, val = _series_from_rows(warm_rows, "Timestamp", start_dt, end_dt, groupby,
                                     filt=lambda r: (r.get("Source","") or "").lower().startswith("email"))
        _update_chart("-CH_WARM_EMAIL-", lbl, val, f"Warm from Email per {groupby}")

        lbl, val = _series_from_rows(warm_rows, "Timestamp", start_dt, end_dt, groupby,
                                     filt=lambda r: (r.get("Source","") or "").lower().startswith("dialer"))
        _update_chart("-CH_WARM_DIALER-", lbl, val, f"Warm from Dialer per {groupby}")

    analytics_refresh_ui(window, "", "", "Week")

    # ---------- Dialer state ----------
    _dial_current_row = {"idx": None}
    _dial_outcome = {"val": None}  # "green" | "gray" | "red" | None

    def _dial_get_row_base(idx):
        row = dial_sheet.get_row_data(idx) or []
        return {HEADER_FIELDS[i]: (row[i] if i < len(HEADER_FIELDS) else "") for i in range(len(HEADER_FIELDS))}

    def _dial_count_calls(idx):
        row = dial_sheet.get_row_data(idx) or []
        cnt = 0
        first_note = len(HEADER_FIELDS) + 3
        last_note = first_note + 7
        for c in range(first_note, last_note+1):
            try:
                if (row[c] or "").strip():
                    cnt += 1
            except Exception:
                pass
        return cnt

    def _dial_put_note_in_next_slot(idx, note_text):
        row = dial_sheet.get_row_data(idx) or []
        first_note = len(HEADER_FIELDS) + 3
        last_note  = first_note + 7
        for c in range(first_note, last_note+1):
            cell = (row[c] if c < len(row) else "")
            if not (cell or "").strip():
                dial_sheet.set_cell_data(idx, c, note_text)
                dial_sheet.refresh()
                return _dial_count_calls(idx)
        dial_sheet.set_cell_data(idx, last_note, note_text)
        dial_sheet.refresh()
        return 8

    def _dial_update_confirm_enabled():
        note_text = (window["-DIAL_NOTE-"].get() or "").strip()
        ok = (_dial_current_row["idx"] is not None) and (len(note_text) > 0)
        window["-DIAL_CONFIRM-"].update(disabled=not ok, button_color=("white","#2E7D32" if ok else "#444444"))

    # ============================================================
    # Event loop
    # ============================================================

    while True:
        event, values = window.read(timeout=400)
        if event in (sg.WINDOW_CLOSE_ATTEMPTED_EVENT, sg.WIN_CLOSED):
            break

        # poll dialer selection (tksheet doesn't emit PSG events)
        try:
            sel_rows = dial_sheet.get_selected_rows() or []
            if sel_rows:
                if _dial_current_row["idx"] != sel_rows[0]:
                    _dial_current_row["idx"] = sel_rows[0]
                    _dial_update_confirm_enabled()
        except Exception:
            pass

        if event == "-OPENFOLDER-":
            try:
                os.startfile(str(APP_DIR))
            except Exception as e:
                sg.popup_error(f"Open folder error: {e}")

        elif event == "-ADDROWS-":
            try:
                sheet.insert_rows(sheet.get_total_rows(), number_of_rows=10); sheet.refresh()
            except Exception:
                try:
                    sheet.insert_rows(sheet.get_total_rows(), amount=10); sheet.refresh()
                except Exception as e:
                    sg.popup_error(f"Could not add rows: {e}")
            refresh_fire_state()

        elif event == "-DELROWS-":
            try:
                sels = sheet.get_selected_rows() or []
                if sels:
                    for r in sorted(sels, reverse=True):
                        try: sheet.delete_rows(r, number_of_rows=1)
                        except Exception: sheet.delete_rows(r, amount=1)
                    sheet.refresh()
            except Exception as e:
                sg.popup_error(f"Could not delete rows: {e}")
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
            window.close()
            # re-run full main to rebuild UI from disk
            main()
            return

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
                map_lines = (values.get("-MAP-","") or "").splitlines()
                new_map = {}
                for line in map_lines:
                    if "->" in line:
                        left,right = line.split("->",1); left,right = left.strip(), right.strip()
                        if left and right: new_map[left] = right
                save_templates_ini(tpls_out, subs_out, new_map)
                window["-TPL_STATUS-"].update("Templates & mapping saved ✓")
            except Exception as e:
                window["-TPL_STATUS-"].update(f"Save error: {e}")

        elif event == "-RESETTPL-":
            save_templates_ini(DEFAULT_TEMPLATES, DEFAULT_SUBJECTS, DEFAULT_MAP)
            window.close()
            main()
            return

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

        # ----- Email Results marking → Warm / No Interest -----
        elif event == "-MARK_GREEN-":
            sels = values.get("-RSTABLE-", [])
            if sels:
                idx = sels[0]
                rows = load_results_rows_sorted()
                if 0 <= idx < len(rows):
                    set_status(rows[idx]["Ref"], "green")
                    try:
                        add_warm_from_result(rows[idx], note="Marked Green on Email Results")
                    except Exception as e:
                        print("Warm add error:", e)
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
                    try:
                        add_no_interest_from_result(rows[idx], note="Marked Red on Email Results", no_contact_flag=0)
                    except Exception as e:
                        print("No-interest add error:", e)
                rows = load_results_rows_sorted()
                data = [[r.get("Ref",""), r.get("Email",""), r.get("Company",""), r.get("Industry",""),
                         r.get("DateSent",""), r.get("DateReplied",""), r.get("Status",""), r.get("Subject","")] for r in rows]
                window["-RSTABLE-"].update(values=data)
                refresh_results_metrics()

        # ----- Dialer outcome buttons -----
        elif event == "-DIAL_SET_GREEN-":
            if _dial_current_row["idx"] is not None:
                r = _dial_current_row["idx"]
                first_dot = len(HEADER_FIELDS)
                dial_sheet.set_cell_data(r, first_dot+0, "●")
                dial_sheet.set_cell_data(r, first_dot+1, "")
                dial_sheet.set_cell_data(r, first_dot+2, "")
                _dial_outcome["val"] = "green"
                dial_sheet.refresh()
                _dial_update_confirm_enabled()

        elif event == "-DIAL_SET_GRAY-":
            if _dial_current_row["idx"] is not None:
                r = _dial_current_row["idx"]
                first_dot = len(HEADER_FIELDS)
                dial_sheet.set_cell_data(r, first_dot+0, "")
                dial_sheet.set_cell_data(r, first_dot+1, "●")
                dial_sheet.set_cell_data(r, first_dot+2, "")
                _dial_outcome["val"] = "gray"
                dial_sheet.refresh()
                _dial_update_confirm_enabled()

        elif event == "-DIAL_SET_RED-":
            if _dial_current_row["idx"] is not None:
                r = _dial_current_row["idx"]
                first_dot = len(HEADER_FIELDS)
                dial_sheet.set_cell_data(r, first_dot+0, "")
                dial_sheet.set_cell_data(r, first_dot+1, "")
                dial_sheet.set_cell_data(r, first_dot+2, "●")
                _dial_outcome["val"] = "red"
                dial_sheet.refresh()
                _dial_update_confirm_enabled()

        elif event == "-DIAL_NOTE-":
            _dial_update_confirm_enabled()

        elif event == "-DIAL_CONFIRM-":
            if _dial_current_row["idx"] is None:
                window["-DIAL_MSG-"].update("Pick a row first.")
            else:
                r = _dial_current_row["idx"]
                base = _dial_get_row_base(r)
                note_text = (values.get("-DIAL_NOTE-","") or "").strip()
                if not note_text:
                    window["-DIAL_MSG-"].update("Type a note.")
                else:
                    outcome = _dial_outcome["val"] or "gray"
                    try:
                        # persist dial
                        dialer_save_call(base, outcome, note_text)
                        calls_after = _dial_put_note_in_next_slot(r, note_text)

                        # update warm/no-interest
                        if outcome == "green":
                            pass  # already added to warm in dialer_save_call
                        elif outcome == "red":
                            add_no_interest(base, note_text, no_contact_flag=0, source="Dialer")
                        else:
                            if calls_after >= 8:
                                add_no_interest(base, "No Contact after 8 calls. " + note_text, no_contact_flag=1, source="Dialer")

                        window["-DIAL_MSG-"].update("Saved ✓")
                        window["-DIAL_NOTE-"].update("")
                        first_dot = len(HEADER_FIELDS)
                        dial_sheet.set_cell_data(r, first_dot+0, "")
                        dial_sheet.set_cell_data(r, first_dot+1, "")
                        dial_sheet.set_cell_data(r, first_dot+2, "")
                        dial_sheet.refresh()
                        _dial_outcome["val"] = None
                        _dial_update_confirm_enabled()
                    except Exception as e:
                        window["-DIAL_MSG-"].update(f"Save error: {e}")

        # ----- Warm tab events -----
        elif event == "-WARM_EXPORT-":
            path = sg.popup_get_file("Save warm_leads.csv", save_as=True, default_extension=".csv",
                                     file_types=(("CSV","*.csv"),), no_window=True)
            if path:
                try:
                    with WARM_LEADS_PATH.open("r", encoding="utf-8") as src, open(path,"w",encoding="utf-8") as dst:
                        dst.write(src.read())
                    window["-WARM_STATUS-"].update("Exported ✓")
                except Exception as e:
                    window["-WARM_STATUS-"].update(f"Export error: {e}")

        elif event == "-WARM_RELOAD-":
            ensure_warm_file()
            warm_rows = []
            with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
                rdr = csv.DictReader(f)
                for r in rdr: warm_rows.append(r)
            warm_table_data = [[r.get(h, "") for h in WARM_FIELDS] for r in warm_rows]
            window["-WARM_TABLE-"].update(values=warm_table_data)
            window["-WARM_STATUS-"].update("Reloaded ✓")

        # ----- Analytics presets & apply -----
        elif event == "-AN_PRESET_ALL-":
            window["-AN_FROM-"].update("")
            window["-AN_TO-"].update("")
            analytics_refresh_ui(window, "", "", values.get("-AN_GROUP-","Week"))
            window["-AN_STATUS-"].update("Updated ✓")

        elif event == "-AN_PRESET_WEEK-":
            today = datetime.now().date()
            start = today - timedelta(days=today.weekday())
            end = start + timedelta(days=6)
            fs, ts = start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")
            window["-AN_FROM-"].update(fs); window["-AN_TO-"].update(ts)
            analytics_refresh_ui(window, fs, ts, values.get("-AN_GROUP-","Week"))
            window["-AN_STATUS-"].update("Updated ✓")

        elif event == "-AN_PRESET_MONTH-":
            today = datetime.now().date()
            start = today.replace(day=1)
            next_start = start.replace(year=start.year+1, month=1, day=1) if start.month==12 else start.replace(month=start.month+1, day=1)
            end = next_start - timedelta(days=1)
            fs, ts = start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")
            window["-AN_FROM-"].update(fs); window["-AN_TO-"].update(ts)
            analytics_refresh_ui(window, fs, ts, values.get("-AN_GROUP-","Week"))
            window["-AN_STATUS-"].update("Updated ✓")

        elif event == "-AN_PRESET_YEAR-":
            today = datetime.now().date()
            start = today.replace(month=1, day=1)
            end = today.replace(month=12, day=31)
            fs, ts = start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")
            window["-AN_FROM-"].update(fs); window["-AN_TO-"].update(ts)
            analytics_refresh_ui(window, fs, ts, values.get("-AN_GROUP-","Week"))
            window["-AN_STATUS-"].update("Updated ✓")

        elif event == "-AN_APPLY-":
            fs = (values.get("-AN_FROM-","") or "").strip()
            ts = (values.get("-AN_TO-","") or "").strip()
            grp = values.get("-AN_GROUP-","Week")
            try:
                analytics_refresh_ui(window, fs, ts, grp)
                window["-AN_STATUS-"].update("Updated ✓")
            except Exception as e:
                window["-AN_STATUS-"].update(f"Error: {e}")

        # ----- Updater button (placeholder; wire to GitHub later) -----
        elif event == "-UPDATE-":
            sg.popup("Updater placeholder — when your GitHub repo/releases are ready, this button will pull the newest EXE.\n\nFor now, rebuild with PyInstaller and replace the EXE to update.")

        # Keep the Fire button state fresh
        refresh_fire_state()

    window.close()

# ============================================================
# Entry
# ============================================================

if __name__ == "__main__":
    ensure_app_files()
    # main() was defined in Part 3 and builds/mounts the UI and grids
    # Recreate the same flow here to pass objects into the event loop wrapper

    # We call main() once to set up the window and mount tksheet grids,
    # but since main() contains that build already, simply call it.
    # (If you prefer the explicit split, you could refactor; this preserves your structure.)
    try:
        main()
    except Exception as e:
        sg.popup_error(f"Fatal error starting app: {e}")
