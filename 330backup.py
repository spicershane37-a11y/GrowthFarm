# ===== CHUNK 1 / 7 — START =====
# deathstar.py — The Death Star (GUI)
# v2025-10-02 — Email + Dialer + Warm Leads + Plain Paste RC + Resizable Columns + Keyboard Nav
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

APP_VERSION = "2025-10-02"

# -------------------- App data locations --------------------
APP_DIR      = Path(os.environ.get("APPDATA", str(Path.home()))) / "DeathStarApp"  # Roaming
CSV_PATH     = APP_DIR / "kybercrystals.csv"
STATE_PATH   = APP_DIR / "annihilated_planets.txt"
TPL_PATH     = APP_DIR / "templates.ini"
RESULTS_PATH = APP_DIR / "results.csv"

# NEW: campaigns config file (separate from templates.ini to avoid clobbering)
CAMPAIGNS_INI = APP_DIR / "campaigns.ini"

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

# Customers sheet headers (FINAL ORDER with address block + simplified fields)
# NOTE: Lat/Lon are included directly after ZIP so the Map tab can read them.
CUSTOMER_FIELDS = [
    "Company",
    "Prospect Name",
    "Phone #",
    "Email",
    "Industry",
    "Address",
    "City",
    "State",
    "ZIP",
    "Lat",      # <-- added
    "Lon",      # <-- added
    "CLTV",
    "Sales/Day",
    "Reorder?",
    "First Order",
    "Last Order",
    "Days",
    "First Contact",
    "Days To Close",
    "Sku's",
    "Notes",
]

# Safety guard: if another build overrode CUSTOMER_FIELDS without Lat/Lon, append them.
if "Lat" not in CUSTOMER_FIELDS:
    CUSTOMER_FIELDS.insert(9, "Lat")
if "Lon" not in CUSTOMER_FIELDS:
    # ensure Lon immediately follows Lat
    lat_index = CUSTOMER_FIELDS.index("Lat")
    if "Lon" in CUSTOMER_FIELDS:
        pass
    else:
        CUSTOMER_FIELDS.insert(lat_index + 1, "Lon")

# -------------------- UI tuning -----------------------------
START_ROWS = 200
DEFAULT_COL_WIDTH = 140

# -------------------- Template defaults (SAFE FALLBACKS) ----
# These guards prevent crashes if the constants aren’t defined elsewhere.
if "DEFAULT_SUBJECT" not in globals():
    DEFAULT_SUBJECT = "Quick intro from YOUR COMPANY"

if "DEFAULT_TEMPLATES" not in globals():
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

if "DEFAULT_SUBJECTS" not in globals():
    DEFAULT_SUBJECTS = {
        "default":      DEFAULT_SUBJECT,
        "butcher_shop": DEFAULT_SUBJECT,
        "farm_orchard": DEFAULT_SUBJECT,
    }

if "DEFAULT_MAP" not in globals():
    DEFAULT_MAP = {}

# -------------------- Campaign defaults ---------------------
# Up to 3 steps; delays are from the *DateSent* in results.csv.
DEFAULT_CAMPAIGN_STEPS = [
    {"subject": DEFAULT_SUBJECTS.get("default", DEFAULT_SUBJECT), "body": DEFAULT_TEMPLATES["default"], "delay_days": "0"},
    {"subject": DEFAULT_SUBJECTS.get("default", DEFAULT_SUBJECT), "body": "", "delay_days": "3"},
    {"subject": DEFAULT_SUBJECTS.get("default", DEFAULT_SUBJECT), "body": "", "delay_days": "7"},
]
DEFAULT_CAMPAIGN_SETTINGS = {
    "send_to_dialer_after": "1",   # 1 = True, 0 = False
}

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
    ensure_campaigns_ini()  # NEW

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
# Campaigns persistence (separate INI)
# ============================================================
# ---------- Campaigns step normalization ----------

def _coerce_step_dict(step):
    """
    Convert any step representation (dict / tuple / list / str / None) into a
    normalized dict: {"enabled": bool, "subject": str, "body": str, "delay_days": int}
    """
    def _to_int(x, default=0):
        try:
            s = str(x).strip()
            return int(s) if s else default
        except Exception:
            return default

    if isinstance(step, dict):
        return {
            "enabled": str(step.get("enabled", "1")).strip().lower() not in ("0", "false", "no"),
            "subject": step.get("subject", "") or "",
            "body": step.get("body", "") or "",
            "delay_days": _to_int(step.get("delay_days", 0), 0),
        }

    if isinstance(step, (list, tuple)):
        # Heuristic: [enabled, subject, body, delay_days]
        enabled = True
        subject = ""
        body = ""
        delay_days = 0
        if len(step) > 0: enabled = str(step[0]).strip().lower() not in ("0", "false", "no", "")
        if len(step) > 1: subject = step[1] or ""
        if len(step) > 2: body = step[2] or ""
        if len(step) > 3: delay_days = _to_int(step[3], 0)
        return {
            "enabled": bool(enabled),
            "subject": subject,
            "body": body,
            "delay_days": delay_days,
        }

    if isinstance(step, str):
        # Treat as "body only"
        return {"enabled": True, "subject": "", "body": step, "delay_days": 0}

    # Fallback empty
    return {"enabled": False, "subject": "", "body": "", "delay_days": 0}


def normalize_campaign_steps(steps):
    """
    Ensure we have exactly 3 normalized step dicts.
    """
    norm = []
    steps = steps or []
    for i in range(min(3, len(steps))):
        norm.append(_coerce_step_dict(steps[i]))
    while len(norm) < 3:
        norm.append({"enabled": False, "subject": "", "body": "", "delay_days": 0})
    return norm


def normalize_campaign_settings(settings):
    """
    Single-source campaign settings (all strings "0"/"1"):
    - send_to_dialer_after: "0"/"1"  (default "1")
    - auto_sync_outlook:    "0"/"1"  (default "0")
    - hourly_campaign_runner:"0"/"1" (default "1")
    """
    defaults = {
        "send_to_dialer_after": "1",   # move non-responders to Dialer automatically
        "auto_sync_outlook": "0",      # Outlook auto-sync OFF by default (user can enable in UI)
        "hourly_campaign_runner": "1", # run campaign scheduler hourly
    }
    if isinstance(settings, dict):
        out = defaults.copy()
        out.update({
            "send_to_dialer_after": str(settings.get("send_to_dialer_after", defaults["send_to_dialer_after"])).strip(),
            "auto_sync_outlook":    str(settings.get("auto_sync_outlook",    defaults["auto_sync_outlook"])).strip(),
            "hourly_campaign_runner": str(settings.get("hourly_campaign_runner", defaults["hourly_campaign_runner"])).strip(),
        })
        return out
    return defaults


def ensure_campaigns_ini():
    """Create campaigns.ini with defaults if missing."""
    if CAMPAIGNS_INI.exists():
        return
    # Use defaults if available in globals; else create simple blanks
    try:
        steps = DEFAULT_CAMPAIGN_STEPS
    except NameError:
        steps = [
            {"subject": "", "body": "", "delay_days": 0},
            {"subject": "", "body": "", "delay_days": 0},
            {"subject": "", "body": "", "delay_days": 0},
        ]
    try:
        settings = DEFAULT_CAMPAIGN_SETTINGS
    except NameError:
        settings = {"send_to_dialer_after":"1","auto_sync_outlook":"0","hourly_campaign_runner":"1"}
    save_campaigns_ini(steps, settings)


def save_campaigns_ini(steps_list, settings_dict):
    """
    Persist up to 3 steps and simple settings.
    steps_list: list of dicts with keys {subject, body, delay_days}
    settings_dict: {"send_to_dialer_after": "0/1", "auto_sync_outlook":"0/1", "hourly_campaign_runner":"0/1"}
    """
    cfg = configparser.ConfigParser()

    # Steps (normalize before writing to be safe)
    steps_list = normalize_campaign_steps(steps_list)
    for i in range(1, 4):
        sec = f"step{i}"
        s = steps_list[i-1]
        cfg[sec] = {
            "enabled":    "1" if s.get("enabled", True) else "0",
            "subject":    s.get("subject", "") or "",
            "body":       s.get("body", "") or "",
            "delay_days": str(s.get("delay_days", 0)),
        }

    # Settings (normalize)
    settings_dict = normalize_campaign_settings(settings_dict or {})
    cfg["settings"] = {
        "send_to_dialer_after": settings_dict.get("send_to_dialer_after", "1"),
        "auto_sync_outlook":    settings_dict.get("auto_sync_outlook", "0"),
        "hourly_campaign_runner": settings_dict.get("hourly_campaign_runner", "1"),
    }

    with CAMPAIGNS_INI.open("w", encoding="utf-8") as f:
        cfg.write(f)


def load_campaigns_ini():
    """Return (steps_list, settings_dict)."""
    ensure_campaigns_ini()
    cfg = configparser.ConfigParser()
    cfg.read(CAMPAIGNS_INI, encoding="utf-8")

    steps = []
    for i in range(1, 4):
        sec = f"step{i}"
        s = cfg[sec] if sec in cfg else {}
        # accept older files that didn’t have "enabled"
        enabled_raw = (s.get("enabled", "1") if isinstance(s, dict) else "1")
        steps.append({
            "enabled":    str(enabled_raw).strip() not in ("0","false","no"),
            "subject":    s.get("subject", "") if isinstance(s, dict) else "",
            "body":       s.get("body", "") if isinstance(s, dict) else "",
            "delay_days": int(str(s.get("delay_days","0")).strip() or "0") if isinstance(s, dict) else 0,
        })

    # Settings
    if "settings" in cfg:
        raw = dict(cfg["settings"])
    else:
        try:
            raw = dict(DEFAULT_CAMPAIGN_SETTINGS)
        except NameError:
            raw = {}
    settings = normalize_campaign_settings(raw)

    return steps, settings
# ---------- Multi-campaign storage (named by niche/industry) ----------

# campaigns.ini layout (backward compatible):
# [index]
# keys = default,butcher shop,farm market
#
# [campaign:default]
# subject1=...
# body1=...
# delay1=3
# subject2=...
# body2=...
# delay2=5
# subject3=...
# body3=...
# delay3=10
# send_to_dialer_after=1
# auto_sync_outlook=0
# hourly_campaign_runner=1
#
# [campaign:butcher shop]
# ...same fields...

def _campaign_section_name(key: str) -> str:
    return f"campaign:{(key or '').strip()}"

def _read_campaign_cfg():
    cfg = configparser.ConfigParser()
    cfg.read(CAMPAIGNS_INI, encoding="utf-8")
    return cfg

def list_campaign_keys():
    """Return a sorted list of campaign keys (niche names)."""
    ensure_campaigns_ini()
    cfg = _read_campaign_cfg()
    keys_csv = (cfg.get("index", "keys", fallback="") or "").strip()
    keys = [k.strip() for k in keys_csv.split(",") if k.strip()]
    # also pick up any sections that start with campaign:
    for sec in cfg.sections():
        if sec.lower().startswith("campaign:"):
            k = sec.split(":",1)[1].strip()
            if k and k not in keys:
                keys.append(k)
    # if empty, add "default"
    if not keys:
        keys = ["default"]
    return sorted(keys, key=lambda s: s.lower())

def _write_index(cfg, keys):
    cfg.setdefault("index", {})
    cfg["index"]["keys"] = ",".join(keys)

def load_campaign_by_key(key: str):
    """
    Return (steps, settings) for a named campaign key.
    Falls back to defaults if key/section missing.
    """
    ensure_campaigns_ini()
    cfg = _read_campaign_cfg()
    sec = _campaign_section_name(key)
    if sec not in cfg:
        # backward-compat migrate root single-campaign if present
        try:
            steps, settings = load_campaigns_ini()
            save_campaign_by_key(key, steps, settings)
            cfg = _read_campaign_cfg()
        except Exception:
            pass

    s = cfg[sec] if sec in cfg else {}
    # Read fields
    steps = []
    for i in range(1, 4):
        steps.append({
            "subject": s.get(f"subject{i}", ""),
            "body": s.get(f"body{i}", ""),
            "delay_days": s.get(f"delay{i}", "0"),
        })
    settings = {
        "send_to_dialer_after": s.get("send_to_dialer_after", "1"),
        "auto_sync_outlook": s.get("auto_sync_outlook", "0"),
        "hourly_campaign_runner": s.get("hourly_campaign_runner", "1"),
    }
    # normalize
    steps = normalize_campaign_steps(steps)
    settings = normalize_campaign_settings(settings)
    return steps, settings

def save_campaign_by_key(key: str, steps, settings):
    """Save (or create) a campaign section for this key and update index."""
    ensure_campaigns_ini()
    cfg = _read_campaign_cfg()
    sec = _campaign_section_name(key)
    if sec not in cfg:
        cfg.add_section(sec)

    steps = normalize_campaign_steps(steps)
    settings = normalize_campaign_settings(settings)

    for i, st in enumerate(steps, start=1):
        cfg[sec][f"subject{i}"] = st.get("subject","")
        cfg[sec][f"body{i}"]    = st.get("body","")
        cfg[sec][f"delay{i}"]   = str(st.get("delay_days","0"))

    # settings
    cfg[sec]["send_to_dialer_after"]   = "1" if settings.get("send_to_dialer_after") else "0"
    cfg[sec]["auto_sync_outlook"]      = "1" if settings.get("auto_sync_outlook") else "0"
    cfg[sec]["hourly_campaign_runner"] = "1" if settings.get("hourly_campaign_runner") else "0"

    # index
    keys = list_campaign_keys()
    if key not in keys:
        keys.append(key)
    _write_index(cfg, keys)

    with CAMPAIGNS_INI.open("w", encoding="utf-8") as f:
        cfg.write(f)

def delete_campaign_by_key(key: str):
    """Delete a campaign section and remove it from index; keep at least 'default'."""
    ensure_campaigns_ini()
    cfg = _read_campaign_cfg()
    sec = _campaign_section_name(key)
    if sec in cfg:
        cfg.remove_section(sec)
    keys = [k for k in list_campaign_keys() if k != key]
    if not keys:
        keys = ["default"]
    _write_index(cfg, keys)
    with CAMPAIGNS_INI.open("w", encoding="utf-8") as f:
        cfg.write(f)

def summarize_campaign_for_table(key: str):
    """Return a compact row describing a campaign for the UI table."""
    steps, settings = load_campaign_by_key(key)
    enabled = [i+1 for i, st in enumerate(steps) if st.get("subject") or st.get("body")]
    delays = [str(st.get("delay_days",0)) for st in steps]
    return [
        key,
        ", ".join(map(str, enabled)) or "—",
        " / ".join(delays),
        "Yes" if settings.get("send_to_dialer_after") else "No",
        "Yes" if settings.get("auto_sync_outlook") else "No",
        "Yes" if settings.get("hourly_campaign_runner") else "No",
    ]
# ----------------------------------------------------------------------



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
    """Ensure warm_leads.csv exists under the v2 schema (WARM_V2_FIELDS)."""
    WARM_LEADS_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not WARM_LEADS_PATH.exists():
        with WARM_LEADS_PATH.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(WARM_V2_FIELDS)
        return

    # Migrate any existing file to v2 header
    rows = []
    with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        existing_fields = rdr.fieldnames or []
        rows = list(rdr) if rdr.fieldnames else []

    if existing_fields != WARM_V2_FIELDS:
        with WARM_LEADS_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=WARM_V2_FIELDS)
            w.writeheader()
            for r in rows:
                w.writerow({h: r.get(h, "") for h in WARM_V2_FIELDS})

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
    Safely request column resizing & related features without touching existing bindings.
    Works across multiple tksheet versions by trying features one-by-one.
    """
    wanted = (
        "column_width_resize",   # tksheet 6.x
        "column_resize",         # older alias
        "resize_columns",        # older alias
        "drag_select",
        "column_drag_and_drop",
    )
    # Try to enable in one call first
    try:
        sheet_obj.enable_bindings(wanted)
        return
    except Exception:
        pass
    # Fall back: try each flag individually
    for fl in wanted:
        try:
            sheet_obj.enable_bindings((fl,))
        except Exception:
            pass

# ============================================================
# Keyboard navigation helpers for tksheet
# ============================================================

def _enable_keyboard_nav(sheet_obj):
    """Enable arrow-key navigation and Tab/Shift-Tab movement even when not editing."""
    try:
        total_rows = sheet_obj.get_total_rows()
    except Exception:
        total_rows = 0
    try:
        total_cols = sheet_obj.get_total_columns()
    except Exception:
        # Try headers length fallback
        try:
            total_cols = len(getattr(sheet_obj, "headers", []))
        except Exception:
            total_cols = 0

    def _clamp(v, lo, hi):
        return max(lo, min(hi, v))

    def _current():
        try:
            r, c = sheet_obj.get_currently_selected()
        except Exception:
            r, c = 0, 0
        r = _clamp(r, 0, max(0, total_rows - 1))
        c = _clamp(c, 0, max(0, total_cols - 1))
        return r, c

    def _select(r, c):
        r = _clamp(r, 0, max(0, total_rows - 1))
        c = _clamp(c, 0, max(0, total_cols - 1))
        # Try multiple APIs to set selection
        for fn in (
            getattr(sheet_obj, "select_cell", None),
            getattr(sheet_obj, "set_currently_selected", None),
        ):
            if callable(fn):
                try:
                    fn(r, c)
                    break
                except Exception:
                    pass
        try:
            sheet_obj.see(r, c)
        except Exception:
            pass
        try:
            sheet_obj.refresh()
        except Exception:
            pass

    def _mv(dr, dc):
        r, c = _current()
        nr, nc = r + dr, c + dc
        # Tab-like wrap across columns then rows
        if nc >= total_cols:
            nr += 1
            nc = 0
        elif nc < 0:
            nr -= 1
            nc = max(0, total_cols - 1)
        _select(nr, nc)
        return "break"

    # Bindings on main table widget (MT) and sheet itself for redundancy
    for w in filter(None, (getattr(sheet_obj, "MT", None), sheet_obj)):
        try:
            w.bind("<Left>",  lambda e: _mv(0, -1))
            w.bind("<Right>", lambda e: _mv(0, 1))
            w.bind("<Up>",    lambda e: _mv(-1, 0))
            w.bind("<Down>",  lambda e: _mv(1, 0))
            # Tab navigation regardless of edit state
            w.bind("<Tab>",        lambda e: _mv(0, 1))
            w.bind("<ISO_Left_Tab>", lambda e: _mv(0, -1))  # Shift+Tab on some platforms
            w.bind("<Shift-Tab>",  lambda e: _mv(0, -1))
        except Exception:
            pass

# ===== CHUNK 1 / 7 — END =====
# ===== CHUNK 2 / 7 — START =====
# ============================================================
# GUI
# ============================================================

def main():
    print(">>> ENTERING main()")
    ensure_app_files()
    templates, subjects, mapping = load_templates_ini()

    # ---- Campaigns: load keys and default campaign (niche/industry) ----
    try:
        campaign_keys = list_campaign_keys()
    except Exception:
        campaign_keys = ["default"]
    current_campaign_key = (campaign_keys[0] if campaign_keys else "default")

    try:
        camp_steps, camp_settings = load_campaign_by_key(current_campaign_key)
    except Exception:
        camp_steps, camp_settings = ([], {})
    camp_steps = normalize_campaign_steps(camp_steps)
    camp_settings = normalize_campaign_settings(camp_settings)

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
        sg.Button("Update", key="-UPDATE-", button_color=("white", "#444444"))
    ]

    # ================== SCOREBOARD WIDGETS ==================
    # ---- DAILY ACTIVITY SCOREBOARD (2 columns) ----
    da_header = [[sg.Text("DAILY ACTIVITY TRACKER",
                          text_color="#9EE493",
                          font=("Segoe UI", 16, "bold"),
                          justification="center",
                          expand_x=True)]]
    da_left = [
        [sg.Text("CALLS:",          text_color="#CCCCCC"), sg.Text("0",     key="-DA_CALLS-",  text_color="#A0FFA0")],
        [sg.Text("EMAILS:",         text_color="#CCCCCC"), sg.Text("0",     key="-DA_EMAILS-", text_color="#A0FFA0")],
        [sg.Text("NEW WARM LEADS:", text_color="#CCCCCC"), sg.Text("0",     key="-DA_WARMS-",  text_color="#A0FFA0")],
    ]
    da_right = [
        [sg.Text("NEW ACCOUNTS:",   text_color="#CCCCCC"), sg.Text("0",     key="-DA_NEWCUS-", text_color="#A0FFA0")],
        [sg.Text("DAILY SALES:",    text_color="#CCCCCC"), sg.Text("$0.00", key="-DA_SALES-",  text_color="#A0FFA0")],
    ]
    daily_scoreboard = sg.Frame(
        "",
        da_header + [[sg.Column(da_left, pad=(6, 6)), sg.Text("   "), sg.Column(da_right, pad=(6, 6))]],
        relief=sg.RELIEF_GROOVE, border_width=2,
        background_color="#1B1B1B", title_color="#9EE493",
        expand_x=False, expand_y=False
    )

    # ---- MONTHLY RESULTS SCOREBOARD (ONE COLUMN) ----
    mo_header = [[sg.Text("MONTHLY RESULTS",
                          text_color="#9EE493",
                          font=("Segoe UI", 16, "bold"),
                          justification="center",
                          expand_x=True)]]
    mo_col = [
        [sg.Text("NEW WARM LEADS:", text_color="#CCCCCC"), sg.Text("0",     key="-MO_WARMS-",  text_color="#A0FFA0")],
        [sg.Text("NEW CUSTOMERS:",  text_color="#CCCCCC"), sg.Text("0",     key="-MO_NEWCUS-", text_color="#A0FFA0")],
        [sg.Text("TOTAL SALES:",    text_color="#CCCCCC"), sg.Text("$0.00", key="-MO_SALES-",  text_color="#A0FFA0")],
    ]
    monthly_scoreboard = sg.Frame(
        "",
        mo_header + [[sg.Column(mo_col, pad=(6, 6))]],
        relief=sg.RELIEF_GROOVE, border_width=2,
        background_color="#1B1B1B", title_color="#9EE493",
        expand_x=False, expand_y=False
    )
    # ================== /SCOREBOARD WIDGETS ==================

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
        sg.Button("Fire the Death Star", key="-FIRE-", size=(25, 2), disabled=True, button_color=("white", "#700000")),
        sg.Text(" (disabled: add valid NEW leads)", key="-FIRE_HINT-", text_color="#BBBBBB")
    ]

    leads_tab = [
        [leads_host],
        [sg.Text("Columns / placeholders:", text_color="#CCCCCC")],
        [sg.Text(", ".join(HEADER_FIELDS), text_color="#9EE493", font=("Consolas", 9))],
        [sg.Column([leads_buttons_row1], pad=(0, 0))],
        [sg.Column([leads_buttons_row2], pad=(0, 0))],
    ]

    # -------- Email Campaigns tab (empty state + editor + saved list) --------
    def _step_row(i, step):
        body_h = 6 if i == 1 else 5
        return [
            [sg.Text(f"Step {i}", text_color="#CCCCCC")],
            [sg.Column([[sg.Text("Subject", text_color="#9EE493")],
                        [sg.Input(default_text=step.get("subject", ""), key=f"-CAMP_SUBJ_{i}-", size=(48, 1), enable_events=True)]], pad=(0, 0)),
             sg.Text("   "),
             sg.Column([[sg.Text("Body", text_color="#9EE493")],
                        [sg.Multiline(default_text=step.get("body", ""), key=f"-CAMP_BODY_{i}-",
                                      size=(90, body_h), font=("Consolas", 10),
                                      text_color="#EEE", background_color="#111", enable_events=True)]], pad=(0, 0), expand_x=True)],
            [sg.Text("Delay (days) after previous send:", text_color="#CCCCCC"),
             sg.Input(default_text=str(step.get("delay_days", "0")), key=f"-CAMP_DELAY_{i}-", size=(6, 1), enable_events=True)],
            [sg.HorizontalSeparator(color="#333333")]
        ]

    camp_rows = []
    for idx in range(1, 4):
        step = camp_steps[idx-1] if idx-1 < len(camp_steps) else {"subject": "", "body": "", "delay_days": "0"}
        camp_rows += _step_row(idx, step)

    send_to_dialer_default = str(camp_settings.get("send_to_dialer_after", "1")).strip() in ("1", "true", "yes", "on")

    # Helper: does any saved campaign actually have content?
    def _any_populated_campaign():
        try:
            for k in campaign_keys:
                stps, _st = load_campaign_by_key(k)
                stps = normalize_campaign_steps(stps)
                if any((s.get("subject") or s.get("body")) for s in stps):
                    return True
        except Exception:
            pass
        return False

    any_campaigns_populated = _any_populated_campaign()

    # ---- NEW: compute response rate per campaign (by matching subjects) ----
    def _campaign_response_rate_for_key(key):
        try:
            stps, _st = load_campaign_by_key(key)
            stps = normalize_campaign_steps(stps)
        except Exception:
            return ""
        subjects_set = {(s.get("subject", "") or "").strip()
                        for s in stps
                        if (s.get("subject", "") or "").strip()}
        if not subjects_set:
            return ""
        try:
            rows = load_results_rows_sorted()
        except Exception:
            return ""
        sent = 0
        replied = 0
        for r in rows:
            subj = (r.get("Subject", "") or "").strip()
            if subj in subjects_set:
                if r.get("DateSent"):
                    sent += 1
                if r.get("DateReplied"):
                    replied += 1
        if sent == 0:
            return "0.0%"
        return f"{(replied / sent) * 100:.1f}%"

    # Active campaigns table data (may be empty) + add Resp %
    try:
        _camp_table_rows_base = [summarize_campaign_for_table(k) for k in campaign_keys]
    except Exception:
        _camp_table_rows_base = []
    _camp_table_rows = []
    for i, k in enumerate(campaign_keys):
        try:
            base = _camp_table_rows_base[i] if i < len(_camp_table_rows_base) else [k]
        except Exception:
            base = [k]
        resp = _campaign_response_rate_for_key(k)
        _camp_table_rows.append(base + [resp])

    # Empty state header (shown when no campaigns yet)
    empty_state = [
        [sg.Text("No campaigns available yet.", text_color="#CCCCCC", key="-CAMP_EMPTY_MSG-")],
        [sg.Button("➕  Add New Campaign", key="-CAMP_ADD_NEW-", button_color=("white", "#2E7D32"))]
    ]

    # Editor header (hidden until user clicks Add New or loads one)
    editor_header = [
        [sg.Text("Campaign niche / industry:", text_color="#9EE493"),
         sg.Combo(values=campaign_keys, default_value=current_campaign_key, key="-CAMP_KEY-", size=(28, 1), enable_events=True),
         sg.Button("New",  key="-CAMP_NEW-"),
         sg.Button("Load", key="-CAMP_LOAD-"),
         sg.Button("Save Campaign", key="-CAMP_SAVE-", button_color=("white", "#2E7D32")),
         sg.Button("Delete This Campaign", key="-CAMP_DELETE-", button_color=("white", "#8B0000")),
         sg.Push(),
         sg.Text("", key="-CAMP_STATUS-", text_color="#A0FFA0")]
    ]

    # Compose the campaigns tab content we will make scrollable
    campaigns_tab_content = [
        [sg.Text("Email Campaigns let you schedule up to 3 follow-ups. Drafts are created based on the *DateSent* in Email Results, not draft time.", text_color="#CCCCCC")],
        [sg.Text("If a prospect replies or is marked Green on Email Results, they’re removed from the campaign automatically.", text_color="#AAAAAA")],

        # Empty state section
        [sg.Column(empty_state, key="-CAMP_EMPTY_WRAP-", visible=not any_campaigns_populated)],

        # Editor section
        [sg.Column(
            editor_header
            + camp_rows
            + [[sg.Checkbox("Send to Dialer automatically if they complete the campaign without replying",
                            key="-CAMP_SEND_TO_DIALER-", default=send_to_dialer_default, text_color="#EEEEEE")],
               [sg.Button("Reset to Defaults", key="-CAMP_RESET-"),
                sg.Button("Reload", key="-CAMP_RELOAD-")]],
            key="-CAMP_EDITOR_WRAP-", visible=any_campaigns_populated, expand_x=True
        )],

        [sg.HorizontalSeparator(color="#4CAF50")],
        [sg.Text("Saved Campaigns", text_color="#9EE493")],
        [sg.Table(values=_camp_table_rows,
                  headings=["Campaign", "Enabled Steps", "Delays (days)", "To Dialer", "Auto Sync", "Hourly Runner", "Resp %"],
                  auto_size_columns=False, col_widths=[24, 14, 18, 10, 10, 14, 8], justification="left", num_rows=8,
                  key="-CAMP_TABLE-", enable_events=True, alternating_row_color="#2a2a2a",
                  text_color="#EEE", background_color="#111", header_text_color="#FFF", header_background_color="#333")],
        [sg.Button("Refresh List", key="-CAMP_REFRESH_LIST-")]
    ]

    # >>> SCROLLABLE wrapper for the entire Email Campaigns tab <<<
    campaigns_tab = [
        [sg.Column(
            campaigns_tab_content,
            size=(980, 620),
            scrollable=True,
            vertical_scroll_only=True,
            expand_x=True,
            expand_y=True,
        )]
    ]

    # -------- Email Results tab --------
    def results_table_data():
        rows = load_results_rows_sorted()
        data = [[r.get("Ref", ""), r.get("Email", ""), r.get("Company", ""), r.get("Industry", ""),
                 r.get("DateSent", ""), r.get("DateReplied", ""), r.get("Status", ""), r.get("Subject", "")] for r in rows]
        return rows, data

    rs_rows, rs_data = results_table_data()
    results_tab = [
        [sg.Text("Sync replies from Outlook; tag Green (good), Gray (neutral), Red (negative).", text_color="#CCCCCC")],
        [sg.Text("Lookback days:", text_color="#CCCCCC"), sg.Input("60", key="-LOOKBACK-", size=(6, 1)),
         sg.Button("Sync from Outlook", key="-SYNC-"),
         sg.Checkbox("Auto Sync (hourly)", key="-AUTO_SYNC-", default=False, text_color="#EEEEEE"),
         sg.Text("", key="-RS_STATUS-", text_color="#A0FFA0")],
        [sg.Table(values=rs_data, headings=["Ref", "Email", "Company", "Industry", "DateSent", "DateReplied", "Status", "Subject"],
                  auto_size_columns=False, col_widths=[10, 26, 26, 14, 18, 18, 8, 40], justification="left", num_rows=15,
                  key="-RSTABLE-", enable_events=True, alternating_row_color="#2a2a2a",
                  text_color="#EEE", background_color="#111", header_text_color="#FFF", header_background_color="#333")],
        [sg.Button("Mark Green", key="-MARK_GREEN-", button_color=("white", "#2E7D32")),
         sg.Button("Mark Gray",  key="-MARK_GRAY-",  button_color=("black", "#DDDDDD")),
         sg.Button("Mark Red",   key="-MARK_RED-",   button_color=("white", "#C62828")),
         sg.Text("   Warm Leads:", text_color="#A0A0A0"), sg.Text("0", key="-WARM-", text_color="#9EE493"),
         sg.Text("   Replies:", text_color="#A0A0A0"), sg.Text("0 / 0", key="-REPLRATE-", text_color="#FFFFFF")]
    ]

    # -------- Dialer tab (grid + outcome dots + notes) --------
    ensure_dialer_files()
    ensure_no_interest_file()

    # use 🙁 so emoji sizes are uniform
    DIALER_EXTRA_COLS = ["🙂", "😐", "🙁", "Note1", "Note2", "Note3", "Note4", "Note5", "Note6", "Note7", "Note8"]
    DIALER_HEADERS = HEADER_FIELDS + DIALER_EXTRA_COLS

    try:
        dialer_matrix = load_dialer_leads_matrix()
    except Exception:
        dialer_matrix = load_dialer_matrix_from_email_csv()
        if not dialer_matrix:
            dialer_matrix = [[""] * len(HEADER_FIELDS) for _ in range(50)]
        dialer_matrix = [row + ["○", "○", "○"] + ([""] * 8) for row in dialer_matrix]

    if len(dialer_matrix) < 100:
        dialer_matrix += [[""] * len(HEADER_FIELDS) + ["○", "○", "○"] + ([""] * 8)
                          for _ in range(100 - len(dialer_matrix))]

    dialer_host = sg.Frame(
        "DIALER GRID",
        [[sg.Text("Loading dialer grid…", key="-DIAL_LOAD-", text_color="#9EE493")]],
        relief=sg.RELIEF_GROOVE, border_width=2, background_color="#1B1B1B",
        title_color="#9EE493", expand_x=True, expand_y=True, key="-DIAL_HOST-",
    )

    dialer_controls_right = [
        [sg.Text("Outcome:", text_color="#CCCCCC")],
        [sg.Button("🟢 Green", key="-DIAL_SET_GREEN-", button_color=("white", "#2E7D32"), size=(14, 1))],
        [sg.Button("⚪ Gray",  key="-DIAL_SET_GRAY-",  button_color=("black", "#DDDDDD"), size=(14, 1))],
        [sg.Button("🔴 Red",   key="-DIAL_SET_RED-",   button_color=("white", "#C62828"), size=(14, 1))],
        [sg.Text("Note:", text_color="#CCCCCC", pad=((0, 0), (10, 0)))],
        [sg.Multiline(key="-DIAL_NOTE-", size=(28, 6), font=("Consolas", 10), background_color="#111", text_color="#EEE")],
        [sg.Button("Confirm Call", key="-DIAL_CONFIRM-", size=(16, 2), disabled=True, button_color=("white", "#444444"))],
        [sg.Text("", key="-DIAL_MSG-", text_color="#A0FFA0", size=(28, 2))]
    ]

    dialer_buttons_under = [
        sg.Button("Add 100 Rows", key="-DIAL_ADD100-"),
    ]

    dialer_tab = [
        [sg.Column([[dialer_host],
                    [sg.Column([dialer_buttons_under], pad=(0, 0))]],
                   expand_x=True, expand_y=True),
         sg.Column(dialer_controls_right, vertical_alignment="top", pad=((10, 0), (0, 0)))]
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
        [sg.Button("🟢 Green", key="-WARM_SET_GREEN-", button_color=("white", "#2E7D32"), size=(14, 1))],
        [sg.Button("⚪ Gray",  key="-WARM_SET_GRAY-",  button_color=("black", "#DDDDDD"), size=(14, 1))],
        [sg.Button("🔴 Red",   key="-WARM_SET_RED-",   button_color=("white", "#C62828"), size=(14, 1))],
        [sg.Text("Note:", text_color="#CCCCCC", pad=((0, 0), (10, 0)))],
        [sg.Multiline(key="-WARM_NOTE-", size=(28, 6), font=("Consolas", 10), background_color="#111", text_color="#EEE")],
        [sg.Button("Confirm", key="-WARM_CONFIRM-", size=(16, 2), disabled=True, button_color=("white", "#444444"))],
        [sg.Text("", key="-WARM_STATUS_SIDE-", text_color="#A0FFA0", size=(28, 2))],
    ]

    warm_buttons_under = [
        sg.Button("Export Warm Leads CSV", key="-WARM_EXPORT-"),
        sg.Button("Reload Warm", key="-WARM_RELOAD-"),
        sg.Button("Add 100 Rows", key="-WARM_ADD100-"),
        sg.Button("→ Confirm New Customer", key="-WARM_MARK_CUSTOMER-", button_color=("white", "#2E7D32")),
        sg.Text("", key="-WARM_STATUS-", text_color="#A0FFA0"),
    ]

    warm_tab = [
        [sg.Column([[warm_host],
                    [sg.Column([warm_buttons_under], pad=(0, 0))]],
                   expand_x=True, expand_y=True),
         sg.Column(warm_controls_right, vertical_alignment="top", pad=((10, 0), (0, 0)))]
    ]

    # -------- Customers tab (grid + analytics panel on right) --------
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
        sg.Button("Export Customers CSV", key="-CUST_EXPORT-"),
        sg.Button("Reload Customers", key="-CUST_RELOAD-"),
        sg.Button("Add 50 Rows", key="-CUST_ADD50-"),
        sg.Button("Add Order", key="-CUST_ADD_ORDER-", button_color=("white", "#2E7D32")),
        sg.Text("", key="-CUST_STATUS-", text_color="#A0FFA0")
    ]

    an_customer = [
        [sg.Text("CUSTOMER ANALYTICS", text_color="#9EE493")],
        [sg.Text("Total Sales"),  sg.Text("0.00", key="-AN_TOTALSALES-", text_color="#A0FFA0")],
        [sg.Text("CAC"),          sg.Text("0.00", key="-AN_CAC-",        text_color="#A0FFA0")],
        [sg.Text("LTV"),          sg.Text("0.00", key="-AN_LTV-",        text_color="#A0FFA0")],
        [sg.Text("CAC : LTV"),    sg.Text("1 : 0", key="-AN_CACLTV-",    text_color="#A0FFA0")],
        [sg.Text("Reorder Rate"), sg.Text("0%",   key="-AN_REORDER-",    text_color="#A0FFA0")],
    ]

    an_pipeline = [
        [sg.HorizontalSeparator(color="#4CAF50")],
        [sg.Text("PIPELINE ANALYTICS", text_color="#9EE493")],
        [sg.Text("Total Warm Leads"),  sg.Text("0",  key="-AN_WARMS-",     text_color="#A0FFA0")],
        [sg.Text("New Customers"),     sg.Text("0",  key="-AN_NEWCUS-",    text_color="#A0FFA0")],
        [sg.Text("Close Rate"),        sg.Text("0%", key="-AN_CLOSERATE-", text_color="#A0FFA0")],
    ]

    analytics_panel = sg.Frame(
        "",
        [[sg.Column(an_customer, pad=(6, 6), expand_x=True, expand_y=False)],
         [sg.Column(an_pipeline, pad=(6, 0), expand_x=True, expand_y=False)]],
        relief=sg.RELIEF_GROOVE, border_width=2,
        background_color="#1B1B1B", title_color="#9EE493",
        expand_x=False, expand_y=False
    )

    customers_tab = [
        [sg.Column([[customers_host],
                    [sg.Column([customers_buttons_under], pad=(0, 0))]],
                   expand_x=True, expand_y=True),
         sg.Column([[analytics_panel]],
                   vertical_alignment="top",
                   pad=((10, 0), (0, 0)),
                   size=(320, 340))]
    ]

    # -------- Map tab (launches an interactive HTML map) --------
    map_tab = [
        [sg.Text("Customer Map", text_color="#9EE493", font=("Segoe UI", 14, "bold"))],
        [sg.Text("Opens a live Leaflet map with pins for each geocoded customer (Company, CLTV, Sales/Day).",
                 text_color="#CCCCCC")],
        [sg.Button("🗺️ Open Customer Map", key="-OPEN_MAP-", button_color=("white", "#2D6CDF"), size=(24, 2)),
         sg.Text("", key="-MAP_STATUS-", text_color="#A0FFA0")]
    ]

    # -------- Compose full layout --------
    SB_LEFT_PAD = (500, 0)
    SB_TOP_PAD  = (50, 6)

    scoreboards_row = [
        sg.Column(
            [[daily_scoreboard, sg.Text("  "), monthly_scoreboard]],
            pad=(SB_LEFT_PAD, SB_TOP_PAD),
            background_color="#202020",
            expand_x=False, expand_y=False
        )
    ]

    layout = [
        top_bar,
        scoreboards_row,
        [sg.TabGroup([[sg.Tab("Email Leads",     leads_tab,     expand_x=True, expand_y=True),
                       sg.Tab("Email Campaigns", campaigns_tab, expand_x=True, expand_y=True),
                       sg.Tab("Email Results",   results_tab,   expand_x=True, expand_y=True),
                       sg.Tab("Dialer",          dialer_tab,    expand_x=True, expand_y=True),
                       sg.Tab("Warm Leads",      warm_tab,      expand_x=True, expand_y=True),
                       sg.Tab("Customers",       customers_tab, expand_x=True, expand_y=True),
                       sg.Tab("Map",             map_tab,       expand_x=True, expand_y=True)]],
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

    # ========== Email Leads sheet ==========
    host_frame_tk = leads_host.Widget
    for child in host_frame_tk.winfo_children():
        try:
            child.destroy()
        except Exception:
            pass
    sheet_holder = sg.tk.Frame(host_frame_tk, bg="#111111")
    sheet_holder.pack(side="top", fill="both", expand=True)

    def _count_emails_sent_for_addr(addr: str) -> int:
        a = (addr or "").strip().lower()
        if not a:
            return 0
        try:
            rows = load_results_rows_sorted()
        except Exception:
            return 0
        n = 0
        for r in rows:
            if (r.get("Email", "") or "").strip().lower() == a and (r.get("DateSent") or "").strip():
                n += 1
        return min(n, 3)

    existing = []
    if CSV_PATH.exists():
        with CSV_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                existing.append([r.get(h, "") for h in HEADER_FIELDS])

    LEADS_HEADERS_DISPLAY = HEADER_FIELDS + ["Emails Sent"]
    try:
        email_idx = HEADER_FIELDS.index("Email")
    except Exception:
        email_idx = 0

    if existing:
        base = existing + [[""] * len(HEADER_FIELDS) for _ in range(max(0, START_ROWS - len(existing)))]
    else:
        base = [[""] * len(HEADER_FIELDS) for _ in range(START_ROWS)]

    data_display = []
    for row in base:
        try:
            addr = row[email_idx] if email_idx is not None and email_idx < len(row) else ""
        except Exception:
            addr = ""
        sent_n = _count_emails_sent_for_addr(addr)
        data_display.append(list(row) + [str(sent_n)])

    sheet = Sheet(
        sheet_holder,
        data=data_display,
        headers=LEADS_HEADERS_DISPLAY,
        show_x_scrollbar=True,
        show_y_scrollbar=True
    )
    sheet.enable_bindings((
        "single_select",
        "arrowkeys", "tab_key", "shift_tab_key",
        "drag_select", "copy", "cut", "delete", "undo",
        "edit_cell", "return_edit_cell", "select_all",
        "right_click_popup_menu", "column_width_resize", "column_resize", "resize_columns"
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
    for c in range(len(LEADS_HEADERS_DISPLAY)):
        try:
            width = 120 if LEADS_HEADERS_DISPLAY[c] == "Emails Sent" else DEFAULT_COL_WIDTH
            sheet.column_width(c, width=width)
        except Exception:
            pass

    _bind_plaintext_paste_for_tksheet(sheet, window.TKroot)
    _ensure_rc_menu_plain_paste(sheet, window.TKroot)
    _enable_column_resizing(sheet)

    def _apply_emails_sent_coloring(_sheet):
        try:
            total_rows = _sheet.get_total_rows()
        except Exception:
            total_rows = len(data_display)
        emails_col = len(LEADS_HEADERS_DISPLAY) - 1
        for r in range(max(0, total_rows)):
            try:
                v = _sheet.get_cell_data(r, emails_col)
            except Exception:
                v = ""
            try:
                n = int(str(v).strip() or "0")
            except Exception:
                n = 0
            bg = "#FFFFFF"; fg = "#000000"
            if n == 1:
                bg = "#EEEEEE"
            elif n == 2:
                bg = "#CFCFCF"
            elif n >= 3:
                bg = "#A6A6A6"
            try:
                _sheet.highlight_cells(row=r, column=emails_col, bg=bg, fg=fg)
            except Exception:
                pass
        try:
            _sheet.refresh()
        except Exception:
            pass

    _apply_emails_sent_coloring(sheet)

    # ========== Dialer grid ==========
    dial_host_tk = dialer_host.Widget
    for child in dial_host_tk.winfo_children():
        try:
            child.destroy()
        except Exception:
            pass
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
        "single_select",
        "arrowkeys", "tab_key", "shift_tab_key",
        "drag_select", "copy", "cut", "delete", "undo",
        "edit_cell", "return_edit_cell", "select_all",
        "right_click_popup_menu", "column_width_resize", "column_resize", "resize_columns"
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

    def _idx(colname, default=None):
        try:
            return DIALER_HEADERS.index(colname)
        except Exception:
            return default

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
            width = 42
        if first_note <= c <= last_note:
            width = 120
        try:
            dial_sheet.column_width(c, width=width)
        except Exception:
            pass

    try:
        outcome_cols = [first_dot, first_dot + 1, first_dot + 2]
        try:
            dial_sheet.align_columns(columns=outcome_cols, align="center")
        except Exception:
            try:
                for col in outcome_cols:
                    dial_sheet.column_align(col, align="center")
            except Exception:
                try:
                    dial_sheet.set_column_alignments(outcome_cols, align="center")
                except Exception:
                    pass
    except Exception:
        pass

    _bind_plaintext_paste_for_tksheet(dial_sheet, window.TKroot)
    _ensure_rc_menu_plain_paste(dial_sheet, window.TKroot)
    _enable_column_resizing(dial_sheet)

    # ========== Warm grid ==========
    warm_host_tk = warm_host.Widget
    for child in warm_host_tk.winfo_children():
        try:
            child.destroy()
        except Exception:
            pass
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
        "single_select",
        "arrowkeys", "tab_key", "shift_tab_key",
        "drag_select", "copy", "cut", "delete", "undo",
        "edit_cell", "return_edit_cell", "select_all",
        "right_click_popup_menu", "column_width_resize", "column_resize", "resize_columns"
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
        if name in ("Company", "Prospect Name"): width = 180
        if name in ("Phone #", "Rep", "Samples?"): width = 90
        if name in ("Email", "Google Reviews", "Industry", "Location"): width = 160
        if name.endswith("Date"): width = 110
        if name in ("First Contact", "Timestamp", "Cost ($)"): width = 120
        try:
            warm_sheet.column_width(c, width=width)
        except Exception:
            pass

    _bind_plaintext_paste_for_tksheet(warm_sheet, window.TKroot)
    _ensure_rc_menu_plain_paste(warm_sheet, window.TKroot)
    _enable_column_resizing(warm_sheet)

    # ========== Customers grid ==========
    cust_host_tk = customers_host.Widget
    for child in cust_host_tk.winfo_children():
        try:
            child.destroy()
        except Exception:
            pass
    cust_holder = sg.tk.Frame(customers_host.Widget, bg="#111111")
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

    # --- Auto-calc Days & Sales/Day (Days from First Order only; CLTV left blank if empty)
    try:
        from datetime import datetime as _dt

        def _parse_date_local(s):
            """Return date() from many common formats, incl. 2-digit years and MD without year."""
            s = (s or "").strip()
            if not s:
                return None
            fmts = (
                "%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%Y/%m/%d",
                "%m/%d/%y", "%m-%d-%y", "%y-%m-%d",
                "%m/%d", "%m-%d"
            )
            for fmt in fmts:
                try:
                    d = _dt.strptime(s, fmt)
                    if fmt in ("%m/%d", "%m-%d"):
                        d = d.replace(year=_dt.now().year)
                    return d.date()
                except Exception:
                    continue
            return None

        def _money_to_float_local(val):
            s = (val or "").strip().replace(",", "").replace("$", "")
            if not s:
                return 0.0
            try:
                return float(s)
            except Exception:
                return 0.0

        def _float_to_money_local(x):
            try:
                return f"{float(x):.2f}"
            except Exception:
                return ""

        idx = {name: i for i, name in enumerate(CUSTOMER_FIELDS)}
        i_company   = idx.get("Company")
        i_first     = idx.get("First Order")
        i_last      = idx.get("Last Order")
        i_cltv      = idx.get("CLTV")
        i_days      = idx.get("Days")
        i_salesday  = idx.get("Sales/Day")

        today = _dt.now().date()

        for row in customers_matrix:
            company = (row[i_company] if i_company is not None and i_company < len(row) else "").strip()
            if not company:
                continue

            # 1) Refresh CLTV / First / Last from orders.csv (if present)
            try:
                stats = compute_customer_order_stats(company)
            except Exception:
                stats = None

            if stats:
                if stats.get("first_order_date") and i_first is not None:
                    row[i_first] = stats["first_order_date"].strftime("%Y-%m-%d")
                if stats.get("last_order_date") and i_last is not None:
                    row[i_last] = stats["last_order_date"].strftime("%Y-%m-%d")
                if i_cltv is not None:
                    cltv_from_orders = float(stats.get("cltv", 0.0) or 0.0)
                    if cltv_from_orders > 0:
                        row[i_cltv] = _float_to_money_local(cltv_from_orders)

            # 2) Compute Days from First Order ONLY
            first_dt = _parse_date_local(row[i_first] if i_first is not None else "")
            if first_dt and i_days is not None:
                days = max(1, (today - first_dt).days)
                row[i_days] = str(days)
            else:
                days = None
                if i_days is not None:
                    row[i_days] = ""

            # 3) Compute Sales/Day strictly from CLTV ÷ Days
            if i_salesday is not None:
                if days:
                    cltv_raw = (row[i_cltv] if i_cltv is not None else "")
                    cltv_val = _money_to_float_local(cltv_raw)
                    row[i_salesday] = _float_to_money_local(cltv_val / float(days)) if (days and cltv_val > 0) else ""
                else:
                    row[i_salesday] = ""

            # 4) Ensure CLTV cell is money-formatted ONLY if non-empty
            if i_cltv is not None:
                raw = (row[i_cltv] or "").strip()
                if raw != "":
                    row[i_cltv] = _float_to_money_local(_money_to_float_local(raw))

    except Exception:
        pass

    # ---- Mount CustomerSheet
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
        "single_select",
        "arrowkeys", "tab_key", "shift_tab_key",
        "drag_select", "copy", "cut", "delete", "undo",
        "edit_cell", "return_edit_cell", "select_all",
        "right_click_popup_menu", "column_width_resize", "column_resize", "resize_columns"
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
        if name in ("Company", "Prospect Name"): width = 180
        if name in ("Phone #", "Rep", "Samples?"): width = 90
        if name in ("Email", "Google Reviews", "Industry", "Location"): width = 160
        if name in ("Opening Order $", "Customer Since", "Timestamp"): width = 150
        try:
            customer_sheet.column_width(c, width=width)
        except Exception:
            pass

    # Freeze the first column (Company) – try several APIs for different tksheet versions
    _froze_ok = False
    for fn_name, arg in (
        ("freeze_columns", 1),        # modern API
        ("set_frozen_columns", 1),    # alt API
        ("freeze", (1, 0)),           # some builds expect (cols, rows)
    ):
        try:
            fn = getattr(customer_sheet, fn_name, None)
            if callable(fn):
                if isinstance(arg, tuple):
                    fn(*arg)
                else:
                    fn(arg)
                _froze_ok = True
                break
        except Exception:
            pass
    if not _froze_ok:
        try:
            customer_sheet.set_options(frozen_columns=1)  # legacy fallback
        except Exception:
            pass

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

# ===== CHUNK 2 / 7 — END =====
# ===== CHUNK 3 / 7 — START =====
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
            hdr = HEADER_FIELDS + ["🙂","😐","🙁"] + [f"Note{i}" for i in range(1,9)]
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
    expected = HEADER_FIELDS + ["🙂","😐","🙁"] + [f"Note{i}" for i in range(1,9)]
    idx_map = [hdr.index(h) if h in hdr else None for h in expected]
    for row in raw[1:]:
        out = []
        for i, idx in enumerate(idx_map):
            if idx is None:
                if len(HEADER_FIELDS) <= i < len(HEADER_FIELDS)+3:
                    out.append("○")
                else:
                    out.append("")
            else:
                out.append(row[idx] if idx < len(row) else "")
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
    """
    Ensure customers.csv exists with the UPDATED schema.
    If it exists but headers differ, migrate by mapping any matching columns by name.
    """
    if not CUSTOMERS_PATH.exists():
        _backup(CUSTOMERS_PATH)
        with CUSTOMERS_PATH.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(CUSTOMER_FIELDS)
        return

    # If exists, check header and migrate if needed
    try:
        with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            old_fields = rdr.fieldnames or []
            if old_fields != CUSTOMER_FIELDS:
                rows = list(rdr)  # read existing
            else:
                rows = None
    except Exception:
        rows = None

    if rows is not None:
        # Build migrated matrix aligned to CUSTOMER_FIELDS
        migrated = []
        for r in rows:
            migrated.append([r.get(h, "") for h in CUSTOMER_FIELDS])
        _backup(CUSTOMERS_PATH)
        _atomic_write_csv(CUSTOMERS_PATH, CUSTOMER_FIELDS, migrated)


# --- Derived fields for customers (Days, Sales/Day, normalized CLTV) ---

def _derive_customer_fields(row_dict: dict):
    """Compute Days and Sales/Day from First Order and CLTV. Returns updates dict (strings)."""
    # Parse First Order date (supports YYYY-MM-DD, MM/DD/YYYY, etc.)
    first_order_str = (row_dict.get("First Order", "") or "").strip()
    first_dt = _parse_date_mmddyyyy(first_order_str)

    # CLTV as float (handles $, commas)
    cltv_f = _money_to_float(row_dict.get("CLTV", ""))

    days_val = None
    sales_per_day = None
    if first_dt:
        try:
            days_val = max(1, (datetime.now().date() - first_dt).days)
        except Exception:
            days_val = None
    if days_val:
        sales_per_day = cltv_f / float(days_val) if days_val else None

    updates = {}
    # Keep CLTV normalized to money string (so file stays consistent)
    updates["CLTV"] = _float_to_money(cltv_f) if cltv_f is not None else row_dict.get("CLTV", "")
    if days_val is not None:
        updates["Days"] = str(int(days_val))
    if sales_per_day is not None:
        updates["Sales/Day"] = _float_to_money(sales_per_day)
    return updates


def load_customers_matrix():
    """Load customers.csv into a matrix matching CUSTOMER_FIELDS, deriving missing fields on the fly."""
    ensure_customers_file()
    rows = []
    with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        if not rdr.fieldnames:
            return rows
        for r in rdr:
            # Derive values for display if missing or stale (non-destructive; doesn't write back here)
            try:
                derived = _derive_customer_fields(r)
                r = {**r, **{k: (derived.get(k) or r.get(k, "")) for k in ("CLTV","Days","Sales/Day")}}
            except Exception:
                pass
            rows.append([r.get(h, "") for h in CUSTOMER_FIELDS])
    return rows


def save_customers_matrix(matrix):
    """Save matrix to customers.csv with backup + atomic replace, recomputing derived fields per row."""
    ensure_customers_file()
    # Convert to dicts, recompute derived fields, then write
    out_rows = []
    for row in matrix:
        rd = {h: (row[i] if i < len(row) else "") for i, h in enumerate(CUSTOMER_FIELDS)}
        try:
            rd.update(_derive_customer_fields(rd))
        except Exception:
            pass
        out_rows.append([rd.get(h, "") for h in CUSTOMER_FIELDS])
    _backup(CUSTOMERS_PATH)
    _atomic_write_csv(CUSTOMERS_PATH, CUSTOMER_FIELDS, out_rows)


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
            # Also derive dependent fields based on new values
            try:
                r.update(_derive_customer_fields(r))
            except Exception:
                pass
            found = True
            break

    if not found:
        # If company not found, append a new row with only provided fields.
        new_row = {h:"" for h in CUSTOMER_FIELDS}
        new_row["Company"] = company or ""
        for k,v in updates.items():
            if k in CUSTOMER_FIELDS:
                new_row[k] = v
        try:
            new_row.update(_derive_customer_fields(new_row))
        except Exception:
            pass
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
    # Prepare updates (align to CUSTOMER_FIELDS: "First Order", "Last Order")
        # Prepare updates
    updates = {}
    if stats["first_order_date"]:
        updates["First Order"] = stats["first_order_date"].strftime("%Y-%m-%d")   # <-- was First Order Date
    if stats["last_order_date"]:
        updates["Last Order"] = stats["last_order_date"].strftime("%Y-%m-%d")     # <-- was Last Order Date
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
# Campaigns helpers (per-ref state via campaigns.csv)
# (Legacy single-INI campaign schema removed)
# ============================================================

# Files
CAMPAIGNS_PATH = APP_DIR / "campaigns.csv"     # per-ref state

CAMPAIGNS_HEADERS = ["Ref","Email","Company","CampaignKey","Stage","DivertToDialer"]


def ensure_campaigns_file():
    if not CAMPAIGNS_PATH.exists():
        with CAMPAIGNS_PATH.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(CAMPAIGNS_HEADERS)


def _read_campaign_rows():
    ensure_campaigns_file()
    rows = []
    with CAMPAIGNS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            rows.append(r)
    return rows


def _campaigns_write_rows(rows):
    """Writer used by Chunk 7 helpers."""
    _backup(CAMPAIGNS_PATH)
    with CAMPAIGNS_PATH.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CAMPAIGNS_HEADERS)
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h,"") for h in CAMPAIGNS_HEADERS})

# Back-compat alias (if any code still referenced the old name)
_write_campaign_rows = _campaigns_write_rows


def upsert_campaign_row(ref_short, email, company, campaign_key, stage=0, divert_to_dialer=0):
    rows = _read_campaign_rows()
    ref_l = (ref_short or "").lower()
    found = False
    for r in rows:
        if (r.get("Ref","") or "").lower() == ref_l:
            r["Email"] = email or r.get("Email","")
            r["Company"] = company or r.get("Company","")
            r["CampaignKey"] = campaign_key or r.get("CampaignKey","default")
            r["Stage"] = str(stage)
            r["DivertToDialer"] = "1" if int(divert_to_dialer or 0) else "0"
            found = True
            break
    if not found:
        rows.append({
            "Ref": ref_short, "Email": email or "", "Company": company or "",
            "CampaignKey": campaign_key or "default", "Stage": str(stage),
            "DivertToDialer": "1" if int(divert_to_dialer or 0) else "0"
        })
    _campaigns_write_rows(rows)


def remove_campaign_by_ref(ref_short):
    rows = _read_campaign_rows()
    ref_l = (ref_short or "").lower()
    rows = [r for r in rows if (r.get("Ref","") or "").lower() != ref_l]
    _campaigns_write_rows(rows)


def get_campaign_row(ref_short):
    ref_l = (ref_short or "").lower()
    for r in _read_campaign_rows():
        if (r.get("Ref","") or "").lower() == ref_l:
            return r
    return None


def set_campaign_stage(ref_short, new_stage):
    rows = _read_campaign_rows()
    ref_l = (ref_short or "").lower()
    for r in rows:
        if (r.get("Ref","") or "").lower() == ref_l:
            r["Stage"] = str(new_stage)
            break
    _campaigns_write_rows(rows)


# ============================================================
# Outlook helpers (extended for single draft by ref)
# ============================================================


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


def outlook_draft_one(row_dict, subject_text, body_text, ref_short):
    """Create a single Outlook draft to row_dict['Email'] with given subject/body and [ref:xxxx]."""
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

    body_html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"></head>
    <body style="margin:0;padding:0;"><div style="font-family:Segoe UI, Arial, sans-serif; font-size:14px; line-height:1.5; color:#111;">
    {blocks_to_html(body_text)}<!-- ref:{ref_short} --></div></body></html>"""
    msg = drafts_root.Items.Add("IPM.Note")
    msg.To = row_dict.get("Email","")
    msg.Subject = f"{subject_text} [ref:{ref_short}]"
    msg.BodyFormat = 2
    msg.HTMLBody = body_html
    msg.Save()
    msg.Move(target_folder)


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

# Remainder (state/results/sync) from original:
REF_RE = re.compile(r"\[ref:([0-9a-f]{6,12})\]", re.IGNORECASE)


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


def _results_lookup_by_ref():
    """Quick map: ref_lower -> row dict (results.csv)."""
    rows = load_results_rows_sorted()
    return { (r.get("Ref","") or "").lower(): r for r in rows }


def _results_dates_for_ref(ref_short):
    """Return (sent_dt, replied_dt) as datetime or (None,None)."""
    r = _results_lookup_by_ref().get((ref_short or "").lower())
    def _p(s):
        s = (s or "").strip()
        if not s: return None
        dt = None
        # try chunk 4's parser later; quick attempt here:
        for fmt in ("%Y-%m-%d %H:%M:%S","%Y-%m-%d",
                    "%m/%d/%Y %I:%M:%S %p","%m/%d/%Y %I:%M %p","%m/%d/%Y"):
            try:
                return datetime.strptime(s, fmt)
            except Exception:
                pass
        return None
    if not r:
        return (None, None)
    return (_p(r.get("DateSent","")), _p(r.get("DateReplied","")))


def _lead_row_from_email_company(email, company):
    """Try to find an original lead row in kybercrystals.csv for placeholders."""
    if not CSV_PATH.exists():
        return None
    email_l = (email or "").strip().lower()
    comp_l  = (company or "").strip().lower()
    with CSV_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            if email_l and (r.get("Email","").strip().lower() == email_l):
                return r
        if comp_l:
            f.seek(0); next(rdr, None)  # rewind
            for r in rdr:
                if (r.get("Company","").strip().lower() == comp_l):
                    return r
    return None


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
    # Stamp last sync time for Daily Activity popup
    try:
        LAST_SYNC_PATH.write_text(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), encoding="utf-8")
    except Exception:
        pass
    return len(sent_map), len(reply_map)

# ===== CHUNK 3 / 7 — END =====
# ===== CHUNK 4 / 7 — START =====
# ============================================================
# Shared helpers (no hard dependency on live tksheet objects)
# ============================================================
from datetime import datetime as _dt

# --- Sentinels to prevent NameError at import time ---
try:
    sheet
except NameError:
    sheet = None
try:
    dial_sheet
except NameError:
    dial_sheet = None
try:
    warm_sheet
except NameError:
    warm_sheet = None
try:
    customer_sheet
except NameError:
    customer_sheet = None
# ----------------------------------------------------

# File used to display "Email data last synced" in the Daily Activity view
LAST_SYNC_PATH = APP_DIR / "last_outlook_sync.txt"

def _today_date():
    return _dt.now().date()

def _fmt_money(x):
    try:
        return f"{float(x):.2f}"
    except Exception:
        return "0.00"

def _parse_any_datetime(s):
    """Best-effort parse across formats used in results, dialer, warm, orders."""
    s = (s or "").strip()
    if not s:
        return None
    fmts = [
        "%Y-%m-%d %H:%M:%S",     # our own writes
        "%Y-%m-%d",              # ISO date
        "%m/%d/%Y %I:%M:%S %p",  # Outlook strings
        "%m/%d/%Y %I:%M %p",     # Outlook alt (no seconds)
        "%m/%d/%Y",              # date only
    ]
    for fmt in fmts:
        try:
            return _dt.strptime(s, fmt)
        except Exception:
            pass
    alt = s.replace("-", "/")
    if alt != s:
        for fmt in ("%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M %p", "%m/%d/%Y"):
            try:
                return _dt.strptime(alt, fmt)
            except Exception:
                pass
    return None

def _read_last_sync_str():
    try:
        if LAST_SYNC_PATH.exists():
            return LAST_SYNC_PATH.read_text(encoding="utf-8").strip()
    except Exception:
        pass
    return "—"

# ============================================================
# Daily Activity computation (reads CSVs only)
# ============================================================
def compute_daily_activity(target_date=None):
    """Return dict with metrics for the given date (default: today)."""
    d = target_date or _today_date()

    # Calls (dialer_results.csv)
    calls_total = calls_green = calls_gray = calls_red = 0
    try:
        if DIALER_RESULTS_PATH.exists():
            with DIALER_RESULTS_PATH.open("r", encoding="utf-8", newline="") as f:
                rdr = csv.DictReader(f)
                for r in rdr:
                    ts = _parse_any_datetime(r.get("Timestamp",""))
                    if ts and ts.date() == d:
                        calls_total += 1
                        oc = (r.get("Outcome","") or "").strip().lower()
                        if oc == "green": calls_green += 1
                        elif oc == "gray": calls_gray += 1
                        elif oc == "red":  calls_red  += 1
    except Exception:
        pass

    # Emails sent (results.csv) — depends on last Outlook sync
    emails_sent = 0
    try:
        if RESULTS_PATH.exists():
            with RESULTS_PATH.open("r", encoding="utf-8", newline="") as f:
                rdr = csv.DictReader(f)
                for r in rdr:
                    ds = _parse_any_datetime(r.get("DateSent",""))
                    if ds and ds.date() == d:
                        emails_sent += 1
    except Exception:
        pass

    # New Warm Leads (warm_leads.csv): prefer 'First Contact', else legacy 'Timestamp'
    new_warm = 0
    try:
        if WARM_LEADS_PATH.exists():
            with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
                rdr = csv.DictReader(f)
                fc_field = "First Contact" if "First Contact" in (rdr.fieldnames or []) else "Timestamp"
                for r in rdr:
                    fc_val = r.get(fc_field, "")
                    dt = _parse_any_datetime(fc_val)
                    if dt and dt.date() == d:
                        new_warm += 1
    except Exception:
        pass

    # New Accounts (customers.csv): Customer Since OR First Order Date
    new_accounts = 0
    try:
        if CUSTOMERS_PATH.exists():
            with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
                rdr = csv.DictReader(f)
                for r in rdr:
                    cs = (r.get("Customer Since","") or "").strip()
                    fo = (r.get("First Order Date","") or "").strip()
                    ok = False
                    if cs:
                        dt = _parse_any_datetime(cs)
                        ok = (dt and dt.date() == d)
                    if not ok and fo:
                        dt = _parse_any_datetime(fo)
                        ok = (dt and dt.date() == d)
                    if ok:
                        new_accounts += 1
    except Exception:
        pass

    # Daily Sales (orders.csv) — sum + count
    orders_count = 0
    sales_sum = 0.0
    try:
        ensure_orders_file()
        if ORDERS_PATH.exists():
            with ORDERS_PATH.open("r", encoding="utf-8", newline="") as f:
                rdr = csv.DictReader(f)
                for r in rdr:
                    od = _parse_any_datetime(r.get("Order Date",""))
                    if od and od.date() == d:
                        orders_count += 1
                        try:
                            sales_sum += float(str(r.get("Amount","") or "0").replace(",",""))
                        except Exception:
                            pass
    except Exception:
        pass

    return {
        "date": d.strftime("%Y-%m-%d"),
        "calls_total": calls_total,
        "calls_green": calls_green,
        "calls_gray": calls_gray,
        "calls_red":  calls_red,
        "emails_sent": emails_sent,
        "new_warm": new_warm,
        "new_accounts": new_accounts,
        "orders_count": orders_count,
        "sales_sum": sales_sum,
        "last_sync": _read_last_sync_str(),
    }

# ============================================================
# Daily Activity popup (pure UI)
# ============================================================
def show_daily_activity_popup():
    m = compute_daily_activity()
    header = f"Daily Activity — {m['date']}"
    layout = [
        [sg.Text(header, text_color="#9EE493", font=("Segoe UI", 12, "bold")),
         sg.Push(),
         sg.Text("Email data last synced:", text_color="#CCCCCC"),
         sg.Text(m["last_sync"], key="-DA_LASTSYNC-", text_color="#FFFFFF")],
        [sg.HorizontalSeparator(color="#4CAF50")],
        [sg.Column([
            [sg.Text("Daily Calls:", size=(16,1), text_color="#CCCCCC"),
             sg.Text(f"{m['calls_total']}  (🟢 {m['calls_green']}  ⚪ {m['calls_gray']}  🔴 {m['calls_red']})", key="-DA_CALLS-", text_color="#FFFFFF")],
            [sg.Text("Daily Emails:", size=(16,1), text_color="#CCCCCC"),
             sg.Text(str(m["emails_sent"]), key="-DA_EMAILS-", text_color="#FFFFFF")],
            [sg.Text("New Warm Leads:", size=(16,1), text_color="#CCCCCC"),
             sg.Text(str(m["new_warm"]), key="-DA_NEWWARM-", text_color="#FFFFFF")],
            [sg.Text("New Accounts:", size=(16,1), text_color="#CCCCCC"),
             sg.Text(str(m["new_accounts"]), key="-DA_NEWACCTS-", text_color="#FFFFFF")],
            [sg.Text("Daily Sales:", size=(16,1), text_color="#CCCCCC"),
             sg.Text(f"${_fmt_money(m['sales_sum'])}  ({m['orders_count']} orders)", key="-DA_SALES-", text_color="#A0FFA0")],
        ], pad=(0,0), expand_x=True)],
        [sg.Push(), sg.Button("Refresh", key="-DA_REFRESH-"), sg.Button("Close", key="-DA_CLOSE-")]
    ]
    win = sg.Window("Daily Activity", layout, modal=True, keep_on_top=True, finalize=True)
    while True:
        ev, _vals = win.read(timeout=60000)
        if ev in (sg.WINDOW_CLOSE_ATTEMPTED_EVENT, sg.WIN_CLOSED, "-DA_CLOSE-"):
            break
        if ev == "-DA_REFRESH-":
            mm = compute_daily_activity()
            win["-DA_LASTSYNC-"].update(mm["last_sync"])
            win["-DA_CALLS-"].update(f"{mm['calls_total']}  (🟢 {mm['calls_green']}  ⚪ {mm['calls_gray']}  🔴 {mm['calls_red']})")
            win["-DA_EMAILS-"].update(str(mm["emails_sent"]))
            win["-DA_NEWWARM-"].update(str(mm["new_warm"]))
            win["-DA_NEWACCTS-"].update(str(mm["new_accounts"]))
            win["-DA_SALES-"].update(f"${_fmt_money(mm['sales_sum'])}  ({mm['orders_count']} orders)")
    win.close()


# ============================================================
# Campaign queue processor (draft follow-ups + divert to Dialer)
# ============================================================
def _campaign_get_lead_row_for_ref(crow):
    """Return a dict of header->value for a lead row that matches campaign row."""
    enr = _lead_row_from_email_company(crow.get("Email",""), crow.get("Company",""))
    if not enr:
        # fallback minimal dict
        enr = {h:"" for h in HEADER_FIELDS}
        enr["Email"] = crow.get("Email","")
        enr["Company"] = crow.get("Company","")
    return enr

def _campaign_stage_from_results_if_needed(ref_short, cur_stage):
    """
    If Stage==0 but results.csv already has a DateSent, auto-bump to Stage 1.
    This keeps things consistent when app restarts after sending E1.
    """
    try:
        if int(cur_stage or 0) > 0:
            return int(cur_stage)
    except Exception:
        pass
    # Check results cache for DateSent
    try:
        res_map = _read_results_by_ref()
        r = res_map.get((ref_short or "").strip().lower())
        sent_dt = _results_sent_dt(r) if r else None
        return 1 if sent_dt else 0
    except Exception:
        return int(cur_stage or 0)

def process_campaign_queue():
    """
    Runs quick checks:
      - If a ref has DateReplied -> remove from campaigns.
      - If stage==0 and DateSent exists -> set stage=1.
      - If stage==1 and due and no reply -> draft E2 via _draft_next_stage_stub, stage=2.
      - If stage==2 and due and no reply -> draft E3 via _draft_next_stage_stub, stage=3.
      - If stage==3 and no reply and divert flag -> push to Dialer & remove.
    """
    ensure_campaigns_file()
    rows = _read_campaign_rows()
    changed = False

    # cache results for quick lookups
    try:
        res_map = _read_results_by_ref()
    except Exception:
        res_map = {}

    for r in rows[:]:
        ref = r.get("Ref","")
        key = r.get("CampaignKey","default")
        # CSV field wins if set; otherwise fall back to campaign settings later
        divert_csv = r.get("DivertToDialer", "")
        try:
            stage = int(r.get("Stage","0") or 0)
        except Exception:
            stage = 0

        # reply?
        res = res_map.get((ref or "").strip().lower())
        if res and _results_replied(res):
            # remove from campaign immediately
            rows.remove(r)
            changed = True
            continue

        # bring stage up to 1 if first email actually sent
        new_stage = _campaign_stage_from_results_if_needed(ref, stage)
        if new_stage != stage:
            r["Stage"] = str(new_stage); stage = new_stage; changed = True

        # Load per-key campaign (steps + settings)
        try:
            steps, settings = load_campaign_by_key(key)
            steps = normalize_campaign_steps(steps)
            settings = normalize_campaign_settings(settings)
        except Exception:
            steps, settings = normalize_campaign_steps([]), normalize_campaign_settings({})

        # compute effective divert setting
        try:
            divert_effective = (str(divert_csv).strip() in ("1","true","True"))
            if str(divert_csv).strip() == "":
                divert_effective = (settings.get("send_to_dialer_after") in ("1", True))
        except Exception:
            divert_effective = (settings.get("send_to_dialer_after") in ("1", True))

        # compute if next is due -> delegate drafting to the stub that enforces delays
        if stage == 1:
            # next: E2
            lead = _campaign_get_lead_row_for_ref(r)
            try:
                drafted = globals().get("_draft_next_stage_stub", lambda *a, **k: False)(
                    ref, lead.get("Email",""), lead.get("Company",""), key, 2
                )
                if drafted:
                    r["Stage"] = "2"
                    changed = True
            except Exception:
                pass

        elif stage == 2:
            # next: E3
            lead = _campaign_get_lead_row_for_ref(r)
            try:
                drafted = globals().get("_draft_next_stage_stub", lambda *a, **k: False)(
                    ref, lead.get("Email",""), lead.get("Company",""), key, 3
                )
                if drafted:
                    r["Stage"] = "3"
                    changed = True
            except Exception:
                pass

        elif stage >= 3:
            # completed all emails; if no reply and divert, push to Dialer once then remove
            if not (res and _results_replied(res)) and divert_effective:
                lead = _campaign_get_lead_row_for_ref(r)
                try:
                    ensure_dialer_leads_file()
                    cur = load_dialer_leads_matrix()
                    # Build one row for dialer grid (header-aligned + ○ ○ ○ + notes)
                    base = [lead.get(h,"") for h in HEADER_FIELDS]
                    cur.append(base + ["○","○","○"] + ([""]*8))
                    save_dialer_leads_matrix(cur)
                except Exception:
                    pass
            rows.remove(r)
            changed = True

    if changed:
        _write_campaign_rows(rows)

# ===== CHUNK 4 / 7 — END =====
# ===== CHUNK 5a / 7 — START =====
# ============================================================
# main_after_mount helpers (Part A)
# - Scoreboard helpers
# - Dialer helper shims
# - Warm helpers
# - Campaigns UI helpers
# - Live-sheet extractors
# - Dialer state + save helpers
# (No event loop here)
# ============================================================

def main_after_mount(window, sheet, dial_sheet, leads_host, dialer_host, templates, subjects, mapping, warm_sheet, customer_sheet):
    # ---------- Scoreboard helpers ----------
    from datetime import datetime as _dt
    import time, csv, os, json
    from pathlib import Path

    # Paths (sidecar for hidden geo persistence if Lat/Lon cols aren’t present)
    try:
        base_dir = APP_DIR
    except Exception:
        base_dir = Path.cwd()
    CUSTOMER_GEO_PATH = base_dir / "customers_geo.csv"

    def _safe_get(window, key):
        try:
            return window[key]
        except Exception:
            return None

    def _fmt_money(val):
        try:
            return f"${float(val):.2f}"
        except Exception:
            return "$0.00"

    def _parse_any_date(s):
        if not s:
            return None
        s = str(s).strip()
        if not s:
            return None
        try:
            if "_parse_any_datetime" in globals() and callable(globals()["_parse_any_datetime"]):
                dt = globals()["_parse_any_datetime"](s)
                if dt:
                    return dt.date()
        except Exception:
            pass
        s2 = s.replace(",", " ")
        fmts = [
            "%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m-%d-%Y",
            "%m/%d", "%m-%d",
            "%m/%d/%Y %I:%M %p", "%m/%d/%Y %H:%M",
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %I:%M %p", "%m-%d-%Y %I:%M %p"
        ]
        for fmt in fmts:
            try:
                dt = _dt.strptime(s2, fmt)
                if fmt in ("%m/%d", "%m-%d"):
                    dt = dt.replace(year=_dt.now().year)
                return dt.date()
            except Exception:
                continue
        try:
            from dateutil import parser as _p
            return _p.parse(s2).date()
        except Exception:
            return None

    def _file_rows(path):
        try:
            if path.exists():
                with path.open("r", encoding="utf-8", newline="") as f:
                    rdr = csv.DictReader(f)
                    for r in rdr:
                        yield r
        except Exception:
            return

    def _float_val(s):
        try:
            return float(str(s).replace("$", "").replace(",", "").strip() or "0")
        except Exception:
            return 0.0

    def compute_daily_metrics():
        today = _dt.now().date()
        calls = 0
        for r in _file_rows(DIALER_RESULTS_PATH):
            d = _parse_any_date(r.get("Timestamp", ""))
            if d == today:
                calls += 1
        emails = 0
        for r in _file_rows(RESULTS_PATH):
            d = _parse_any_date(r.get("DateSent", ""))
            if d == today:
                emails += 1
        new_warm = 0
        for r in _file_rows(WARM_LEADS_PATH):
            d = _parse_any_date(r.get("First Contact", "") or r.get("Timestamp", ""))
            if d == today:
                new_warm += 1
        new_accounts = 0
        for r in _file_rows(CUSTOMERS_PATH):
            d = _parse_any_date(r.get("Customer Since", ""))
            if d == today:
                new_accounts += 1
        daily_sales = 0.0
        for r in _file_rows(ORDERS_PATH):
            d = _parse_any_date(r.get("Order Date", ""))
            if d == today:
                daily_sales += _float_val(r.get("Amount", ""))
        return {"calls": calls, "emails": emails, "new_warm": new_warm, "new_accounts": new_accounts, "daily_sales": daily_sales}

    def compute_monthly_metrics():
        now = _dt.now()
        y, m = now.year, now.month
        mo_warm = 0
        for r in _file_rows(WARM_LEADS_PATH):
            d = _parse_any_date(r.get("First Contact", "") or r.get("Timestamp", ""))
            if d and d.year == y and d.month == m:
                mo_warm += 1
        mo_newcus = 0
        for r in _file_rows(CUSTOMERS_PATH):
            d = _parse_any_date(r.get("Customer Since", ""))
            if d and d.year == y and d.month == m:
                mo_newcus += 1
        mo_sales = 0.0
        for r in _file_rows(ORDERS_PATH):
            d = _parse_any_date(r.get("Order Date", ""))
            if d and d.year == y and d.month == m:
                mo_sales += _float_val(r.get("Amount", ""))
        return {"mo_warm": mo_warm, "mo_newcus": mo_newcus, "mo_sales": mo_sales}

    def update_scoreboards(window):
        try:
            d = compute_daily_metrics()
            m = compute_monthly_metrics()
            if (el := _safe_get(window, "-DA_CALLS-")): el.update(str(d["calls"]))
            if (el := _safe_get(window, "-DA_EMAILS-")): el.update(str(d["emails"]))
            if (el := _safe_get(window, "-DA_WARMS-")): el.update(str(d["new_warm"]))
            if (el := _safe_get(window, "-DA_NEWCUS-")): el.update(str(d["new_accounts"]))
            if (el := _safe_get(window, "-DA_SALES-")): el.update(_fmt_money(d["daily_sales"]))
            if (el := _safe_get(window, "-MO_WARMS-")): el.update(str(m["mo_warm"]))
            if (el := _safe_get(window, "-MO_NEWCUS-")): el.update(str(m["mo_newcus"]))
            if (el := _safe_get(window, "-MO_SALES-")): el.update(_fmt_money(m["mo_sales"]))
        except Exception:
            pass  # never crash UI on scoreboard refresh

    def _start_scoreboard_timer(window, interval_ms=4925):
        """Use Tk after() to refresh scoreboards periodically."""
        try:
            def _tick():
                try:
                    update_scoreboards(window)
                finally:
                    try:
                        window.TKroot.after(interval_ms, _tick)
                    except Exception:
                        pass
            update_scoreboards(window)
            window.TKroot.after(interval_ms, _tick)
        except Exception:
            try:
                update_scoreboards(window)
            except Exception:
                pass

    # -------- Dialer helper shims --------
    def dialer_cols_info(_headers=None):
        first_dot = len(HEADER_FIELDS)
        last_dot = first_dot + 2
        first_note = len(HEADER_FIELDS) + 3
        last_note = first_note + 7
        return {"first_dot": first_dot, "last_dot": last_dot,
                "first_note": first_note, "last_note": last_note}

    _DOT_BG = {"green": "#2E7D32", "gray": "#9E9E9E", "red": "#C62828"}
    _DOT_FG = {"green": "#FFFFFF", "gray": "#000000", "red": "#FFFFFF"}

    def dialer_clear_dot_highlights(sheet_obj, row, cols):
        try:
            for c in range(cols["first_dot"], cols["last_dot"] + 1):
                sheet_obj.highlight_cells(row=row, column=c, bg=None, fg=None)
        except Exception:
            pass

    def dialer_colorize_outcome(sheet_obj, row, outcome, cols=None):
        if cols is None:
            cols = dialer_cols_info(None)
        base = cols["first_dot"]
        try:
            for i in range(3):
                sheet_obj.set_cell_data(row, base + i, "○")
            idx = {"green": 0, "gray": 1, "red": 2}[outcome]
            c = base + idx
            try:
                sheet_obj.highlight_cells(row=row, column=c, bg=_DOT_BG[outcome], fg=_DOT_FG[outcome])
            except Exception:
                pass
            sheet_obj.set_cell_data(row, c, "●")
            sheet_obj.refresh()
        except Exception:
            pass

    def dialer_next_empty_note_col(sheet_obj, row, cols=None):
        if cols is None:
            cols = dialer_cols_info(None)
        try:
            r = sheet_obj.get_row_data(row) or []
        except Exception:
            return None
        for c in range(cols["first_note"], cols["last_note"] + 1):
            if c >= len(r) or not (r[c] or "").strip():
                return c
        return None

    def dialer_move_to_next_row(sheet_obj, current_row):
        try:
            total = sheet_obj.get_total_rows()
        except Exception:
            total = 0
        nxt = current_row + 1 if total == 0 else min(current_row + 1, max(0, total - 1))
        try:
            sheet_obj.set_currently_selected(nxt, 0)
            sheet_obj.see(nxt, 0)
        except Exception:
            pass
        return nxt

    # -------- Warm helpers --------
    def warm_get_col_index_map():
        idx = {name: i for i, name in enumerate(WARM_V2_FIELDS)}
        return {
            "cost": idx.get("Cost ($)"),
            "timestamp": idx.get("Timestamp"),
            "first_call": idx.get("Call 1"),
            "last_call": idx.get("Call 15"),
        }

    def warm_next_empty_call_col(row_values, col_map):
        if not row_values:
            return None
        c1, cN = col_map.get("first_call"), col_map.get("last_call")
        if c1 is None or cN is None:
            return None
        for c in range(c1, cN + 1):
            cell = row_values[c] if c < len(row_values) else ""
            if not (cell or "").strip():
                return c
        return None

    def warm_format_cost(val):
        s = (val or "").strip().replace(",", "")
        if not s:
            return ""
        try:
            return f"{float(s):.2f}"
        except Exception:
            return s

    # ---------- Campaigns UI helpers ----------
    def _camp_toggle_empty_vs_editor(window, show_editor: bool):
        try:
            window["-CAMP_EMPTY_WRAP-"].update(visible=not show_editor)
            window["-CAMP_EDITOR_WRAP-"].update(visible=show_editor)
        except Exception:
            pass

    def _camp_blank_steps():
        return [
            {"subject": "", "body": "", "delay_days": 0},
            {"subject": "", "body": "", "delay_days": 0},
            {"subject": "", "body": "", "delay_days": 0},
        ]

    def _camp_default_settings():
        return {"send_to_dialer_after": "1"}

    def camp_read_editor(window):
        steps = []
        for i in (1, 2, 3):
            subj = window[f"-CAMP_SUBJ_{i}-"].get() if f"-CAMP_SUBJ_{i}-" in window.AllKeysDict else ""
            body = window[f"-CAMP_BODY_{i}-"].get() if f"-CAMP_BODY_{i}-" in window.AllKeysDict else ""
            delay = window[f"-CAMP_DELAY_{i}-"].get() if f"-CAMP_DELAY_{i}-" in window.AllKeysDict else "0"
            try:
                delay_i = int(str(delay).strip() or "0")
            except Exception:
                delay_i = 0
            steps.append({"subject": subj or "", "body": body or "", "delay_days": max(0, delay_i)})
        steps = normalize_campaign_steps(steps)

        send_to_dialer = False
        if "-CAMP_SEND_TO_DIALER-" in window.AllKeysDict:
            send_to_dialer = bool(window["-CAMP_SEND_TO_DIALER-"].get())
        settings = normalize_campaign_settings({"send_to_dialer_after": "1" if send_to_dialer else "0"})
        return steps, settings

    def camp_write_editor(window, steps, settings):
        steps = normalize_campaign_steps(steps or [])
        for idx, st in enumerate(steps, start=1):
            if f"-CAMP_SUBJ_{idx}-" in window.AllKeysDict:
                window[f"-CAMP_SUBJ_{idx}-"].update(st.get("subject", ""))
            if f"-CAMP_BODY_{idx}-" in window.AllKeysDict:
                window[f"-CAMP_BODY_{idx}-"].update(st.get("body", ""))
            if f"-CAMP_DELAY_{idx}-" in window.AllKeysDict:
                window[f"-CAMP_DELAY_{idx}-"].update(str(st.get("delay_days", 0)))
        settings = normalize_campaign_settings(settings or {})
        if "-CAMP_SEND_TO_DIALER-" in window.AllKeysDict:
            window["-CAMP_SEND_TO_DIALER-"].update(bool(settings.get("send_to_dialer_after") in ("1", True)))

    def _camp_refresh_combo_and_table(window):
        try:
            keys = list_campaign_keys()
        except Exception:
            keys = []
        if not keys:
            keys = ["default"]
        current = None
        try:
            if "-CAMP_KEY-" in window.AllKeysDict:
                current = (window["-CAMP_KEY-"].get() or "").strip()
        except Exception:
            current = None
        if not current or current not in keys:
            current = keys[0]
        if "-CAMP_KEY-" in window.AllKeysDict:
            window["-CAMP_KEY-"].update(values=keys, value=current)
        try:
            rows = [summarize_campaign_for_table(k) for k in keys]
        except Exception:
            rows = []
        if "-CAMP_TABLE-" in window.AllKeysDict:
            window["-CAMP_TABLE-"].update(values=rows)

        populated = False
        try:
            for k in keys:
                stps, _ = load_campaign_by_key(k)
                stps = normalize_campaign_steps(stps)
                if any((s.get("subject") or s.get("body")) for s in stps):
                    populated = True
                    break
        except Exception:
            populated = False
        _camp_toggle_empty_vs_editor(window, show_editor=populated)

    def _camp_prompt_new_key():
        import re
        key = sg.popup_get_text(
            "Name your new campaign (e.g., 'butcher shop', 'farm market'):",
            title="Add New Campaign",
        )
        if not key:
            return None
        key = key.strip()
        if not key:
            return None
        key = re.sub(r"\s+", " ", key)
        return key

    def _camp_load_into_editor_by_key(window, key):
        try:
            steps, settings = load_campaign_by_key(key)
        except Exception:
            steps, settings = _camp_blank_steps(), _camp_default_settings()
        camp_write_editor(window, steps, settings)
        try:
            window["-CAMP_KEY-"].update(value=key)
        except Exception:
            pass
        _camp_toggle_empty_vs_editor(window, True)

    # ---------- Live-sheet extractors ----------
    def matrix_from_sheet():
        if sheet is None or not hasattr(sheet, "get_sheet_data"):
            return []
        raw = sheet.get_sheet_data() or []
        trimmed = []
        for row in raw:
            row = (list(row) + [""] * len(HEADER_FIELDS))[:len(HEADER_FIELDS)]
            trimmed.append([str(c) for c in row])
        while trimmed and not any((cell or "").strip() for cell in trimmed[-1]):
            trimmed.pop()
        return trimmed

    def warm_matrix_from_sheet_v2():
        if warm_sheet is None or not hasattr(warm_sheet, "get_sheet_data"):
            return []
        raw = warm_sheet.get_sheet_data() or []
        trimmed = []
        for row in raw:
            row = (list(row) + [""] * len(WARM_V2_FIELDS))[:len(WARM_V2_FIELDS)]
            trimmed.append([str(c) for c in row])
        while trimmed and not any((cell or "").strip() for cell in trimmed[-1]):
            trimmed.pop()
        return trimmed

    def customers_matrix_from_sheet():
        if customer_sheet is None or not hasattr(customer_sheet, "get_sheet_data"):
            return []
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
    state = {"row": None, "outcome": None, "note_col_by_row": {}, "colored_row": None}

    def _row_selected(sheet_obj):
        if sheet_obj is None:
            return None
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
        try:
            return (window["-DIAL_NOTE-"].get() or "").strip()
        except Exception:
            return ""

    def _confirm_enabled():
        r = state["row"]
        if r is None:
            return False
        have_outcome = state["outcome"] in ("green", "gray", "red")
        have_text = bool(_current_note_text())
        sticky = state["note_col_by_row"].get(r)
        have_slot = (sticky is not None) or (dialer_next_empty_note_col(dial_sheet, r, cols) is not None)
        return have_outcome and have_text and have_slot

    def _update_confirm_button():
        ok = _confirm_enabled()
        try:
            window["-DIAL_CONFIRM-"].update(disabled=not ok, button_color=("white", "#2E7D32" if ok else "#444444"))
        except Exception:
            pass

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
                dial_sheet.set_cell_data(r, f + i, "○")
            dial_sheet.set_cell_data(r, f + {"green": 0, "gray": 1, "red": 2}[which], "●")
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

    def _save_dialer_grid_to_csv():
        try:
            data = dial_sheet.get_sheet_data() or []
        except Exception:
            data = []
        expected_len = len(HEADER_FIELDS) + 3 + 8
        matrix = []
        for row in data:
            r = (list(row) + [""] * expected_len)[:expected_len]
            for i in range(len(HEADER_FIELDS), len(HEADER_FIELDS) + 3):
                r[i] = r[i] if (r[i] or "").strip() else "○"
            matrix.append(r)
        save_dialer_leads_matrix(matrix)

    # prime the dialer selected row
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
    warm_state = {"row": None, "outcome": None}
    warm_cols = warm_get_col_index_map()

    def _warm_selected_row():
        return _row_selected(warm_sheet)

    def _warm_note_text():
        try:
            return (window["-WARM_NOTE-"].get() or "").strip()
        except Exception:
            return ""

    def _warm_confirm_enabled():
        r = warm_state["row"]
        if r is None: return False
        if warm_state["outcome"] not in ("green", "gray", "red"): return False
        if not _warm_note_text(): return False
        try:
            row_vals = warm_sheet.get_row_data(r) or []
        except Exception:
            row_vals = []
        return warm_next_empty_call_col(row_vals, warm_cols) is not None

    def _warm_update_confirm_button():
        ok = _warm_confirm_enabled()
        try:
            window["-WARM_CONFIRM-"].update(disabled=not ok, button_color=("white", "#2E7D32" if ok else "#444444"))
        except Exception:
            pass

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
        ci = warm_cols.get("cost")
        if ci is None or warm_sheet is None:
            return
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
        try:
            total = warm_sheet.get_total_rows()
        except Exception:
            total = 0
        for r in range(total):
            _warm_cost_normalize_in_row(r)
        matrix = warm_matrix_from_sheet_v2()
        save_warm_leads_matrix_v2(matrix)

# ===== CHUNK 5a / 7 — END (continue in 5b) =====
# ===== CHUNK 5b / 7 — START =====
# Continuation of main_after_mount from 5a

    # ============================================================
    # Customers helpers (save / selection / add order / analytics)
    # ============================================================
    def _customer_selected_row():
        return _row_selected(customer_sheet)

    def _cust_idx(name, default=None):
        try:
            return CUSTOMER_FIELDS.index(name)
        except Exception:
            return default

    def _popup_add_order(company):
        comp_disp = company or "(unknown)"
        layout = [
            [sg.Text(f"Add Order for: {comp_disp}", text_color="#9EE493")],
            [sg.Text("Amount ($):", size=(12,1)), sg.Input(key="-AO_AMOUNT-", size=(20,1))],
            [sg.Text("Order Date:", size=(12,1)), sg.Input(_dt.now().strftime("%Y-%m-%d"), key="-AO_DATE-", size=(20,1)),
             sg.Text(" (YYYY-MM-DD or MM/DD/YYYY)", text_color="#AAAAAA")],
            [sg.Push(), sg.Button("Cancel"), sg.Button("Add", button_color=("white","#2E7D32"))]
        ]
        win = sg.Window("Add Order", layout, modal=True, finalize=True)
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

    def _safe_update(key, text):
        try:
            if key in window.AllKeysDict:
                window[key].update(text)
        except Exception:
            pass

    # ---------- Customers autosave + display money formatting ----------
    def _customers_commit_edits():
        try:
            if hasattr(customer_sheet, "end_edit_cell"):
                customer_sheet.end_edit_cell()
        except Exception:
            pass
        try:
            if hasattr(customer_sheet, "MT") and hasattr(customer_sheet.MT, "end_edit_cell"):
                customer_sheet.MT.end_edit_cell()
        except Exception:
            pass
        try:
            for attr in ("quit_edit", "close_text_editor", "stop_editing"):
                fn = getattr(getattr(customer_sheet, "MT", customer_sheet), attr, None)
                if callable(fn):
                    try:
                        fn()
                    except Exception:
                        pass
        except Exception:
            pass

    def _customers_snapshot():
        try:
            raw = customer_sheet.get_sheet_data() or []
        except Exception:
            raw = []
        trimmed = []
        for row in raw:
            r = (list(row) + [""] * len(CUSTOMER_FIELDS))[:len(CUSTOMER_FIELDS)]
            trimmed.append([str(x) for x in r])
        while trimmed and not any((cell or "").strip() for cell in trimmed[-1]):
            trimmed.pop()
        return tuple(tuple(r) for r in trimmed)

    def _customers_display_refresh_money():
        idx_cltv = _cust_idx("CLTV")
        idx_spd  = _cust_idx("Sales/Day")
        if idx_cltv is None and idx_spd is None:
            return
        try:
            total = customer_sheet.get_total_rows()
        except Exception:
            total = 0
        editing = None
        try:
            editing = getattr(getattr(customer_sheet, "MT", customer_sheet), "text_editor_loc", None)
        except Exception:
            editing = None

        for r in range(total):
            try:
                row = customer_sheet.get_row_data(r) or []
            except Exception:
                row = []
            for idx in (idx_cltv, idx_spd):
                if idx is None:
                    continue
                try:
                    if editing and isinstance(editing, (tuple, list)) and len(editing) >= 2 and (r, idx) == tuple(editing):
                        continue
                except Exception:
                    pass
                s = str(row[idx] if idx < len(row) else "").strip()
                if not s:
                    disp = ""
                else:
                    try:
                        v = float(s.replace("$", "").replace(",", ""))
                        disp = f"${v:,.2f}"
                    except Exception:
                        disp = s
                try:
                    customer_sheet.set_cell_data(r, idx, disp)
                except Exception:
                    pass
        try:
            customer_sheet.refresh()
        except Exception:
            pass

    _customers_last_hash = {"val": None}

    # ------------ Geocoding helpers (auto-run during autosave) ------------
    __geo_cache_mem = {}
    __geo_last_time = [0.0]

    def _compose_addr_from_row(row_vals):
        def g(name):
            i = _cust_idx(name)
            return (row_vals[i] if i is not None and i < len(row_vals) else "").strip()
        addr = g("Address")
        city = g("City")
        state = g("State")
        zipc = g("ZIP")
        if not any([addr, city, state, zipc]):
            loc = g("Location")
            if loc:
                return loc
            return ""
        parts = [addr, city, state, zipc]
        return ", ".join([p for p in parts if p])

    def _geocode_address(addr):
        if not addr:
            return None
        key = addr.lower().strip()
        if key in __geo_cache_mem:
            return __geo_cache_mem[key]
        gap = time.time() - __geo_last_time[0]
        if gap < 1.1:
            time.sleep(1.1 - gap)
        try:
            import urllib.parse, urllib.request, json as _json
            url = "https://nominatim.openstreetmap.org/search?" + urllib.parse.urlencode({
                "q": addr, "format": "json", "limit": 1, "addressdetails": 0
            })
            req = urllib.request.Request(url, headers={"User-Agent": "DeathStarCRM/1.0 (contact: you@example.com)"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = _json.loads(resp.read().decode("utf-8", errors="ignore"))
            __geo_last_time[0] = time.time()
            if data:
                lat = float(data[0]["lat"]); lon = float(data[0]["lon"])
                __geo_cache_mem[key] = (lat, lon)
                return (lat, lon)
        except Exception:
            __geo_last_time[0] = time.time()
            return None
        return None

    def _load_geo_sidecar():
        mp = {}
        if CUSTOMER_GEO_PATH.exists():
            try:
                with CUSTOMER_GEO_PATH.open("r", encoding="utf-8", newline="") as f:
                    rdr = csv.DictReader(f)
                    for r in rdr:
                        k = (r.get("AddressKey","") or "").strip().lower()
                        if k and (r.get("Lat") and r.get("Lon")):
                            try:
                                mp[k] = (float(r["Lat"]), float(r["Lon"]))
                            except Exception:
                                pass
            except Exception:
                pass
        return mp

    def _save_geo_sidecar(mapping):
        rows = []
        for k, (la, lo) in mapping.items():
            rows.append({"AddressKey": k, "Lat": f"{la:.6f}", "Lon": f"{lo:.6f}"})
        try:
            _backup(CUSTOMER_GEO_PATH)
        except Exception:
            pass
        try:
            with CUSTOMER_GEO_PATH.open("w", encoding="utf-8", newline="") as f:
                w = csv.DictWriter(f, fieldnames=["AddressKey","Lat","Lon"])
                w.writeheader()
                for r in rows:
                    w.writerow(r)
        except Exception:
            pass

    __geo_sidecar = _load_geo_sidecar()

    def _customers_geocode_row_if_needed(row_vals):
        i_lat = _cust_idx("Lat")
        i_lon = _cust_idx("Lon")
        lat_now = (row_vals[i_lat] if i_lat is not None and i_lat < len(row_vals) else "").strip() if i_lat is not None else ""
        lon_now = (row_vals[i_lon] if i_lon is not None and i_lon < len(row_vals) else "").strip() if i_lon is not None else ""
        need_coords = not (lat_now and lon_now)
        if not need_coords:
            return
        addr = _compose_addr_from_row(row_vals)
        if not addr:
            return
        key = addr.lower().strip()
        if key in __geo_sidecar:
            la, lo = __geo_sidecar[key]
        else:
            got = _geocode_address(addr)
            if not got:
                return
            la, lo = got
            __geo_sidecar[key] = (la, lo)
            _save_geo_sidecar(__geo_sidecar)
        if i_lat is not None:
            try: customer_sheet.set_cell_data(customer_sheet.get_currently_selected()[0] if False else 0, 0, "")
            except Exception: pass
        if i_lat is not None:
            try: customer_sheet.set_cell_data(row_vals_index, i_lat, f"{la:.6f}")
            except Exception: pass
        if i_lon is not None:
            try: customer_sheet.set_cell_data(row_vals_index, i_lon, f"{lo:.6f}")
            except Exception: pass

    # ---------- Orders helpers used by analytics ----------
    def _orders_count_by_company():
        counts = {}
        if ORDERS_PATH.exists():
            with ORDERS_PATH.open("r", encoding="utf-8", newline="") as f:
                rdr = csv.DictReader(f)
                for r in rdr:
                    comp = (r.get("Company","") or "").strip()
                    amt = r.get("Amount","") or ""
                    try:
                        val = float(str(amt).replace(",", "").strip() or "0")
                    except Exception:
                        val = 0.0
                    if comp:
                        c, s = counts.get(comp, (0, 0.0))
                        counts[comp] = (c+1, s+val)
        return counts

    def refresh_customer_analytics():
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
                            samples_sum += float((r.get("Cost ($)","") or "0").replace(",", "").strip() or "0")
                        except Exception:
                            pass
        except Exception:
            pass

        new_customers = 0
        total_customers = 0
        ltv_vals = []
        total_sales = 0.0
        reorder_yes = 0

        orders_counts = _orders_count_by_company()
        idx_reorder = _cust_idx("Reorder?")
        idx_company = _cust_idx("Company")

        try:
            rows = customer_sheet.get_sheet_data() or []
        except Exception:
            rows = []

        try:
            if CUSTOMERS_PATH.exists():
                with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
                    rdr = csv.DictReader(f)
                    disk_rows = list(rdr)
            else:
                disk_rows = []
        except Exception:
            disk_rows = []

        disk_changed = False

        for r_idx, row in enumerate(rows):
            rec = {CUSTOMER_FIELDS[i]: (row[i] if i < len(CUSTOMER_FIELDS) else "") for i in range(len(CUSTOMER_FIELDS))}
            comp = (rec.get("Company","") or "").strip()
            if not any((rec.get(h,"") or "").strip() for h in CUSTOMER_FIELDS):
                continue
            total_customers += 1
            if (rec.get("Customer Since","") or "").strip():
                new_customers += 1
            try:
                v = float((rec.get("CLTV","") or "").replace("$","").replace(",","" ).strip() or "0")
                if v > 0:
                    ltv_vals.append(v)
                    total_sales += v
            except Exception:
                pass
            is_yes_now = ((rec.get("Reorder?","") or "").strip().lower() == "yes")
            oc = orders_counts.get(comp, (0, 0.0))[0] if comp else 0
            if oc >= 2 and not is_yes_now:
                if idx_reorder is not None:
                    try:
                        customer_sheet.set_cell_data(r_idx, idx_reorder, "Yes")
                        is_yes_now = True
                    except Exception:
                        pass
                for drow in disk_rows:
                    if (drow.get("Company","") or "").strip() == comp:
                        if (drow.get("Reorder?","") or "").strip().lower() != "yes":
                            drow["Reorder?"] = "Yes"
                            disk_changed = True
            if is_yes_now:
                reorder_yes += 1

        if disk_changed:
            try:
                _backup(CUSTOMERS_PATH)
                _atomic_write_csv(CUSTOMERS_PATH, CUSTOMER_FIELDS, [[r.get(h,"") for h in CUSTOMER_FIELDS] for r in disk_rows])
            except Exception:
                pass

        close_rate = (new_customers / warm_leads * 100.0) if warm_leads else 0.0
        cac = (samples_sum / new_customers) if new_customers else 0.0
        avg_ltv = (sum(ltv_vals) / len(ltv_vals)) if ltv_vals else 0.0
        reorder_rate = (reorder_yes / total_customers * 100.0) if total_customers else 0.0

        _safe_update("-AN_WARMS-", str(warm_leads))
        _safe_update("-AN_NEWCUS-", str(new_customers))
        _safe_update("-AN_CLOSERATE-", f"{close_rate:.1f}%")
        _safe_update("-AN_TOTALSALES-", f"{total_sales:.2f}")
        _safe_update("-AN_CAC-", f"{cac:.2f}")
        _safe_update("-AN_LTV-", f"{avg_ltv:.2f}")
        ratio = (avg_ltv / cac) if cac > 0 else 0.0
        _safe_update("-AN_CACLTV-", f"1 : {ratio:.2f}" if cac > 0 else "1 : 0")
        _safe_update("-AN_REORDER-", f"{reorder_rate:.1f}%")

    # Start periodic autosave for Customers and format money in-grid initially
    try:
        _customers_display_refresh_money()
    except Exception:
        pass
    try:
        _start_customers_autosave(window, interval_ms=4000)
    except Exception:
        pass

    # ============================================================
    # AUTOSAVE: debounce helpers and sheet bindings (DEDUPED & FIXED)
    # ============================================================
    class _Debounced:
        def __init__(self, tkroot):
            self.tkroot = tkroot
            self._after = None
        def call(self, delay_ms, fn):
            try:
                if self._after is not None:
                    self.tkroot.after_cancel(self._after)
            except Exception:
                pass
            try:
                self._after = self.tkroot.after(delay_ms, fn)
            except Exception:
                fn()

    _cust_deb = _Debounced(window.TKroot)
    _warm_deb = _Debounced(window.TKroot)
    _dial_deb = _Debounced(window.TKroot)

    def _customers_recompute_and_save():
        """Normalize money cols, recompute, geocode if needed, then save."""
        try:
            idx = {name: i for i, name in enumerate(CUSTOMER_FIELDS)}
            i_company  = idx.get("Company")
            i_first    = idx.get("First Order")
            i_since    = idx.get("Customer Since")
            i_cltv     = idx.get("CLTV")
            i_days     = idx.get("Days")
            i_salesday = idx.get("Sales/Day")
            i_lat      = idx.get("Lat")
            i_lon      = idx.get("Lon")

            def _pdate(s):
                s = (s or "").strip()
                if not s: return None
                for fmt in ("%Y-%m-%d","%m/%d/%Y","%m-%d-%Y","%Y/%m/%d","%m/%d","%m-%d"):
                    try:
                        d = _dt.strptime(s, fmt)
                        if fmt in ("%m/%d","%m-%d"):
                            d = d.replace(year=_dt.now().year)
                        return d.date()
                    except Exception:
                        pass
                return None
            def _m2f(x):
                s = (str(x) if x is not None else "").replace("$","").replace(",","" ).strip()
                try: return float(s) if s else 0.0
                except Exception: return 0.0
            def _f2m(v):
                try: return f"{float(v):.2f}"
                except Exception: return ""

            total = customer_sheet.get_total_rows() if hasattr(customer_sheet,"get_total_rows") else 0
            for r in range(total):
                try:
                    row = customer_sheet.get_row_data(r) or []
                except Exception:
                    continue

                global row_vals_index
                row_vals_index = r
                try:
                    lat_now = (row[i_lat] if (i_lat is not None and i_lat < len(row)) else "").strip() if i_lat is not None else ""
                    lon_now = (row[i_lon] if (i_lon is not None and i_lon < len(row)) else "").strip() if i_lon is not None else ""
                except Exception:
                    lat_now = lon_now = ""
                if not (lat_now and lon_now):
                    _customers_geocode_row_if_needed(row)
                    try:
                        row = customer_sheet.get_row_data(r) or row
                    except Exception:
                        pass

                company = (row[i_company] if i_company is not None and i_company < len(row) else "").strip()
                if not company:
                    if i_cltv is not None and i_cltv < len(row) and (row[i_cltv] or "").strip() == "0.00":
                        try: customer_sheet.set_cell_data(r, i_cltv, "")
                        except Exception: pass
                    continue

                if i_cltv is not None and i_cltv < len(row):
                    val = (row[i_cltv] or "").strip()
                    if val != "":
                        try:
                            customer_sheet.set_cell_data(r, i_cltv, _f2m(_m2f(val)))
                        except Exception:
                            pass

                first_s  = row[i_first] if i_first is not None and i_first < len(row) else ""
                since_s  = row[i_since] if i_since is not None and i_since < len(row) else ""
                first_dt = _pdate(first_s) or _pdate(since_s)
                if i_days is not None:
                    try:
                        if first_dt:
                            days = max(1, (_dt.now().date() - first_dt).days)
                            customer_sheet.set_cell_data(r, i_days, str(days))
                        else:
                            customer_sheet.set_cell_data(r, i_days, "")
                    except Exception:
                        pass

                if i_salesday is not None:
                    try:
                        cltv_v = _m2f(customer_sheet.get_cell_data(r, i_cltv)) if i_cltv is not None else 0.0
                    except Exception:
                        cltv_v = 0.0
                    try:
                        days_v = int(float(customer_sheet.get_cell_data(r, i_days) or "0"))
                    except Exception:
                        days_v = 0
                    try:
                        if cltv_v > 0 and days_v > 0:
                            customer_sheet.set_cell_data(r, i_salesday, _f2m(cltv_v / days_v))
                        else:
                            customer_sheet.set_cell_data(r, i_salesday, "")
                    except Exception:
                        pass

            _save_customers_grid_to_csv()
            refresh_customer_analytics()
        except Exception:
            try:
                _save_customers_grid_to_csv()
                refresh_customer_analytics()
            except Exception:
                pass

    def _warm_save_debounced():
        _warm_deb.call(600, _save_warm_grid_to_csv_v2)

    def _dial_save_debounced():
        _dial_deb.call(600, _save_dialer_grid_to_csv)

    def _cust_save_debounced():
        _cust_deb.call(600, _customers_recompute_and_save)

    try:
        if hasattr(customer_sheet, "extra_bindings"):
            customer_sheet.extra_bindings([
                ("end_edit_cell", lambda *a, **k: _cust_save_debounced()),
                ("paste",         lambda *a, **k: _cust_save_debounced()),
                ("row_add",       lambda *a, **k: _cust_save_debounced()),
                ("row_delete",    lambda *a, **k: _cust_save_debounced()),
            ])
    except Exception:
        pass

    try:
        if hasattr(warm_sheet, "extra_bindings"):
            warm_sheet.extra_bindings([
                ("end_edit_cell", lambda *a, **k: _warm_save_debounced()),
                ("paste",         lambda *a, **k: _warm_save_debounced()),
                ("row_add",       lambda *a, **k: _warm_save_debounced()),
                ("row_delete",    lambda *a, **k: _warm_save_debounced()),
            ])
    except Exception:
        pass

    try:
        if hasattr(dial_sheet, "extra_bindings"):
            dial_sheet.extra_bindings([
                ("end_edit_cell", lambda *a, **k: _dial_save_debounced()),
                ("paste",         lambda *a, **k: _dial_save_debounced()),
                ("row_add",       lambda *a, **k: _dial_save_debounced()),
                ("row_delete",    lambda *a, **k: _dial_save_debounced()),
            ])
    except Exception:
        pass

    # ---- Start periodic autosave and scoreboard updates ----
    try:
        _start_scoreboard_timer(window, interval_ms=5000)
    except Exception:
        pass

    try:
        _customers_recompute_and_save()
    except Exception:
        pass

# ===== CHUNK 5b / 7 — END (Chunk 6 begins next, not here) =====
# ===== CHUNK 6 / 7 — START =====
    # ============================================================
    # Prime UI (Email Results stats + analytics + scoreboards)
    # ============================================================

    # Helper: recompute & recolor the "Emails Sent" display column on the Leads grid
    def _emails_sent_for(addr: str) -> int:
        a = (addr or "").strip().lower()
        if not a:
            return 0
        try:
            rows = load_results_rows_sorted()
        except Exception:
            return 0
        n = 0
        for r in rows:
            if (r.get("Email", "") or "").strip().lower() == a and (r.get("DateSent") or "").strip():
                n += 1
        return min(n, 3)

    def _refresh_emails_sent_column_impl():
        """
        Rebuild the last column ("Emails Sent") based on results.csv and apply the
        same grayscale heat styling used at mount time.
        """
        try:
            if sheet is None:
                return
            # Last column is the extra "Emails Sent" appended to HEADER_FIELDS
            try:
                total_rows = sheet.get_total_rows()
            except Exception:
                total_rows = 0

            # Figure out where Email column lives in the data portion
            try:
                email_idx = HEADER_FIELDS.index("Email")
            except Exception:
                email_idx = 0

            emails_col = len(HEADER_FIELDS)  # display col index for "Emails Sent"
            for r in range(max(0, total_rows)):
                try:
                    row_vals = sheet.get_row_data(r) or []
                except Exception:
                    row_vals = []
                addr = row_vals[email_idx] if email_idx < len(row_vals) else ""
                n = _emails_sent_for(addr)

                # Update cell text
                try:
                    sheet.set_cell_data(r, emails_col, str(n))
                except Exception:
                    pass

                # Apply coloring: 0 white, 1 light gray, 2 gray, 3+ dark gray
                bg = "#FFFFFF"; fg = "#000000"
                if n == 1:
                    bg = "#EEEEEE"
                elif n == 2:
                    bg = "#CFCFCF"
                elif n >= 3:
                    bg = "#A6A6A6"
                try:
                    sheet.highlight_cells(row=r, column=emails_col, bg=bg, fg=fg)
                except Exception:
                    pass

            try:
                sheet.refresh()
            except Exception:
                pass
        except Exception:
            # Never crash UI during a cosmetic refresh
            pass

    # attach as a method so we can call window._refresh_emails_sent_column() safely
    try:
        setattr(window, "_refresh_emails_sent_column", _refresh_emails_sent_column_impl)
    except Exception:
        pass

    # ------------------------------------------------------------
    # MAP helpers (read customers, build Leaflet HTML, open browser)
    # ------------------------------------------------------------
    def _first_nonempty(d, keys, default=""):
        for k in keys:
            v = d.get(k, "")
            if (v or "").strip():
                return v
        return default

    def _money_fmt(x):
        s = (str(x) if x is not None else "").replace("$", "").replace(",", "").strip()
        if not s:
            return ""
        try:
            return f"${float(s):,.2f}"
        except Exception:
            return str(x)

    # Compose an "address key" the same way Chunk 5 geocoder does ("Address, City, State, ZIP" or "Location")
    def _compose_addr_key_from_rowdict(row):
        addr = (row.get("Address") or "").strip()
        city = (row.get("City") or "").strip()
        state = (row.get("State") or "").strip()
        zipc = (row.get("ZIP") or "").strip()
        if any([addr, city, state, zipc]):
            parts = [p for p in (addr, city, state, zipc) if (p or "").strip()]
            key = ", ".join(parts)
        else:
            key = (row.get("Location") or "").strip()
        return key.lower().strip()

    # NEW: Merge coordinates from customers_geo.csv -> customers.csv (fills blank Lat/Lon)
    def _merge_sidecar_coords_into_customers():
        try:
            GEO_PATH = APP_DIR / "customers_geo.csv"
        except Exception:
            GEO_PATH = Path.cwd() / "customers_geo.csv"

        # 1) Load sidecar map: AddressKey -> (lat, lon)
        sidecar = {}
        if GEO_PATH.exists():
            try:
                with GEO_PATH.open("r", encoding="utf-8", newline="") as f:
                    rdr = csv.DictReader(f)
                    # Support both schemas:
                    #  - AddressKey,Lat,Lon   (from Chunk 5 geocoder)
                    #  - lat,lon,company/...  (fallback — we won't have the address key then)
                    fields = [h.lower() for h in (rdr.fieldnames or [])]
                    if "addresskey".lower() in fields and "lat" in fields and "lon" in fields:
                        for r in rdr:
                            k = (r.get("AddressKey") or r.get("addresskey") or "").strip().lower()
                            la = r.get("Lat") or r.get("lat") or ""
                            lo = r.get("Lon") or r.get("lon") or ""
                            try:
                                la_f = float(la); lo_f = float(lo)
                                if k:
                                    sidecar[k] = (la_f, lo_f)
                            except Exception:
                                pass
                    else:
                        # No AddressKey schema; we can’t safely map rows -> skip merge
                        sidecar = {}
            except Exception:
                sidecar = {}

        if not sidecar:
            return  # nothing to merge

        # 2) Read customers.csv, fill blanks where we have a sidecar match
        if not CUSTOMERS_PATH.exists():
            return
        try:
            with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
                rdr = csv.DictReader(f)
                fieldnames = list(rdr.fieldnames or [])
                rows = list(rdr)
        except Exception:
            return

        # Ensure Lat/Lon in header (append if missing)
        changed = False
        if "Lat" not in fieldnames:
            fieldnames.append("Lat"); changed = True
        if "Lon" not in fieldnames:
            fieldnames.append("Lon"); changed = True

        for r in rows:
            lat_now = (r.get("Lat") or "").strip()
            lon_now = (r.get("Lon") or "").strip()
            if lat_now and lon_now:
                continue  # already filled

            key = _compose_addr_key_from_rowdict(r)
            if not key:
                continue
            hit = sidecar.get(key)
            if not hit:
                continue
            la, lo = hit
            r["Lat"] = f"{la:.6f}"
            r["Lon"] = f"{lo:.6f}"
            changed = True

        if not changed:
            return

        # 3) Write back (with backup), so the map & grid can see coordinates
        try:
            _backup(CUSTOMERS_PATH)
        except Exception:
            pass
        try:
            with CUSTOMERS_PATH.open("w", encoding="utf-8", newline="") as f:
                w = csv.DictWriter(f, fieldnames=fieldnames)
                w.writeheader()
                for r in rows:
                    w.writerow(r)
        except Exception:
            # Best-effort; we won’t crash UI
            pass

    def _load_customers_for_map():
        """Return (records, skipped_count). Pulls coords from main CSV (Lat/Lon)."""
        recs = []
        skipped = 0

        def _first_nonempty_row(row, keys):
            for k in keys:
                v = (row.get(k, "") or "").strip()
                if v:
                    return v
            return ""

        def _money_fmt_local(s):
            s = (s or "").strip().replace("$", "").replace(",", "")
            if not s:
                return ""
            try:
                return f"${float(s):.2f}"
            except Exception:
                return ""

        try:
            # Prefer lat/lon columns in customers.csv (after we just merged sidecar coords)
            if CUSTOMERS_PATH.exists():
                with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
                    rdr = csv.DictReader(f)
                    for r in rdr:
                        lat = _first_nonempty_row(r, ["Lat", "Latitude", "lat", "LAT"])
                        lon = _first_nonempty_row(r, ["Lon", "Lng", "Longitude", "longitude", "lon", "LON"])
                        try:
                            lat_f = float(str(lat)) if lat else None
                            lon_f = float(str(lon)) if lon else None
                        except Exception:
                            lat_f = lon_f = None

                        if lat_f is None or lon_f is None:
                            skipped += 1
                            continue

                        company = _first_nonempty_row(r, ["Company"])
                        cltv    = _money_fmt_local(_first_nonempty_row(r, ["CLTV"]))
                        spd     = _money_fmt_local(_first_nonempty_row(r, ["Sales/Day", "Sales per Day", "Sales Per Day"]))

                        popup_html = (
                            f"<b>{company or '(Unnamed)'}</b><br/>"
                            f"CLTV: {cltv or '$0.00'}<br/>"
                            f"Sales/Day: {spd or '$0.00'}"
                        )

                        recs.append({
                            "lat": lat_f, "lon": lon_f,
                            "company": company, "cltv": cltv, "spd": spd,
                            "popup_html": popup_html
                        })

        except Exception as e:
            print("map loader error:", e)

        return recs, skipped

    def _write_leaflet_map_html(recs, outfile):
        """Write a standalone Leaflet HTML file with pins for recs."""
        # Fallback center: average of coords if possible, else CONUS
        if recs:
            try:
                avg_lat = sum(r["lat"] for r in recs) / len(recs)
                avg_lon = sum(r["lon"] for r in recs) / len(recs)
            except Exception:
                avg_lat, avg_lon = 39.5, -98.35
        else:
            avg_lat, avg_lon = 39.5, -98.35

        # Basic Leaflet page
        markers_js = []
        for r in recs:
            popup_safe = r["popup_html"].replace("\\", "\\\\").replace("`", "\\`")
            markers_js.append(
                f"L.marker([{r['lat']:.6f}, {r['lon']:.6f}]).addTo(map)"
                f".bindPopup(`{popup_safe}`);"
            )
        markers_blob = "\n        ".join(markers_js)

        html = f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Customer Map</title>
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
  <style>
    html, body, #map {{ height: 100%; margin: 0; padding: 0; background:#111; }}
    .leaflet-popup-content-wrapper, .leaflet-popup-tip {{ background:#222; color:#eee; }}
  </style>
</head>
<body>
  <div id="map"></div>
  <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
  <script>
    var map = L.map('map').setView([{avg_lat:.6f}, {avg_lon:.6f}], {12 if len(recs)==1 else 5});
    L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
      maxZoom: 19,
      attribution: '&copy; OpenStreetMap'
    }}).addTo(map);

    {markers_blob}
  </script>
</body>
</html>"""
        try:
            with open(outfile, "w", encoding="utf-8") as f:
                f.write(html)
            return True
        except Exception:
            return False

    def _open_customer_map(window):
        """Build HTML map and open it. Updates -MAP_STATUS- label."""
        try:
            out_path = APP_DIR / "customer_map.html"
        except Exception:
            out_path = Path.cwd() / "customer_map.html"

        # Ensure customers.csv has coordinates merged from sidecar before building the map
        try:
            _merge_sidecar_coords_into_customers()
        except Exception:
            pass

        recs, skipped = _load_customers_for_map()
        if not recs and skipped == 0:
            try:
                window["-MAP_STATUS-"].update("No customers yet.")
            except Exception:
                pass
            return

        ok = _write_leaflet_map_html(recs, str(out_path))
        if not ok:
            try:
                window["-MAP_STATUS-"].update("Error writing map HTML.")
            except Exception:
                pass
            return

        # Try to open in default browser
        try:
            import webbrowser
            webbrowser.open(str(out_path))
        except Exception:
            pass

        msg = f"Opened map ({len(recs)} pin(s){', skipped ' + str(skipped) + ' without coords' if skipped else ''})."
        try:
            window["-MAP_STATUS-"].update(msg)
        except Exception:
            pass

    def refresh_fire_state():
        matrix = matrix_from_sheet()
        seen = load_state_set()
        new_count = 0
        for row in matrix:
            d = dict_from_row(row)
            if not valid_email(d.get("Email", "")):
                continue
            fp = row_fingerprint_from_dict(d)
            if fp not in seen:
                new_count += 1
        if new_count > 0:
            window["-FIRE-"].update(disabled=False, button_color=("white", "#C00000"))
            window["-FIRE_HINT-"].update(f" Ready: {new_count} new lead(s).")
        else:
            window["-FIRE-"].update(disabled=True, button_color=("white", "#700000"))
            window["-FIRE_HINT-"].update(" (no NEW leads; already drafted or no valid emails)")

    def refresh_results_metrics():
        rows = load_results_rows_sorted()
        total_sent = sum(1 for r in rows if r.get("DateSent"))
        total_replied = sum(1 for r in rows if r.get("DateReplied"))
        warm = sum(1 for r in rows if (r.get("Status", "").lower() == "green"))
        window["-WARM-"].update(str(warm))
        window["-REPLRATE-"].update(f"{total_replied} / {total_sent}")

    refresh_fire_state()
    refresh_results_metrics()
    # NEW: update Emails Sent column/colors at startup
    try:
        window._refresh_emails_sent_column()
    except Exception:
        pass

    _update_confirm_button()
    _warm_update_confirm_button()
    try:
        refresh_customer_analytics()
    except Exception:
        pass

    # Start scoreboard auto-refresh (every 5 seconds)
    _start_scoreboard_timer(window, interval_ms=5000)

    # NEW: one-time merge of sidecar coords into customers.csv at startup
    try:
        _merge_sidecar_coords_into_customers()
    except Exception:
        pass

    # ============================================================
    # Event loop
    # ============================================================
    while True:
        event, values = window.read(timeout=300)
        if event in (sg.WINDOW_CLOSE_ATTEMPTED_EVENT, sg.WIN_CLOSED):
            break

        # Selection change tracking
        r_now = _row_selected(dial_sheet)
        if r_now is not None and r_now != state["row"]:
            _set_working_row(r_now)
            state["note_col_by_row"].setdefault(r_now, None)
            _update_confirm_button()

        r_warm_now = _warm_selected_row()
        if r_warm_now is not None and r_warm_now != warm_state["row"]:
            _warm_set_row(r_warm_now)

        # ---------- Top bar ----------
        if event == "-UPDATE-":
            # Manual refresh of scoreboards + analytics
            try:
                update_scoreboards(window)
                refresh_customer_analytics()
            except Exception as e:
                popup_error(f"Refresh error: {e}")

        # ---------------- Email Leads tab ----------------
        if event == "-OPENFOLDER-":
            try:
                os.startfile(str(APP_DIR))
            except Exception as e:
                popup_error(f"Open folder error: {e}")

        elif event == "-ADDROWS-":
            if sheet is None:
                popup_error("Leads sheet not initialized.")
            else:
                try:
                    sheet.insert_rows(sheet.get_total_rows(), number_of_rows=10)
                    sheet.refresh()
                except Exception:
                    try:
                        sheet.insert_rows(sheet.get_total_rows(), amount=10)
                        sheet.refresh()
                    except Exception as e:
                        popup_error(f"Could not add rows: {e}")
            refresh_fire_state()
            # NEW: refresh Emails Sent display after structure change
            try:
                window._refresh_emails_sent_column()
            except Exception:
                pass

        elif event == "-DELROWS-":
            if sheet is None:
                popup_error("Leads sheet not initialized.")
            else:
                try:
                    sels = sheet.get_selected_rows() or []
                    if sels:
                        for r in sorted(sels, reverse=True):
                            try:
                                sheet.delete_rows(r, 1)
                            except Exception:
                                try:
                                    sheet.delete_rows(r)
                                except Exception:
                                    try:
                                        sheet.del_rows(r, 1)
                                    except Exception:
                                        pass
                        sheet.refresh()
                except Exception as e:
                    popup_error(f"Could not delete rows: {e}")
            refresh_fire_state()
            # NEW: refresh Emails Sent display after structure change
            try:
                window._refresh_emails_sent_column()
            except Exception:
                pass

        elif event == "-SAVECSV-":
            try:
                save_matrix_to_csv(matrix_from_sheet())
                window["-STATUS-"].update("Saved CSV")
            except Exception as e:
                window["-STATUS-"].update(f"Save error: {e}")

        # ---------------- Email Campaigns tab (NEW handlers) ----------------
        elif event in ("-CAMP_ADD_NEW-", "-CAMP_NEW-"):
            new_key = _camp_prompt_new_key()
            if new_key:
                # Update combo to include the new key (not yet saved)
                try:
                    keys = list_campaign_keys()
                except Exception:
                    keys = []
                if new_key not in keys:
                    keys.append(new_key)
                if "-CAMP_KEY-" in window.AllKeysDict:
                    window["-CAMP_KEY-"].update(values=keys, value=new_key)
                camp_write_editor(window, _camp_blank_steps(), _camp_default_settings())
                _camp_toggle_empty_vs_editor(window, True)
                try:
                    window["-CAMP_STATUS-"].update("New campaign ready. Fill in fields and click Save.")
                except Exception:
                    pass

        elif event == "-CAMP_LOAD-" or event == "-CAMP_KEY-":
            key = (values.get("-CAMP_KEY-") or "").strip()
            if not key:
                continue
            _camp_load_into_editor_by_key(window, key)
            try:
                window["-CAMP_STATUS-"].update(f"Loaded '{key}'.")
            except Exception:
                pass

        elif event == "-CAMP_SAVE-":
            key = (values.get("-CAMP_KEY-") or "").strip()
            if not key:
                popup_error("Provide a campaign name first (use New).")
                continue
            steps, settings = camp_read_editor(window)
            try:
                save_campaign_by_key(key, steps, settings)
                window["-CAMP_STATUS-"].update("Saved ✓")
            except Exception as e:
                window["-CAMP_STATUS-"].update(f"Save error: {e}")
            _camp_refresh_combo_and_table(window)
            _camp_toggle_empty_vs_editor(window, True)

        elif event == "-CAMP_DELETE-":
            key = (values.get("-CAMP_KEY-") or "").strip()
            if not key:
                continue
            yn = sg.popup_yes_no(f"Delete campaign '{key}'? This cannot be undone.")
            if yn == "Yes":
                try:
                    delete_campaign_by_key(key)
                    window["-CAMP_STATUS-"].update("Deleted ✓")
                except Exception as e:
                    window["-CAMP_STATUS-"].update(f"Delete error: {e}")
                _camp_refresh_combo_and_table(window)
                # If none left, hide editor
                try:
                    keys_left = list_campaign_keys()
                except Exception:
                    keys_left = []
                if not keys_left:
                    _camp_toggle_empty_vs_editor(window, False)
                else:
                    # Load first remaining
                    _camp_load_into_editor_by_key(window, keys_left[0])

        elif event == "-CAMP_RELOAD-":
            key = (values.get("-CAMP_KEY-") or "").strip()
            if key:
                _camp_load_into_editor_by_key(window, key)
                try:
                    window["-CAMP_STATUS-"].update("Reloaded ✓")
                except Exception:
                    pass

        elif event == "-CAMP_RESET-":
            camp_write_editor(window, _camp_blank_steps(), _camp_default_settings())
            try:
                window["-CAMP_STATUS-"].update("Reset fields. (Not saved yet)")
            except Exception:
                pass

        elif event == "-CAMP_REFRESH_LIST-":
            _camp_refresh_combo_and_table(window)
            try:
                window["-CAMP_STATUS-"].update("Refreshed ✓")
            except Exception:
                pass

        elif event == "-CAMP_TABLE-":
            # Load the selected campaign into the editor
            try:
                sel = values.get("-CAMP_TABLE-", [])
                if sel:
                    idx = sel[0]
                    keys = list_campaign_keys()
                    rows = [summarize_campaign_for_table(k) for k in keys]
                    if 0 <= idx < len(rows):
                        key = rows[idx][0]
                        if "-CAMP_KEY-" in window.AllKeysDict:
                            window["-CAMP_KEY-"].update(value=key)
                        _camp_load_into_editor_by_key(window, key)
                        try:
                            window["-CAMP_STATUS-"].update(f"Loaded '{key}' from list.")
                        except Exception:
                            pass
            except Exception:
                pass

        # ---------------- Email Results tab ----------------
        elif event == "-SYNC-":
            if not require_pywin32():
                window["-RS_STATUS-"].update("pywin32 missing (Outlook COM). Install pywin32.")
                continue
            try:
                # NOTE: correct key name is -LOOKBACK- (not -LOOKBACK_)
                days = int((values.get("-LOOKBACK-", "") or "60").strip())
            except Exception:
                days = 60
            window["-RS_STATUS-"].update("Syncing…")
            try:
                s_count, r_count = outlook_sync_results(days)
                rows = load_results_rows_sorted()
                data = [[r.get("Ref", ""), r.get("Email", ""), r.get("Company", ""), r.get("Industry", ""),
                         r.get("DateSent", ""), r.get("DateReplied", ""), r.get("Status", ""), r.get("Subject", "")] for r in rows]
                window["-RSTABLE-"].update(values=data)
                refresh_results_metrics()
                window["-RS_STATUS-"].update(f"Synced: {s_count} sent refs; {r_count} replies.")
                # After syncing, also refresh scoreboards (emails / warm leads may change)
                try:
                    update_scoreboards(window)
                except Exception:
                    pass
                # NEW: refresh Emails Sent column/colors after results changed
                try:
                    window._refresh_emails_sent_column()
                except Exception:
                    pass
            except Exception as e:
                window["-RS_STATUS-"].update(f"Sync error: {e}")

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
                data = [[r.get("Ref", ""), r.get("Email", ""), r.get("Company", ""), r.get("Industry", ""),
                         r.get("DateSent", ""), r.get("DateReplied", ""), r.get("Status", ""), r.get("Subject", "")] for r in rows]
                window["-RSTABLE-"].update(values=data)
                refresh_results_metrics()
                try:
                    refresh_customer_analytics()
                    update_scoreboards(window)
                except Exception:
                    pass

        elif event == "-MARK_GRAY-":
            sels = values.get("-RSTABLE-", [])
            if sels:
                idx = sels[0]
                rows = load_results_rows_sorted()
                if 0 <= idx < len(rows):
                    set_status(rows[idx]["Ref"], "gray")
                rows = load_results_rows_sorted()
                data = [[r.get("Ref", ""), r.get("Email", ""), r.get("Company", ""), r.get("Industry", ""),
                         r.get("DateSent", ""), r.get("DateReplied", ""), r.get("Status", ""), r.get("Subject", "")] for r in rows]
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
                data = [[r.get("Ref", ""), r.get("Email", ""), r.get("Company", ""), r.get("Industry", ""),
                         r.get("DateSent", ""), r.get("DateReplied", ""), r.get("Status", ""), r.get("Subject", "")] for r in rows]
                window["-RSTABLE-"].update(values=data)
                refresh_results_metrics()

        # ---------------- Dialer tab ----------------
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
                                    if xv:
                                        dial_sheet.MT.xview_moveto(xv[0])
                                    if yv:
                                        dial_sheet.MT.yview_moveto(yv[0])
                                except Exception:
                                    pass

                        dialer_save_call(base, outcome, note_text)
                        if outcome == "red":
                            add_no_interest(base, note_text, no_contact_flag=0, source="Dialer")
                        elif outcome == "gray":
                            filled = 0
                            row_vals2 = dial_sheet.get_row_data(r) or []
                            for k in range(cols["first_note"], cols["last_note"] + 1):
                                if k < len(row_vals2) and (row_vals2[k] or "").strip():
                                    filled += 1
                            if filled >= 8:
                                add_no_interest(base, "No Contact after 8 calls. " + note_text, no_contact_flag=1, source="Dialer")

                        window["-DIAL_MSG-"].update("Saved ✓")
                        window["-DIAL_NOTE-"].update("")
                        state["outcome"] = None
                        state["note_col_by_row"].pop(r, None)

                        if outcome in ("green", "red"):
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
                            _save_dialer_grid_to_csv()
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
                            _save_dialer_grid_to_csv()
                            new_row = dialer_move_to_next_row(dial_sheet, r)
                            _set_working_row(new_row)

                        _update_confirm_button()

                    except Exception as e:
                        window["-DIAL_MSG-"].update(f"Save error: {e}")

        elif event == "-DIAL_ADD100-":
            try:
                add = [[""] * len(HEADER_FIELDS) + ["○", "○", "○"] + ([""] * 8) for _ in range(100)]
                try:
                    cur = dial_sheet.get_sheet_data() or []
                except Exception:
                    cur = []
                dial_sheet.set_sheet_data((cur or []) + add)
                dial_sheet.refresh()
                _save_dialer_grid_to_csv()
            except Exception as e:
                window["-DIAL_MSG-"].update(f"Add rows error: {e}")

        # ---------------- Warm tab ----------------
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
            if warm_state["outcome"] not in ("green", "gray", "red"):
                window["-WARM_STATUS-"].update("Choose an outcome first.")
                continue
            note_text = _warm_note_text()
            if not note_text:
                window["-WARM_STATUS-"].update("Type a note.")
                continue

            try:
                row_vals = warm_sheet.get_row_data(r) or []
            except Exception:
                row_vals = []

            # Normalize Cost ($) only; DO NOT touch creation date ("First Contact")
            _warm_cost_normalize_in_row(r)

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

            try:
                warm_sheet.refresh()
            except Exception:
                pass
            _save_warm_grid_to_csv_v2()

            outcome = warm_state["outcome"]
            wmap = {WARM_V2_FIELDS[i]: (row_vals[i] if i < len(WARM_V2_FIELDS) else "") for i in range(len(WARM_V2_FIELDS))}
            base = {
                "Email": wmap.get("Email", ""),
                "First Name": (wmap.get("Prospect Name", "") or "").split(" ")[0] if wmap.get("Prospect Name") else "",
                "Last Name": " ".join((wmap.get("Prospect Name", "") or "").split(" ")[1:]) if wmap.get("Prospect Name") else "",
                "Company": wmap.get("Company", ""),
                "Industry": wmap.get("Industry", ""),
                "Phone": wmap.get("Phone #", ""),
                "City": (wmap.get("Location", "") or "").split(",")[0] if wmap.get("Location") else "",
                "State": (wmap.get("Location", "") or "").split(",")[-1].strip() if wmap.get("Location") and "," in wmap.get("Location") else "",
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
                try:
                    refresh_customer_analytics()
                except Exception:
                    pass
            except Exception as e:
                window["-WARM_STATUS-"].update(f"Save error: {e}")

        elif event == "-WARM_EXPORT-":
            path = sg.popup_get_file("Save warm_leads.csv", save_as=True, default_extension=".csv",
                                     file_types=(("CSV", "*.csv"),), no_window=True)
            if path:
                try:
                    _save_warm_grid_to_csv_v2()
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
                try:
                    refresh_customer_analytics()
                except Exception:
                    pass
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
            warm_row = {WARM_V2_FIELDS[i]: (row[i] if i < len(WARM_V2_FIELDS) else "") for i in range(len(WARM_V2_FIELDS))}
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

            cust = {h: "" for h in CUSTOMER_FIELDS}
            cust["Company"] = warm_row.get("Company", "")
            cust["Prospect Name"] = warm_row.get("Prospect Name", "")
            cust["Phone #"] = warm_row.get("Phone #", "")
            cust["Email"] = warm_row.get("Email", "")
            cust["Location"] = warm_row.get("Location", "")
            cust["Industry"] = warm_row.get("Industry", "")
            cust["Google Reviews"] = warm_row.get("Google Reviews", "")
            cust["Rep"] = warm_row.get("Rep", "")
            cust["Samples?"] = warm_row.get("Samples?", "")
            cust["Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            if "Opening Order $" in CUSTOMER_FIELDS:
                cust["Opening Order $"] = amt
            if "Customer Since" in CUSTOMER_FIELDS:
                cust["Customer Since"] = datetime.now().strftime("%Y-%m-%d")

            if "Notes" in CUSTOMER_FIELDS:
                last_note = ""
                for i in range(15, 0, -1):
                    v = warm_row.get(f"Call {i}", "")
                    if (v or "").strip():
                        last_note = v
                        break
                cust["Notes"] = last_note

            # ---- Lat/Lon: pre-fill if possible, else leave blank (autosave/geocoder or sidecar merge will fill) ----
            if ("Lat" in CUSTOMER_FIELDS) or ("Lon" in CUSTOMER_FIELDS):
                lat_val, lon_val = "", ""

                # Try to build the best address we can from warm_row
                addr_parts = []
                for key in ("Address", "City", "State", "ZIP"):
                    val = warm_row.get(key, "")
                    if (val or "").strip():
                        addr_parts.append(str(val).strip())
                if addr_parts:
                    addr = ", ".join(addr_parts)
                else:
                    # Fallback to Location field
                    addr = (cust.get("Location") or "").strip()

                # If the geocode sidecar is available, try to look it up now
                try:
                    key = addr.lower().strip()
                    # __geo_sidecar is defined in Chunk 5 when the app starts
                    if key and "__geo_sidecar" in globals() and isinstance(globals()["__geo_sidecar"], dict):
                        hit = globals()["__geo_sidecar"].get(key)
                        if hit and isinstance(hit, (tuple, list)) and len(hit) >= 2:
                            lat_val = f"{float(hit[0]):.6f}"
                            lon_val = f"{float(hit[1]):.6f}"
                except Exception:
                    # It's fine to miss; merge helper or autosave will fill later
                    pass

                if "Lat" in CUSTOMER_FIELDS:
                    cust["Lat"] = lat_val
                if "Lon" in CUSTOMER_FIELDS:
                    cust["Lon"] = lon_val

            try:
                existing = []
                if CUSTOMERS_PATH.exists():
                    with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
                        rdr = csv.DictReader(f)
                        for r in rdr:
                            existing.append([r.get(h, "") for h in CUSTOMER_FIELDS])
                existing.append([cust.get(h, "") for h in CUSTOMER_FIELDS])
                _backup(CUSTOMERS_PATH)
                _atomic_write_csv(CUSTOMERS_PATH, CUSTOMER_FIELDS, existing)
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
                try:
                    refresh_customer_analytics()
                except Exception:
                    pass
            except Exception as e:
                window["-WARM_STATUS-"].update(f"Move error (customers): {e}")

        # ---------------- Customers tab ----------------
        elif event == "-CUST_ADD50-":
            try:
                add = [[""] * len(CUSTOMER_FIELDS) for _ in range(50)]
                cur = customer_sheet.get_sheet_data() or []
                customer_sheet.set_sheet_data((cur or []) + add)
                customer_sheet.refresh()
                _save_customers_grid_to_csv()
                try:
                    refresh_customer_analytics()
                except Exception:
                    pass
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

            try:
                row_vals = customer_sheet.get_row_data(r_sel) or []
            except Exception:
                row_vals = []
            idx_company = _cust_idx("Company", 0)
            company = row_vals[idx_company] if idx_company is not None and idx_company < len(row_vals) else ""
            if not (company or "").strip():
                window["-CUST_STATUS-"].update("Company is required on the selected row.")
                continue

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

            stats = compute_customer_order_stats(company)
            idx_fod  = _cust_idx("First Order")
            idx_lod  = _cust_idx("Last Order")
            idx_cltv = _cust_idx("CLTV")
            idx_days = _cust_idx("Days")
            idx_spd  = _cust_idx("Sales/Day")

            def _set_if(idx, val):
                if idx is None:
                    return
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

            _save_customers_grid_to_csv()
            window["-CUST_STATUS-"].update("Order added ✓")

            try:
                refresh_customer_analytics()
            except Exception:
                pass

        elif event == "-CUST_SAVE-":
            _save_customers_grid_to_csv()
            try:
                refresh_customer_analytics()
            except Exception:
                pass

        elif event == "-CUST_EXPORT-":
            path = sg.popup_get_file("Save customers.csv", save_as=True, default_extension=".csv",
                                     file_types=(("CSV", "*.csv"),), no_window=True)
            if path:
                try:
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
                try:
                    refresh_customer_analytics()
                except Exception:
                    pass
            except Exception as e:
                window["-CUST_STATUS-"].update(f"Reload error: {e}")

        # ---------------- Map tab ----------------
        elif event == "-OPEN_MAP-":
            _open_customer_map(window)

        # Keep analytics and fire button fresh
        refresh_fire_state()
        try:
            refresh_customer_analytics()
        except Exception:
            pass

    window.close()
# ===== CHUNK 6 / 7 — END =====
# ===== CHUNK 7 / 7 — START =====
# ============================================================
# Email Campaigns: enroll helpers and drafting logic
# (integrates with hourly task from Chunk 5 via _draft_next_stage_stub)
# Uses the multi-campaign (per niche/industry) storage wired in earlier:
#   load_campaign_by_key(key)  -> (steps, settings)
#   normalize_campaign_steps(steps) with keys: subject, body, delay_days
# ============================================================

# ---------- Campaign enrollment & maintenance ----------
def campaigns_enroll(ref_short: str, email: str, company: str,
                     campaign_key: str = "default",
                     divert_to_dialer: bool = True):
    """
    Add a row to campaigns.csv if not already present. Stage starts at 0.
    Stage meanings:
      0: queued before first DateSent is known
      1: E1 sent (DateSent exists); next due -> E2
      2: E2 sent; next due -> E3
      3: E3 sent; campaign complete (hourly task will remove; may divert to Dialer)
    """
    ensure_campaigns_file()
    ref_l = (ref_short or "").strip().lower()
    rows = []
    exists = False
    if CAMPAIGNS_PATH.exists():
        with CAMPAIGNS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                if (r.get("Ref","") or "").strip().lower() == ref_l:
                    exists = True
                rows.append(r)
    if not exists:
        rows.append({
            "Ref": ref_short or "",
            "Email": email or "",
            "Company": company or "",
            "CampaignKey": campaign_key or "default",
            "Stage": "0",
            "DivertToDialer": "1" if divert_to_dialer else "0",
        })
        _campaigns_write_rows(rows)

def campaigns_is_enrolled(ref_short: str) -> bool:
    ensure_campaigns_file()
    ref_l = (ref_short or "").strip().lower()
    if not CAMPAIGNS_PATH.exists():
        return False
    with CAMPAIGNS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            if (r.get("Ref","") or "").strip().lower() == ref_l:
                return True
    return False

# ---------- Outlook single draft helper ----------
def _ensure_outlook_folder_drafts_sub(session, name: str):
    """Returns a subfolder under Drafts named `name` (creates if missing)."""
    store = pick_store(session)
    drafts_root = store.GetDefaultFolder(16)  # olFolderDrafts
    for i in range(1, drafts_root.Folders.Count + 1):
        f = drafts_root.Folders.Item(i)
        if (f.Name or "").lower() == (name or "").strip().lower():
            return f
    return drafts_root.Folders.Add(name or "Death Star")

def _draft_one_outlook(ref_short: str, email: str, subj_text: str, body_text: str):
    """Create a single Outlook draft under the Death Star subfolder."""
    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    session = outlook.GetNamespace("MAPI")
    target_folder = _ensure_outlook_folder_drafts_sub(session, DEATHSTAR_SUBFOLDER)
    body_html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"></head>
    <body style="margin:0;padding:0;">
      <div style="font-family:Segoe UI, Arial, sans-serif; font-size:14px; line-height:1.5; color:#111;">
        {blocks_to_html(body_text)}
        <!-- ref:{ref_short} -->
      </div>
    </body></html>"""
    msg = target_folder.Items.Add("IPM.Note")
    msg.To = email or ""
    msg.Subject = f"{subj_text} [ref:{ref_short}]"
    msg.BodyFormat = 2
    msg.HTMLBody = body_html
    msg.Save()
    # update results.csv cache for visibility
    try:
        upsert_result(ref_short, email or "", "", "", subj_text)
    except Exception:
        pass
    return True

# ---------- Placeholders & row dict ----------
def _rowdict_for_placeholders(results_row: dict):
    """Prefer the original lead row (for First Name, etc.)."""
    lead = _lead_row_from_email_company(results_row.get("Email",""), results_row.get("Company",""))
    if lead:
        return lead
    d = {h: "" for h in HEADER_FIELDS}
    d["Email"] = results_row.get("Email","")
    d["Company"] = results_row.get("Company","")
    d["Industry"] = results_row.get("Industry","")
    return d

# ---------- Due-date logic anchored to DateSent ----------
def _days_since(dt) -> int:
    if not dt:
        return 0
    try:
        return max(0, (datetime.now() - dt).days)
    except Exception:
        return 0

def _get_step_delays_for_key(campaign_key: str):
    """
    Read delays from the per-key campaign definition.
    Returns tuple: (delay_e2_days, delay_e3_days)
    where:
      delay_e2_days = delay_days of step 2
      delay_e3_days = delay_days of step 3
    """
    try:
        steps, _settings = load_campaign_by_key(campaign_key or "default")
        steps = normalize_campaign_steps(steps)
        d2 = int(str(steps[1].get("delay_days", 0)).strip() or "0")  # step 2 delay
        d3 = int(str(steps[2].get("delay_days", 0)).strip() or "0")  # step 3 delay
        return (max(0, d2), max(0, d3))
    except Exception:
        return (3, 7)  # safe fallback

def _is_due_for_next(results_row: dict, next_stage: int, campaign_key: str) -> bool:
    """
    Decide if it's time to draft E2/E3 anchored to last DateSent.
    - next_stage 2: days_since(DateSent) >= delay(step2)
    - next_stage 3: days_since(DateSent) >= delay(step2) + delay(step3)
    """
    sent_dt = _results_sent_dt(results_row)
    if not sent_dt:
        return False
    d2, d3 = _get_step_delays_for_key(campaign_key)
    elapsed = _days_since(sent_dt)
    if next_stage == 2:
        return elapsed >= d2
    if next_stage == 3:
        return elapsed >= (d2 + d3)
    return False

def _get_subject_body_for_stage(campaign_key: str, stage_num: int):
    """
    Pull subject/body for the given stage (1..3) from the selected campaign.
    Returns (subject, body). Falls back to simple templates if missing.
    """
    steps, _settings = load_campaign_by_key(campaign_key or "default")
    steps = normalize_campaign_steps(steps)
    idx = max(1, min(3, stage_num)) - 1
    subj = (steps[idx].get("subject") or "").strip()
    body = (steps[idx].get("body") or "").strip()
    if not subj:
        subj = ["Quick hello for {Company}",
                "Following up for {Company}",
                "Worth a quick chat about {Company}?"][idx]
    if not body:
        body_defaults = [
            "Hi {First Name},\n\nWanted to share something relevant to {Company}.\n\nCheers,\nMe",
            "Hi {First Name},\n\nCircling back in case my note missed you.\n\nBest,\nMe",
            "Hi {First Name},\n\nLast follow-up from me—open to a quick call?\n\nThanks,\nMe",
        ]
        body = body_defaults[idx]
    return subj, body

# ---------- Missing helpers for multi-campaign migration ----------
def _read_results_by_ref():
    """ref_lower -> results row dict"""
    rows = load_results_rows_sorted()
    return { (r.get("Ref","") or "").lower(): r for r in rows }

def _results_replied(r: dict) -> bool:
    return bool(_parse_any_datetime(r.get("DateReplied","")))

def _results_sent_dt(r: dict):
    return _parse_any_datetime(r.get("DateSent",""))

# ---------- Public: replaces stub referenced by hourly task ----------
def draft_next_stage_from_config(ref: str, email: str, company: str,
                                 campaign_key: str, next_stage: int) -> bool:
    """
    Returns True iff we created a draft for `next_stage` (1..3).
    Preconditions:
      - pywin32 available
      - Ref exists in results.csv
      - Not replied
      - For stages > 1, delay satisfied based on DateSent and campaign delays
    """
    try:
        if not require_pywin32():
            return False

        res_map = _read_results_by_ref()
        r = res_map.get((ref or "").lower())
        if not r:
            return False
        if _results_replied(r):
            return False

        # Ensure a target email; prefer arg, fall back to results.csv
        target_email = (email or "").strip() or (r.get("Email","") or "").strip()
        if not target_email:
            return False

        # Stage 1 drafts typically created at send time; only gate delays for 2/3
        if next_stage in (2, 3):
            if not _is_due_for_next(r, next_stage, campaign_key):
                return False

        subj_tpl, body_tpl = _get_subject_body_for_stage(campaign_key, next_stage)
        rowd = _rowdict_for_placeholders(r)
        subj_text = apply_placeholders(subj_tpl, rowd)
        body_text = apply_placeholders(body_tpl, rowd)

        _draft_one_outlook(ref, target_email, subj_text, body_text)
        return True
    except Exception:
        return False

# Hot-swap the stub used in CHUNK 5 so the hourly runner actually drafts
try:
    globals()["_draft_next_stage_stub"] = draft_next_stage_from_config
except Exception:
    pass

# ---------- Convenience enrollment helpers ----------
def campaigns_enroll_from_results_row(res_row: dict, campaign_key="default", divert_to_dialer=True):
    ref = res_row.get("Ref","") or ""
    email = res_row.get("Email","") or ""
    company = res_row.get("Company","") or ""
    if not ref:
        return
    campaigns_enroll(ref, email, company, campaign_key, divert_to_dialer)

def campaigns_bulk_enroll_from_status(status="gray", campaign_key="default", divert_to_dialer=True, max_rows=2000):
    rows = load_results_rows_sorted()
    count = 0
    for r in rows:
        if count >= max_rows:
            break
        if _results_replied(r):
            continue
        if (r.get("Status","") or "").strip().lower() == (status or "").lower():
            campaigns_enroll_from_results_row(r, campaign_key, divert_to_dialer)
            count += 1
    return count


# ============================================================
# Entry (must be last in the file, after all chunks)
# ============================================================
if __name__ == "__main__":
    import traceback
    try:
        main()
    except Exception as e:
        # Print full traceback to console so we see the exact file/line
        traceback.print_exc()
        try:
            popup_error(f"Fatal error starting app: {e}")
        except Exception:
            pass
        input("\n\n[ERROR] Press Enter to exit...")

# ===== CHUNK 7 / 7 — END =====




