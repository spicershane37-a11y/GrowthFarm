# gf_store.py
# Unified storage layer for GrowthFarm
# - App folder & paths
# - Email Leads / Warm v2 / Dialer (grid + results) / Customers / Orders / Results
# - Templates & Multi-campaign INI
# - Per-ref campaign state (campaigns.csv)
# - Helper utilities (placeholders, money/date parsing)
#
# Safe to import from any UI module. Keep logic elsewhere.

from __future__ import annotations

import csv
import os
from pathlib import Path
from datetime import datetime, date
import configparser
from typing import List, Dict, Iterable, Tuple, Optional
import re
import sys  # used to locate sidecar app.ini

# ----------------------------
# App directory & file paths
# ----------------------------
# Reads a sidecar app.ini next to the launched script/EXE to decide the roaming data folder:
#   [app]
#   data_dir=GrowthFarm Test Account
#
# If app.ini is missing, default to "GrowthFarm".

def _sidecar_ini_path() -> Path:
    try:
        return Path(sys.argv[0]).resolve().parent / "app.ini"
    except Exception:
        return Path.cwd() / "app.ini"

def _data_dir_name_from_ini(default_name: str = "GrowthFarm") -> str:
    ini = _sidecar_ini_path()
    if ini.exists():
        cfg = configparser.ConfigParser()
        try:
            cfg.read(ini, encoding="utf-8")
            name = (cfg.get("app", "data_dir", fallback=default_name) or "").strip()
            if name:
                return name
        except Exception:
            pass
    return default_name

APP_NAME: str = _data_dir_name_from_ini("GrowthFarm")
# Store under Roaming on Windows; if APPDATA missing (non-Windows), fall back to Home.
APP_DIR: Path = Path(os.environ.get("APPDATA", str(Path.home()))) / APP_NAME

def get_app_dir() -> Path:
    """Return the resolved roaming data directory as a Path (used by login/import)."""
    return APP_DIR

# Core files (grids/state)
EMAIL_LEADS_PATH   = APP_DIR / "email_leads.csv"
RESULTS_PATH       = APP_DIR / "results.csv"
WARM_LEADS_PATH    = APP_DIR / "warm_leads.csv"
NO_INTEREST_PATH   = APP_DIR / "no_interest.csv"

CUSTOMERS_PATH     = APP_DIR / "customers.csv"
ORDERS_PATH        = APP_DIR / "orders.csv"

# Dialer
DIALER_RESULTS_PATH = APP_DIR / "dialer_results.csv"  # call log
DIALER_LEADS_PATH   = APP_DIR / "dialer_leads.csv"    # dialer grid storage

# Templates / campaigns (config)
TEMPLATES_INI      = APP_DIR / "templates.ini"
CAMPAIGNS_INI      = APP_DIR / "campaigns.ini"

# Per-ref campaign state (stage tracking)
CAMPAIGNS_PATH     = APP_DIR / "campaigns.csv"        # columns: CAMPAIGNS_HEADERS below

STATE_PATH         = APP_DIR / "state.txt"            # optional “seen” set for email drafts, etc.

# ----------------------------
# Default headers / templates
# ----------------------------
HEADER_FIELDS = [
    "Email","First Name","Last Name","Company","Industry","Phone",
    "Address","City","State","Reviews","Website","Notes",
]

# Warm v2 (flattened; 15 “Call i” slots)
WARM_V2_FIELDS = [
    "Company","Prospect Name","Phone #","Email",
    "Location","Industry","Google Reviews","Rep","Samples?","Timestamp",
    "Cost ($)",
    *[f"Call {i}" for i in range(1, 16)],
    "First Contact",
]

# Back-compat: expose WARM_FIELDS as an alias to WARM_V2_FIELDS
try:
    WARM_FIELDS
except NameError:
    WARM_FIELDS = WARM_V2_FIELDS

# Customers (Lat/Lon for Map + derived fields)
CUSTOMER_FIELDS = [
    "Company","Prospect Name","Phone #","Email","Industry",
    "Address","City","State","ZIP","Lat","Lon",
    "CLTV","Sales/Day","Reorder?","First Order","Last Order","Days",
    "First Contact","Days To Close","Sku's","Notes",
]

DEFAULT_TEMPLATES = {
    "default": (
        "Hey {First Name},\n\n"
        "My name is YOUR NAME with YOUR COMPANY. We help {Industry} MAIN GOAL. "
        "If it’s useful, I can share examples or send over a couple of samples.\n\n"
        "Thanks,\nYOUR NAME\nYOUR COMPANY\nPHONE\nWEBSITE"
    )
}
DEFAULT_SUBJECTS = {"default": "Quick intro from YOUR COMPANY"}
DEFAULT_MAP = {}

# Campaign defaults (3 steps)
DEFAULT_CAMPAIGN_STEPS = [
    {"enabled": True,  "subject": DEFAULT_SUBJECTS["default"], "body": DEFAULT_TEMPLATES["default"], "delay_days": 0},
    {"enabled": False, "subject": "", "body": "", "delay_days": 3},
    {"enabled": False, "subject": "", "body": "", "delay_days": 7},
]
DEFAULT_CAMPAIGN_SETTINGS = {
    "send_to_dialer_after": "1",
    "auto_sync_outlook": "0",
    "hourly_campaign_runner": "1",
}

# Per-ref CSV schema
CAMPAIGNS_HEADERS = ["Ref","Email","Company","CampaignKey","Stage","DivertToDialer"]

# ----------------------------
# Small utilities
# ----------------------------
def _atomic_write_csv(path: Path, headers: List[str], rows: Iterable[Iterable[str]]):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    # FIX: newline must be "" (was "}")
    with tmp.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        for row in rows:
            w.writerow(list(row)[:len(headers)])
    tmp.replace(path)

def _read_csv_matrix(path: Path, headers: List[str]) -> List[List[str]]:
    if not path.exists():
        return []
    out = []
    with path.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        if not rdr.fieldnames:
            return []
        for r in rdr:
            out.append([r.get(h, "") for h in headers])
    return out

def _write_csv_matrix(path: Path, headers: List[str], matrix: List[List[str]]):
    rows = []
    for row in matrix:
        rows.append([row[i] if i < len(row) else "" for i in range(len(headers))])
    _atomic_write_csv(path, headers, rows)

def _ensure_file_with_header(path: Path, headers: List[str]):
    if not path.exists():
        _atomic_write_csv(path, headers, [])

# Backups
BACKUP_DIR = APP_DIR / "_backups"
def _backup(path: Path):
    try:
        if path.exists() and path.is_file():
            BACKUP_DIR.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            bak = BACKUP_DIR / f"{path.stem}.{stamp}{path.suffix}.bak"
            with path.open("rb") as s, bak.open("wb") as d:
                d.write(s.read())
    except Exception:
        pass

# ----------------------------
# Public: ensure & basic IO
# ----------------------------
def ensure_no_interest_file() -> None:
    """Ensure no_interest.csv exists with the canonical schema."""
    if not NO_INTEREST_PATH.exists():
        NO_INTEREST_PATH.parent.mkdir(parents=True, exist_ok=True)
        with NO_INTEREST_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow([
                "Timestamp",
                "Email","First Name","Last Name",
                "Company","Industry","Phone",
                "City","State","Website",
                "Note","Source","NoContactFlag"
            ])

def append_no_interest(row_dict: Dict[str,str], note: str, no_contact_flag: int, source: str) -> None:
    """Append a single no-interest record (single source of truth)."""
    ensure_no_interest_file()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with NO_INTEREST_PATH.open("a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            ts,
            row_dict.get("Email",""), row_dict.get("First Name",""), row_dict.get("Last Name",""),
            row_dict.get("Company",""), row_dict.get("Industry",""), row_dict.get("Phone",""),
            row_dict.get("City",""), row_dict.get("State",""), row_dict.get("Website",""),
            note, source, int(no_contact_flag or 0)
        ])

def ensure_app_files():
    """Create APP_DIR and seed files if missing."""
    APP_DIR.mkdir(parents=True, exist_ok=True)

    _ensure_file_with_header(EMAIL_LEADS_PATH, HEADER_FIELDS)
    _ensure_file_with_header(RESULTS_PATH, ["Ref","Email","Company","Industry","DateSent","DateReplied","Status","Subject"])
    _ensure_file_with_header(WARM_LEADS_PATH, WARM_V2_FIELDS)

    # Use canonical no_interest schema
    ensure_no_interest_file()

    _ensure_file_with_header(CUSTOMERS_PATH, CUSTOMER_FIELDS)
    _ensure_file_with_header(ORDERS_PATH, ["Company","Order Date","Amount"])

    # dialer files
    ensure_dialer_files()       # call log
    ensure_dialer_leads_file()  # dialer grid

    # templates.ini (very small)
    if not TEMPLATES_INI.exists():
        cfg = configparser.ConfigParser()
        cfg["templates"] = DEFAULT_TEMPLATES
        cfg["subjects"]  = DEFAULT_SUBJECTS
        cfg["map"]       = DEFAULT_MAP
        with TEMPLATES_INI.open("w", encoding="utf-8") as f:
            cfg.write(f)

    # campaigns.ini / campaigns.csv
    if not CAMPAIGNS_INI.exists():
        save_campaigns_ini(DEFAULT_CAMPAIGN_STEPS, DEFAULT_CAMPAIGN_SETTINGS)
    if not CAMPAIGNS_PATH.exists():
        _ensure_file_with_header(CAMPAIGNS_PATH, CAMPAIGNS_HEADERS)

    if not STATE_PATH.exists():
        STATE_PATH.write_text("", encoding="utf-8")

# Back-compat alias (older modules import this name)
ensure_app_dirs_and_files = ensure_app_files

# ----------------------------
# Email Leads (matrix IO)
# ----------------------------
def load_email_leads_matrix() -> List[List[str]]:
    return _read_csv_matrix(EMAIL_LEADS_PATH, HEADER_FIELDS)

def save_email_leads_matrix(matrix: List[List[str]]):
    _write_csv_matrix(EMAIL_LEADS_PATH, HEADER_FIELDS, matrix)

# ----------------------------
# Results (dict helpers)
# ----------------------------
def load_results_rows_sorted() -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    if RESULTS_PATH.exists():
        with RESULTS_PATH.open("r", encoding="utf-8", newline="") as f:
            rows = list(csv.DictReader(f))
    def sk(r): return (r.get("DateReplied",""), r.get("DateSent",""))
    rows.sort(key=sk, reverse=True)
    return rows

def upsert_result(ref_short: str, email: str, company: str, industry: str, subject: str,
                  sent_dt: str = "", replied_dt: str = ""):
    rows = load_results_rows_sorted()
    idx = next((i for i, x in enumerate(rows) if x.get("Ref") == ref_short), None)
    rec = {
        "Ref": ref_short, "Email": email or "", "Company": company or "", "Industry": industry or "",
        "DateSent": sent_dt or (rows[idx]["DateSent"] if idx is not None else ""),
        "DateReplied": replied_dt or (rows[idx]["DateReplied"] if idx is not None else ""),
        "Status": rows[idx]["Status"] if idx is not None else "",
        "Subject": subject or (rows[idx]["Subject"] if idx is not None else ""),
    }
    if idx is None: rows.append(rec)
    else: rows[idx] = rec
    _atomic_write_csv(
        RESULTS_PATH,
        ["Ref","Email","Company","Industry","DateSent","DateReplied","Status","Subject"],
        ([r.get(h,"") for h in ["Ref","Email","Company","Industry","DateSent","DateReplied","Status","Subject"]] for r in rows)
    )

def set_status(ref_short: str, status: str):
    rows = load_results_rows_sorted()
    for r in rows:
        if r.get("Ref") == ref_short:
            r["Status"] = status
            break
    _atomic_write_csv(
        RESULTS_PATH,
        ["Ref","Email","Company","Industry","DateSent","DateReplied","Status","Subject"],
        ([r.get(h,"") for h in ["Ref","Email","Company","Industry","DateSent","DateReplied","Status","Subject"]] for r in rows)
    )

# ----------------------------
# Warm Leads (matrix IO) + migration
# ----------------------------
def ensure_warm_file():
    """Ensure warm_leads.csv exists under the v2 schema (WARM_V2_FIELDS)."""
    WARM_LEADS_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not WARM_LEADS_PATH.exists():
        _atomic_write_csv(WARM_LEADS_PATH, WARM_V2_FIELDS, [])
        return

    # migrate any existing file to v2 header (best-effort)
    try:
        with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            existing_fields = rdr.fieldnames or []
            rows = list(rdr) if rdr.fieldnames else []
    except Exception:
        existing_fields, rows = [], []

    if existing_fields and existing_fields != WARM_V2_FIELDS:
        _backup(WARM_LEADS_PATH)
        with WARM_LEADS_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=WARM_V2_FIELDS)
            w.writeheader()
            for r in rows:
                w.writerow({h: r.get(h, "") for h in WARM_V2_FIELDS})

def load_warm_leads_matrix_v2() -> List[List[str]]:
    ensure_warm_file()
    return _read_csv_matrix(WARM_LEADS_PATH, WARM_V2_FIELDS)

def save_warm_leads_matrix_v2(matrix: List[List[str]]):
    _backup(WARM_LEADS_PATH)
    _write_csv_matrix(WARM_LEADS_PATH, WARM_V2_FIELDS, matrix)

# ----------------------------
# Dialer (grid CSV + call log)
# ----------------------------
EMOJI_GREEN = "🙂"
EMOJI_GRAY  = "😐"
EMOJI_RED   = "🙁"      # canonical
EMOJI_RED_LEGACY = "☹️"  # legacy seen in older files

def ensure_dialer_files():
    if not DIALER_RESULTS_PATH.exists():
        with DIALER_RESULTS_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow([
                "Timestamp","Outcome","Email","First Name","Last Name","Company","Industry",
                "Phone","Address","City","State","Reviews","Website","Note"
            ])

def ensure_dialer_leads_file():
    if not DIALER_LEADS_PATH.exists():
        with DIALER_LEADS_PATH.open("w", encoding="utf-8", newline="") as f:
            hdr = HEADER_FIELDS + [EMOJI_GREEN, EMOJI_GRAY, EMOJI_RED] + [f"Note{i}" for i in range(1,9)]
            csv.writer(f).writerow(hdr)

def load_dialer_leads_matrix() -> List[List[str]]:
    """Load dialer grid rows. Accept legacy ☹️ header; normalize to 🙁 in memory."""
    ensure_dialer_leads_file()
    with DIALER_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
        raw = list(csv.reader(f))
    if not raw:
        return []
    hdr = raw[0]
    expected = HEADER_FIELDS + [EMOJI_GREEN, EMOJI_GRAY, EMOJI_RED] + [f"Note{i}" for i in range(1,9)]
    header_lookup = {h: (hdr.index(h) if h in hdr else None) for h in expected}
    if header_lookup[EMOJI_RED] is None and EMOJI_RED_LEGACY in hdr:
        header_lookup[EMOJI_RED] = hdr.index(EMOJI_RED_LEGACY)
    idx_map = [header_lookup[h] for h in expected]
    out = []
    for row in raw[1:]:
        new = []
        for i, idx in enumerate(idx_map):
            if idx is None:
                # outcome dots or notes fallback
                if len(HEADER_FIELDS) <= i < len(HEADER_FIELDS) + 3:
                    new.append("○")
                else:
                    new.append("")
            else:
                new.append(row[idx] if idx < len(row) else "")
        out.append(new)
    return out

def save_dialer_leads_matrix(matrix: List[List[str]]):
    _backup(DIALER_LEADS_PATH)
    headers = HEADER_FIELDS + [EMOJI_GREEN, EMOJI_GRAY, EMOJI_RED] + [f"Note{i}" for i in range(1,9)]
    _write_csv_matrix(DIALER_LEADS_PATH, headers, matrix)

# ----------------------------
# Customers & Orders
# ----------------------------
def _parse_date(s: str) -> Optional[date]:
    s = (s or "").strip()
    if not s:
        return None
    fmts = ("%Y-%m-%d","%m/%d/%Y","%m-%d-%Y","%Y/%m/%d","%m/%d/%y","%m-%d-%y","%m/%d","%m-%d")
    for fmt in fmts:
        try:
            d = datetime.strptime(s, fmt)
            if fmt in ("%m/%d","%m-%d"):
                d = d.replace(year=datetime.now().year)
            return d.date()
        except Exception:
            continue
    return None

def _money_to_float(val: str) -> float:
    s = (val or "").strip().replace(",", "").replace("$","")
    if not s:
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0

def _float_to_money(x) -> str:
    try:
        return f"{float(x):.2f}"
    except Exception:
        return ""

def compute_customer_order_stats(company: str) -> Dict[str, object]:
    """
    Aggregate order stats for a company.
    Returns: cltv (float), first_order_date (date|None), last_order_date (date|None),
             days_since_first (int|None), sales_per_day (float|None), order_count (int)
    """
    total = 0.0
    dates: List[date] = []
    order_count = 0
    if ORDERS_PATH.exists():
        with ORDERS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                if (r.get("Company","") or "").strip().lower() == (company or "").strip().lower():
                    order_count += 1
                    total += _money_to_float(r.get("Amount",""))
                    d = _parse_date(r.get("Order Date",""))
                    if d:
                        dates.append(d)
    first_d = min(dates) if dates else None
    last_d  = max(dates) if dates else None
    days_since_first = (datetime.now().date() - first_d).days if first_d else None
    sales_per_day = (total / float(days_since_first)) if days_since_first and days_since_first > 0 else None
    return {
        "cltv": total,
        "first_order_date": first_d,
        "last_order_date": last_d,
        "days_since_first": days_since_first,
        "sales_per_day": sales_per_day,
        "order_count": order_count,
    }

def ensure_customers_file():
    if not CUSTOMERS_PATH.exists():
        _atomic_write_csv(CUSTOMERS_PATH, CUSTOMER_FIELDS, [])
        return
    # migrate header if needed
    with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        old = rdr.fieldnames or []
        rows = list(rdr) if old else []
    if old and old != CUSTOMER_FIELDS:
        _backup(CUSTOMERS_PATH)
        _atomic_write_csv(CUSTOMERS_PATH, CUSTOMER_FIELDS, ([r.get(h,"") for h in CUSTOMER_FIELDS] for r in rows))

def _derive_customer_fields(row_dict: Dict[str,str]) -> Dict[str,str]:
    """
    Derive CLTV/Days/Sales/Day from orders.
    - Sales/Day is ONLY shown if order_count >= 2 (reorder).
    - Reorder? is auto-treated as true if order_count >= 2 (even if the cell is blank).
    - First/Last Order are backfilled from orders if present.
    """
    company = (row_dict.get("Company","") or "").strip()
    stats = compute_customer_order_stats(company) if company else {
        "cltv": 0.0, "first_order_date": None, "last_order_date": None,
        "days_since_first": None, "sales_per_day": None, "order_count": 0
    }

    # Prefer computed CLTV from orders, fall back to what's in the cell
    cltv_in_cell = _money_to_float(row_dict.get("CLTV",""))
    cltv = stats["cltv"] if (stats["cltv"] or 0) > 0 else cltv_in_cell

    out: Dict[str, str] = {}

    # Normalize CLTV to money format if > 0
    out["CLTV"] = _float_to_money(cltv) if cltv > 0 else ""

    # Backfill First/Last Order from orders
    if stats["first_order_date"]:
        out["First Order"] = stats["first_order_date"].strftime("%Y-%m-%d")
    if stats["last_order_date"]:
        out["Last Order"] = stats["last_order_date"].strftime("%Y-%m-%d")

    # Days since first order
    if stats["days_since_first"] is not None:
        out["Days"] = str(stats["days_since_first"])

    # Reorder logic: treat as true if >= 2 orders, or if the cell is already Yes
    order_count = int(stats.get("order_count", 0) or 0)
    reorder_cell = (row_dict.get("Reorder?","").strip().lower() in ("yes","y","1","true"))
    reorder_effective = reorder_cell or (order_count >= 2)

    # Sales/Day gating
    if reorder_effective and order_count >= 2 and stats["sales_per_day"] is not None:
        out["Sales/Day"] = _float_to_money(stats["sales_per_day"])
        # If the Reorder? cell was blank, auto-mark Yes for convenience
        if not row_dict.get("Reorder?","").strip():
            out["Reorder?"] = "Yes"
    else:
        # FORCE blank when < 2 orders or not a reorder
        out["Sales/Day"] = ""

    return out

def load_customers_matrix() -> List[List[str]]:
    ensure_customers_file()
    rows: List[List[str]] = []
    with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            try:
                r.update(_derive_customer_fields(r))
            except Exception:
                pass
            rows.append([r.get(h, "") for h in CUSTOMER_FIELDS])
    return rows

def save_customers_matrix(matrix: List[List[str]]):
    ensure_customers_file()
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

def append_order_row(company: str, order_date: str, amount: str):
    """Append an order, then recompute CLTV/Days/Sales/Day on the customer row."""
    company = (company or "").strip()
    if not company:
        raise ValueError("Company is required for orders.")
    # Normalize inputs
    d = _parse_date(order_date) or datetime.now().date()
    amt = _money_to_float(amount)
    # Write order
    with ORDERS_PATH.open("a", encoding="utf-8", newline="") as f:
        csv.writer(f).writerow([company, d.strftime("%Y-%m-%d"), _float_to_money(amt)])
    # Recompute + update customer row
    stats = compute_customer_order_stats(company)
    updates = {}
    if stats["first_order_date"]: updates["First Order"] = stats["first_order_date"].strftime("%Y-%m-%d")
    if stats["last_order_date"]:  updates["Last Order"]  = stats["last_order_date"].strftime("%Y-%m-%d")
    updates["CLTV"] = _float_to_money(stats["cltv"])
    updates["Days"] = str(stats["days_since_first"]) if stats["days_since_first"] is not None else ""
    # Gate Sales/Day: only when order_count >= 2
    if int(stats.get("order_count", 0) or 0) >= 2 and stats["sales_per_day"] is not None:
        updates["Sales/Day"] = _float_to_money(stats["sales_per_day"])
        # If they’ve hit 2+ orders, ensure Reorder? shows Yes
        updates["Reorder?"] = "Yes"
    else:
        updates["Sales/Day"] = ""
    update_customer_row_fields_by_company(company, updates)

def update_customer_row_fields_by_company(company: str, updates: Dict[str,str]):
    ensure_customers_file()
    try:
        with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            rows = list(rdr); flds = rdr.fieldnames or CUSTOMER_FIELDS
    except Exception:
        rows, flds = [], CUSTOMER_FIELDS
    comp_l = (company or "").strip().lower()
    found = False
    for r in rows:
        if (r.get("Company","") or "").strip().lower() == comp_l:
            for k,v in updates.items():
                if k in CUSTOMER_FIELDS:
                    r[k] = v
            try:
                r.update(_derive_customer_fields(r))
            except Exception:
                pass
            found = True
            break
    if not found:
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

# ----------------------------
# Templates / Campaigns (INI)
# ----------------------------
def load_templates_ini() -> Tuple[Dict[str,str], Dict[str,str], Dict[str,str]]:
    cfg = configparser.ConfigParser()
    cfg.read(TEMPLATES_INI, encoding="utf-8")
    tpls = dict(cfg["templates"]) if "templates" in cfg else dict(DEFAULT_TEMPLATES)
    subs = dict(cfg["subjects"])  if "subjects"  in cfg else dict(DEFAULT_SUBJECTS)
    mp   = dict(cfg["map"])       if "map"       in cfg else dict(DEFAULT_MAP)
    # backfill defaults (non-destructive)
    for k,v in DEFAULT_TEMPLATES.items(): tpls.setdefault(k, v)
    for k,v in DEFAULT_SUBJECTS.items():  subs.setdefault(k, v)
    return tpls, subs, mp

def _coerce_step_dict(step):
    def _to_int(x, default=0):
        try:
            s = str(x).strip()
            return int(s) if s else default
        except Exception:
            return default
    if isinstance(step, dict):
        return {
            "enabled": str(step.get("enabled", "1")).strip().lower() not in ("0","false","no"),
            "subject": step.get("subject","") or "",
            "body": step.get("body","") or "",
            "delay_days": _to_int(step.get("delay_days", 0), 0),
        }
    if isinstance(step, (list, tuple)):
        enabled = str(step[0]).strip().lower() not in ("0","false","no","") if len(step)>0 else True
        subject = step[1] if len(step)>1 else ""
        body = step[2] if len(step)>2 else ""
        delay_days = _to_int(step[3], 0) if len(step)>3 else 0
        return {"enabled": bool(enabled), "subject": subject or "", "body": body or "", "delay_days": delay_days}
    if isinstance(step, str):
        return {"enabled": True, "subject": "", "body": step, "delay_days": 0}
    return {"enabled": False, "subject": "", "body": "", "delay_days": 0}

def normalize_campaign_steps(steps):
    norm = []
    steps = steps or []
    for i in range(min(3, len(steps))):
        norm.append(_coerce_step_dict(steps[i]))
    while len(norm) < 3:
        norm.append({"enabled": False, "subject": "", "body": "", "delay_days": 0})
    return norm

def normalize_campaign_settings(settings):
    defaults = {
        "send_to_dialer_after": "1",
        "auto_sync_outlook": "0",
        "hourly_campaign_runner": "1",
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

def save_campaigns_ini(steps_list, settings_dict):
    cfg = configparser.ConfigParser()
    steps_list = normalize_campaign_steps(steps_list)
    for i in range(1, 4):
        s = steps_list[i-1]
        cfg[f"step{i}"] = {
            "enabled":    "1" if s.get("enabled") else "0",
            "subject":    s.get("subject",""),
            "body":       s.get("body",""),
            "delay_days": str(s.get("delay_days", 0)),
        }
    st = normalize_campaign_settings(settings_dict or {})
    cfg["settings"] = {
        "send_to_dialer_after":   st.get("send_to_dialer_after","1"),
        "auto_sync_outlook":      st.get("auto_sync_outlook","0"),
        "hourly_campaign_runner": st.get("hourly_campaign_runner","1"),
    }
    with CAMPAIGNS_INI.open("w", encoding="utf-8") as f:
        cfg.write(f)

def load_campaigns_ini():
    cfg = configparser.ConfigParser()
    cfg.read(CAMPAIGNS_INI, encoding="utf-8")
    steps = []
    for i in range(1, 4):
        s = cfg[f"step{i}"] if f"step{i}" in cfg else {}
        steps.append({
            "enabled":    str((s.get("enabled","1") if isinstance(s, dict) else "1")).strip() not in ("0","false","no"),
            "subject":    s.get("subject","") if isinstance(s, dict) else "",
            "body":       s.get("body","") if isinstance(s, dict) else "",
            "delay_days": int(str(s.get("delay_days","0")).strip() or "0") if isinstance(s, dict) else 0,
        })
    st = cfg["settings"] if "settings" in cfg else {}
    settings = normalize_campaign_settings(st if isinstance(st, dict) else {})
    return steps, settings

# Multi-campaign sections (campaign:<key>) with index
def _campaign_section_name(key: str) -> str:
    return f"campaign:{(key or '').strip()}"

def _read_campaign_cfg():
    cfg = configparser.ConfigParser()
    cfg.read(CAMPAIGNS_INI, encoding="utf-8")
    return cfg

def list_campaign_keys():
    cfg = _read_campaign_cfg()
    keys_csv = (cfg.get("index", "keys", fallback="") or "").strip()
    keys = [k.strip() for k in keys_csv.split(",") if k.strip()]
    for sec in cfg.sections():
        if sec.lower().startswith("campaign:"):
            k = sec.split(":",1)[1].strip()
            if k and k not in keys:
                keys.append(k)
    if not keys:
        keys = ["default"]
    return sorted(keys, key=lambda s: s.lower())

def _write_index(cfg, keys):
    if "index" not in cfg:
        cfg.add_section("index")
    cfg["index"]["keys"] = ",".join(keys)

def load_campaign_by_key(key: str):
    cfg = _read_campaign_cfg()
    sec = _campaign_section_name(key)
    if sec not in cfg:
        steps, settings = load_campaigns_ini()
        save_campaign_by_key(key, steps, settings)
        cfg = _read_campaign_cfg()
    s = cfg[sec] if sec in cfg else {}
    steps = []
    for i in range(1, 4):
        steps.append({"subject": s.get(f"subject{i}", ""), "body": s.get(f"body{i}", ""), "delay_days": s.get(f"delay{i}", "0")})
    settings = {
        "send_to_dialer_after": s.get("send_to_dialer_after", "1"),
        "auto_sync_outlook": s.get("auto_sync_outlook", "0"),
        "hourly_campaign_runner": s.get("hourly_campaign_runner", "1"),
    }
    return normalize_campaign_steps(steps), normalize_campaign_settings(settings)

def save_campaign_by_key(key: str, steps, settings):
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
    cfg[sec]["send_to_dialer_after"]   = "1" if settings.get("send_to_dialer_after") else "0"
    cfg[sec]["auto_sync_outlook"]      = "1" if settings.get("auto_sync_outlook") else "0"
    cfg[sec]["hourly_campaign_runner"] = "1" if settings.get("hourly_campaign_runner") else "0"
    keys = list_campaign_keys()
    if key not in keys: keys.append(key)
    _write_index(cfg, keys)
    with CAMPAIGNS_INI.open("w", encoding="utf-8") as f:
        cfg.write(f)

def delete_campaign_by_key(key: str):
    cfg = _read_campaign_cfg()
    sec = _campaign_section_name(key)
    if sec in cfg:
        cfg.remove_section(sec)
    keys = [k for k in list_campaign_keys() if k != key] or ["default"]
    _write_index(cfg, keys)
    with CAMPAIGNS_INI.open("w", encoding="utf-8") as f:
        cfg.write(f)

def summarize_campaign_for_table(key: str):
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

# ----------------------------
# Per-ref campaign state (campaigns.csv)
# ----------------------------
def ensure_campaigns_file():
    if not CAMPAIGNS_PATH.exists():
        _atomic_write_csv(CAMPAIGNS_PATH, CAMPAIGNS_HEADERS, [])

def _read_campaign_rows():
    ensure_campaigns_file()
    with CAMPAIGNS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        return list(rdr)

def _campaigns_write_rows(rows: List[Dict[str,str]]):
    _backup(CAMPAIGNS_PATH)
    with CAMPAIGNS_PATH.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CAMPAIGNS_HEADERS)
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h,"") for h in CAMPAIGNS_HEADERS})

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
            r["Stage"] = str(new_stage); break
    _campaigns_write_rows(rows)

# ----------------------------
# Placeholder utils
# ----------------------------
PLACEHOLDER_RE = re.compile(r"\{([^}]+)\}")

def normalize_header_map(row_dict: Dict[str,str]):
    norm: Dict[str,str] = {}
    for k, v in (row_dict or {}).items():
        if not k: continue
        k2 = k.strip()
        if not k2: continue
        norm[k2] = v
        norm[k2.lower()] = v
        norm[k2.replace(" ","_").lower()] = v
    return norm

def apply_placeholders(text: str, row_dict: Dict[str,str], profile=None) -> str:
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

def dict_from_row(row: List[str]) -> Dict[str,str]:
    return {HEADER_FIELDS[i]: (row[i] if i < len(HEADER_FIELDS) else "") for i in range(len(HEADER_FIELDS))}

