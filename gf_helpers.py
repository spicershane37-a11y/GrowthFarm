# gf_helpers.py
# Consolidated helpers from Chunks 3 + 4 (non-UI only)

from __future__ import annotations
import csv, re, html, hashlib, time, os
from datetime import datetime, timedelta
from pathlib import Path

# -----------------------------------------------------------------------------------
# Safe fallbacks for globals that are usually defined in your bootstrap / chunk 1
# -----------------------------------------------------------------------------------
if "APP_DIR" not in globals():
    APP_DIR = Path.cwd() / "GrowthFarmData"
APP_DIR.mkdir(parents=True, exist_ok=True)

if "CSV_PATH" not in globals():
    CSV_PATH = APP_DIR / "kybercrystals.csv"

if "WARM_LEADS_PATH" not in globals():
    WARM_LEADS_PATH = APP_DIR / "warm_leads.csv"

if "CUSTOMERS_PATH" not in globals():
    CUSTOMERS_PATH = APP_DIR / "customers.csv"

if "ORDERS_PATH" not in globals():
    ORDERS_PATH = APP_DIR / "orders.csv"

if "RESULTS_PATH" not in globals():
    RESULTS_PATH = APP_DIR / "results.csv"

if "STATE_PATH" not in globals():
    STATE_PATH = APP_DIR / "state.txt"

if "DIALER_RESULTS_PATH" not in globals():
    DIALER_RESULTS_PATH = APP_DIR / "dialer_results.csv"

if "HEADER_FIELDS" not in globals():
    # Minimal safe header set – your real list should override this at import time
    HEADER_FIELDS = ["Company","First Name","Last Name","Email","Industry","Phone","City","State","Website"]

if "CUSTOMER_FIELDS" not in globals():
    CUSTOMER_FIELDS = [
        "Company","Prospect Name","Phone #","Email","Location","Industry","Google Reviews","Rep",
        "Samples?","Timestamp","Customer Since","First Order","Last Order","CLTV","Days","Sales/Day",
        "Reorder?","Lat","Lon","Notes","Opening Order $"
    ]

if "TARGET_MAILBOX_HINT" not in globals():
    TARGET_MAILBOX_HINT = ""  # optional string to pick Outlook store

if "DEATHSTAR_SUBFOLDER" not in globals():
    DEATHSTAR_SUBFOLDER = "Death Star"

if "DEFAULT_SUBJECT" not in globals():
    DEFAULT_SUBJECT = "Quick hello"

# -----------------------------------------------------------------------------------
# CSV I/O for primary leads grid
# -----------------------------------------------------------------------------------
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

# -----------------------------------------------------------------------------------
# Backups + atomic writes
# -----------------------------------------------------------------------------------
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

# -----------------------------------------------------------------------------------
# Warm Leads v2 schema helpers
# -----------------------------------------------------------------------------------
def _build_warm_v2_fields():
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
    rows = []
    if not WARM_LEADS_PATH.exists():
        return rows
    with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            rows.append([r.get(h, "") for h in WARM_V2_FIELDS])
    return rows

def save_warm_leads_matrix_v2(matrix):
    _backup(WARM_LEADS_PATH)
    _atomic_write_csv(WARM_LEADS_PATH, WARM_V2_FIELDS, matrix)

# -----------------------------------------------------------------------------------
# Dialer leads CSV helpers
# -----------------------------------------------------------------------------------
DIALER_LEADS_PATH = APP_DIR / "dialer_leads.csv"

EMOJI_GREEN = "🙂"
EMOJI_GRAY  = "😐"
EMOJI_RED   = "🙁"
EMOJI_RED_LEGACY = "☹️"

def ensure_dialer_leads_file():
    if not DIALER_LEADS_PATH.exists():
        with DIALER_LEADS_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            hdr = HEADER_FIELDS + [EMOJI_GREEN, EMOJI_GRAY, EMOJI_RED] + [f"Note{i}" for i in range(1,9)]
            w.writerow(hdr)

def load_dialer_leads_matrix():
    ensure_dialer_leads_file()
    rows = []
    with DIALER_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.reader(f)
        raw = list(rdr)
    if not raw:
        return rows
    hdr = raw[0]
    expected = HEADER_FIELDS + [EMOJI_GREEN, EMOJI_GRAY, EMOJI_RED] + [f"Note{i}" for i in range(1,9)]
    header_lookup = {h: (hdr.index(h) if h in hdr else None) for h in expected}
    if header_lookup[EMOJI_RED] is None and EMOJI_RED_LEGACY in hdr:
        header_lookup[EMOJI_RED] = hdr.index(EMOJI_RED_LEGACY)
    idx_map = [header_lookup[h] for h in expected]
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
    _backup(DIALER_LEADS_PATH)
    headers = HEADER_FIELDS + [EMOJI_GREEN, EMOJI_GRAY, EMOJI_RED] + [f"Note{i}" for i in range(1,9)]
    _atomic_write_csv(DIALER_LEADS_PATH, headers, matrix)

# -----------------------------------------------------------------------------------
# Customers CSV helpers (+ derived fields)
# -----------------------------------------------------------------------------------
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

    try:
        with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            old_fields = rdr.fieldnames or []
            if old_fields != CUSTOMER_FIELDS:
                rows = list(rdr)
            else:
                rows = None
    except Exception:
        rows = None

    if rows is not None:
        migrated = []
        for r in rows:
            migrated.append([r.get(h, "") for h in CUSTOMER_FIELDS])
        _backup(CUSTOMERS_PATH)
        _atomic_write_csv(CUSTOMERS_PATH, CUSTOMER_FIELDS, migrated)

def _parse_date_mmddyyyy(s):
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d","%m/%d/%Y","%m-%d-%Y","%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
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

def _derive_customer_fields(row_dict: dict):
    """Compute Days and Sales/Day from First Order and CLTV. Returns updates dict (strings)."""
    first_order_str = (row_dict.get("First Order", "") or "").strip()
    first_dt = _parse_date_mmddyyyy(first_order_str)
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
    updates["CLTV"] = _float_to_money(cltv_f) if cltv_f is not None else row_dict.get("CLTV", "")
    if days_val is not None:
        updates["Days"] = str(int(days_val))
    if sales_per_day is not None:
        updates["Sales/Day"] = _float_to_money(sales_per_day)
    return updates

def load_customers_matrix():
    """Load customers.csv; derive missing fields for display."""
    ensure_customers_file()
    rows = []
    with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        if not rdr.fieldnames:
            return rows
        for r in rdr:
            try:
                derived = _derive_customer_fields(r)
                r = {**r, **{k: (derived.get(k) or r.get(k, "")) for k in ("CLTV","Days","Sales/Day")}}
            except Exception:
                pass
            rows.append([r.get(h, "") for h in CUSTOMER_FIELDS])
    return rows

def save_customers_matrix(matrix):
    """Recompute derived fields per row and save."""
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

def update_customer_row_fields_by_company(company, updates: dict):
    """Upsert customer row (by exact Company, case-insensitive)."""
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

# -----------------------------------------------------------------------------------
# Orders helpers (for analytics / CLTV updates)
# -----------------------------------------------------------------------------------
def ensure_orders_file():
    if not ORDERS_PATH.exists():
        with ORDERS_PATH.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(["Company","Order Date","Amount"])

def append_order_row(company: str, order_date: str, amount: str):
    """
    Append an order then recompute/update the customer's First/Last/CLTV/Days/Sales/Day.
    """
    ensure_orders_file()
    company = (company or "").strip()
    if not company:
        raise ValueError("Company is required for orders.")
    dt = _parse_date_mmddyyyy(order_date) or datetime.now().date()
    amt_f = _money_to_float(amount)

    with ORDERS_PATH.open("a", encoding="utf-8", newline="") as f:
        csv.writer(f).writerow([company, dt.strftime("%Y-%m-%d"), _float_to_money(amt_f)])

    stats = compute_customer_order_stats(company)
    updates = {}
    if stats["first_order_date"]:
        updates["First Order"] = stats["first_order_date"].strftime("%Y-%m-%d")
    if stats["last_order_date"]:
        updates["Last Order"] = stats["last_order_date"].strftime("%Y-%m-%d")
    updates["CLTV"] = _float_to_money(stats["cltv"])
    updates["Days"] = str(stats["days_since_first"]) if stats["days_since_first"] is not None else ""
    updates["Sales/Day"] = _float_to_money(stats["sales_per_day"]) if stats["sales_per_day"] is not None else ""
    update_customer_row_fields_by_company(company, updates)

def compute_customer_order_stats(company: str):
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
        days_since_first = max(1, (datetime.now().date() - first_dt).days)
        sales_per_day = total / float(days_since_first) if days_since_first else None

    return {
        "first_order_date": first_dt,
        "last_order_date": last_dt,
        "cltv": total,
        "days_since_first": days_since_first,
        "sales_per_day": sales_per_day,
    }

# -----------------------------------------------------------------------------------
# Utilities (placeholders, keys, fingerprints)
# -----------------------------------------------------------------------------------
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

# -----------------------------------------------------------------------------------
# Campaigns helpers (per-ref state via campaigns.csv)
# -----------------------------------------------------------------------------------
CAMPAIGNS_PATH = APP_DIR / "campaigns.csv"
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
    _backup(CAMPAIGNS_PATH)
    with CAMPAIGNS_PATH.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CAMPAIGNS_HEADERS)
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h,"") for h in CAMPAIGNS_HEADERS})
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

# -----------------------------------------------------------------------------------
# Outlook helpers (drafting + results sync)
# -----------------------------------------------------------------------------------
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
    return { (r.get("Ref","") or "").lower(): r for r in load_results_rows_sorted() }

def _results_dates_for_ref(ref_short):
    """Return (sent_dt, replied_dt) as datetime or (None,None)."""
    r = _results_lookup_by_ref().get((ref_short or "").lower())
    def _p(s):
        s = (s or "").strip()
        if not s: return None
        for fmt in ("%Y-%m-%d %H:%M:%S","%Y-%m-%d",
                    "%m/%d/%Y %I:%M:%S %p","%m/%d/%Y %I:%M %p","%m/%d/%Y"):
            try:
                return datetime.strptime(s, fmt)
            except Exception:
                pass
        alt = s.replace("-", "/")
        for fmt in ("%m/%d/%Y %I:%M:%S %p","%m/%d/%Y %I:%M %p","%m/%d/%Y"):
            try:
                return datetime.strptime(alt, fmt)
            except Exception:
                pass
        return None
    if not r:
        return (None, None)
    return (_p(r.get("DateSent","")), _p(r.get("DateReplied","")))

def _lead_row_from_email_company(email, company):
    """Try to find the original lead row for placeholders."""
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
            f.seek(0); next(rdr, None)
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
    try:
        LAST_SYNC_PATH.write_text(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), encoding="utf-8")
    except Exception:
        pass
    return len(sent_map), len(reply_map)

def upsert_result(ref_short, email, company, industry, subject):
    """Convenience updater for results cache when drafting."""
    rows = load_results_rows_sorted()
    byref = { (r.get("Ref","") or "").lower(): r for r in rows }
    key = (ref_short or "").lower()
    r = byref.get(key)
    if not r:
        r = {"Ref": ref_short, "Email": email, "Company": company, "Industry": industry,
             "DateSent": "", "DateReplied": "", "Status": "", "Subject": subject or ""}
        rows.append(r)
    else:
        r["Email"] = email or r.get("Email","")
        r["Company"] = company or r.get("Company","")
        r["Industry"] = industry or r.get("Industry","")
        r["Subject"] = subject or r.get("Subject","")
    with RESULTS_PATH.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["Ref","Email","Company","Industry","DateSent","DateReplied","Status","Subject"])
        w.writeheader(); w.writerows(rows)

# -----------------------------------------------------------------------------------
# Chunk 4: shared time parsing + daily activity (non-UI)
# -----------------------------------------------------------------------------------
LAST_SYNC_PATH = APP_DIR / "last_outlook_sync.txt"

def _today_date():
    return datetime.now().date()

def _fmt_money(x):
    try:
        return f"{float(x):.2f}"
    except Exception:
        return "0.00"

def _parse_any_datetime(s):
    s = (s or "").strip()
    if not s:
        return None
    fmts = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y",
    ]
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    alt = s.replace("-", "/")
    if alt != s:
        for fmt in ("%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M %p", "%m/%d/%Y"):
            try:
                return datetime.strptime(alt, fmt)
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

    # Emails sent (results.csv)
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

    # New Warm Leads
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

    # New Accounts (customers.csv)
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

    # Daily Sales (orders.csv)
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

# -----------------------------------------------------------------------------------
# Chunk 4: Campaign queue processor (non-UI; relies on campaign config elsewhere)
# -----------------------------------------------------------------------------------
def _campaign_get_lead_row_for_ref(crow):
    enr = _lead_row_from_email_company(crow.get("Email",""), crow.get("Company",""))
    if not enr:
        enr = {h:"" for h in HEADER_FIELDS}
        enr["Email"] = crow.get("Email","")
        enr["Company"] = crow.get("Company","")
    return enr

def _campaign_stage_from_results_if_needed(ref_short, cur_stage):
    try:
        if int(cur_stage or 0) > 0:
            return int(cur_stage)
    except Exception:
        pass
    try:
        res_map = _read_results_by_ref()
        r = res_map.get((ref_short or "").strip().lower())
        sent_dt = _results_sent_dt(r) if r else None  # defined in campaigns chunk normally
        return 1 if sent_dt else 0
    except Exception:
        return int(cur_stage or 0)

def process_campaign_queue():
    """
    - If a ref has DateReplied -> remove from campaigns.
    - If stage==0 and DateSent exists -> set stage=1.
    - If stage==1 and due and no reply -> draft E2 via _draft_next_stage_stub, stage=2.
    - If stage==2 and due and no reply -> draft E3 via _draft_next_stage_stub, stage=3.
    - If stage==3 and no reply and divert flag -> push to Dialer & remove.
    """
    ensure_campaigns_file()
    rows = _read_campaign_rows()
    changed = False

    try:
        # These helpers live in campaigns config module; import at runtime if needed
        from gf_campaigns import load_campaign_by_key, normalize_campaign_steps, normalize_campaign_settings
        from gf_campaigns import _read_results_by_ref, _results_replied, _results_sent_dt
    except Exception:
        # Fallback to local if available
        pass

    try:
        res_map = _read_results_by_ref()
    except Exception:
        res_map = {}

    for r in rows[:]:
        ref = r.get("Ref","")
        key = r.get("CampaignKey","default")
        divert_csv = r.get("DivertToDialer", "")
        try:
            stage = int(r.get("Stage","0") or 0)
        except Exception:
            stage = 0

        res = res_map.get((ref or "").strip().lower())
        try:
            replied = _results_replied(res) if res else False
        except Exception:
            replied = False
        if replied:
            rows.remove(r); changed = True; continue

        new_stage = _campaign_stage_from_results_if_needed(ref, stage)
        if new_stage != stage:
            r["Stage"] = str(new_stage); stage = new_stage; changed = True

        try:
            steps, settings = load_campaign_by_key(key)
            steps = normalize_campaign_steps(steps)
            settings = normalize_campaign_settings(settings)
        except Exception:
            steps, settings = [], {}

        try:
            divert_effective = (str(divert_csv).strip() in ("1","true","True"))
            if str(divert_csv).strip() == "":
                divert_effective = (settings.get("send_to_dialer_after") in ("1", True))
        except Exception:
            divert_effective = (settings.get("send_to_dialer_after") in ("1", True))

        if stage == 1:
            lead = _campaign_get_lead_row_for_ref(r)
            try:
                drafted = globals().get("_draft_next_stage_stub", lambda *a, **k: False)(
                    ref, lead.get("Email",""), lead.get("Company",""), key, 2
                )
                if drafted:
                    r["Stage"] = "2"; changed = True
            except Exception:
                pass

        elif stage == 2:
            lead = _campaign_get_lead_row_for_ref(r)
            try:
                drafted = globals().get("_draft_next_stage_stub", lambda *a, **k: False)(
                    ref, lead.get("Email",""), lead.get("Company",""), key, 3
                )
                if drafted:
                    r["Stage"] = "3"; changed = True
            except Exception:
                pass

        elif stage >= 3:
            try:
                replied = _results_replied(res) if res else False
            except Exception:
                replied = False
            if not replied and divert_effective:
                lead = _campaign_get_lead_row_for_ref(r)
                try:
                    ensure_dialer_leads_file()
                    cur = load_dialer_leads_matrix()
                    base = [lead.get(h,"") for h in HEADER_FIELDS]
                    cur.append(base + ["○","○","○"] + ([""]*8))
                    save_dialer_leads_matrix(cur)
                except Exception:
                    pass
            rows.remove(r); changed = True

    if changed:
        _write_campaign_rows(rows)
