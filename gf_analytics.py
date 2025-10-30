# gf_analytics.py
# Live analytics updater for:
#   • Customer Analytics (right-side)
#   • Pipeline Analytics (right-side, persistent counters)
#   • Daily Activity Tracker (top-left)  -> includes Calls from calls_log.csv
#   • Monthly Results (top-left)
#
# Exposes:
#   init_analytics(window, interval_ms=1500)
#   increment_warm_generated(n=1)
#   increment_new_customer(n=1)
#   log_call(source, outcome, note, company="", prospect="", email="", phone="")
#
# "Call" definition (current):
#   Any time a note is saved in either the Dialer tab or the Warm Leads tab.
#   We append a row to calls_log.csv so exports & management reports are easy.

from __future__ import annotations

import csv
import json
from datetime import datetime, date, timezone
from pathlib import Path
from typing import Dict, List, Optional

# --- timezone handling with safe fallbacks ---
try:
    from zoneinfo import ZoneInfo  # stdlib (needs tzdata on Windows)
except Exception:
    ZoneInfo = None  # type: ignore

def _detect_local_tz():
    """Return a tzinfo safely:
    1) Try America/Indiana/Indianapolis
    2) Fallback to US/Eastern
    3) Fallback to system local tz
    4) Finally UTC
    """
    # 1) Preferred Indiana zone
    if ZoneInfo:
        try:
            return ZoneInfo("America/Indiana/Indianapolis")
        except Exception:
            pass
        # 2) US/Eastern as a close fallback
        try:
            return ZoneInfo("US/Eastern")
        except Exception:
            pass
    # 3) System local tz
    try:
        return datetime.now().astimezone().tzinfo or timezone.utc
    except Exception:
        pass
    # 4) UTC last resort
    return timezone.utc

_LOCAL_TZ = _detect_local_tz()

# Data sources
from gf_store import (
    APP_DIR,
    CUSTOMERS_PATH,
    ORDERS_PATH,
    WARM_LEADS_PATH,
    RESULTS_PATH,          # emails sent log for daily count
    CUSTOMER_FIELDS,
    WARM_V2_FIELDS,
)

# ---------- persistent counters (in this file) ----------
_COUNTERS_PATH = APP_DIR / "analytics_counters.json"

# ---------- calls log (persistent CSV) ----------
CALLS_LOG_PATH = APP_DIR / "calls_log.csv"
CALLS_HEADERS = ["Timestamp", "Source", "Outcome", "Note", "Company", "Prospect", "Email", "Phone"]


# ==============================
# Calls log helpers
# ==============================
def _ensure_calls_log():
    try:
        CALLS_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        if not CALLS_LOG_PATH.exists():
            with CALLS_LOG_PATH.open("w", encoding="utf-8", newline="") as f:
                w = csv.writer(f)
                w.writerow(CALLS_HEADERS)
    except Exception:
        pass


def log_call(
    source: str,
    outcome: str,
    note: str,
    company: str = "",
    prospect: str = "",
    email: str = "",
    phone: str = "",
) -> None:
    """
    Append a call record. Call this from Dialer/Warm when a note is saved.
    - source: "dialer" | "warm" (free text allowed)
    - outcome: "green" | "gray" | "red" | anything (stored as-is)
    """
    _ensure_calls_log()
    try:
        ts = datetime.now(_LOCAL_TZ).isoformat(timespec="seconds")
        row = [ts, str(source or ""), str(outcome or ""), str(note or ""),
               str(company or ""), str(prospect or ""), str(email or ""), str(phone or "")]
        with CALLS_LOG_PATH.open("a", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(row)
    except Exception:
        # Swallow errors; analytics should never crash the app.
        pass


# ==============================
# CSV helpers
# ==============================
def _safe_read_dicts(path: Path) -> List[Dict[str, str]]:
    if not path.exists():
        return []
    try:
        with path.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            return list(rdr) if rdr.fieldnames else []
    except Exception:
        return []


def _row_has_payload(row: Dict[str, str], core_fields: List[str]) -> bool:
    for k in core_fields:
        if (row.get(k, "") or "").strip():
            return True
    return False


# ==============================
# Persistent pipeline counters
# ==============================
def _seed_values() -> Dict[str, int]:
    warm_rows = _safe_read_dicts(WARM_LEADS_PATH)
    warm_core = ["Company", "Prospect Name", "Email", "Phone #"]
    warm_generated = 0
    for r in warm_rows:
        if _row_has_payload(r, warm_core) and (r.get("Timestamp", "") or "").strip():
            warm_generated += 1

    cust_rows = _safe_read_dicts(CUSTOMERS_PATH)
    new_customers = 0
    for r in cust_rows:
        company = (r.get("Company", "") or "").strip()
        first_order = (r.get("First Order", "") or "").strip()
        first_contact = (r.get("First Contact", "") or "").strip()
        if company and (first_order or first_contact):
            new_customers += 1

    return {"warm_generated": warm_generated, "new_customers": new_customers}


def _save_counters(counters: Dict[str, int]) -> None:
    try:
        _COUNTERS_PATH.parent.mkdir(parents=True, exist_ok=True)
        with _COUNTERS_PATH.open("w", encoding="utf-8") as f:
            json.dump(
                {
                    "warm_generated": int(counters.get("warm_generated", 0) or 0),
                    "new_customers": int(counters.get("new_customers", 0) or 0),
                    "last_update": datetime.now(_LOCAL_TZ).isoformat(timespec="seconds"),
                },
                f,
                indent=2,
            )
    except Exception:
        pass


def _load_counters() -> Dict[str, int]:
    try:
        if _COUNTERS_PATH.exists():
            with _COUNTERS_PATH.open("r", encoding="utf-8") as f:
                data = json.load(f) or {}
                wg = max(0, int(data.get("warm_generated", 0) or 0))
                nc = max(0, int(data.get("new_customers", 0) or 0))
                return {"warm_generated": wg, "new_customers": nc}
    except Exception:
        pass
    seed = _seed_values()
    _save_counters(seed)
    return seed


def ensure_seeded() -> None:
    if not _COUNTERS_PATH.exists():
        _save_counters(_seed_values())
    _ensure_calls_log()  # make sure calls log exists at startup


def get_totals() -> Dict[str, int]:
    return _load_counters()


def increment_warm_generated(n: int = 1) -> None:
    try:
        n = int(n)
    except Exception:
        n = 1
    if n <= 0:
        return
    c = _load_counters()
    c["warm_generated"] = int(c.get("warm_generated", 0)) + n
    _save_counters(c)


def increment_new_customer(n: int = 1) -> None:
    try:
        n = int(n)
    except Exception:
        n = 1
    if n <= 0:
        return
    c = _load_counters()
    c["new_customers"] = int(c.get("new_customers", 0)) + n
    _save_counters(c)


# ==============================
# Date/time & money helpers
# ==============================
def _money_to_float(val: str) -> float:
    s = (val or "").strip().replace(",", "").replace("$", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def _float_to_money(x: str | float) -> str:
    try:
        return f"{float(x):,.2f}"
    except Exception:
        return "0.00"


def _parse_date(s: str) -> Optional[date]:
    """Loose date-only parser (kept for compatibility)."""
    if not s:
        return None
    s = s.strip()
    fmts = (
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m-%d-%Y",
        "%Y/%m/%d",
        "%m/%d/%y",
        "%m-%d-%y",
        "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %I:%M %p",
    )
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None


def _parse_any_dt_local(s: str | None) -> Optional[datetime]:
    """
    Parse common datetime strings and return an *aware* datetime in local tz.
    Handles:
      • ISO8601 with/without tz  (2025-10-28T09:12:00-04:00, ...Z, or no tz)
      • RFC 2822 (email-style)   (Tue, 28 Oct 2025 09:12:00 -0400)
      • The looser formats supported by _parse_date
    """
    if not s:
        return None
    s = s.strip()

    # ISO8601 first (fromisoformat handles many shapes)
    try:
        dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=_LOCAL_TZ)
        return dt.astimezone(_LOCAL_TZ)
    except Exception:
        pass

    # RFC 2822 / email style
    try:
        from email.utils import parsedate_to_datetime
        dt = parsedate_to_datetime(s)
        if dt is not None:
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=_LOCAL_TZ)
            return dt.astimezone(_LOCAL_TZ)
    except Exception:
        pass

    # Fallback: your previous date-only parser (assume local midnight)
    d = _parse_date(s)
    if d:
        return datetime(d.year, d.month, d.day, tzinfo=_LOCAL_TZ)
    return None


def _mtime(path: Path) -> Optional[float]:
    try:
        return path.stat().st_mtime
    except Exception:
        return None


# ==============================
# Customer Analytics (right pane)
# ==============================
def _compute_customer_metrics() -> Dict[str, str]:
    # Orders: total sales + per-company count
    orders = _safe_read_dicts(ORDERS_PATH)
    total_sales = 0.0
    orders_by_company = {}
    for r in orders:
        company = (r.get("Company", "") or "").strip().lower()
        amt = _money_to_float(r.get("Amount", ""))
        total_sales += amt
        orders_by_company.setdefault(company, 0)
        orders_by_company[company] += 1

    # Customers: CLTV & reorder
    customers = _safe_read_dicts(CUSTOMERS_PATH)
    cltvs = []
    with_reorder = 0
    total_customers = 0
    for r in customers:
        total_customers += 1
        cltvs.append(_money_to_float(r.get("CLTV", "")))
        company = (r.get("Company", "") or "").strip().lower()
        explicit = (r.get("Reorder?", "") or "").strip().lower()
        has_two_orders = orders_by_company.get(company, 0) >= 2
        if explicit in ("yes", "y", "true", "1") or has_two_orders:
            with_reorder += 1

    avg_ltv = (sum(cltvs) / len(cltvs)) if cltvs else 0.0
    reorder_rate = (with_reorder / total_customers * 100.0) if total_customers > 0 else 0.0

    # CAC: sum of warm costs / customers
    warm_rows = _safe_read_dicts(WARM_LEADS_PATH)
    total_cost = 0.0
    for r in warm_rows:
        total_cost += _money_to_float(r.get("Cost ($)", ""))

    cac = total_cost / max(1, total_customers)
    ratio_rhs = (avg_ltv / cac) if cac > 0 else 0.0  # LTV:CAC numeric value

    return {
        "total_sales": _float_to_money(total_sales),
        "cac": _float_to_money(cac),
        "ltv": _float_to_money(avg_ltv),
        "ratio": f"{ratio_rhs:.1f}",  # UI label is "LTV : CAC"
        "reorder": f"{reorder_rate:.0f}%",
    }


# ==============================
# Pipeline Analytics (right pane)
# ==============================
def _compute_pipeline_metrics() -> Dict[str, str]:
    ensure_seeded()
    totals = get_totals()
    warm_generated = int(totals.get("warm_generated", 0) or 0)
    new_customers = int(totals.get("new_customers", 0) or 0)
    close_rate = (new_customers / warm_generated * 100.0) if warm_generated > 0 else 0.0
    return {
        "warms": str(warm_generated),
        "new_customers": str(new_customers),
        "close_rate": f"{close_rate:.0f}%",
    }


# ==============================
# Daily Activity & Monthly Results (top-left)
# ==============================
def _calls_count_for_day(day: date) -> int:
    _ensure_calls_log()
    rows = _safe_read_dicts(CALLS_LOG_PATH)
    if not rows:
        return 0
    count = 0
    for r in rows:
        ts = (r.get("Timestamp", "") or "").strip()
        dt = _parse_any_dt_local(ts)
        if dt and dt.date() == day:
            count += 1
    return count


def _calls_count_for_month(year: int, month: int) -> int:
    _ensure_calls_log()
    rows = _safe_read_dicts(CALLS_LOG_PATH)
    if not rows:
        return 0
    count = 0
    for r in rows:
        ts = (r.get("Timestamp", "") or "").strip()
        dt = _parse_any_dt_local(ts)
        if dt and dt.year == year and dt.month == month:
            count += 1
    return count


def _compute_daily_metrics() -> Dict[str, str]:
    # “Today” is based on the local business timezone
    today_local = datetime.now(_LOCAL_TZ).date()

    # Emails sent today (robust parse + local tz + de-dupe by {To, Subject, date})
    emails = 0
    rows_results = _safe_read_dicts(RESULTS_PATH)
    seen = set()
    for r in rows_results:
        dt = _parse_any_dt_local(r.get("DateSent") or r.get("Date") or "")
        if not dt:
            continue
        if dt.date() == today_local:
            key = (
                (r.get("To") or "").strip().lower(),
                (r.get("Subject") or "").strip(),
                dt.date().isoformat(),
            )
            if key in seen:
                continue
            seen.add(key)
            emails += 1

    # Sales today
    sales_today = 0.0
    for r in _safe_read_dicts(ORDERS_PATH):
        dt = _parse_any_dt_local(r.get("Order Date") or r.get("Date") or "")
        if dt and dt.date() == today_local:
            sales_today += _money_to_float(r.get("Amount", ""))

    # New warm leads today
    warms_today = 0
    warm_rows = _safe_read_dicts(WARM_LEADS_PATH)
    ts_field = "First Contact" if warm_rows and "First Contact" in (warm_rows[0].keys()) else "Timestamp"
    for r in warm_rows:
        dt = _parse_any_dt_local(r.get(ts_field, ""))
        if dt and dt.date() == today_local:
            warms_today += 1

    # New accounts (customers) today
    newcus_today = 0
    for r in _safe_read_dicts(CUSTOMERS_PATH):
        cs = _parse_any_dt_local(r.get("Customer Since") or r.get("First Order") or "")
        if cs and cs.date() == today_local:
            newcus_today += 1

    # Calls today (from calls_log.csv)
    calls_today = _calls_count_for_day(today_local)

    return {
        "calls": str(calls_today),
        "emails": str(emails),
        "warms": str(warms_today),
        "newcus": str(newcus_today),
        "sales": f"${_float_to_money(sales_today)}",
    }


def _compute_monthly_metrics() -> Dict[str, str]:
    now = datetime.now(_LOCAL_TZ)
    month_warms = 0
    month_newcus = 0
    month_sales = 0.0

    # Warm leads this month
    warm_rows = _safe_read_dicts(WARM_LEADS_PATH)
    ts_field = "First Contact" if warm_rows and "First Contact" in (warm_rows[0].keys()) else "Timestamp"
    for r in warm_rows:
        dt = _parse_any_dt_local(r.get(ts_field, ""))
        if dt and dt.year == now.year and dt.month == now.month:
            month_warms += 1

    # New customers this month
    for r in _safe_read_dicts(CUSTOMERS_PATH):
        cs = _parse_any_dt_local(r.get("Customer Since") or r.get("First Order") or "")
        if cs and cs.year == now.year and cs.month == now.month:
            month_newcus += 1

    # Sales this month
    for r in _safe_read_dicts(ORDERS_PATH):
        dt = _parse_any_dt_local(r.get("Order Date") or r.get("Date") or "")
        if dt and dt.year == now.year and dt.month == now.month:
            month_sales += _money_to_float(r.get("Amount", ""))

    # Calls this month (from calls_log.csv) – available if you add a UI label
    calls_this_month = _calls_count_for_month(now.year, now.month)

    return {
        "warms": str(month_warms),
        "newcus": str(month_newcus),
        "sales": f"${_float_to_money(month_sales)}",
        "calls": str(calls_this_month),
    }


# ==============================
# Apply to window
# ==============================
def _apply_customer_metrics_to_window(window, m: Dict[str, str]) -> None:
    try: window["-AN_TOTALSALES-"].update(m.get("total_sales", "0.00"))
    except Exception: pass
    try: window["-AN_CAC-"].update(m.get("cac", "0.00"))
    except Exception: pass
    try: window["-AN_LTV-"].update(m.get("ltv", "0.00"))
    except Exception: pass
    try: window["-AN_CACLTV-"].update(m.get("ratio", "0.0"))
    except Exception: pass
    try: window["-AN_REORDER-"].update(m.get("reorder", "0%"))
    except Exception: pass


def _apply_pipeline_metrics_to_window(window, m: Dict[str, str]) -> None:
    try: window["-AN_WARMS-"].update(m.get("warms", "0"))
    except Exception: pass
    try: window["-AN_NEWCUS-"].update(m.get("new_customers", "0"))
    except Exception: pass
    try: window["-AN_CLOSERATE-"].update(m.get("close_rate", "0%"))
    except Exception: pass


def _apply_daily_to_window(window, m: Dict[str, str]) -> None:
    try: window["-DA_CALLS-"].update(m.get("calls", "0"))
    except Exception: pass
    try: window["-DA_EMAILS-"].update(m.get("emails", "0"))
    except Exception: pass
    try: window["-DA_WARMS-"].update(m.get("warms", "0"))
    except Exception: pass
    try: window["-DA_NEWCUS-"].update(m.get("newcus", "0"))
    except Exception: pass
    try: window["-DA_SALES-"].update(m.get("sales", "$0.00"))
    except Exception: pass


def _apply_monthly_to_window(window, m: Dict[str, str]) -> None:
    try: window["-MO_WARMS-"].update(m.get("warms", "0"))
    except Exception: pass
    try: window["-MO_NEWCUS-"].update(m.get("newcus", "0"))
    except Exception: pass
    try: window["-MO_SALES-"].update(m.get("sales", "$0.00"))
    except Exception: pass
    # If you add a Monthly Calls label later (e.g., key "-MO_CALLS-"), uncomment:
    # try: window["-MO_CALLS-"].update(m.get("calls", "0"))
    # except Exception: pass


# ==============================
# Watcher / entry point
# ==============================
_LAST_MTIMES = {"warm": None, "cust": None, "orders": None, "results": None, "calls": None}

def _refresh_all(window) -> None:
    try:
        _apply_customer_metrics_to_window(window, _compute_customer_metrics())
        _apply_pipeline_metrics_to_window(window, _compute_pipeline_metrics())
        _apply_daily_to_window(window, _compute_daily_metrics())
        _apply_monthly_to_window(window, _compute_monthly_metrics())
    except Exception:
        pass


def _files_changed() -> bool:
    changed = False
    for key, path in (
        ("warm", WARM_LEADS_PATH),
        ("cust", CUSTOMERS_PATH),
        ("orders", ORDERS_PATH),
        ("results", RESULTS_PATH),
        ("calls", CALLS_LOG_PATH),
    ):
        mt = _mtime(path)
        global _LAST_MTIMES
        if _LAST_MTIMES[key] is None:
            _LAST_MTIMES[key] = mt
        elif mt != _LAST_MTIMES[key]:
            _LAST_MTIMES[key] = mt
            changed = True
    return changed


def init_analytics(window, interval_ms: int = 1500) -> None:
    ensure_seeded()
    _refresh_all(window)

    def _tick():
        try:
            if _files_changed():
                _refresh_all(window)
            else:
                # Even if files haven't changed, pipeline counters may have been bumped.
                _apply_pipeline_metrics_to_window(window, _compute_pipeline_metrics())
        finally:
            try:
                window.TKroot.after(interval_ms, _tick)
            except Exception:
                pass

    try:
        window.TKroot.after(interval_ms, _tick)
    except Exception:
        pass

