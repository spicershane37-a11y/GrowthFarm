# gf_customers.py
# Customers tab grid + persistence + simple "Add Order" helper
# + live analytics (customer + pipeline) with a small file-change watcher.
from __future__ import annotations

from typing import Optional, List, Dict, Tuple
from datetime import datetime, date
import csv
import re
from pathlib import Path

from gf_store import (
    CUSTOMER_FIELDS,
    CUSTOMERS_PATH,
    ORDERS_PATH,
    WARM_LEADS_PATH,
    load_customers_matrix,
    save_customers_matrix,
    append_order_row,
    compute_customer_order_stats,  # (unused here but kept for compatibility)
)

try:
    from tksheet import Sheet
except Exception:
    Sheet = None  # guarded by caller (gf_ui_logic)

_SHEET_CUSTOMERS: Optional[Sheet] = None
_WINDOW = None

# -----------------------------
# Helpers
# -----------------------------
def _ensure_sheet() -> Sheet:
    if _SHEET_CUSTOMERS is None:
        raise RuntimeError("Customers sheet is not mounted yet.")
    return _SHEET_CUSTOMERS

def _persist_customers_now() -> bool:
    """Write the entire customers sheet to customers.csv (keeps exact width)."""
    if _SHEET_CUSTOMERS is None:
        return False
    try:
        rows = _SHEET_CUSTOMERS.get_sheet_data(return_copy=True) or []
    except Exception:
        rows = _SHEET_CUSTOMERS.get_sheet_data() or []
    width = len(CUSTOMER_FIELDS)
    fixed = [(list(r) + [""] * width)[:width] for r in rows]
    save_customers_matrix(fixed)
    return True

def _sanitize_amount(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    s = s.replace(",", "")
    if s.startswith("$"):
        s = s[1:]
    if not re.match(r"^\d+(\.\d{1,2})?$", s):
        return ""
    return s

def _unbind_default_paste_and_rc(sheet: Sheet) -> None:
    """Remove Ctrl+V and default right-click popup so app-level paste owns behavior."""
    # Unbind paste from all relevant widgets
    for w in filter(None, (sheet, getattr(sheet, "MT", None),
                           getattr(sheet, "RI", None), getattr(sheet, "CH", None),
                           getattr(sheet, "Toplevel", None))):
        try:
            w.unbind("<Control-v>"); w.unbind("<Control-V>")
            w.unbind("<Control-Shift-v>"); w.unbind("<Control-Shift-V>")
            w.unbind("<Command-v>"); w.unbind("<Command-V>")
        except Exception:
            pass
    # Kill default RC popup so there’s no “menu paste” path to row 1
    try:
        sheet.disable_bindings(("right_click_popup_menu",))
    except Exception:
        pass

# -----------------------------
# Analytics core
# -----------------------------
def _parse_date(s: str) -> Optional[date]:
    s = (s or "").strip()
    if not s:
        return None
    fmts = ("%Y-%m-%d","%m/%d/%Y","%m-%d-%Y","%Y/%m/%d","%m/%d/%y","%m-%d-%y")
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None

def _money_to_float(val: str) -> float:
    s = (val or "").strip().replace(",", "").replace("$","")
    try:
        return float(s) if s else 0.0
    except Exception:
        return 0.0

def _float_to_money(x) -> str:
    try:
        return f"{float(x):.2f}"
    except Exception:
        return "0.00"

def _orders_by_company() -> Dict[str, List[Tuple[date, float]]]:
    out: Dict[str, List[Tuple[date, float]]] = {}
    if ORDERS_PATH.exists():
        with ORDERS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                comp = (r.get("Company","") or "").strip()
                d = _parse_date(r.get("Order Date",""))
                amt = _money_to_float(r.get("Amount",""))
                if comp and d:
                    out.setdefault(comp, []).append((d, amt))
    return out

def _warm_cost_by_company() -> Dict[str, float]:
    """Sum Cost ($) from warm_leads.csv grouped by Company."""
    out: Dict[str, float] = {}
    if WARM_LEADS_PATH.exists():
        with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                comp = (r.get("Company","") or "").strip()
                cost = _money_to_float(r.get("Cost ($)",""))
                if comp and cost > 0:
                    out[comp] = out.get(comp, 0.0) + cost
    return out

def _load_customers_rows() -> List[Dict[str,str]]:
    rows: List[Dict[str,str]] = []
    if CUSTOMERS_PATH.exists():
        with CUSTOMERS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                rows.append(r)
    return rows

def _month_bounds(today: date) -> Tuple[date, date]:
    start = today.replace(day=1)
    if start.month == 12:
        nxt = start.replace(year=start.year+1, month=1, day=1)
    else:
        nxt = start.replace(month=start.month+1, day=1)
    end = nxt  # exclusive
    return start, end

def _in_month(d: Optional[date], start: date, end: date) -> bool:
    return (d is not None) and (start <= d < end)

def _compute_customer_analytics() -> Dict[str,str]:
    customers = _load_customers_rows()
    orders_map = _orders_by_company()
    warm_cost = _warm_cost_by_company()

    total_sales = 0.0
    for comp, orders in orders_map.items():
        for _d, amt in orders:
            total_sales += (amt or 0.0)

    cltvs = [ _money_to_float(r.get("CLTV","")) for r in customers ]
    avg_ltv = (sum(cltvs) / len(cltvs)) if cltvs else 0.0

    cost_vals = []
    cust_companies = {(r.get("Company","") or "").strip() for r in customers if (r.get("Company","") or "").strip()}
    for comp in cust_companies:
        if warm_cost.get(comp, 0.0) > 0:
            cost_vals.append(warm_cost[comp])
    avg_cac = (sum(cost_vals) / len(cost_vals)) if cost_vals else 0.0

    ratio_str = "1 : 0"
    if avg_cac > 0:
        ratio = avg_ltv / avg_cac
        ratio_str = f"1 : {ratio:.1f}"

    reorder_count = 0
    cust_count = 0
    for comp in cust_companies:
        cnt = len(orders_map.get(comp, []))
        if cnt >= 2:
            reorder_count += 1
        if cnt >= 1 or comp:
            cust_count += 1
    reorder_rate = (reorder_count / cust_count * 100.0) if cust_count > 0 else 0.0

    return {
        "total_sales": f"{total_sales:,.2f}",
        "avg_cac": f"{avg_cac:,.2f}",
        "avg_ltv": f"{avg_ltv:,.2f}",
        "ratio": ratio_str,
        "reorder": f"{reorder_rate:.0f}%",
    }

def _compute_pipeline_analytics() -> Dict[str,str]:
    warm_total = 0
    if WARM_LEADS_PATH.exists():
        with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            for r in rdr:
                if (r.get("Company","") or "").strip():
                    warm_total += 1

    customers = _load_customers_rows()
    orders_map = _orders_by_company()

    today = datetime.now().date()
    m_start, m_end = _month_bounds(today)
    new_this_month = 0
    for r in customers:
        comp = (r.get("Company","") or "").strip()
        ords = orders_map.get(comp, [])
        first_date = min([d for (d, _amt) in ords], default=None)
        if _in_month(first_date, m_start, m_end):
            new_this_month += 1

    total_customers = len({(r.get("Company","") or "").strip() for r in customers if (r.get("Company","") or "").strip()})
    close_rate = (total_customers / warm_total * 100.0) if warm_total > 0 else 0.0

    return {
        "warms": str(warm_total),
        "newcus": str(new_this_month),
        "close": f"{close_rate:.0f}%",
    }

def update_customer_analytics_in_ui(window) -> None:
    an = _compute_customer_analytics()
    try: window["-AN_TOTALSALES-"].update(an["total_sales"])
    except Exception: pass
    try: window["-AN_CAC-"].update(an["avg_cac"])
    except Exception: pass
    try: window["-AN_LTV-"].update(an["avg_ltv"])
    except Exception: pass
    try: window["-AN_CACLTV-"].update(an["ratio"])
    except Exception: pass
    try: window["-AN_REORDER-"].update(an["reorder"])
    except Exception: pass

    pl = _compute_pipeline_analytics()
    try: window["-AN_WARMS-"].update(pl["warms"])
    except Exception: pass
    try: window["-AN_NEWCUS-"].update(pl["newcus"])
    except Exception: pass
    try: window["-AN_CLOSERATE-"].update(pl["close"])
    except Exception: pass

# -----------------------------
# Lightweight watcher (recompute analytics)
# -----------------------------
_LAST_MTIMES = {"customers": None, "orders": None, "warm": None}
_WATCH_STARTED = False
_WATCH_INTERVAL_MS = 2000

def _mtime_or_none(p: Path) -> Optional[float]:
    try:
        return p.stat().st_mtime
    except Exception:
        return None

def _changed() -> bool:
    changed = False
    for key, path in (("customers", CUSTOMERS_PATH), ("orders", ORDERS_PATH), ("warm", WARM_LEADS_PATH)):
        mt = _mtime_or_none(path)
        prev = _LAST_MTIMES.get(key)
        if prev is None:
            _LAST_MTIMES[key] = mt
        else:
            if mt is not None and mt != prev:
                _LAST_MTIMES[key] = mt
                changed = True
    return changed

def _start_watch(window):
    global _WATCH_STARTED
    if _WATCH_STARTED:
        return
    _WATCH_STARTED = True

    for key, path in (("customers", CUSTOMERS_PATH), ("orders", ORDERS_PATH), ("warm", WARM_LEADS_PATH)):
        _LAST_MTIMES[key] = _mtime_or_none(path)

    def _tick():
        try:
            if _changed():
                update_customer_analytics_in_ui(window)
        finally:
            try:
                window.TKroot.after(_WATCH_INTERVAL_MS, _tick)
            except Exception:
                pass

    try:
        window.TKroot.after(_WATCH_INTERVAL_MS, _tick)
    except Exception:
        pass

# -----------------------------
# Public mounting API
# -----------------------------
def mount_customers_grid(window, parent_elem) -> Optional[Sheet]:
    """Mount a tksheet in the Customers tab, loading from customers.csv if present."""
    global _SHEET_CUSTOMERS, _WINDOW
    _WINDOW = window
    if Sheet is None:
        return None

    parent = parent_elem.Widget

    # Clear children
    try:
        for c in parent.winfo_children():
            c.destroy()
    except Exception:
        pass

    try:
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)
    except Exception:
        pass

    matrix = load_customers_matrix() or [[""] * len(CUSTOMER_FIELDS)]
    sheet = Sheet(
        parent,
        data=[list(r) for r in matrix],
        headers=list(CUSTOMER_FIELDS),
        show_x_scrollbar=True,
        show_y_scrollbar=True,
    )
    sheet.grid(row=0, column=0, sticky="nsew")

    # Enable only safe bindings (NO default RC menu; NO internal paste)
    sheet.enable_bindings((
        "single_select", "row_select", "drag_select",
        "edit_cell", "arrowkeys", "tab",
        "copy", "cut", "delete",  # intentionally omit "paste"
        # omit "right_click_popup_menu" to avoid menu-paste
    ))

    # Ensure internal paste & RC popup are disabled so app-level paste owns it
    _unbind_default_paste_and_rc(sheet)

    # Autosave when a cell edit finishes
    try:
        sheet.extra_bindings([
            ("end_edit_cell", lambda _e: (_persist_customers_now(), update_customer_analytics_in_ui(window))),
        ])
    except Exception:
        pass

    _SHEET_CUSTOMERS = sheet

    # Kick analytics immediately and start watcher
    try:
        update_customer_analytics_in_ui(window)
    except Exception:
        pass
    try:
        _start_watch(window)
    except Exception:
        pass

    return sheet

# -----------------------------
# Public event API
# -----------------------------
def handle_customers_events(window, event, values):
    """
    Customers buttons (match layout keys):
      -CUST_EXPORT-        : save grid -> customers.csv
      -CUST_RELOAD-        : reload customers.csv -> grid
      -CUST_ADD50-         : add 50 blank rows (auto-save)
      -CUST_ADD_ORDER-     : quick popup to add an order row (company/date/amount)
    """
    if _SHEET_CUSTOMERS is None:
        return

    if event == "-CUST_EXPORT-":
        ok = _persist_customers_now()
        window["-CUST_STATUS-"].update("Exported customers.csv" if ok else "Export failed.")
        try: update_customer_analytics_in_ui(window)
        except Exception: pass

    elif event == "-CUST_RELOAD-":
        mtx = load_customers_matrix()
        if mtx:
            _SHEET_CUSTOMERS.set_sheet_data(mtx, reset_col_positions=True, reset_row_positions=True)
            _SHEET_CUSTOMERS.headers(list(CUSTOMER_FIELDS))
            _SHEET_CUSTOMERS.refresh()
            window["-CUST_STATUS-"].update("Reloaded.")
        else:
            window["-CUST_STATUS-"].update("No data found.")
        try: update_customer_analytics_in_ui(window)
        except Exception: pass

    elif event == "-CUST_ADD50-":
        try:
            _SHEET_CUSTOMERS.insert_rows(rows=50, idx="end")
        except Exception:
            _SHEET_CUSTOMERS.insert_rows(_SHEET_CUSTOMERS.get_total_rows(), amount=50)
        _SHEET_CUSTOMERS.refresh()
        _persist_customers_now()
        window["-CUST_STATUS-"].update("Added 50 rows.")
        try: update_customer_analytics_in_ui(window)
        except Exception: pass

    elif event == "-CUST_ADD_ORDER-":
        import PySimpleGUI as sg
        company = sg.popup_get_text("Company name for the order:", title="Add Order") or ""
        if not company.strip():
            window["-CUST_STATUS-"].update("Order canceled (no company).")
            return
        today = datetime.now().strftime("%Y-%m-%d")
        date_s = sg.popup_get_text(f"Order date (YYYY-MM-DD):", default_text=today, title="Add Order") or today
        amount_s = sg.popup_get_text("Amount (e.g. 199.99):", title="Add Order") or ""
        amt = _sanitize_amount(amount_s)
        if not amt:
            window["-CUST_STATUS-"].update("Order canceled (invalid amount).")
            return
        try:
            append_order_row(company.strip(), date_s.strip(), amt)
            window["-CUST_STATUS-"].update("Order recorded.")
        except Exception as e:
            window["-CUST_STATUS-"].update(f"Order error: {e}")
            return
        try: update_customer_analytics_in_ui(window)
        except Exception: pass

# -----------------------------
# Persist on demand / exit hook
# -----------------------------
def persist_customers_now() -> bool:
    return _persist_customers_now()
